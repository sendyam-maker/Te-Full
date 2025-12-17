VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc24g0 
   AutoRedraw      =   -1  'True
   Caption         =   "請款單整批列印"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   350
   ClientWidth     =   7040
   ClipControls    =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   7040
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Word(&W)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   3600
      Style           =   1  '圖片外觀
      TabIndex        =   34
      Top             =   2970
      Width           =   3075
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   6030
      ScaleHeight     =   460
      ScaleWidth      =   650
      TabIndex        =   32
      Top             =   4800
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.TextBox txtOutMode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2070
      MaxLength       =   1
      TabIndex        =   30
      Text            =   "1"
      Top             =   5130
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtAdd 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1620
      MaxLength       =   1
      TabIndex        =   13
      Top             =   4260
      Width           =   705
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2010
      TabIndex        =   14
      Top             =   4710
      Width           =   3765
   End
   Begin VB.TextBox txtCopy 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   12
      Text            =   "3"
      Top             =   3840
      Width           =   705
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2070
      Width           =   612
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1428
      TabIndex        =   11
      Top             =   3465
      Width           =   4455
   End
   Begin VB.OptionButton Option2 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   600
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   450
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   2970
      Width           =   3075
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   0
      Top             =   240
      Width           =   612
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1440
      TabIndex        =   2
      Top             =   600
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
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
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   1575
      _ExtentX        =   2787
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label16 
      Caption         =   $"Frmacc24g0.frx":0000
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   3210
      TabIndex        =   36
      Top             =   3840
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "     　                     ( 5.中文請款單)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   2640
      Width           =   7065
   End
   Begin VB.Label Label14 
      Caption         =   "注意:本功能將依畫面上< 請款對象> 及<客戶編號>的設定判斷是否產生電子檔!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   2385
      TabIndex        =   33
      Top             =   30
      Width           =   4335
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "請款單輸出方式：    (1:印表機 2:電子檔)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   405
      TabIndex        =   31
      Top             =   5160
      Visible         =   0   'False
      Width           =   4065
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "     　                     ( 3.加總明細首頁 4.加總明細項目)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   29
      Top             =   2400
      Width           =   7065
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "列印地址條：         (Y : 是)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   420
      TabIndex        =   28
      Top             =   4290
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "地址條印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   27
      Top             =   4725
      Width           =   1725
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "列印份數：         (份)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   450
      TabIndex        =   26
      Top             =   3870
      Width           =   2865
   End
   Begin VB.Label lbl2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3090
      TabIndex        =   25
      Top             =   1710
      Width           =   3495
   End
   Begin VB.Label lbl1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3090
      TabIndex        =   24
      Top             =   1350
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "列印方式：        (1.加總首頁 2.加總明細( 3+4) )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   2100
      Width           =   6465
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   22
      Top             =   1710
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   21
      Top             =   3465
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   19
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "請款編號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "請款對象："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   16
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "請款日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "系統類別："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc24g0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/2 日期欄已修改
Option Explicit

Public adoacc1k0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adoacc1l0 As New ADODB.Recordset
Dim strSql As String
Dim strSQL1 As String
Dim strNo As String
Dim douAmount As Double
Dim strAmount As String
Dim intLength As Integer
Dim intCounter As Integer
Dim douUSDollar As Double
Dim strLanguage As String
Dim strCurr As String
Dim intRecords As Integer
Private Const intInit As Integer = 500
Dim strSystemType As String
Dim prnPrint As Printer
Dim strPrint As String
Dim strProperty As String
Dim strSystemName As String
'Add By Cheng 2003/03/17
Dim m_intPage As Integer '頁數
Dim m_strDetailKind As String
'Add By Cheng 2003/03/21
Dim m_strOriExcRate As Double '原始匯率
Dim m_strUSD As Double '美金金額
Dim m_strA1K02 As String '請款日期 Added by Morgan 2014/9/23
'Add By Cheng 2003/03/31
Dim m_strSQLA As String
'Add By Cheng 2003/04/29
'Modified by Lydia 2020/09/16 調整變數名稱
Dim strConPA As String
Dim strConTM As String
Dim strConLC As String
Dim strConSP As String
'end 2020/09/16
'Add By Cheng 2003/04/29
Dim m_strDBNOFrom As String '請款單起號
Dim m_strDBNOTo As String '請款單迄號
'Add By Cheng 2004/04/27
Dim m_strA1K01 As String '請款單號
'End
'Add by Morgan 2008/4/1
Dim m_bPrinter As Boolean, m_iPages As Integer, m_Device
Dim m_stCaseNos As String, m_stCaseNo As String, m_strDN As String
Dim m_stFileName As String
Dim m_EFilePath As String
Dim m_bolEmail As Boolean, m_bolPaper As Boolean 'Add by Morgan 2010/6/29 Email同時寄紙本

'Added by Morgan 2013/1/2
Dim m_DNCurr As String '請款幣別
Dim m_bPrint2Pdf As Boolean
Dim m_iPrintCurrType As Integer '列印幣別格式:1.純台幣 2.台幣+外幣合計 3.純外幣 4.外幣+美金合計

'Added by Morgan 2014/9/9
Private Type INVITEM
   iType As Single '資料類別 1.表頭 2.主題 3.請款項目 4.合計 5.表尾 6.頁碼
   IBottomLine As Boolean '下邊線
   IText1 As String '欄位1
   IBold1 As Boolean '粗體1
   IULine1 As Boolean '底線1
   IText2 As String '欄位2
   IBold2 As Boolean '粗體2
   IULine2 As Boolean '底線2
   IText3 As String '欄位3
   IBold3 As Boolean '粗體3
   IULine3 As Boolean '底線3
   IText4 As String '欄位4
   IBold4 As Boolean '粗體4
   IULine4 As Boolean '底線4
   IText5 As String '欄位5
   IBold5 As Boolean '粗體5
   IULine5 As Boolean '底線5
End Type
Dim m_Item() As INVITEM '列印項目
Dim m_bolNewDoc As Boolean
Dim m_bolWord As Boolean
Dim m_DocName As String
Dim m_bBatchRule As Boolean '是否適用整批列印規則(更新單筆外幣請款金額為四捨五入到小數第二位，外幣合計無條件捨去到整數)
'end 2014/9/9
Dim strLang2 As String, bolChineseDB As Boolean 'Added by Lydia 2015/04/14 +中文請款單(客戶定稿語文)


Private Sub Command2_Click(Index As Integer)
   'Added by Lydia 2015/08/10
   'Modified by Lydia 2015/08/19 改變訊息
'   If Text6.Text = "5" And Index <> 0 Then
'       MsgBox "中文請款單請選擇列印!", , MsgText(5)
'       Exit Sub
'   End If
   If Text6.Text = "5" Then
     If Index = 0 Then
       MsgBox "中文請款單請選擇Word!", , MsgText(5)
       Exit Sub
     Else
       Index = 0 '中文請款單只產生Word,可是程式寫在列印
     End If
   End If
   
   If FormCheck = False Then
'      MsgBox MsgText(195), , MsgText(5)
      Exit Sub
   End If
   '選擇請款日期
   If Option1.Value Then
      If MaskEdBox1.Text = MsgText(29) Then
         MsgBox MsgText(195), , MsgText(5)
         Exit Sub
      End If
      If MaskEdBox1.Text = MsgText(29) Then
         MsgBox MsgText(195), , MsgText(5)
         Exit Sub
      End If
   End If
   Screen.MousePointer = vbHourglass
   For Each prnPrint In Printers
      If prnPrint.DeviceName = Combo1 Then
         Set Printer = prnPrint
         Exit For
      End If
   Next
   
   'Modify by Morgan 2010/6/29
   ''Add by Morgan 2008/4/2 以請款對象及客戶編號判斷是否產生電子檔
   ''請款對象
   ''Modify by Morgan 2009/12/30　改呼叫共用函數以免不一致
   ''If GetEMailSet(Text1) = True Then
   'If PUB_GetEMailFlag(Text2, , , , Text1) = True Then
   '   txtOutMode = "2"
   'End If
   ''客戶
   'If txtOutMode = "1" And Text5 <> "" Then
   '   'Modify by Morgan 2009/12/30　改呼叫共用函數以免不一致
   '   'If GetEMailSet(Text5) = True Then
   '   If PUB_GetEMailFlag(Text2, , , , Text5) = True Then
   '      txtOutMode = "2"
   '   End If
   'End If
   ''end 2008/4/2
   If txtOutMode = "1" Then
      'Modified by Morgan 2014/5/30 加判斷是否有設 D/N e化
      m_bolEmail = PUB_GetEMailFlag(Text2, , , m_bolPaper, Text1, True)
      If Not m_bolEmail And Text5 <> "" Then
         m_bolEmail = PUB_GetEMailFlag(Text2, , , m_bolPaper, Text5, True)
      End If
   End If
   'end 2010/6/29
   
   'Modify By Sindy 2011/3/3
   'PUB_CheckDNMemo Text1 'Add by Morgan 2008/6/11 D/N備註提醒
   PUB_CheckDNMemo Text1, , Text2 'Add by Morgan 2008/6/11 D/N備註提醒
   
   PrintData Index
   
   If strCon10 <> MsgText(602) Then 'Added by Morgan 2014/10/22 +判斷有資料才要跑
      'Modified by Morgan 2016/10/11 外商人員列印時也要存電子檔(特殊請款單都要,因為財務處會要)--陳金蓮
      'If Index = 0 Then
      If Index = 0 And Pub_StrUserSt03 <> "F12" Then
      'end 2016/10/11
         'Add by Morgan 2008/4/2 若產生電子檔時還要另外列印一份紙本
         'Modify by Morgan 2010/6/29
         'If m_bPrinter = False Then
         '   txtOutMode = "1"
         '   txtCopy = "1"
         If txtOutMode = "1" And m_bPrinter <> True Then
            '未設定同時印紙本
            If Not m_bolPaper Then
               txtCopy = "1"
            End If
            m_bolEmail = False
         'end 2010/6/29
            PrintData
            MsgBox "電子檔已存於 [ " & m_EFilePath & " ]！"
         End If
         'end 2008/4/2
      End If
   
   'If strCon10 <> MsgText(602) Then 'Removed by Morgan 2014/10/22
      FormClear
   End If
   For Each prnPrint In Printers
      If prnPrint.DeviceName = strPrint Then
         Set Printer = prnPrint
      End If
   Next
   Screen.MousePointer = vbDefault
   StatusView MsgText(100)
End Sub


Private Sub Form_Activate()
'edit by nickc 2007/02/08
'   '93.3.16 ADD BY SONIA
'   If IsObject(mdiMain) Then
'      mdiMain.toolshow
'   End If
'   '93.3.16 END
'edit by nickc 2007/02/08 不用 dll 了
   Dim formCnt As Integer
   For formCnt = 0 To Forms.Count - 1
       If UCase(Forms(formCnt).Name) = "MDIMAIN" Then
             Forms(formCnt).ToolShow
             Exit For
       End If
   Next
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      StatusView MsgText(100)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
'Add By Cheng 2003/03/17
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
    'Modify By Cheng 2003/03/17
'   Me.Width = 5250
'   Me.Height = 3200
   Me.Width = 7155
   'Modified by Lydia 2015/04/14
   'Me.Height = 5685
   Me.Height = 6015
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
    'Modify By Cheng 2003/03/21
    '預設點選請款編號
'   FormEnabled1
    Me.Option2.Value = True
   FormEnabled2
   
   'Modified by Morgan 2017/11/8 設定印表機改呼叫公用函數,原程式移除
   PUB_SetPrinter Me.Name, Combo1, strPrint
   PUB_SetPrinter Me.Name, Combo2, , False
   'end 2017/11/8
   
   StatusView MsgText(100)
      
   If Pub_StrUserSt03 = "M51" Then
      Label13.Visible = True
      txtOutMode.Visible = True
      Label16.Visible = True 'Added by Lydia 2020/11/09
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Modify By Cheng 2003/03/24
    '若是否列印地址條上"Y"
    If Me.txtAdd.Text = "Y" Then
        'Add By Cheng 2003/01/29
        '列印地址條
        PUB_PrintAddressList strUserNum, Me.Combo2.Text
        '刪除地址條列表資料
        PUB_DeleteAddressList strUserNum
        '初始化序號
        pub_AddressListSN = 0
        '將印表機設回預設印表機
        For Each prnPrint In Printers
           If prnPrint.DeviceName = strPrint Then
              Set Printer = prnPrint
           End If
        Next
    End If
    'Add By Cheng 2003/03/17
    '若印表機變動, 則更新列印設定
    If Me.Combo1.Text <> Me.Combo1.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
    End If
    '若印表機變動, 則更新列印設定
    If Me.Combo2.Text <> Me.Combo2.Tag Then
        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
    End If
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc24g0 = Nothing
End Sub

Private Sub Option1_Click()
   If Option1.Value Then
      FormEnabled1
   Else
      FormEnabled2
   End If
'   strSystemType = Text2
'   FormClear
'   Text2 = strSystemType
End Sub

Private Sub Option2_Click()
   If Option1.Value Then
      FormEnabled1
   Else
      FormEnabled2
   End If
'   strSystemType = Text2
'   FormClear
'   Text2 = strSystemType
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
   CloseIme
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    Select Case Len(Text1)
        Case 6
            Text1 = AfterZero(Text1)
        Case 8
            Text1 = Text1 & "0"
    End Select
    'Add By Cheng 2003/03/17
    If Me.Text1.Text <> "" Then
        Me.Text1.Text = Left(Me.Text1.Text & "000000000", 9)
        'Modify by Morgan 2010/6/29
        'Me.lbl1.Caption = GetFagentEngName(Me.Text1.Text)
        lbl1.Caption = GetFagentEngName(Text1.Text, Text2.Text, strExc(1))
        If Text2 <> "" And strExc(1) <> "" Then
            txtCopy = strExc(1)
        End If
        'end 2010/6/29
        If Me.lbl1.Caption = "" Then
            MsgBox "請款對象輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            Text1_GotFocus
        End If
    Else
        Me.lbl1.Caption = ""
    End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
   CloseIme
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Add by Morgan 2010/6/29
Private Sub Text2_Validate(Cancel As Boolean)
   If Text2 <> "" Then
      If Text1 <> "" Then
         GetFagentEngName Text1.Text, Text2.Text, strExc(1)
         If strExc(1) <> "" Then
            txtCopy = strExc(1)
         End If
      End If
      If Text5 <> "" Then
         GetCustomerEngName Text5.Text, Text2.Text, strExc(1)
         If strExc(1) <> "" Then
            txtCopy = strExc(1)
         End If
      End If
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
   CloseIme
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   TextInverse Text4
   CloseIme
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
    Text2 = ""
    Text1 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text6 = ""
    MaskEdBox1.Mask = ""
    MaskEdBox1.Text = ""
    MaskEdBox1.Mask = DFormat
    MaskEdBox2.Mask = ""
    MaskEdBox2.Text = ""
    MaskEdBox2.Mask = DFormat
    Me.lbl1.Caption = ""
    Me.Lbl2.Caption = ""
    Text2.SetFocus
    'Add By Cheng 2003/03/25
    '預設列印方式
    If Me.Option1.Value Then
        Me.Text6.Text = "1"
    Else
        Me.Text6.Text = "2"
    End If
    
    txtCopy = "3"
End Sub

'****************************************************
' 列印明細資料
' 外商要依請款對象+客戶代號跳頁, 其他系統類別不要
'****************************************************
Private Sub PrintData(Optional Index As Integer = 0)
Dim strDescription As String
Dim strDes(3) As String
Dim intCols As Integer
Dim StrSQLa As String
Dim strPrintText As String
Dim strArrCaseNo '本所案號陣列
Dim rsA As New ADODB.Recordset
Dim rsB As New ADODB.Recordset
Dim StrSqlB As String
Dim strPrintText1 As String
Dim lngBoxX As Long, lngBoxY As Long 'Add by Morgan 2006/7/5
'Added by Morgan 2013/1/2
Dim bCanPDF As Boolean '是否有安裝 PDFCreator
Dim bolSetDone As Boolean

'add by nickc 2007/02/08
Dim lngAmount As Long
'Add by Morgan 2008/4/2
Dim stCaseNo As String
Dim strKey As String 'Add By Sindy 2020/7/24
Dim strConPA_b As String 'Added by Morgan 2022/12/16
Dim strFormula As String 'Add By Sindy 2025/4/8
    
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/22 清除查詢印表記錄檔欄位
    
    strSql = MsgText(601)
    '系統類別
    If Text2 <> MsgText(601) Then
        strSql = strSql & " and a1k13 = '" & Text2 & "'"
        pub_QL05 = pub_QL05 & ";" & Label4 & Text2 'Add By Sindy 2010/12/22
    End If
    '選擇請款對象
    If Option1.Value Then
        '請款日期
        If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
            strSql = strSql & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
        End If
        If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
            strSql = strSql & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
        End If
        If (MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29)) Or _
           (MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29)) Then
           pub_QL05 = pub_QL05 & ";" & Label1 & MaskEdBox1 & "-" & MaskEdBox2 'Add By Sindy 2010/12/22
        End If
    '選擇請款編號
    Else
        '請款編號
        If Text3 <> MsgText(601) Then
            strSql = strSql & " and a1k01 >= '" & Text3 & "'"
        End If
        If Text4 <> MsgText(601) Then
            strSql = strSql & " and a1k01 <= '" & Text4 & "'"
        End If
        If Text3 <> MsgText(601) Or Text4 <> MsgText(601) Then
           pub_QL05 = pub_QL05 & ";" & Label3 & Text3 & "-" & Text4 'Add By Sindy 2010/12/22
        End If
    End If
    '請款對象
    If Text1 <> MsgText(601) Then
        strSql = strSql & " and a1k28 = '" & ChangeCustomerL(Text1) & "'"
        pub_QL05 = pub_QL05 & ";" & Label2(0) & Text1 & lbl1 'Add By Sindy 2010/12/22
    End If
    '客戶編號
    strConPA = "": strConTM = "": strConLC = "": strConSP = ""  'Added by Lydia 2020/09/16
    If Me.Text5.Text <> "" Then
        strConPA = " and PA26='" & ChangeCustomerL(Me.Text5.Text) & "' "
        strConTM = " and TM23='" & ChangeCustomerL(Me.Text5.Text) & "' "
        strConLC = " and LC11='" & ChangeCustomerL(Me.Text5.Text) & "' "
        strConSP = " and SP08='" & ChangeCustomerL(Me.Text5.Text) & "' "
        pub_QL05 = pub_QL05 & ";" & Label8 & Text5 & Lbl2 'Add By Sindy 2010/12/22
    End If
    'Added by Lydia 2015/04/14 +中文請款單
    If Trim(Text6) <> "" Then
        pub_QL05 = pub_QL05 & ";" & Left(Label9, 5) & Text6 & "(1.加總首頁 2.加總明細(3+4) 3.加總明細首頁 4.加總明細項目 5.中文請款單)" 'Add By Sindy 2010/12/22
    End If
    If Trim(txtCopy) <> "" Then
        pub_QL05 = pub_QL05 & ";" & Left(Label10, 5) & txtCopy  'Add By Sindy 2010/12/22
    End If
    If Trim(txtAdd) <> "" Then
        pub_QL05 = pub_QL05 & ";" & Left(Label11, 6) & txtAdd   'Add By Sindy 2010/12/22
    End If
    If Trim(txtOutMode) = "1" Then
        pub_QL05 = pub_QL05 & ";" & Left(Label13, 8) & "1:印表機" 'Add By Sindy 2010/12/22
    Else
        pub_QL05 = pub_QL05 & ";" & Left(Label13, 8) & "2:電子檔" 'Add By Sindy 2010/12/22
    End If
    
    strConPA_b = strConPA 'Added by Morgan 2022/12/16
    
    'Makr by Lydia 2020/11/09 內商Y513450北京正理要整批結清部份請款單
    '----109/9/8 請款單號共128筆; 109/11/9 確認已結清
    'strSql = strSql & " and a1k01 in ('X10911625','X10909228','X10911710','X10911711','X10909229','X10911626','X10911627'," & _
            "'X10911628','X10911708','X10911709','X10911848','X10911200','X10911201','X10911202'," & _
            "'X10909800','X10911351','X10911203','X10909801','X10911114','X10912034','X10911707'," & _
            "'X10911706','X10912773','X10911702','X10911696','X10911697','X10911698','X10911699'," & _
            "'X10911700','X10911701','X10912033','X10911704','X10911705','X10912774','X10911703'," & _
            "'X10912032'," & _
            "'X10707371','X10715473','X10806322','X10909482','X10808243','X10808243','X10808244'," & _
            "'X10907586','X10816428','X10816429','X10816430','X10817362','X10817363','X10906288'," & _
            "'X10908240','X10908019','X10818379','X10820242','X10820250','X10820252','X10820253'," & _
            "'X10820254','X10820251','X10820246','X10820247','X10820248','X10820249','X10911422'," & _
            "'X10910554','X10909225','X10909226','X10909227','X10909222','X10910546','X10909223'," & _
            "'X10909224','X10910551','X10910552','X10910547','X10910549','X10910550','X10905609'," & _
            "'X10905611','X10907931','X10907932','X10907933','X10907934','X10907935','X10907936'," & _
            "'X10907937','X10907938','X10907939','X10907940','X10907941','X10907942','X10907943'," & _
            "'X10907944','X10907945','X10907946','X10907947','X10907948','X10907949','X10907950'," & _
            "'X10907951','X10907952','X10907953','X10907954','X10907955','X10907956','X10907957'," & _
            "'X10907958','X10907959','X10907960','X10907237','X10907238','X10907961','X10907962'," & _
            "'X10907963','X10907964','X10907966','X10907967','X10907968','X10907969','X10907970'," & _
            "'X10907971','X10907972','X10907973','X10907975','X10907976','X10907977','X10907978'," & _
            "'X10907979') "
    '----109/11/9 結清105/1/1~109/9/23的欠款,排除下列12筆
    'strSql = strSql & " and a1k01 not in ('X10802757','X10802758','X10802759','X10803335','X10803336','X10910098','X10912500','X10912501','X10912502','X10914104','X10914106','X10914107') "
    'Memo by Lydia 2020/11/09 要注意請款單列印在抓明細時是用項次A1L02，若遇到xxx99規費的項次先於xxx服務費會造成表格定位錯誤。
    'end 2020/11/09
    
    adoacc1k0.CursorLocation = adUseClient
    '列印方式選擇--加總首頁
    'Modify by Morgan 2008/5/27 信頭改抓列印對象 a1k28-->a1k27
    'Modify By Sindy 2011/3/7 +fa108
    'Modify by Morgan 2011/5/25 +fa70,去掉group by 語法,因為 union 會distinct 且排序由 order by 決定
    'Modified by Morgan 2013/1/2 +a1k33
    'Modified by Morgan 2018/4/27 +a1k28
    'Modified by Morgan 2023/4/10 +
    If Me.Text6.Text = "1" Then
        If Me.Text2.Text = "FCT" Or Me.Text2.Text = "S" Then
            '專利
            StrSQLa = "select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, pa26 as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, patent where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) = fa02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & _
                                " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, pa26 as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, patent where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA
            '商標
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, tm23 as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, trademark where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) = fa02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & _
                                " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, tm23 as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, trademark where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM
            '法務
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, lc11 as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, lawcase where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) = fa02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & _
                                " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, lc11 as cuno,cu148 as fa108,cu102 fa70,a1k18,a1k33,a1k28 from acc1k0, customer, lawcase where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC
            '服務
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, sp08 as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, servicepractice where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) = fa02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & _
                                " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, sp08 as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, servicepractice where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP
            StrSQLa = StrSQLa & " Order By 1, 24 "
        Else
            '專利
            StrSQLa = "select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, '' as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, patent where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) = fa02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & _
                                " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, '' as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, patent where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA
            '商標
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, '' as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, trademark where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) = fa02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & _
                                " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, '' as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, trademark where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM
            '法務
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, '' as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, lawcase where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) = fa02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & _
                                " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, '' as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, lawcase where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC
            '服務
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, '' as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, servicepractice where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) = fa02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & _
                                " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, '' as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, servicepractice where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) = cu02 and (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP
            StrSQLa = StrSQLa & " Order By 1, 24 "
        End If
        adoacc1k0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   '列印方式選擇--加總明細
   Else
        If Me.Text2.Text = "FCT" Or Me.Text2.Text = "S" Then
            '專利
            StrSQLa = "select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, pa26 as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, patent where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & _
                         " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, pa26 as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, patent where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA
            '商標
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, tm23 as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, trademark where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & _
                         " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, tm23 as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, trademark where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM
            '法務
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, lc11 as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, lawcase where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & _
                         " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, lc11 as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, lawcase where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC
            '服務
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, sp08 as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, servicepractice where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & _
                         " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, sp08 as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, servicepractice where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP
            StrSQLa = StrSQLa & " Order By 1, 24 "
        Else
            '專利
            StrSQLa = "select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, '' as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, patent where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & _
                         " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, '' as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, patent where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA
            '商標
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, '' as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, trademark where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & _
                         " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, '' as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, trademark where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM
            '法務
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, '' as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, lawcase where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & _
                         " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, '' as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, lawcase where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC
            '服務
            StrSQLa = StrSQLa & " union select a1k27, a1k13, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, fa04, fa18, a1k02, fa43, a1k10, '' as cuno, fa108,fa70,a1k18,a1k33,a1k28 from acc1k0, fagent, servicepractice where substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & _
                         " union select a1k27, a1k13, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu24 as fa18, cu66 as fa33, cu25 as fa19, cu67 as fa34, cu26 as fa20, cu68 as fa35, cu27 as fa21, cu69 as fa36, cu28 as fa22, cu06 as fa06, cu29 as fa23, cu04 as fa04, cu24 as fa18, a1k02, cu76 as fa43, a1k10, '' as cuno,cu148 as fa108,cu102 as fa70,a1k18,a1k33,a1k28 from acc1k0, customer, servicepractice where substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP
            StrSQLa = StrSQLa & " Order By 1, 24 "
        End If
        adoacc1k0.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   End If
   If adoacc1k0.RecordCount = 0 Then
      InsertQueryLog (0) 'Add By Sindy 2010/12/22
      strCon10 = MsgText(602)
      MsgBox MsgText(28), , MsgText(5)
      adoacc1k0.Close
      Exit Sub
   Else
      InsertQueryLog (adoacc1k0.RecordCount) 'Add By Sindy 2010/12/22
   End If
    
   'Added by Morgan 2013/1/2 電子檔改pdf格式
   '檢查是否有安裝PDFCreator
   bCanPDF = False
'未完成--程式流程要改
'   If PrinterIndex("PDFCreator") >= 0 Then
'      bCanPDF = True
'   End If
   m_bPrint2Pdf = False

   'end 2013/1/2
   
   m_bolNewDoc = True 'Added by Morgan 2014/9/15
   m_DocName = "" 'Added by Morgan 2014/9/18
   m_stCaseNos = ""
   m_stCaseNo = ""
   m_strDN = ""
   m_stFileName = ""
   
   '預設請款單格式為加總明細的格式1
   m_strDetailKind = "1"
   lngAmount = 0
   intLength = 0
   douAmount = 0
   douUSDollar = 0
   intCounter = 4
   
   strNo = ""
   strCon10 = ""
   m_strUSD = "0"
   m_strA1K01 = ""
   
'Added by Morgan 2014/9/17
m_bolWord = False
'Modify By Sindy 2015/7/13 +外專都帶信頭
'If Index = 0 Then
'Modified by Morgan 2016/10/11 外商人員列印時也要存電子檔(特殊請款單都要,因為財務處會要)--陳金蓮
'If Index = 0 And Not (Text2 = "FCP" Or Text2 = "FG") Then
If Index = 0 And Not (Text2 = "FCP" Or Text2 = "FG") And Pub_StrUserSt03 <> "F12" Then
'2015/7/13 END
'end 2014/9/17

   'Add by Morgan 2008/4/2 判斷是否產生電子檔
   'Modify by Morgan 2010/6/29
   'If Text6 = "2" And txtOutMode = "2" Then
   If Text6 = "2" And (txtOutMode = "2" Or m_bolEmail) Then
      
      'Added by Morgan 2013/1/2 電子檔改PDF格式
      If bCanPDF = True Then
         m_bPrinter = True
         m_bPrint2Pdf = True
         Set m_Device = Printer
         Load frmPDF
         MyNewPage True
      Else
      'end 2013/1/2
      
         m_bPrinter = False
         Set m_Device = Picture1
         m_Device.AutoRedraw = True
         m_Device.Height = 16836
         m_Device.Width = 11904
         DelPic
         
      End If 'Added by Morgan 2013/1/2
      
   Else
      m_bPrinter = True
      Set m_Device = Printer
      m_Device.Orientation = 1
   End If
   'end 2008/4/2

'Added by Morgan 2014/9/17
Else
   m_bolWord = True
   
   'Added by Morgan 2016/10/11 外商人員列印時也要存電子檔(特殊請款單都要,因為財務處會要)--陳金蓮
   m_bPrinter = False
   If Index = 0 And Pub_StrUserSt03 = "F12" Then
      If txtOutMode = "1" Then
         m_bPrinter = True
      End If
   End If
   'end 2016/10/11
         
   'Add By Sindy 2015/7/13 為了後面辨識是不是要開Word純列印
   If (Text2 = "FCP" Or Text2 = "FG") Then
      If Index = 0 Then
         m_bPrinter = True
         txtOutMode = "1"
      Else
         txtOutMode = "2"
      End If
   End If
   '2015/7/13 END
End If
'end 2014/9/17

   m_intPage = 1
   
   strLang2 = "": bolChineseDB = False 'Move by Lydia 2020/09/16 從下面移上來
   
If Not m_bolWord Then 'Added by Morgan 2014/9/17

   m_Device.FontSize = 12
   '設定列印份數
   If m_bPrinter = True Then
      'Added by Morgan 2013/1/2 電子檔改PDF格式
      If m_bPrint2Pdf = True Then
         m_Device.Copies = 1
      Else
      'end 2013/1/2
         m_Device.Copies = IIf(Val("0" & Me.txtCopy.Text) = "0", 3, Val("0" & Me.txtCopy.Text))
         Me.txtCopy.Text = m_Device.Copies
      End If 'Added by Morgan 2013/1/2
   End If
   
    m_Device.Font = "Times New Roman"
    m_Device.Font.Italic = False
    m_Device.Font.Bold = False
    
End If 'Added by Morgan 2014/9/17
   
   Do While adoacc1k0.EOF = False
   
      'Modied by Morgan 2013/1/2
      '新舊請款單以是否有設定列印幣別格式(a1k33)判斷
   
'      'Modify By Sindy 2011/3/7
'      If CheckSys("" & adoacc1k0.Fields("a1k13").Value) = "2" Or _
'         CheckSys("" & adoacc1k0.Fields("a1k13").Value) = "6" Then
'         If IsNull(adoacc1k0.Fields("fa108").Value) = False Then
'            strCurr = adoacc1k0.Fields("fa108").Value
'         Else
'            strCurr = MsgText(601)
'         End If
'      '2011/3/7 End
'      Else
'         If IsNull(adoacc1k0.Fields("fa43").Value) = False Then
'            strCurr = adoacc1k0.Fields("fa43").Value
'         Else
'            strCurr = MsgText(601)
'         End If
'      End If
      
      m_DNCurr = "" & adoacc1k0.Fields("a1k18").Value
      If Not IsNull(adoacc1k0("a1k33")) Then
         m_iPrintCurrType = Val(adoacc1k0("a1k33"))
      Else
         'Modify By Sindy 2016/11/29 + , , "" & adoacc1k0.Fields("a1k14").Value, "" & adoacc1k0.Fields("a1k15").Value, "" & adoacc1k0.Fields("a1k16").Value
         'Mofieid by Morgan 2016/12/1 會有錯(明細重複列印,總計多了件數倍)先還原,若個案有特別需另外控制
         'Modified by Morgan 2018/4/27
         'm_iPrintCurrType = PUB_GetDefaultCurrPrintType("" & adoacc1k0.Fields("a1k13").Value, "" & adoacc1k0.Fields("a1k27").Value, m_DNCurr)
         m_iPrintCurrType = PUB_GetDefaultCurrPrintType("" & adoacc1k0.Fields("a1k13").Value, "" & adoacc1k0.Fields("a1k28").Value, m_DNCurr, , , , , "" & adoacc1k0.Fields("a1k27").Value)
         'm_iPrintCurrType = PUB_GetDefaultCurrPrintType("" & adoacc1k0.Fields("a1k13").Value, "" & adoacc1k0.Fields("a1k27").Value, m_DNCurr, , _
            "" & adoacc1k0.Fields("a1k14").Value, "" & adoacc1k0.Fields("a1k15").Value, "" & adoacc1k0.Fields("a1k16").Value)
         '2016/11/29 END
      End If
      
      'Added by Lydia 2020/09/16 共同條件+請款日期a1k02 (因為版面是以同一日期為一份)
      'Modifed by Morgan 2022/12/16
      'strExc(1) = IIf("" & adoacc1k0.Fields("cuno").Value <> "", "and PA26='" & adoacc1k0.Fields("cuno").Value & "' ", "")
      strExc(1) = strConPA_b
      'end 2022/12/6
      strConPA = strExc(1) & " and a1k02=" & adoacc1k0.Fields("a1k02").Value & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and pa26='" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null "
      strConTM = Replace(strExc(1), "PA26", "TM23") & " and a1k02=" & adoacc1k0.Fields("a1k02").Value & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and tm23='" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null "
      strConLC = Replace(strExc(1), "PA26", "LC11") & " and a1k02=" & adoacc1k0.Fields("a1k02").Value & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and lc11='" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null "
      strConSP = Replace(strExc(1), "PA26", "SP08") & " and a1k02=" & adoacc1k0.Fields("a1k02").Value & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and sp08='" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null "
      'end 2022/12/16
      'end 2020/09/16
      
      'Removed by Morgan 2016/12/6 已有非美金,先開放User視需要自行修改內容
      ''原程式沒有考慮非美金請款,暫維持非美金用2(台幣+美金合計)格式
      'Select Case m_iPrintCurrType
      'Case 3, 4 '純外幣,外幣+美金合計
      '   If m_DNCurr <> "USD" Then
      '      m_iPrintCurrType = 2
      '   End If
      'End Select
      'end 2016/12/6
      
      'end 2013/1/2
      
        '列印方式選擇--加總明細
        If Me.Text6.Text <> "1" Then
'         '若列印對象及申請人不同
         'Modified by Lydia 2020/09/16 +請款日期a1k02 (因為版面是以同一日期為一份)
         'If strNo <> adoacc1k0.Fields("a1k27").Value & adoacc1k0.Fields("cuno").Value Then
         '   If douAmount <> 0 Then
         If strNo <> adoacc1k0.Fields("a1k27").Value & adoacc1k0.Fields("cuno").Value & adoacc1k0.Fields("a1k02").Value Then
            If douAmount <> 0 Or (strNo <> "" And Me.Text6.Text = "3") Then  '有明細金額 或 只印加總明細首頁
         'end 2020/09/16
               '若選擇加總明細或加總明細項目
               If Me.Text6.Text = "2" Or Me.Text6.Text = "4" Then
                  
                  If Not m_bolWord Then 'Added by Morgan 2014/9/17
                     m_Device.Line (0 + intInit, 1500 + intCounter * 300 - 200)-(10200 + intInit, 1500 + intCounter * 300 - 200)
                  End If
                  
                   PrintSum
                   
                   'Modify by Morgan 2011/7/27 新信紙有信尾要上移
                   'intCounter = 45
                   intCounter = 43
                   
                  If Not m_bolWord Then 'Added by Morgan 2014/9/17
                     m_Device.CurrentX = 4500 + intInit
                     m_Device.CurrentY = 1500 + intCounter * 300
                     m_Device.Print "P." & m_intPage
                  'Added by Morgan 2014/9/17
                  Else
                     SetWordArray m_Item, intCounter, 1, "P." & m_intPage, 6
                  End If
                  'END 2014/9/17
                   
                   '若選擇加總明細首頁
                   If Me.Text6.Text = "4" Then
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        runWord
                     Else
                     'end 2014/9/17
                       If m_bPrinter = True Then
                          'Modified by Morgan 2013/1/2
                          'm_Device.NewPage
                          MyNewPage
                       Else
                          SetPic m_intPage
                       End If
                       m_intPage = 1
                     End If
                   End If
                   m_strA1K01 = ""
               End If
               douAmount = 0
               douUSDollar = 0
               m_strUSD = "0"
               '若選擇加總明細或加總明細首頁
               If Me.Text6.Text = "2" Or Me.Text6.Text = "3" Then
                  'Added by Morgan 2014/9/17
                  If m_bolWord Then
                     runWord
                  Else
                  'end 2014/9/17
                     If m_bPrinter = True Then
                        'Modified by Morgan 2013/1/2
                        'm_Device.NewPage
                        MyNewPage
                     Else
                        SetPic m_intPage
                     End If
                  End If
                  m_intPage = 1
               End If
            End If
            m_strDetailKind = "1"
            intCounter = 4
            If Me.Text6.Text = "2" Or Me.Text6.Text = "3" Then PrintHead
            'Modified by Lydia 2020/09/16 +請款日期a1k02 (因為版面是以同一日期為一份)
            'strNo = "" & adoacc1k0.Fields("a1k27").Value & adoacc1k0.Fields("cuno").Value
            strNo = "" & adoacc1k0.Fields("a1k27").Value & adoacc1k0.Fields("cuno").Value & adoacc1k0.Fields("a1k02").Value
         End If
         
         '第一頁
         m_strDetailKind = "1"
         adoquery.CursorLocation = adUseClient
         'Modified by Lydia 2015/04/14 +案號
         'strLang2 = "": bolChineseDB = False 'Move by Lydia 2020/09/16
         'Modified by Lydia 2020/09/16 共同條件+請款日期a1k02 (因為版面是以同一日期為一份)
         'StrSQLa = "select pa77 as Yno, pa01||'-'||pa02||'-'||pa03||'-'||pa04 as Caseno, pa11 as Ano, a1k01 as Dno,pa01 as sno1,pa02 as sno2,pa03 as sno3,pa04 as sno4 from acc1k0, patent where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and pa26='" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & strConPA & " group by pa77, pa01||'-'||pa02||'-'||pa03||'-'||pa04, pa11, a1k01,pa01,pa02,pa03,pa04 "
         'StrSQLa = StrSQLa & " union select tm45 as Yno, tm01||'-'||tm02||'-'||tm03||'-'||tm04 as Caseno, nvl(tm15,tm12) as Ano, a1k01 as Dno,tm01 as sno1,tm02 as sno2,tm03 as sno3,tm04 as sno4 from acc1k0, trademark where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and tm23 = '" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & strConTM & " group by tm45, tm01||'-'||tm02||'-'||tm03||'-'||tm04, nvl(tm15,tm12),a1k01,tm01,tm02,tm03,tm04 "
         'StrSQLa = StrSQLa & " union select lc23 as Yno, lc01||'-'||lc02||'-'||lc03||'-'||lc04 as Caseno, '' as Ano, a1k01 as Dno,lc01 as sno1,lc02 as sno2,lc03 as sno3,lc04 as sno4 from acc1k0, lawcase where a1k13 = lc01 and a1k14 = lc02 and a1k15 = lc03 and a1k16 = lc04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and lc11 = '" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & strConLC & " group by lc23, lc01||'-'||lc02||'-'||lc03||'-'||lc04, a1k01,lc01,lc02,lc03,lc04 "
         'StrSQLa = StrSQLa & " union select sp27 as Yno, sp01||'-'||sp02||'-'||sp03||'-'||sp04 as Caseno, sp11 as Ano, a1k01 as Dno,sp01 as sno1,sp02 as sno2,sp03 as sno3,sp04 as sno4 from acc1k0, servicepractice where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and sp08 = '" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & strConSP & " group by sp27, sp01||'-'||sp02||'-'||sp03||'-'||sp04,sp11,a1k01,sp01,sp02,sp03,sp04 "
         StrSQLa = "select pa77 as Yno, pa01||'-'||pa02||'-'||pa03||'-'||pa04 as Caseno, pa11 as Ano, a1k01 as Dno,pa01 as sno1,pa02 as sno2,pa03 as sno3,pa04 as sno4 from acc1k0, patent where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 " & strSql & strConPA & " group by pa77, pa01||'-'||pa02||'-'||pa03||'-'||pa04, pa11, a1k01,pa01,pa02,pa03,pa04 "
         StrSQLa = StrSQLa & " union select tm45 as Yno, tm01||'-'||tm02||'-'||tm03||'-'||tm04 as Caseno, nvl(tm15,tm12) as Ano, a1k01 as Dno,tm01 as sno1,tm02 as sno2,tm03 as sno3,tm04 as sno4 from acc1k0, trademark where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 " & strSql & strConTM & " group by tm45, tm01||'-'||tm02||'-'||tm03||'-'||tm04, nvl(tm15,tm12),a1k01,tm01,tm02,tm03,tm04 "
         StrSQLa = StrSQLa & " union select lc23 as Yno, lc01||'-'||lc02||'-'||lc03||'-'||lc04 as Caseno, '' as Ano, a1k01 as Dno,lc01 as sno1,lc02 as sno2,lc03 as sno3,lc04 as sno4 from acc1k0, lawcase where a1k13 = lc01 and a1k14 = lc02 and a1k15 = lc03 and a1k16 = lc04 " & strSql & strConLC & " group by lc23, lc01||'-'||lc02||'-'||lc03||'-'||lc04, a1k01,lc01,lc02,lc03,lc04 "
         StrSQLa = StrSQLa & " union select sp27 as Yno, sp01||'-'||sp02||'-'||sp03||'-'||sp04 as Caseno, sp11 as Ano, a1k01 as Dno,sp01 as sno1,sp02 as sno2,sp03 as sno3,sp04 as sno4 from acc1k0, servicepractice where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 " & strSql & strConSP & " group by sp27, sp01||'-'||sp02||'-'||sp03||'-'||sp04,sp11,a1k01,sp01,sp02,sp03,sp04 "
         'end 2020/09/16
        '依請款編號排序
         StrSQLa = StrSQLa & " order by 4 "
         adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            m_bBatchRule = ChkBatchRule(adoquery.Fields("DNO").Value) 'Added by Morgan 2014/9/19

            intRecords = adoquery.RecordCount
            'Modified by Morgan 2013/1/2
            'If m_bPrinter = True Then
            If m_bPrinter = True And m_bPrint2Pdf = False Then
            'end 2013/1/2
            
               '新增地址條列表資料
               If Me.txtAdd.Text = "Y" Then
                   If "" & adoquery("Caseno").Value <> "" Then
                       strArrCaseNo = Split(adoquery("Caseno").Value, "-")
                       pub_AddressListSN = pub_AddressListSN + 1
                       PUB_AddNewAddressList strUserNum, "" & strArrCaseNo(0), "" & strArrCaseNo(1), "" & strArrCaseNo(2), "" & strArrCaseNo(3), "" & pub_AddressListSN, "0"
                   End If
               End If
               
            'Add by Morgan 2008/4/2
            Else
               stCaseNo = Replace("" & adoquery("Caseno").Value, "-0-00", "")
               stCaseNo = Replace("" & stCaseNo, "-", "")
               If stCaseNo <> m_stCaseNo Then
                  CopyImg
                  m_stCaseNo = stCaseNo
               End If
               m_strDN = "" & adoquery("Dno").Value
               m_stCaseNos = ""
            End If
            'end 2008/4/2
         Else
            intRecords = 0
         End If
         
         Do While adoquery.EOF = False
            'Add By Cheng 2004/04/27
            m_strA1K01 = m_strA1K01 & "'" & adoquery.Fields("DNO").Value & "',"
            'End
            
            'Modified by Lydia 2015/04/14 抓收件人定稿語言
            If Me.Text6.Text = "5" Then  '中文整批請款單
                strExc(10) = Frmacc2480.GetLanguage("" & adoquery.Fields("sno1").Value, "" & adoquery.Fields("sno2").Value, "" & adoquery.Fields("sno3").Value, "" & adoquery.Fields("sno4").Value, "" & adoquery.Fields("dno").Value)
                If strExc(10) = "1" Then
                   If bolChineseDB = False Then bolChineseDB = True
                   strLang2 = strLang2 & adoquery.Fields("Dno") & ","
                End If
                'Modified by Lydia 2020/09/16
                'If adoquery.AbsolutePosition < adoquery.RecordCount Then
                '   GoTo ChiJumpNext
                'Else
                '   GoTo ChiJumpLast
                'End If
                GoTo ChiJumpNext
                'end 2020/09/16
            End If
            'end 2015/04/14
            
            '若選擇加總明細或加總明細首頁
            If Me.Text6.Text = "2" Or Me.Text6.Text = "3" Then
                If intCounter >= 40 Then
                     intCounter = intCounter + 1
                     
                     If Not m_bolWord Then 'Added by Morgan 2014/9/17
                        m_Device.Line (0 + intInit, 1500 + (intCounter) * 300 - 200)-(10200 + intInit, 1500 + (intCounter) * 300 - 200)
                     End If
                     
                     'Modify by Morgan 2011/7/27 新信紙有信尾要上移
                     'intCounter = 45
                     intCounter = 43
                     
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 1, "P." & m_intPage, 6
                     Else
                     'END 2014/9/17
                     
                        m_Device.CurrentX = 4500 + intInit
                        m_Device.CurrentY = 1500 + (intCounter) * 300
                        m_Device.Print "P." & m_intPage
                        
                     End If 'Added by Morgan 2014/9/17
                     
                     intCounter = intCounter + 1
                     
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 1, "Continued", 6
                     Else
                     'END 2014/9/17
                     
                        m_Device.CurrentX = 4200 + intInit
                        m_Device.CurrentY = 1500 + (intCounter) * 300
                        m_Device.Print "Continued"
                     
                     End If 'Added by Morgan 2014/9/17
                     
                     
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        runWord
                     Else
                     'end 2014/9/17
                        If m_bPrinter = True Then
                           'Modified by Morgan 2013/1/2
                           'm_Device.NewPage
                           MyNewPage
                        Else
                           SetPic m_intPage
                        End If
                     End If
                     m_intPage = m_intPage + 1
                     intCounter = 4
                     PrintHead
                End If
                '若為FCT案加印商標名稱
                If "" & adoacc1k0.Fields("a1k13").Value = "FCT" Or "" & adoacc1k0.Fields("a1k13").Value = "S" Then
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 1, "*", 3
                     Else
                     'END 2014/9/17
                     
                        m_Device.CurrentX = 0 + 350
                        m_Device.CurrentY = 1500 + intCounter * 300
                        m_Device.Print "*"
                        
                     End If 'Added by Morgan 2014/9/17
                     
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, GetCaseName("" & adoquery.Fields("CaseNo").Value), 3
                     Else
                     'END 2014/9/17
                     
                        m_Device.CurrentX = 0 + intInit
                        m_Device.CurrentY = 1500 + intCounter * 300
                        m_Device.Print GetCaseName("" & adoquery.Fields("CaseNo").Value)
                        
                     End If 'Added by Morgan 2014/9/17
                     
                     intCounter = intCounter + 1
                End If
               
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, "" & adoquery.Fields("Ano").Value, 3
               Else
               'END 2014/9/17
               
                  m_Device.CurrentX = 0 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print "" & adoquery.Fields("Ano").Value
                  
               End If 'Added by Morgan 2014/9/17
               
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 3, "" & adoquery.Fields("Dno").Value, 3
               Else
               'END 2014/9/17
               
                  m_Device.CurrentX = 2500 + intInit - 500 + 220
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print "" & adoquery.Fields("Dno").Value
                  
               End If 'Added by Morgan 2014/9/17
               
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 4, Replace("" & adoquery.Fields("CaseNo").Value, "-0-00", ""), 3
               Else
               'END 2014/9/17
               
                  m_Device.CurrentX = 4500 + intInit - 1000 + 1020
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print Replace("" & adoquery.Fields("CaseNo").Value, "-0-00", "")
                  
               End If 'Added by Morgan 2014/9/17
               
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 5, "" & adoquery.Fields("Yno").Value, 3
               Else
               'END 2014/9/17
               
                  m_Device.CurrentX = 6500 + intInit - 500 + 790
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print "" & adoquery.Fields("Yno").Value
                  
               End If 'Added by Morgan 2014/9/17
                
               'Add by Morgan 2008/4/2
               If m_bPrinter = False Then
                  strExc(1) = Replace("" & adoquery("Caseno").Value, "-0-00", "")
                  strExc(1) = Replace("" & strExc(1), "-", "")
                  If strExc(1) <> m_stCaseNo Then
                     m_stCaseNos = m_stCaseNos & strExc(1) & ","
                  End If
               End If
               'end 2008/4/2
            End If
            'Add By Cheng 2003/03/21
            '取得原始匯率
            GetRateAndUSD "" & adoquery.Fields("Dno").Value
            '若選擇加總明細或加總明細首頁
            If Me.Text6.Text = "2" Or Me.Text6.Text = "3" Then
                intCols = intCols + 1
                intCounter = intCounter + 1
            End If
ChiJumpNext:
            adoquery.MoveNext
        Loop
        adoquery.Close
        
        'Added by Morgan 2014/9/16
        '設定特殊請款單, 記錄整批列印單號
        UpdateBatch Left(m_strA1K01, Len(m_strA1K01) - 1)
        'end 2014/9/16
        
        '若選擇加總明細或加總明細首頁
        If Me.Text6.Text = "2" Or Me.Text6.Text = "3" Then
            intCounter = intCounter + 1
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter - 1, 2, "", 4
               SetWordArray m_Item, intCounter, 2, "", 4
            Else
            'END 2014/9/17
            
               m_Device.Line (0 + intInit, 1500 + (intCounter) * 300 - 200)-(10200 + intInit, 1500 + (intCounter) * 300 - 200)
               
            End If 'Added by Morgan 2014/9/17
            
            intCounter = intCounter + 1
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, "TOTAL: " & intRecords & " cases", 4
            Else
            'END 2014/9/17
            
               m_Device.CurrentX = 2500 + intInit
               m_Device.CurrentY = 1500 + (intCounter) * 300
               m_Device.Print "TOTAL: " & intRecords & " cases"
               
            End If 'Added by Morgan 2014/9/17
            
            'Modify by Morgan 2011/7/27 新信紙有信尾要上移
            'intCounter = 45
            intCounter = 43
            
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 1, "P." & m_intPage, 6
            Else
            'END 2014/9/17
            
               m_Device.CurrentX = 4500 + intInit
               m_Device.CurrentY = 1500 + (intCounter) * 300
               m_Device.Print "P." & m_intPage
               
            End If 'Added by Morgan 2014/9/17
            
            intCounter = intCounter + 1
            
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 1, "Continued", 6
            Else
            'END 2014/9/17
            
               m_Device.CurrentX = 4200 + intInit
               m_Device.CurrentY = 1500 + (intCounter) * 300
               m_Device.Print "Continued"
               
            End If 'Added by Morgan 2014/9/17
            
        End If
        'Modified by Lydia 2020/09/16 +中文整批
        If Me.Text6.Text = "3" Or Me.Text6.Text = "5" Then
           GoTo NextRecord
        End If
        
         '加總明細
         If Me.Text6.Text = "2" Then
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               runWord
            Else
            'end 2014/9/17
               If m_bPrinter = True Then
                  'Modified by Morgan 2013/1/2
                  'm_Device.NewPage
                  MyNewPage
               Else
                  SetPic m_intPage
               End If
            End If
            m_intPage = m_intPage + 1
         ElseIf Me.Text6.Text = "4" Then
            m_intPage = 1
         End If
         
         'Add By Sindy 2025/4/8 若a1L16=a1K18直接使用輸入的a1L17 (減少計算上的誤差值)
         '原:sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount
         strFormula = "sum(decode(a1l07,0,decode(a1L16,a1K18,a1L17,trunc((a1l05 - a1l07) / nvl(a1k10, 1))),trunc((a1l05 - a1l07) / nvl(a1k10, 1))))"
         '2025/4/8 END
         
        '第二頁
         m_strDetailKind = "2"
         intCounter = 4
         If Me.Text6.Text = "2" Or Me.Text6.Text = "4" Then PrintHead
         adoquery.CursorLocation = adUseClient
        'Modified by Morgan 2013/1/2 Namount(明細美金)改無條件捨去,之前沒有印過純美金帳單(原程式有誤),所以沒有舊請款單不一致問題 sum((a1l05 - a1l07) / nvl(a1k10, 1))-->trunc((a1l05 - a1l07) / nvl(a1k10, 1))
        'Modified by Lydia 2020/09/16 (保留相同SQL到後面使用) 共同條件+請款日期a1k02 (因為版面是以同一日期為一份)
'        '專利
'         StrSQLa = "select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, fagent, acc1j0, patent where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and pa26='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConPA & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
'                       " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, customer, acc1j0, patent where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and pa26='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConPA & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
'        '商標
'         StrSQLa = StrSQLa & " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, fagent, acc1j0, trademark where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and tm23='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConTM & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
'                       " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, customer, acc1j0, trademark where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and tm23='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConTM & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
'         '法務
'         StrSQLa = StrSQLa & " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, fagent, acc1j0, lawcase where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and lc11='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConLC & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
'                       " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, customer, acc1j0, lawcase where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and lc11='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConLC & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
'        '服務
'         StrSQLa = StrSQLa & " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, fagent, acc1j0, servicepractice where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and sp08='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConSP & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
'                       " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, customer, acc1j0, servicepractice where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and sp08='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConSP & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
        '專利
         StrSQLa = "select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, fagent, acc1j0, patent where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02  and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
                       " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, customer, acc1j0, patent where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02  and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
        '商標
         StrSQLa = StrSQLa & " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, fagent, acc1j0, trademark where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02  and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
                       " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, customer, acc1j0, trademark where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02  and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
         '法務
         StrSQLa = StrSQLa & " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, fagent, acc1j0, lawcase where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02  and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
                       " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, customer, acc1j0, lawcase where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02  and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
        '服務
         StrSQLa = StrSQLa & " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, fagent, acc1j0, servicepractice where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02  and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
                       " union select a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, customer, acc1j0, servicepractice where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02  and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & " group by a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
        'end 2020/09/16
        '依輸入序號由大到小排序
        StrSQLa = StrSQLa & " Order By a1l02 Desc "
        '*******************************************
        'Modified by Lydia 2020/09/16 共同條件+請款日期a1k02 (因為版面是以同一日期為一份)
'        '專利
'         m_strSQLA = "select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, fagent, acc1j0, patent where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and pa26='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConPA & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
'                       " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, customer, acc1j0, patent where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and pa26='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConPA & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
'        '商標
'         m_strSQLA = m_strSQLA & " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, fagent, acc1j0, trademark where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and tm23='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConTM & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
'                       " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, customer, acc1j0, trademark where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and tm23='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConTM & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
'         '法務
'         m_strSQLA = m_strSQLA & " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, fagent, acc1j0, lawcase where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and lc11='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConLC & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
'                       " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, customer, acc1j0, lawcase where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and lc11='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConLC & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
'        '服務
'         m_strSQLA = m_strSQLA & " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, fagent, acc1j0, servicepractice where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and sp08='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConSP & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
'                       " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, sum(trunc((a1l05 - a1l07) / nvl(a1k10, 1))) as Namount from acc1k0, acc1l0, customer, acc1j0, servicepractice where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and sp08='" & adoacc1k0.Fields("cuno").Value & "' ") & " and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "'" & strSql & strConSP & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
        '專利
         m_strSQLA = "select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, fagent, acc1j0, patent where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
                       " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, customer, acc1j0, patent where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
        '商標
         m_strSQLA = m_strSQLA & " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, fagent, acc1j0, trademark where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
                       " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, customer, acc1j0, trademark where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
         '法務
         m_strSQLA = m_strSQLA & " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, fagent, acc1j0, lawcase where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
                       " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, customer, acc1j0, lawcase where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
        '服務
         m_strSQLA = m_strSQLA & " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, fagent, acc1j0, servicepractice where a1k01 = a1l01 and substr(a1k27, 1, 8) = fa01 and substr(a1k27, 9, 1) =  fa02 and a1l03 = a1j01 and a1l04 = a1j02 and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06 " & _
                       " union select a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06, sum(a1l05) as Amount, sum(a1l07) as Discount, sum((a1l05 - a1l07)) as Ramount, " & strFormula & " as Namount from acc1k0, acc1l0, customer, acc1j0, servicepractice where a1k01 = a1l01 and substr(a1k27, 1, 8) = cu01 and substr(a1k27, 9, 1) =  cu02 and a1l03 = a1j01 and a1l04 = a1j02 and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & " group by a1l01, a1l02, a1j01, a1j02, a1j04, a1j05, a1j06, a1j16, a1l06"
        'end 2020/09/16
        m_strSQLA = m_strSQLA & " Order By a1j02 "
        rsA.CursorLocation = adUseClient
        rsA.Open m_strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            StrSqlB = "Delete From ACCRPT425 Where ID ='" & strUserNum & "' "
            cnnConnection.Execute StrSqlB
            While Not rsA.EOF
                StrSqlB = "Select * From ACCRPT425 Where ID='" & strUserNum & "' "
                rsB.CursorLocation = adUseClient
                rsB.Open StrSqlB, cnnConnection, adOpenDynamic, adLockOptimistic
                rsB.AddNew
                rsB("R42501").Value = "" & rsA("a1j02").Value
                rsB("R42502").Value = CDbl(rsA("Amount").Value)
                rsB("R42503").Value = CDbl(rsA("Discount").Value)
                rsB("ID").Value = strUserNum
                rsB("XNo").Value = "" & rsA("a1l01").Value 'Add By Sindy 2020/7/24 + 請款編號
                rsB.UPDATE
                If rsB.State <> adStateClosed Then rsB.Close
                Set rsB = Nothing
                rsA.MoveNext
            Wend
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
         
         'Add By Sindy 2020/7/24
         m_strSQLA = "Select * From ACCRPT425 Where ID='" & strUserNum & "' Order By R42501"
         adoquery.CursorLocation = adUseClient
         adoquery.Open m_strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
         While Not adoquery.EOF
             '檢查有沒有代收代付(XXX98)要合併
             If Right("" & adoquery.Fields("R42501").Value, 2) = "98" Then
                 strKey = Mid("" & adoquery.Fields("R42501").Value, 1, Len("" & adoquery.Fields("R42501").Value) - 2)
                 m_strSQLA = "Update ACCRPT425 Set R42502=R42502+" & Val("" & adoquery.Fields("R42502").Value) & ", R42503=R42503+" & Val("" & adoquery.Fields("R42503").Value) & " Where ID='" & strUserNum & "' And R42501='" & strKey & "' And XNo='" & adoquery.Fields("XNo").Value & "'"
                 cnnConnection.Execute m_strSQLA
                 m_strSQLA = "Delete From ACCRPT425 Where ID='" & strUserNum & "' And R42501='" & adoquery.Fields("R42501").Value & "' And XNo='" & adoquery.Fields("XNo").Value & "'"
                 cnnConnection.Execute m_strSQLA
             End If
             adoquery.MoveNext
         Wend
         adoquery.Close
         '2020/7/24 END
         
        '********************************************
        'adoquery.CursorLocation = adUseClient
        adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly 'Memo by Lydia 2020/09/15 保留相同SQL
        '重新整理請款明細
        If adoquery.RecordCount > 0 Then
            StrSQLa = "Delete From ACCRPT424 Where ID ='" & strUserNum & "' "
            cnnConnection.Execute StrSQLa
            While Not adoquery.EOF
                StrSQLa = "Select * From ACCRPT424 Where ID='" & strUserNum & "' And R42402='" & adoquery("a1j01").Value & "' And R42403='" & adoquery("a1j02").Value & "' "
                rsA.CursorLocation = adUseClient
                rsA.Open StrSQLa, cnnConnection, adOpenDynamic, adLockOptimistic
                '若有資料
                If rsA.EOF = False Then
                    rsA("R42409").Value = CDbl(rsA("R42409").Value) + CDbl(adoquery("Amount").Value)
                    rsA("R42410").Value = CDbl(rsA("R42410").Value) + CDbl(adoquery("Discount").Value)
                    rsA("R42411").Value = CDbl(rsA("R42411").Value) + CDbl(adoquery("Ramount").Value)
                    rsA("R42412").Value = CDbl(rsA("R42412").Value) + CDbl(adoquery("Namount").Value)
                    rsA.UPDATE
                '若無資料
                Else
                    rsA.AddNew
                    rsA("R42401").Value = "" & adoquery("a1l02").Value
                    rsA("R42402").Value = "" & adoquery("a1j01").Value
                    rsA("R42403").Value = "" & adoquery("a1j02").Value
                    rsA("R42404").Value = "" & adoquery("a1j04").Value
                    rsA("R42405").Value = "" & adoquery("a1j05").Value
                    rsA("R42406").Value = "" & adoquery("a1j06").Value
                    rsA("R42407").Value = "" & adoquery("a1j16").Value
                    rsA("R42408").Value = "" & adoquery("a1l06").Value
                    rsA("R42409").Value = CDbl(adoquery("Amount").Value)
                    rsA("R42410").Value = CDbl(adoquery("Discount").Value)
                    rsA("R42411").Value = CDbl(adoquery("Ramount").Value)
                    rsA("R42412").Value = CDbl(adoquery("Namount").Value)
                    rsA("ID").Value = strUserNum
                    rsA.UPDATE
                End If
                If rsA.State <> adStateClosed Then rsA.Close
                Set rsA = Nothing
                adoquery.MoveNext
            Wend
            adoquery.Close
            
            StrSQLa = "Select R42401 AS a1l02, R42402 AS a1j01, R42403 AS a1j02, R42404 AS a1j04, R42405 AS a1j05, R42406 AS a1j06, R42407 AS a1j16, R42408 AS a1l06, R42409 AS Amount, R42410 AS Discount, R42411 AS Ramount, R42412 AS Namount   From ACCRPT424 Where ID='" & strUserNum & "' Order By R42401 "
            'adoquery.CursorLocation = adUseClient
            adoquery.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            'Add By Cheng 2003/09/30
            While Not adoquery.EOF
                'T的03併入01
                If "" & adoquery.Fields("a1j01").Value = "T" And "" & adoquery.Fields("a1j02").Value = "03" Then
                    StrSQLa = "Update ACCRPT424 Set R42409=R42409+" & Val("" & adoquery.Fields("Amount").Value) & ", R42410=R42410+" & Val("" & adoquery.Fields("Discount").Value) & ", R42411=R42411+" & Val("" & adoquery.Fields("Ramount").Value) & ", R42412=R42412+" & Val("" & adoquery.Fields("Namount").Value) & " Where ID='" & strUserNum & "' And R42402='T' And R42403='01' "
                    cnnConnection.Execute StrSQLa
                    StrSQLa = "Delete From ACCRPT424 Where ID='" & strUserNum & "' And R42401='" & adoquery.Fields("a1l02").Value & "' "
                    cnnConnection.Execute StrSQLa
                'Add By Sindy 2020/7/24 檢查有沒有代收代付(XXX98)要合併
                ElseIf Right("" & adoquery.Fields("a1j02").Value, 2) = "98" Then
                    strKey = Mid("" & adoquery.Fields("a1j02").Value, 1, Len("" & adoquery.Fields("a1j02").Value) - 2)
                    StrSQLa = "Update ACCRPT424 Set R42409=R42409+" & Val("" & adoquery.Fields("Amount").Value) & ", R42410=R42410+" & Val("" & adoquery.Fields("Discount").Value) & ", R42411=R42411+" & Val("" & adoquery.Fields("Ramount").Value) & ", R42412=R42412+" & Val("" & adoquery.Fields("Namount").Value) & " Where ID='" & strUserNum & "' And R42402='" & "" & adoquery.Fields("a1j01").Value & "' And R42403='" & strKey & "'"
                    cnnConnection.Execute StrSQLa
                    StrSQLa = "Delete From ACCRPT424 Where ID='" & strUserNum & "' And R42401='" & adoquery.Fields("a1l02").Value & "' "
                    cnnConnection.Execute StrSQLa
                '2020/7/24 END
                End If
                adoquery.MoveNext
            Wend
            'End
            adoquery.Close
            StrSQLa = "Select R42401 AS a1l02, R42402 AS a1j01, R42403 AS a1j02, R42404 AS a1j04, R42405 AS a1j05, R42406 AS a1j06, R42407 AS a1j16, R42408 AS a1l06, R42409 AS Amount, R42410 AS Discount, R42411 AS Ramount, R42412 AS Namount   From ACCRPT424 Where ID='" & strUserNum & "' Order By R42401 "
            adoquery.CursorLocation = adUseClient
            adoquery.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        End If
         
         Do While adoquery.EOF = False
            If intCounter >= 40 Then
               'Modify by Morgan 2011/7/27 新信紙有信尾要上移
               'intCounter = 45
               intCounter = 43
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 1, "P." & m_intPage, 6
               Else
               'END 2014/9/17
               
                  m_Device.CurrentX = 4500 + intInit
                  m_Device.CurrentY = 1500 + (intCounter) * 300
                  m_Device.Print "P." & m_intPage
                  
               End If 'Added by Morgan 2014/9/17
               
               intCounter = intCounter + 1
               
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 1, "Continued", 6
               Else
               'END 2014/9/17
               
                  m_Device.CurrentX = 4200 + intInit
                  m_Device.CurrentY = 1500 + (intCounter) * 300
                  m_Device.Print "Continued"
                  
               End If 'Added by Morgan 2014/9/17
               
               m_strDetailKind = "2"
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  runWord
               Else
               'end 2014/9/17
                  If m_bPrinter = True Then
                     'Modified by Morgan 2013/1/2
                     'm_Device.NewPage
                     MyNewPage
                  Else
                     SetPic m_intPage
                  End If
               End If
               m_intPage = m_intPage + 1
               intCounter = 4
               PrintHead
            End If
            strDescription = ""
            
            If Not m_bolWord Then 'Added by Morgan 2014/9/17
               m_Device.CurrentX = 0 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
            End If

            Select Case strLanguage
               Case "2"
                  '初始化細項算式明細
                  strPrintText1 = ""
                  
                  '若無請款細項科目
                  If IsNull(adoquery.Fields("a1j04").Value) Then
                     'm_Device.Print ""
                  '若有請款細項科目
                  Else
                    '若請款細項科目為"Official fees"
                     If adoquery.Fields("a1j04").Value = "Official fees" Then
                        If "" & adoquery.Fields("a1l06").Value <> "" Then
                            strPrintText = adoquery.Fields("a1j04").Value & PUB_GetA1L06Text("" & adoquery.Fields("a1l06").Value) & strDescription
                        Else
                            strPrintText = adoquery.Fields("a1j04").Value
                        End If
                        strPrintText1 = GetItemDetail("" & adoquery.Fields("a1j02").Value)
                        If CountLength(strPrintText & strPrintText1) <= 70 Then
                           'Added by Morgan 2014/9/17
                           If m_bolWord Then
                              SetWordArray m_Item, intCounter, 1, strPrintText & strPrintText1, 3
                           Else
                           'END 2014/9/17
                           
                              m_Device.Print strPrintText & strPrintText1
                              
                           End If 'Added by Morgan 2014/9/17
                           
                        Else
                           'Added by Morgan 2014/9/17
                           If m_bolWord Then
                              SetWordArray m_Item, intCounter, 1, strPrintText, 3
                           Else
                           'END 2014/9/17
                              m_Device.Print strPrintText
                           End If 'Added by Morgan 2014/9/17
                           
                           If strPrintText1 <> "" Then
                              intCounter = intCounter + 1
                              'Added by Morgan 2014/9/17
                              If m_bolWord Then
                                 SetWordArray m_Item, intCounter, 1, strPrintText1, 3
                              Else
                              'END 2014/9/17
                              
                                 PrintDropLine Trim(strPrintText1), "", intCounter, 70
                                 
                              End If 'Added by Morgan 2014/9/17
                              
                            End If
                            'End
                        End If
                     
                     'Added by Morgan 2016/12/2
                     '使用Word表格時請款項目合併後顯示讓Word自動折行(參照單張請款單模式)
                     ElseIf m_bolWord Then
                        strPrintText1 = GetItemDetail("" & adoquery.Fields("a1j02").Value)
                        strPrintText = adoquery.Fields("a1j04").Value & strDescription
                        If Not IsNull(adoquery.Fields("a1j05").Value) Then
                           strPrintText = strPrintText & " " & adoquery.Fields("a1j05").Value
                        End If
                        If Not IsNull(adoquery.Fields("a1j06").Value) Then
                           strPrintText = strPrintText & " " & adoquery.Fields("a1j06").Value
                        End If
                        strPrintText = strPrintText & strPrintText1
                        SetWordArray m_Item, intCounter, 1, strPrintText, 3
                        
                     '若為其他請款細項科目
                     Else
                        '若項目欄2無資料
                        If IsNull(adoquery.Fields("a1j05").Value) Then
                           strPrintText1 = GetItemDetail("" & adoquery.Fields("a1j02").Value)
                           strPrintText = adoquery.Fields("a1j04").Value & strDescription
                           'Modify by Morgan 2005/10/26 控制折行
                           'm_Device.Print strPrintText & strPrintText1
                           'Added by Morgan 2014/9/17
                           If m_bolWord Then
                              SetWordArray m_Item, intCounter, 1, strPrintText, 3
                           Else
                           'END 2014/9/17
                           
                              PrintDropLine Trim(strPrintText), "", intCounter, 70
                              
                           End If 'Added by Morgan 2014/9/17
                           
                           If Trim(strPrintText1) <> "" Then
                              intCounter = intCounter + 1
                              'Added by Morgan 2014/9/17
                              If m_bolWord Then
                                 SetWordArray m_Item, intCounter, 1, strPrintText1, 3
                              Else
                              'END 2014/9/17
                              
                                 PrintDropLine Trim(strPrintText1), "", intCounter, 70
                                 
                              End If 'Added by Morgan 2014/9/17
                              
                           End If
                           
                        '若項目欄2有資料
                        Else
                           strPrintText = adoquery.Fields("a1j04").Value
                           '列印請款細項科目
                           If strPrintText1 = "" Then
                              'Added by Morgan 2014/9/17
                              If m_bolWord Then
                                 SetWordArray m_Item, intCounter, 1, strPrintText, 3
                              Else
                              'END 2014/9/17
                              
                                 m_Device.Print strPrintText
                                 
                              End If 'Added by Morgan 2014/9/17
                              
                           Else
                                If CountLength(strPrintText & strPrintText1) <= 70 Then
                                    'Added by Morgan 2014/9/17
                                    If m_bolWord Then
                                       SetWordArray m_Item, intCounter, 1, strPrintText & strPrintText1, 3
                                    Else
                                    'END 2014/9/17
                                       m_Device.Print strPrintText & strPrintText1
                                       
                                    End If 'Added by Morgan 2014/9/17
                                Else
                                    'Added by Morgan 2014/9/17
                                    If m_bolWord Then
                                       SetWordArray m_Item, intCounter, 1, strPrintText, 3
                                    Else
                                    'END 2014/9/17
                                       m_Device.Print strPrintText
                                       
                                    End If 'Added by Morgan 2014/9/17
                                    
                                    If strPrintText1 <> "" Then
                                       intCounter = intCounter + 1
                                       'Added by Morgan 2014/9/17
                                       If m_bolWord Then
                                          SetWordArray m_Item, intCounter, 1, strPrintText1, 3
                                       Else
                                       'END 2014/9/17
                                       
                                          PrintDropLine Trim(strPrintText1), "", intCounter, 70
                                          
                                       End If 'Added by Morgan 2014/9/17
                                    End If
                                End If
                                'End
                           End If
                           
                           
'                        End If
                        End If
                        

                     End If
                     
If Not m_bolWord Then 'Added by Morgan 2016/12/2

                     '若項目欄2有資料
                     If IsNull(adoquery.Fields("a1j05").Value) = False Then

                        intCounter = intCounter + 1
                        If Not m_bolWord Then 'Added by Morgan 2014/9/17
                           m_Device.CurrentX = 0 + intInit
                           m_Device.CurrentY = 1500 + intCounter * 300
                        End If 'Added by Morgan 2014/9/17
                        
                        '若項目欄3無資料
                        If IsNull(adoquery.Fields("a1j06").Value) Then
                           strPrintText1 = GetItemDetail("" & adoquery.Fields("a1j02").Value)
                           strPrintText = adoquery.Fields("a1j05").Value & strDescription
                           'Modify by Morgan 2004/7/23
                           '加長度控制
                           'm_Device.Print strPrintText & strPrintText1
                           If CountLength(strPrintText & strPrintText1) <= 70 Then
                              'Added by Morgan 2014/9/17
                              If m_bolWord Then
                                 SetWordArray m_Item, intCounter, 1, strPrintText & strPrintText1, 3
                              Else
                              'END 2014/9/17
                              
                                 m_Device.Print strPrintText & strPrintText1
                                 
                              End If 'Added by Morgan 2014/9/17
                              
                           Else
                              'Added by Morgan 2014/9/17
                              If m_bolWord Then
                                 SetWordArray m_Item, intCounter, 1, strPrintText, 3
                              Else
                              'END 2014/9/17
                                 m_Device.Print strPrintText
                              End If 'Added by Morgan 2014/9/17
                              
                              If strPrintText1 <> "" Then
                                 intCounter = intCounter + 1
                                 'Added by Morgan 2014/9/17
                                 If m_bolWord Then
                                    SetWordArray m_Item, intCounter, 1, strPrintText1, 3
                                 Else
                                 'END 2014/9/17
                                 
                                    PrintDropLine Trim(strPrintText1), "", intCounter, 70
                                    
                                 End If 'Added by Morgan 2014/9/17
                                 
                              End If
                           End If
                        '若項目欄3有資料
                        Else
                           strPrintText = adoquery.Fields("a1j05").Value
                           If strPrintText1 = "" Then
                              'Added by Morgan 2014/9/17
                              If m_bolWord Then
                                 SetWordArray m_Item, intCounter, 1, strPrintText, 3
                              Else
                              'END 2014/9/17
                              
                                 m_Device.Print strPrintText
                                 
                              End If 'Added by Morgan 2014/9/17
                           Else
                              If CountLength(strPrintText & strPrintText1) <= 70 Then
                                 'Added by Morgan 2014/9/17
                                 If m_bolWord Then
                                    SetWordArray m_Item, intCounter, 1, strPrintText & strPrintText1, 3
                                 Else
                                 'END 2014/9/17
                                 
                                    m_Device.Print strPrintText & strPrintText1
                                    
                                 End If 'Added by Morgan 2014/9/17
                              Else
                                 'Added by Morgan 2014/9/17
                                 If m_bolWord Then
                                    SetWordArray m_Item, intCounter, 1, strPrintText, 3
                                 Else
                                 'END 2014/9/17
                                    m_Device.Print strPrintText
                                 End If 'Added by Morgan 2014/9/17
                                 
                                 If strPrintText1 <> "" Then
                                    intCounter = intCounter + 1
                                    'Added by Morgan 2014/9/17
                                    If m_bolWord Then
                                       SetWordArray m_Item, intCounter, 1, strPrintText1, 3
                                    Else
                                    'END 2014/9/17
                                       PrintDropLine Trim(strPrintText1), "", intCounter, 70
                                    End If 'Added by Morgan 2014/9/17
                                 End If
                              End If
                                'End
                            End If
'                        End If
                        End If
                     End If
                     
                     '若項目欄3有資料
                     If IsNull(adoquery.Fields("a1j06").Value) = False Then
                        intCounter = intCounter + 1
                        If Not m_bolWord Then 'Added by Morgan 2014/9/17
                           m_Device.CurrentX = 0 + intInit
                           m_Device.CurrentY = 1500 + intCounter * 300
                        End If
                        strPrintText = adoquery.Fields("a1j06").Value & strDescription
                        strPrintText1 = GetItemDetail("" & adoquery.Fields("a1j02").Value)
                        strPrintText = strPrintText
                        
                        If strPrintText1 = "" Then
                           'Added by Morgan 2014/9/17
                           If m_bolWord Then
                              SetWordArray m_Item, intCounter, 1, strPrintText, 3
                           Else
                           'END 2014/9/17
                              m_Device.Print strPrintText
                           End If 'Added by Morgan 2014/9/17
                           
                        Else
                           If CountLength(strPrintText & strPrintText1) <= 70 Then
                              'Added by Morgan 2014/9/17
                              If m_bolWord Then
                                 SetWordArray m_Item, intCounter, 1, strPrintText & strPrintText1, 3
                              Else
                              'END 2014/9/17
                                 m_Device.Print strPrintText & strPrintText1
                              End If 'Added by Morgan 2014/9/17
                           Else
                              'Added by Morgan 2014/9/17
                              If m_bolWord Then
                                 SetWordArray m_Item, intCounter, 1, strPrintText, 3
                              Else
                              'END 2014/9/17
                                 m_Device.Print strPrintText
                              End If 'Added by Morgan 2014/9/17
                              
                              If strPrintText1 <> "" Then
                                 intCounter = intCounter + 1
                                 'Added by Morgan 2014/9/17
                                 If m_bolWord Then
                                    SetWordArray m_Item, intCounter, 1, strPrintText1, 3
                                 Else
                                 'END 2014/9/17
                                    PrintDropLine Trim(strPrintText1), "", intCounter, 70
                                 End If 'Added by Morgan 2014/9/17
                              End If
                           End If
                           'End
                        End If
'                        End If
                     End If
End If 'Added by Morgan 2016/12/2
                     
                  End If
                  
               Case "3"
                  If IsNull(adoquery.Fields("a1j16").Value) Then
                     'm_Device.Print ""
                  Else
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 1, adoquery.Fields("a1j16").Value & strDescription, 3
                     Else
                     'END 2014/9/17
                        m_Device.Print adoquery.Fields("a1j16").Value & strDescription
                     End If 'Added by Morgan 2014/9/17
                  End If
            End Select
            
            If Not m_bolWord Then 'Added by Morgan 2014/9/17
               m_Device.CurrentX = 7000 + 1200 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
            End If
            
            'Modified by Morgan 2013/1/2
'            Select Case strCurr
'               Case "U"
'                  m_Device.Print "USD"
'               Case Else
'                  m_Device.Print "NTD"
'            End Select
            Select Case m_iPrintCurrType
            Case 3, 4 '純外幣,外幣+美金合計
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, m_DNCurr, 3
               Else
               'END 2014/9/17
                  m_Device.Print m_DNCurr
               End If 'Added by Morgan 2014/9/17
            Case Else
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, "NTD", 3
               Else
               'END 2014/9/17
                  m_Device.Print "NTD"
               End If 'Added by Morgan 2014/9/17
            End Select
            'end 2013/1/2

            If IsNull(adoquery.Fields("Amount").Value) = False Then
               'Modified by Morgan 2013/1/2
'               Select Case strCurr
'                  Case "U"
'                     strAmount = PUB_ChgFormat(Val(adoquery.Fields("Ramount").Value) / Val(adoacc1k0.Fields("a1k10").Value), True)
'                     douAmount = douAmount + Val(adoquery.Fields("Ramount").Value) * Val(adoacc1k0.Fields("a1k10").Value)
'                  Case Else
'                     strAmount = PUB_ChgFormat(adoquery.Fields("Ramount").Value, True)
'                     douAmount = douAmount + Val(adoquery.Fields("Ramount").Value)
'               End Select
'               douUSDollar = douUSDollar + Val(adoquery.Fields("Namount").Value)
               Select Case m_iPrintCurrType
               Case 3, 4 '3.純外幣, 4.外幣+美金合計
                  'Modified by Morgan 2016/12/6 有非美金
                  'If m_DNCurr = "USD" Then
                     '改直接抓 Namount 否則會與合計不符,原程式累計金額有誤一併修正
                     strAmount = Format(Val(adoquery.Fields("Namount").Value), FDollar)
                     douAmount = douAmount + Val(adoquery.Fields("Namount").Value)
                     
                  ''原程式不會有非美金,要用再加
                  'Else
                  '
                  'End If
                  'end 2016/12/6
               Case Else
                  strAmount = PUB_ChgFormat(adoquery.Fields("Ramount").Value, True)
                  douAmount = douAmount + Val(adoquery.Fields("Ramount").Value)
               End Select
               '若請款幣別非美金且原匯率不是放美金匯率時要改
               'douUSDollar = douUSDollar + Val(adoquery.Fields("Namount").Value) 'Removed by Morgan 2016/12/6 目前沒用
               'end 2013/1/2
               
               strAmount = Format(strAmount, FDollar) 'Added by Morgan 2013/1/3 金額都要印.00
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 3, strAmount, 3
               Else
               'END 2014/9/17
                  intLength = m_Device.TextWidth(strAmount)
                  m_Device.CurrentX = 10200 - intLength + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print strAmount
               End If 'Added by Morgan 2014/9/17
            End If
            intCounter = intCounter + 2
            adoquery.MoveNext
         Loop
         adoquery.Close
      '選擇代理人
        '列印方式選擇--加總首頁
      Else
         m_strDetailKind = ""
         PrintHead
         intCounter = 18
         'Added by Morgan 2014/9/17
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 2, "INVOICE COVER SHEET", 3, True, True
         Else
         'END 2014/9/17
            m_Device.FontBold = True
            m_Device.FontItalic = True
            m_Device.FontUnderline = True
            m_Device.FontSize = 16
            m_Device.CurrentX = 1500
            m_Device.CurrentY = 1500 + intCounter * 300
            m_Device.Print "INVOICE COVER SHEET"
            m_Device.FontUnderline = False
         End If 'Added by Morgan 2014/9/17
         
         If Not m_bolWord Then 'Added by Morgan 2014/9/17
            m_Device.CurrentX = 5500
            m_Device.CurrentY = 1500 + intCounter * 300
         End If
         adoquery.CursorLocation = adUseClient
         'Remove by Lydia 2020/09/16 共同條件+請款日期a1k02 (因為版面是以同一日期為一份)
        'If Me.Text2.Text = "FCT" Or Me.Text2.Text = "S" Then
        '    '專利
        '     StrSQLa = "select a1k01 from acc1k0, patent where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & " and pa26='" & adoacc1k0.Fields("cuno").Value & "' "
        '    '商標
        '     StrSQLa = StrSQLa & " union select a1k01 from acc1k0, trademark where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & " and tm23='" & adoacc1k0.Fields("cuno").Value & "' "
        '    '法務
        '     StrSQLa = StrSQLa & " union select a1k01 from acc1k0, lawcase where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & " and lc11='" & adoacc1k0.Fields("cuno").Value & "' "
        '    '服務
        '     StrSQLa = StrSQLa & " union select a1k01 from acc1k0, servicepractice where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & " and sp08='" & adoacc1k0.Fields("cuno").Value & "' "
        'Else
            '專利
             StrSQLa = "select a1k01 from acc1k0, patent where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & " "
            '商標
             StrSQLa = StrSQLa & " union select a1k01 from acc1k0, trademark where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & " "
            '法務
             StrSQLa = StrSQLa & " union select a1k01 from acc1k0, lawcase where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & " "
            '服務
             StrSQLa = StrSQLa & " union select a1k01 from acc1k0, servicepractice where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & " "
        'End If
        'end 2020/09/16
        
        StrSQLa = StrSQLa & " order by a1k01" 'Added by Morgan 2014/9/5
         adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 3, "NO. " & adoquery.Fields("a1k01").Value, 3, True
            Else
            'END 2014/9/17
               m_Device.Print "NO. " & adoquery.Fields("a1k01").Value
            End If 'Added by Morgan 2014/9/17
         Else
            MsgBox MsgText(28), , MsgText(5)
            If m_bPrinter = True Then
               If Not m_bolWord Then 'Added by Morgan 2014/9/17
                  m_Device.KillDoc
               End If
            End If
            adoacc1k0.Close
            adoquery.Close
            Exit Sub
         End If
         adoquery.Close
         intCounter = intCounter + 4
         If Not m_bolWord Then 'Added by Morgan 2014/9/17
            m_Device.FontItalic = False
            m_Device.FontSize = 12
            m_Device.CurrentX = 3500
            m_Device.CurrentY = 1500 + intCounter * 300
         End If
         adoquery.CursorLocation = adUseClient
        'Modify By Cheng 2003/04/29
        'Remove by Lydia 2020/09/16 共同條件+請款日期a1k02 (因為版面是以同一日期為一份)
        'If Me.Text2.Text = "FCT" Or Me.Text2.Text = "S" Then
        '    '專利
        '     StrSQLa = "select sum(a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / nvl(a1k10, 0)), count(a1k01) from acc1k0, patent where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & " and pa26='" & adoacc1k0.Fields("cuno").Value & "' "
        '      '商標
        '     StrSQLa = StrSQLa & " union select sum(a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / nvl(a1k10, 0)), count(a1k01) from acc1k0, trademark where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & " and tm23='" & adoacc1k0.Fields("cuno").Value & "' "
        '    '法務
        '     StrSQLa = StrSQLa & " union select sum(a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / nvl(a1k10, 0)), count(a1k01) from acc1k0, lawcase where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & " and lc11='" & adoacc1k0.Fields("cuno").Value & "' "
        '    '服務
        '     StrSQLa = StrSQLa & " union select sum(a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / nvl(a1k10, 0)), count(a1k01) from acc1k0, servicepractice where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & " and sp08='" & adoacc1k0.Fields("cuno").Value & "' "
        ' Else
            '專利
             StrSQLa = "select sum(a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / nvl(a1k10, 0)), count(a1k01) from acc1k0, patent where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04 " & strSql & strConPA & " "
            '商標
             StrSQLa = StrSQLa & " union select sum(a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / nvl(a1k10, 0)), count(a1k01) from acc1k0, trademark where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04 " & strSql & strConTM & " "
            '法務
             StrSQLa = StrSQLa & " union select sum(a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / nvl(a1k10, 0)), count(a1k01) from acc1k0, lawcase where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04 " & strSql & strConLC & " "
            '服務
             StrSQLa = StrSQLa & " union select sum(a1k08 - nvl(a1k06, 0) - nvl(a1k30, 0) / nvl(a1k10, 0)), count(a1k01) from acc1k0, servicepractice where (a1k29 is null or a1k29 = '') and (a1k12 is null or a1k12 = '') and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 " & strSql & strConSP & " "
        'End If
        'end 2020/09/16
        
         adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
         
         'Modified by Morgan 2013/1/18
         'Select Case strCurr
         '   Case "N"
         Select Case m_iPrintCurrType
            Case 1
         'end 2013/1/18
         
               If adoquery.RecordCount <> 0 Then
                  If IsNull(adoquery.Fields(0).Value) Then
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, "TOTAL: NTD ", 4
                     Else
                     'END 2014/9/17
                        m_Device.Print "TOTAL: NTD "
                     End If 'Added by Morgan 2014/9/17
                  Else
                     'Modified by Morgan 2013/1/3 金額都要印.00
                     'm_Device.Print "TOTAL: NTD " & PUB_ChgFormat(Val(adoquery.Fields(0).Value) * Val(adoacc1k0.Fields("a1k10").Value), True)
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, "TOTAL: NTD " & Format(Val(adoquery.Fields(0).Value) * Val(adoacc1k0.Fields("a1k10").Value), FDollar), 4
                     Else
                     'END 2014/9/17
                        m_Device.Print "TOTAL: NTD " & Format(Val(adoquery.Fields(0).Value) * Val(adoacc1k0.Fields("a1k10").Value), FDollar)
                     End If 'Added by Morgan 2014/9/17
                  End If
                  intCounter = intCounter + 1
                  If IsNull(adoquery.Fields(1).Value) = False Then
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, "( " & adoquery.Fields(1).Value & " cases)", 4
                     Else
                     'END 2014/9/17
                        m_Device.CurrentX = 4500
                        m_Device.CurrentY = 1500 + intCounter * 300
                        m_Device.Print "( " & adoquery.Fields(1).Value & " cases)"
                     End If 'Added by Morgan 2014/9/17
                  End If
               Else
                  'Added by Morgan 2014/9/17
                  If m_bolWord Then
                     SetWordArray m_Item, intCounter, 2, "TOTAL: NTD ", 4
                  Else
                  'END 2014/9/17
                     m_Device.Print "TOTAL: NTD "
                  End If 'Added by Morgan 2014/9/17
               End If
            Case Else
               If adoquery.RecordCount <> 0 Then
                  If IsNull(adoquery.Fields(0).Value) Then
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, "TOTAL: USD ", 4
                     Else
                     'END 2014/9/17
                        m_Device.Print "TOTAL: USD "
                     End If 'Added by Morgan 2014/9/17
                  Else
                     'Modified by Morgan 2013/1/3 金額都要印.00
                     'm_Device.Print "TOTAL: USD " & PUB_ChgFormat(adoquery.Fields(0).Value, True)
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, "TOTAL: USD " & Format(adoquery.Fields(0).Value, FDollar), 4
                     Else
                     'END 2014/9/17
                        m_Device.Print "TOTAL: USD " & Format(adoquery.Fields(0).Value, FDollar)
                     End If 'Added by Morgan 2014/9/17
                  End If
                  intCounter = intCounter + 1
                  If IsNull(adoquery.Fields(1).Value) = False Then
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, "( " & adoquery.Fields(1).Value & " cases)", 4
                     Else
                     'END 2014/9/17
                        m_Device.CurrentX = 4500
                        m_Device.CurrentY = 1500 + intCounter * 300
                        m_Device.Print "( " & adoquery.Fields(1).Value & " cases)"
                     End If 'Added by Morgan 2014/9/17
                  End If
               Else
                  'Added by Morgan 2014/9/17
                  If m_bolWord Then
                     SetWordArray m_Item, intCounter, 2, "TOTAL: USD ", 4
                  Else
                  'END 2014/9/17
                     m_Device.Print "TOTAL: USD "
                  End If 'Added by Morgan 2014/9/17
               End If
         End Select
         adoquery.Close
         intCounter = intCounter + 6
         'Added by Morgan 2014/9/17
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 1, ReportSum(71001), 5
         Else
         'END 2014/9/17
            m_Device.FontBold = False
            m_Device.CurrentX = 500
            m_Device.CurrentY = 1500 + intCounter * 300
            m_Device.Print ReportSum(71001)
         End If 'Added by Morgan 2014/9/17
         
         intCounter = intCounter + 1
         'Added by Morgan 2014/9/17
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 1, ReportSum(73001), 5
         Else
         'END 2014/9/17
            m_Device.CurrentX = 500
            m_Device.CurrentY = 1500 + intCounter * 300
            m_Device.Print ReportSum(73001)
         End If 'Added by Morgan 2014/9/17
      
         'Add by Morgan 2006/7/5 加印 "Wire Transfer Preferred" -- 蘇副總
         'Added by Morgan 2014/9/17
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 2, "Wire", 5
            SetWordArray m_Item, intCounter + 1, 2, "Transfer", 5
            SetWordArray m_Item, intCounter + 2, 2, "Preferred", 5
         Else
         'END 2014/9/17
            lngBoxX = 500 + 7500
            lngBoxY = 1500 + intCounter * 300
            m_Device.DrawWidth = 5
            m_Device.Line (lngBoxX, lngBoxY - 100)-(lngBoxX + 1200, lngBoxY + 900), , B
            m_Device.DrawWidth = 1
            m_Device.CurrentX = lngBoxX + 150
            m_Device.CurrentY = lngBoxY
            m_Device.Print "Wire"
            m_Device.CurrentX = lngBoxX + 150
            m_Device.CurrentY = lngBoxY + 300
            m_Device.Print "Transfer"
            m_Device.CurrentX = lngBoxX + 150
            m_Device.CurrentY = lngBoxY + 600
            m_Device.Print "Preferred"
         End If 'Added by Morgan 2014/9/17
         'end 2006/7/5
         
         intCounter = intCounter + 1
         'Added by Morgan 2014/9/17
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 1, ReportSum(85), 5
         Else
         'END 2014/9/17
            m_Device.CurrentX = 500
            m_Device.CurrentY = 1500 + intCounter * 300
            m_Device.Print ReportSum(85)
         End If 'Added by Morgan 2014/9/17
         
         intCounter = intCounter + 1
         'Added by Morgan 2014/9/17
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 1, ReportSum(74), 5
         Else
         'END 2014/9/17
            m_Device.CurrentX = 500
            m_Device.CurrentY = 1500 + intCounter * 300
            m_Device.Print ReportSum(74)
         End If 'Added by Morgan 2014/9/17
         
'Removed by Morgan 2023/3/7 2013/12/17 已取消此帳戶
'         intCounter = intCounter + 1
'         'Added by Morgan 2014/9/17
'         If m_bolWord Then
'            SetWordArray m_Item, intCounter, 1, ReportSum(129), 5
'         Else
'         'END 2014/9/17
'            m_Device.CurrentX = 500
'            m_Device.CurrentY = 1500 + intCounter * 300
'            m_Device.Print ReportSum(129)
'         End If 'Added by Morgan 2014/9/17
'end 2023/3/7
         
         intCounter = intCounter + 1
         'Added by Morgan 2014/9/17
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 1, ReportSum(121), 5
         Else
         'END 2014/9/17
            m_Device.CurrentX = 500
            m_Device.CurrentY = 1500 + intCounter * 300
            m_Device.Print ReportSum(121)
         End If 'Added by Morgan 2014/9/17
         
         intCounter = intCounter + 1
         'Added by Morgan 2014/9/17
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 1, ReportSum(75), 5
         Else
         'END 2014/9/17
            m_Device.CurrentX = 500
            m_Device.CurrentY = 1500 + intCounter * 300
            m_Device.Print ReportSum(75)
         End If 'Added by Morgan 2014/9/17
         
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select usxr02 from usxrate where usxr01 in (select max(usxr01) from usxrate)", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields("usxr02").Value) Then
               'm_Device.Print ""
            Else
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 1, PUB_ChgFormat(adoquery.Fields("usxr02").Value, True), 5
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 3530
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print PUB_ChgFormat(adoquery.Fields("usxr02").Value, True)
               End If 'Added by Morgan 2014/9/17
            End If
         Else
            'm_Device.Print ""
         End If
         adoquery.Close
         
         intCounter = intCounter + 1
         'Added by Morgan 2014/9/17
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 1, ReportSum(86), 5, True
         Else
         'END 2014/9/17
            m_Device.FontBold = True
            m_Device.CurrentX = 500
            m_Device.CurrentY = 1500 + intCounter * 300
            m_Device.Print ReportSum(86)
            m_Device.FontBold = False
         End If
         
            
         'Added by Morgan 2014/9/17
         If m_bolWord Then
            runWord
         Else
         'end 2014/9/17
            If m_bPrinter = True Then
               'Modified by Morgan 2013/1/2
               'm_Device.NewPage
               MyNewPage
            Else
               SetPic m_intPage
            End If
         End If
         
         m_intPage = m_intPage + 1
         intCounter = 0
      End If
   
NextRecord:
      adoacc1k0.MoveNext
   Loop
   '列印方式選擇--加總明細
   If Me.Text6.Text <> "1" Then
       If Me.Text6.Text = "2" Or Me.Text6.Text = "4" Then
         If Not m_bolWord Then 'Added by Morgan 2014/9/17
            m_Device.Line (0 + intInit, 1500 + intCounter * 300 - 200)-(10200 + intInit, 1500 + intCounter * 300 - 200)
         End If
         PrintSum
         'Modify by Morgan 2011/7/27 新信紙有信尾要上移
         'intCounter = 45
         intCounter = 43
         'Added by Morgan 2014/9/17
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 1, "P." & m_intPage, 6
         Else
         'END 2014/9/17
            m_Device.CurrentX = 4500 + intInit
            m_Device.CurrentY = 1500 + intCounter * 300
            m_Device.Print "P." & m_intPage
         End If 'Added by Morgan 2014/9/17
       End If
   End If
   adoacc1k0.Close
   
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      runWord True
   Else
   'end 2014/9/17
      '加總首頁, 加總明細
      If m_bPrinter = True Then
         m_Device.EndDoc
         'Added by Morgan 2013/1/2
         If m_bPrint2Pdf Then
            frmPDF.EndtProcess
            CopyImg
         End If
         'end 2013/1/2
      Else
         If m_intPage > 1 Then
            SetPic m_intPage
         Else
            SetPic 0
         End If
         CopyImg
      End If
   End If
   
ChiJumpLast:
   'Added by Lydia 2015/04/14 整批請款單列印
   If Me.Text6.Text = "5" Then
      'Modified by Lydia 2020/09/16
      'adoacc1k0.Close
      'adoquery.Close
      If adoacc1k0.State <> adStateClosed Then adoacc1k0.Close
      If adoquery.State <> adStateClosed Then adoquery.Close
      'end 2020/09/16
      If bolChineseDB = True Then
        '單號起
        strExc(5) = Mid(strLang2, 1, InStr(strLang2, ",") - 1)
        If Right(strLang2, 1) = "," Then strLang2 = Mid(strLang2, 1, Len(strLang2) - 1)
        '單號止
        strExc(6) = Right(strLang2, Len(strExc(5)))
        '記錄整批列印單號(A1K32)
        strExc(7) = Replace(strLang2, ",", "','")
        strExc(7) = "'" & strExc(7) & "'"
        UpdateBatch Left(strExc(7), Len(strExc(7)))
        Load Frmacc2480
        With Frmacc2480
           .Text1.Text = strExc(5)
           .Text2.Text = strExc(6)
           .txtOutMode = Me.txtOutMode.Text
           .m_bolChiDB = True
           .m_ChiSys = Me.Text2.Text
           .m_ChiApply = Me.Text1.Text
           .m_ChiCust = Me.Text5.Text
           .m_bEditDoc = True
           .m_bBeCalled = True
           .m_CallPrevForm = Me.Name 'Added by Lydia 2020/01/06 呼叫請款單的程式名稱
           .m_bEMail = True
           .m_Chi2Word = m_bolWord 'Word + 信頭信尾
           .m_SavePath = m_EFilePath
           .m_ChiArrNO = strExc(7) 'Added by Lydia 2015/08/10 傳-整批5.中文請款單的單號
           .Command2_Click
        End With
        Unload Frmacc2480
        strFormName = Me.Name
        tool3_enabled
      End If
   End If
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead()
Dim intRow As Integer
'Dim strCustName As String
Dim strSystemName As String
Dim strCaseName As String
Dim strAppNo As String
Dim strCustNo As String
Dim strYes As String
'Add By Cheng 2003/03/20
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Modified by Morgan 2022/9/1 會超過5個,改動態
'Dim strCustName(4) As String '申請人名稱
Dim strCustName() As String '申請人名稱
'end 2022/9/1
'add by nickc 2007/02/08
Dim ii As Integer
Dim intTop As Integer
Dim strTitle As String 'Added by Morgan 2016/3/15

'Added by Morgan 2016/3/15
If Text1 = "Y52960000" Or Text1 = "Y54391000" Then
   strTitle = "  INVOICE "
Else
   strTitle = "DEBIT NOTE"
End If
'end 2016/3/15
   
   'Added by Morgan 2014/9/9
   Erase m_Item
   ReDim m_Item(1)
   'end 2014/9/9
   
   strYes = ""
   strSQL1 = ""
   'Modify by Morgan 2011/7/27
   'intCounter = 4
   intCounter = 2
    '系統類別
    If Text2 <> MsgText(601) Then
        strSQL1 = strSQL1 & " and a1k13 = '" & Text2 & "'"
    End If
    '選擇請款對象
    If Option1.Value Then
        '請款日期
        If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
           strSQL1 = strSQL1 & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
        End If
        If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
           strSQL1 = strSQL1 & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
        End If
    '選擇請款編號
    Else
        '請款編號
        If Text3 <> MsgText(601) Then
           strSQL1 = strSQL1 & " and a1k01 >= '" & Text3 & "'"
        End If
        If Text4 <> MsgText(601) Then
           strSQL1 = strSQL1 & " and a1k01 <= '" & Text4 & "'"
        End If
    End If
    '請款對象
    If Text1 <> MsgText(601) Then
        strSQL1 = strSQL1 & " and a1k28 = '" & Text1 & "'"
    End If

         rsA.CursorLocation = adUseClient
         'Modify by Morgan 2010/5/10 判斷沒設定稿語文時預設英文否則相同申請人資料會抓到多筆
        'Remove by Lydia 2020/09/16 共同條件+請款日期a1k02 (因為版面是以同一日期為一份)
        'If Me.Text2.Text = "FCT" Or Me.Text2.Text = "S" Then
        '    StrSQLa = "select cu05, cu88, cu89, cu90, nvl(pa85,'2') as Lang, a1k04 from acc1k0, patent, customer where substr(pa26, 1, 8) = cu01 and substr(pa26, 9, 1) = cu02 and pa01 = a1k13 and pa02 = a1k14 and pa03 = a1k15 and pa04 = a1k16 And a1k27='" & adoacc1k0.Fields("a1k27").Value & "' And pa26='" & adoacc1k0.Fields("cuno").Value & "' " & strSQL1 & strConPA & _
                            " union select cu05, cu88, cu89, cu90, nvl(tm53,'2') as Lang, a1k04 from acc1k0, trademark, customer where substr(tm23, 1, 8) = cu01 and substr(tm23, 9, 1) = cu02 and tm01 = a1k13 and tm02 = a1k14 and tm03 = a1k15 and tm04 = a1k16 And a1k27='" & adoacc1k0.Fields("a1k27").Value & "' And tm23='" & adoacc1k0.Fields("cuno").Value & "' " & strSQL1 & strConTM & _
                            " union select cu05, cu88, cu89, cu90, '2' as Lang, a1k04 from acc1k0, lawcase, customer where substr(lc11, 1, 8) = cu01 and substr(lc11, 9, 1) = cu02 and lc01 = a1k13 and lc02 = a1k14 and lc03 = a1k15 and lc04 = a1k16 And a1k27='" & adoacc1k0.Fields("a1k27").Value & "' And lc11='" & adoacc1k0.Fields("cuno").Value & "' " & strSQL1 & strConLC & _
                            " union select cu05, cu88, cu89, cu90, NVL(sp34,'2') as Lang, a1k04 from acc1k0, servicepractice, customer where substr(sp08, 1, 8) = cu01 and substr(sp08, 9, 1) = cu02 and sp01 = a1k13 and sp02 = a1k14 and sp03 = a1k15 and sp04 = a1k16 And a1k27='" & adoacc1k0.Fields("a1k27").Value & "' And sp08='" & adoacc1k0.Fields("cuno").Value & "' " & strSQL1 & strConSP
        'Else
            StrSQLa = "select cu05, cu88, cu89, cu90, NVL(pa85,'2') as Lang, a1k04 from acc1k0, patent, customer where substr(pa26, 1, 8) = cu01 and substr(pa26, 9, 1) = cu02 and pa01 = a1k13 and pa02 = a1k14 and pa03 = a1k15 and pa04 = a1k16  " & strSQL1 & strConPA & _
                            " union select cu05, cu88, cu89, cu90, nvl(tm53,'2') as Lang, a1k04 from acc1k0, trademark, customer where substr(tm23, 1, 8) = cu01 and substr(tm23, 9, 1) = cu02 and tm01 = a1k13 and tm02 = a1k14 and tm03 = a1k15 and tm04 = a1k16 " & strSQL1 & strConTM & _
                            " union select cu05, cu88, cu89, cu90, '2' as Lang, a1k04 from acc1k0, lawcase, customer where substr(lc11, 1, 8) = cu01 and substr(lc11, 9, 1) = cu02 and lc01 = a1k13 and lc02 = a1k14 and lc03 = a1k15 and lc04 = a1k16 " & strSQL1 & strConLC & _
                            " union select cu05, cu88, cu89, cu90, nvl(sp34,'2') as Lang, a1k04 from acc1k0, servicepractice, customer where substr(sp08, 1, 8) = cu01 and substr(sp08, 9, 1) = cu02 and sp01 = a1k13 and sp02 = a1k14 and sp03 = a1k15 and sp04 = a1k16 " & strSQL1 & strConSP
        'End If
        'end 2020/09/16
        
         rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
        
        'Modified by Morgan 2022/9/1
        'strCustName(0) = "": strCustName(1) = "": strCustName(2) = "": strCustName(3) = "": strCustName(4) = ""
        Erase strCustName
        'end 2022/8/1
        
        ii = 0
        ReDim strCustName(ii) 'Added by Morgan 2022/9/1
        While Not rsA.EOF
            If IsNull(rsA.Fields("Lang").Value) = False Then
               strLanguage = rsA.Fields("Lang").Value
            Else
               strLanguage = "2"
            End If
            If IsNull(rsA.Fields("cu05").Value) = False Then
                If "" & rsA("a1k04").Value = "Y" Then
                     intCounter = intCounter + 1 'Added by Morgan 2017/2/14
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, rsA.Fields("cu05").Value
                     Else
                     'END 2014/9/17
                        m_Device.CurrentX = 500 + intInit
                        m_Device.CurrentY = 1500 + intCounter * 300
                        m_Device.Print rsA.Fields("cu05").Value
                     End If 'Added by Morgan 2014/9/17
                End If
                strCustName(ii) = strCustName(ii) & rsA.Fields("cu05").Value
            End If
            If IsNull(rsA.Fields("cu88").Value) = False Then
                If "" & rsA("a1k04").Value = "Y" Then
                     intCounter = intCounter + 1
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, rsA.Fields("cu88").Value
                     Else
                     'END 2014/9/17
                        m_Device.CurrentX = 500 + intInit
                        m_Device.CurrentY = 1500 + intCounter * 300
                        m_Device.Print rsA.Fields("cu88").Value
                     End If 'Added by Morgan 2014/9/17
                End If
                strCustName(ii) = strCustName(ii) & " " & rsA.Fields("cu88").Value
            End If
            If IsNull(rsA.Fields("cu89").Value) = False Then
                If "" & rsA("a1k04").Value = "Y" Then
                     intCounter = intCounter + 1
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, rsA.Fields("cu89").Value
                     Else
                     'END 2014/9/17
                        m_Device.CurrentX = 500 + intInit
                        m_Device.CurrentY = 1500 + intCounter * 300
                        m_Device.Print rsA.Fields("cu89").Value
                     End If 'Added by Morgan 2014/9/17
                End If
                strCustName(ii) = strCustName(ii) & " " & rsA.Fields("cu89").Value
            End If
            If IsNull(rsA.Fields("cu90").Value) = False Then
                If "" & rsA("a1k04").Value = "Y" Then
                     intCounter = intCounter + 1
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, rsA.Fields("cu90").Value
                     Else
                     'END 2014/9/17
                        m_Device.CurrentX = 500 + intInit
                        m_Device.CurrentY = 1500 + intCounter * 300
                        m_Device.Print rsA.Fields("cu90").Value
                     End If 'Added by Morgan 2014/9/17
                End If
                strCustName(ii) = strCustName(ii) & " " & rsA.Fields("cu90").Value
            End If
            If IsNull(rsA.Fields("a1k04").Value) Then
               strYes = MsgText(601)
            Else
               strYes = rsA.Fields("a1k04").Value
            End If
            ii = ii + 1
            ReDim Preserve strCustName(ii) 'Added by Morgan 2022/9/1
            rsA.MoveNext
        Wend
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
'      Else
'         strLanguage = "2"
'         strCustName = ""
'      End If
'   Else
'      strLanguage = "2"
'      strCustName = ""
'   End If
   
   'Modified by Morgan 2017/2/14
   'intCounter = intCounter + 2
   intCounter = intCounter + 1
   If intCounter < 4 Then intCounter = 4
   'end 2017/2/14
   
   If strYes = MsgText(602) Then
      'Added by Morgan 2014/9/17
      If m_bolWord Then
         SetWordArray m_Item, intCounter, 1, "C/O"
      Else
      'END 2014/9/17
         m_Device.CurrentX = 0 + intInit
         m_Device.CurrentY = 1500 + intCounter * 300
         m_Device.Print "C/O"
      End If 'Added by Morgan 2014/9/17
   End If
    'Add By Cheng 2003/03/28
    '設定語言
    If strLanguage = "" Then strLanguage = "2"
   Select Case strLanguage
      Case "2"
         If IsNull(adoacc1k0.Fields("fa05").Value) = False Then
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa05").Value
            Else
            'END 2014/9/17
               m_Device.CurrentX = 500 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
               m_Device.Print adoacc1k0.Fields("fa05").Value
            End If 'Added by Morgan 2014/9/17
         End If
         If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
            intCounter = intCounter + 1
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa63").Value
            Else
            'END 2014/9/17
               m_Device.CurrentX = 500 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
               m_Device.Print adoacc1k0.Fields("fa63").Value
            End If 'Added by Morgan 2014/9/17
         End If
         If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
            intCounter = intCounter + 1
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa64").Value
            Else
            'END 2014/9/17
               m_Device.CurrentX = 500 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
               m_Device.Print adoacc1k0.Fields("fa64").Value
            End If 'Added by Morgan 2014/9/17
         End If
         If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
            intCounter = intCounter + 1
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa65").Value
            Else
            'END 2014/9/17
               m_Device.CurrentX = 500 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
               m_Device.Print adoacc1k0.Fields("fa65").Value
            End If 'Added by Morgan 2014/9/17
         End If
         intCounter = intCounter + 1
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If IsNull(adoacc1k0.Fields("fa18").Value) = False Then
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa18").Value
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 500 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print adoacc1k0.Fields("fa18").Value
               End If 'Added by Morgan 2014/9/17
            End If
         Else
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa32").Value
            Else
            'END 2014/9/17
               m_Device.CurrentX = 500 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
               m_Device.Print adoacc1k0.Fields("fa32").Value
            End If 'Added by Morgan 2014/9/17
         End If
        '選擇請款日期
         If Option1.Value Then
            If MaskEdBox1.Text <> MsgText(29) And MaskEdBox2.Text <> MsgText(29) Then
               strExc(1) = IIf(Month(AFDate(CADate(FCDate(MaskEdBox1.Text)))) = 5, Format(AFDate(CADate(FCDate(MaskEdBox1.Text))), "mmm d, yyyy"), Format(AFDate(CADate(FCDate(MaskEdBox1.Text))), "mmm. d, yyyy")) & " - " & _
                           IIf(Month(AFDate(CADate(FCDate(MaskEdBox2.Text)))) = 5, Format(AFDate(CADate(FCDate(MaskEdBox2.Text))), "mmm d, yyyy"), Format(AFDate(CADate(FCDate(MaskEdBox2.Text))), "mmm. d, yyyy"))
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, 3, 3, strExc(1)
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 6500 + intInit
                  'Modify by Morgan 2011/7/27 新信紙
                  'm_Device.CurrentY = 3000
                  m_Device.CurrentY = 2400
                  m_Device.Print strExc(1)
               End If 'Added by Morgan 2014/9/17
            End If
         '選擇請款編號
         Else
            If "" & adoacc1k0.Fields("a1k02").Value <> "" Then
               '若為五月份
               If Month(AFDate(CADate(adoacc1k0.Fields("a1k02").Value))) = 5 Then
                 strExc(1) = Format(AFDate(CADate(adoacc1k0.Fields("a1k02").Value)), "mmm d, yyyy")
               '若非為五月份
               Else
                   strExc(1) = Format(AFDate(CADate(adoacc1k0.Fields("a1k02").Value)), "mmm. d, yyyy")
               End If
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, 3, 3, strExc(1)
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 6500 + intInit
                  'Modify by Morgan 2011/7/27 新信紙
                  'm_Device.CurrentY = 3000
                  m_Device.CurrentY = 2400
                  m_Device.Print strExc(1)
               End If 'Added by Morgan 2014/9/17
            End If
         End If
                  
         'Added by Morgan 2021/2/22 --陳金蓮
         If Text1 = "Y52543000" Then
            'Modified by Morgan 2023/2/21
            'strExc(1) = "Purchase Order Number: 4300004798"
            strExc(1) = "Vendor Number: 20718822"
            'end 2023/2/21
            If m_bolWord Then
               SetWordArray m_Item, 4, 3, strExc(1)
            Else
               m_Device.CurrentX = 6500 + intInit
               m_Device.CurrentY = 2400 + 300
               m_Device.Print strExc(1)
            End If
         End If
         'end 2019/9/19
         
         'Added by Morgan 2023/4/10
         strExc(1) = PUB_GetACCNO(adoacc1k0.Fields("a1k27").Value, Text2)
         If strExc(1) <> "" Then
            If m_bolWord Then
               SetWordArray m_Item, 7, 3, strExc(1)
            Else
               m_Device.CurrentX = 6500 + intInit
               m_Device.CurrentY = 2400 + 300 * 4
               m_Device.Print strExc(1)
            End If
         End If
         'end 2023/4/10
         
         'Added by Morgan 2022/10/31--陳金蓮
'         If adoacc1k0.Fields("a1k03") = "Y52622000" And (Text2 = "FCT" Or Text2 = "S") Then
'            strExc(1) = "Keltie fee earner: " & adoacc1k0.Fields("Cno").Value
'            If m_bolWord Then
'               SetWordArray m_Item, 5, 3, strExc(1)
'            Else
'               m_Device.CurrentX = 6500 + intInit
'               m_Device.CurrentY = 2400 + 2 * 300
'               m_Device.Print strExc(1)
'            End If
'         End If
         'end 2022/10/31
                  
'         intCounter = intCounter + 1
        'Modify By Cheng 2003/07/09
'         If IsNull(adoacc1k0.Fields("fa33").Value) Then
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If IsNull(adoacc1k0.Fields("fa19").Value) = False Then
               intCounter = intCounter + 1
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa19").Value
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 500 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print adoacc1k0.Fields("fa19").Value
               End If 'Added by Morgan 2014/9/17
            End If
         Else
            If IsNull(adoacc1k0.Fields("fa33").Value) = False Then
               intCounter = intCounter + 1
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa33").Value
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 500 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print adoacc1k0.Fields("fa33").Value
               End If 'Added by Morgan 2014/9/17
            End If
         End If
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If IsNull(adoacc1k0.Fields("fa20").Value) = False Then
               intCounter = intCounter + 1
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa20").Value
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 500 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print adoacc1k0.Fields("fa20").Value
               End If 'Added by Morgan 2014/9/17
            End If
         Else
            If IsNull(adoacc1k0.Fields("fa34").Value) = False Then
               intCounter = intCounter + 1
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa34").Value
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 500 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print adoacc1k0.Fields("fa34").Value
               End If 'Added by Morgan 2014/9/17
            End If
         End If
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If IsNull(adoacc1k0.Fields("fa21").Value) = False Then
               intCounter = intCounter + 1
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa21").Value
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 500 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print adoacc1k0.Fields("fa21").Value
               End If 'Added by Morgan 2014/9/17
            End If
         Else
            If IsNull(adoacc1k0.Fields("fa35").Value) = False Then
               intCounter = intCounter + 1
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa35").Value
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 500 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print adoacc1k0.Fields("fa35").Value
               End If 'Added by Morgan 2014/9/17
            End If
         End If
'         intCounter = intCounter + 1
        'Modify By Cheng 2003/07/09
'         If IsNull(adoacc1k0.Fields("fa36").Value) Then
         If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If IsNull(adoacc1k0.Fields("fa22").Value) = False Then
               intCounter = intCounter + 1
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa22").Value
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 500 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print adoacc1k0.Fields("fa22").Value
               End If 'Added by Morgan 2014/9/17
            End If
            'Add by Morgan 2011/5/25
            '英文地址6
            If IsNull(adoacc1k0.Fields("fa70").Value) = False Then
               intCounter = intCounter + 1
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa70").Value
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 500 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print adoacc1k0.Fields("fa70").Value
               End If 'Added by Morgan 2014/9/17
            End If
            
         Else
            If IsNull(adoacc1k0.Fields("fa36").Value) = False Then
               intCounter = intCounter + 1
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa36").Value
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 500 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print adoacc1k0.Fields("fa36").Value
               End If 'Added by Morgan 2014/9/17
            End If
         End If
            '若請款單格式為加總明細格式1
            If m_strDetailKind = "1" Then
                'Add By Cheng 2003/04/29
                'Modified by Lydia 2020/09/16 共同條件+請款日期a1k02 (因為版面是以同一日期為一份)
                'StrSQLa = "select pa77 as Yno, pa01||'-'||pa02||'-'||pa03||'-'||pa04 as Caseno, pa11 as Ano, a1k01 as Dno from acc1k0, patent where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and pa26='" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & " group by pa77, pa01||'-'||pa02||'-'||pa03||'-'||pa04, pa11, a1k01 "
                'StrSQLa = StrSQLa & " union select tm45 as Yno, tm01||'-'||tm02||'-'||tm03||'-'||tm04 as Caseno, nvl(tm15,tm12) as Ano, a1k01 as Dno from acc1k0, trademark where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and tm23 = '" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & " group by tm45, tm01||'-'||tm02||'-'||tm03||'-'||tm04, nvl(tm15,tm12), a1k01  "
                'StrSQLa = StrSQLa & " union select lc23 as Yno, lc01||'-'||lc02||'-'||lc03||'-'||lc04 as Caseno, '' as Ano, a1k01 as Dno from acc1k0, lawcase where a1k13 = lc01 and a1k14 = lc02 and a1k15 = lc03 and a1k16 = lc04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and lc11 = '" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & " group by lc23, lc01||'-'||lc02||'-'||lc03||'-'||lc04, a1k01 "
                'StrSQLa = StrSQLa & " union select sp27 as Yno, sp01||'-'||sp02||'-'||sp03||'-'||sp04 as Caseno, sp11 as Ano, a1k01 as Dno from acc1k0, servicepractice where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and sp08 = '" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & " group by sp27, sp01||'-'||sp02||'-'||sp03||'-'||sp04, sp11, a1k01 "
                StrSQLa = "select pa77 as Yno, pa01||'-'||pa02||'-'||pa03||'-'||pa04 as Caseno, pa11 as Ano, a1k01 as Dno from acc1k0, patent where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 " & strSql & strConPA & " group by pa77, pa01||'-'||pa02||'-'||pa03||'-'||pa04, pa11, a1k01 "
                StrSQLa = StrSQLa & " union select tm45 as Yno, tm01||'-'||tm02||'-'||tm03||'-'||tm04 as Caseno, nvl(tm15,tm12) as Ano, a1k01 as Dno from acc1k0, trademark where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 " & strSql & strConTM & " group by tm45, tm01||'-'||tm02||'-'||tm03||'-'||tm04, nvl(tm15,tm12), a1k01  "
                StrSQLa = StrSQLa & " union select lc23 as Yno, lc01||'-'||lc02||'-'||lc03||'-'||lc04 as Caseno, '' as Ano, a1k01 as Dno from acc1k0, lawcase where a1k13 = lc01 and a1k14 = lc02 and a1k15 = lc03 and a1k16 = lc04 " & strSql & strConLC & " group by lc23, lc01||'-'||lc02||'-'||lc03||'-'||lc04, a1k01 "
                StrSQLa = StrSQLa & " union select sp27 as Yno, sp01||'-'||sp02||'-'||sp03||'-'||sp04 as Caseno, sp11 as Ano, a1k01 as Dno from acc1k0, servicepractice where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 " & strSql & strConSP & " group by sp27, sp01||'-'||sp02||'-'||sp03||'-'||sp04, sp11, a1k01 "
                'end 2020/09/16
                StrSQLa = StrSQLa & " order by 4 "
                rsA.CursorLocation = adUseClient
                rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                If rsA.RecordCount > 0 Then
                    rsA.MoveFirst
                    m_strDBNOFrom = "" & rsA.Fields(3).Value
                    rsA.MoveLast
                    m_strDBNOTo = "" & rsA.Fields(3).Value
                Else
                    m_strDBNOFrom = ""
                    m_strDBNOTo = ""
                End If
                If rsA.State <> adStateClosed Then rsA.Close
                Set rsA = Nothing
                intCounter = intCounter + 1
                
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, strTitle, 2
                  SetWordArray m_Item, intCounter, 3, "No." & m_strDBNOFrom & " - " & Right(m_strDBNOTo, 3), 2
                  m_DocName = m_strDBNOFrom & "-" & Right(m_strDBNOTo, 3)
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 4000 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  'Modify by Morgan 2009/5/21 迄號改印後三碼
                  m_Device.Print strTitle & "                        No." & m_strDBNOFrom & " - " & Right(m_strDBNOTo, 3)
               End If 'Added by Morgan 2014/9/17
                
                'Add By Cheng 2003/04/29
                If Me.Text2.Text = "FCP" Or Me.Text2.Text = "FG" Then
                    intCounter = intCounter + 1
                End If
                '若系統類別為FCT
                If "" & adoacc1k0.Fields("a1k13").Value = "FCT" Or "" & adoacc1k0.Fields("a1k13").Value = "S" Then
                   rsA.CursorLocation = adUseClient
                   StrSQLa = " And a1k13='" & Me.Text2.Text & "' "
                   '選擇請款日期
                   If Me.Option1 Then
                        StrSQLa = StrSQLa & " And a1k02>=" & Val(Replace(Me.MaskEdBox1.Text, "/", "")) & " "
                        StrSQLa = StrSQLa & " And a1k02<=" & Val(Replace(Me.MaskEdBox2.Text, "/", "")) & " "
                   '選擇請款編號
                   Else
                        StrSQLa = StrSQLa & " And a1k01>='" & Me.Text3.Text & "' "
                        StrSQLa = StrSQLa & " And a1k01<='" & Me.Text4.Text & "' "
                   End If
                   StrSQLa = StrSQLa & " And a1k28='" & Me.Text1.Text & "' "
                   rsA.Open "select cp10, cp01 from caseprogress, acc1k0 where cp60 = a1k01 and ((cp10 >= '101' and cp10 <= '105') or cp10='125') " & StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
                   If rsA.RecordCount <> 0 Then
                      Select Case rsA.Fields("cp01").Value
                         Case "CFT", "FCT", "T"
                            If "" & rsA.Fields("cp10").Value = "101" Then
                               strProperty = "New "
                            Else
                               strProperty = ""
                            End If
                        'Add By Cheng 2003/04/14
                         Case "FCP", "FG"
                            strProperty = ""
                         Case Else
                            strProperty = "New "
                      End Select
                   Else
                      strProperty = ""
                   End If
                  rsA.Close
                  intCounter = intCounter + 1
                  intCounter = intCounter + 1
                  strExc(1) = "Re: " & GetNationName(strLanguage) & " " & strProperty & "Trademark " & IIf(Me.Text2.Text = "S", "Searches ", "Matters ")
                  'Added by Morgan 2014/9/17
                  If m_bolWord Then
                     SetWordArray m_Item, intCounter, 1, "Re: ", 2
                     SetWordArray m_Item, intCounter, 2, GetNationName(strLanguage) & " " & strProperty & "Trademark " & IIf(Me.Text2.Text = "S", "Searches ", "Matters "), 2
                  Else
                  'END 2014/9/17
                     m_Device.Line (0 + intInit, 1500 + intCounter * 300 - 200)-(10200 + intInit, 1500 + intCounter * 300 - 200)
                     m_Device.CurrentX = 0 + intInit
                     m_Device.CurrentY = 1500 + intCounter * 300 + intTop
                     m_Device.Print strExc(1)
                  End If 'Added by Morgan 2014/9/17
                  
                    If "" & strCustName(0) <> "" Then
                        intCounter = intCounter + 1
                        'Added by Morgan 2014/9/17
                        If m_bolWord Then
                           SetWordArray m_Item, intCounter, 2, "In the name of " & strCustName(0), 2
                        Else
                        'END 2014/9/17
                           m_Device.CurrentX = 400 + intInit
                           m_Device.CurrentY = 1500 + intCounter * 300 + intTop
                           If CountLength(strCustName(0)) <= 80 Then
                              m_Device.Print "In the name of " & strCustName(0)
                           Else
                              PrintDropLine "" & strCustName(0), "In the name of ", intCounter, 80
                           End If
                        End If 'Added by Morgan 2014/9/17
                             
                        '其他的申請人
                        'Modified by Morgan 2022/9/1
                        'For ii = 1 To 4
                        For ii = 1 To UBound(strCustName)
                        'end 2022/9/1
                          If strCustName(ii) <> "" Then
                              intCounter = intCounter + 1
                              'Added by Morgan 2014/9/17
                              If m_bolWord Then
                                 SetWordArray m_Item, intCounter, 3, strCustName(ii), 2
                              Else
                              'END 2014/9/17
                                 m_Device.CurrentX = 400 + intInit + m_Device.TextWidth("In the name of ")
                                 m_Device.CurrentY = 1500 + intCounter * 300 + intTop
                                 m_Device.Print strCustName(ii)
                              End If 'Added by Morgan 2014/9/17
                          End If
                        Next ii
                    End If
                    
                     '多空一行
                     intCounter = intCounter + 2
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 1, "Mark", 2
                     Else
                     'END 2014/9/17
                        m_Device.CurrentX = 0 + intInit
                        m_Device.CurrentY = 1500 + intCounter * 300
                        m_Device.Print "Mark"
                     End If 'Added by Morgan 2014/9/17
                     
                     intCounter = intCounter + 1
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, "Reg./Filing No.", 3
                     Else
                     'END 2014/9/17
                        m_Device.CurrentX = 0 + intInit
                        m_Device.CurrentY = 1500 + intCounter * 300
                        m_Device.Print "Reg./Filing No."
                     End If 'Added by Morgan 2014/9/17
                     
                '若系統類別為其他
                Else
                     intCounter = intCounter + 1
                     'Added by Morgan 2014/9/17
                     If m_bolWord Then
                        SetWordArray m_Item, intCounter, 2, "Filing No.", 3
                     Else
                     'END 2014/9/17
                        m_Device.CurrentX = 0 + intInit
                        m_Device.CurrentY = 1500 + intCounter * 300
                        m_Device.Print "Filing No."
                     End If 'Added by Morgan 2014/9/17
                     
                End If
                
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 3, "Debit No.", 3
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 2500 + intInit - 500 + 220
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print "Debit No."
               End If 'Added by Morgan 2014/9/17
               
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 4, "Our Ref", 3
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 4500 + intInit - 1000 + 1020
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print "Our Ref"
               End If 'Added by Morgan 2014/9/17
               
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 5, "Your Ref", 3
                  SetWordArray m_Item, intCounter, 0, "", 3, , , True
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 6500 + intInit - 500 + 790
                  m_Device.CurrentY = 1500 + intCounter * 300
                  m_Device.Print "Your Ref"
               End If 'Added by Morgan 2014/9/17
               
               intCounter = intCounter + 2
               If Not m_bolWord Then 'Added by Morgan 2014/9/17
                  m_Device.Line (0 + intInit, 1500 + intCounter * 300 - 200)-(10200 + intInit, 1500 + intCounter * 300 - 200)
               End If
                
            '若請款單格式為加總明細格式2
            ElseIf m_strDetailKind = "2" Then
                'Add By Cheng 2003/06/06
                'Modified by Lydia 2020/09/16 共同條件+請款日期a1k02 (因為版面是以同一日期為一份)
                'StrSQLa = "select pa77 as Yno, pa01||'-'||pa02||'-'||pa03||'-'||pa04 as Caseno, pa11 as Ano, a1k01 as Dno from acc1k0, patent where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and pa26='" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & " group by pa77, pa01||'-'||pa02||'-'||pa03||'-'||pa04, pa11, a1k01 "
                'StrSQLa = StrSQLa & " union select tm45 as Yno, tm01||'-'||tm02||'-'||tm03||'-'||tm04 as Caseno, nvl(tm15,tm12) as Ano, a1k01 as Dno from acc1k0, trademark where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and tm23 = '" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & " group by tm45, tm01||'-'||tm02||'-'||tm03||'-'||tm04, nvl(tm15,tm12), a1k01  "
                'StrSQLa = StrSQLa & " union select lc23 as Yno, lc01||'-'||lc02||'-'||lc03||'-'||lc04 as Caseno, '' as Ano, a1k01 as Dno from acc1k0, lawcase where a1k13 = lc01 and a1k14 = lc02 and a1k15 = lc03 and a1k16 = lc04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and lc11 = '" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & " group by lc23, lc01||'-'||lc02||'-'||lc03||'-'||lc04, a1k01 "
                'StrSQLa = StrSQLa & " union select sp27 as Yno, sp01||'-'||sp02||'-'||sp03||'-'||sp04 as Caseno, sp11 as Ano, a1k01 as Dno from acc1k0, servicepractice where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and a1k27 = '" & adoacc1k0.Fields("a1k27").Value & "' and a1k13 = '" & adoacc1k0.Fields("a1k13").Value & "' " & IIf(IsNull(adoacc1k0.Fields("cuno").Value), "", " and sp08 = '" & adoacc1k0.Fields("cuno").Value & "' ") & " and (a1k29 <> 'Y' or a1k29 is null) and a1k12 is null" & strSql & " group by sp27, sp01||'-'||sp02||'-'||sp03||'-'||sp04, sp11, a1k01 "
                StrSQLa = "select pa77 as Yno, pa01||'-'||pa02||'-'||pa03||'-'||pa04 as Caseno, pa11 as Ano, a1k01 as Dno from acc1k0, patent where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 " & strSql & strConPA & " group by pa77, pa01||'-'||pa02||'-'||pa03||'-'||pa04, pa11, a1k01 "
                StrSQLa = StrSQLa & " union select tm45 as Yno, tm01||'-'||tm02||'-'||tm03||'-'||tm04 as Caseno, nvl(tm15,tm12) as Ano, a1k01 as Dno from acc1k0, trademark where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 " & strSql & strConTM & " group by tm45, tm01||'-'||tm02||'-'||tm03||'-'||tm04, nvl(tm15,tm12), a1k01  "
                StrSQLa = StrSQLa & " union select lc23 as Yno, lc01||'-'||lc02||'-'||lc03||'-'||lc04 as Caseno, '' as Ano, a1k01 as Dno from acc1k0, lawcase where a1k13 = lc01 and a1k14 = lc02 and a1k15 = lc03 and a1k16 = lc04 " & strSql & strConLC & " group by lc23, lc01||'-'||lc02||'-'||lc03||'-'||lc04, a1k01 "
                StrSQLa = StrSQLa & " union select sp27 as Yno, sp01||'-'||sp02||'-'||sp03||'-'||sp04 as Caseno, sp11 as Ano, a1k01 as Dno from acc1k0, servicepractice where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 " & strSql & strConSP & " group by sp27, sp01||'-'||sp02||'-'||sp03||'-'||sp04, sp11, a1k01 "
                'end 2020/09/16
                StrSQLa = StrSQLa & " order by 4 "
                rsA.CursorLocation = adUseClient
                rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                If rsA.RecordCount > 0 Then
                    rsA.MoveFirst
                    m_strDBNOFrom = "" & rsA.Fields(3).Value
                    rsA.MoveLast
                    m_strDBNOTo = "" & rsA.Fields(3).Value
                Else
                    m_strDBNOFrom = ""
                    m_strDBNOTo = ""
                End If
                If rsA.State <> adStateClosed Then rsA.Close
                Set rsA = Nothing
               intCounter = intCounter + 1
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, intCounter, 2, strTitle, 2
                  SetWordArray m_Item, intCounter, 3, "No." & m_strDBNOFrom & " - " & Right(m_strDBNOTo, 3), 2
                  m_DocName = m_strDBNOFrom & "-" & Right(m_strDBNOTo, 3)
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 4000 + intInit
                  m_Device.CurrentY = 1500 + intCounter * 300
                  'Modify by Morgan 2009/5/21 迄號改印後三碼
                  m_Device.Print strTitle & "                       No." & m_strDBNOFrom & " - " & Right(m_strDBNOTo, 3)
                  m_Device.Line (0 + intInit, 1500 + intCounter * 300 + 300)-(10200 + intInit, 1500 + intCounter * 300 + 300)
               End If 'Added by Morgan 2014/9/17
               
               intCounter = intCounter + 2
            End If
      Case "3"
         If IsNull(adoacc1k0.Fields("fa06").Value) = False Then
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa06").Value & "　" & "御中"
            Else
            'END 2014/9/17
               m_Device.CurrentX = 500 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
               'Modify By Sindy 2010/3/26
               'm_Device.Print adoacc1k0.Fields("fa06").Value
               m_Device.Print adoacc1k0.Fields("fa06").Value & "　" & "御中"
               '2010/3/26 End
            End If 'Added by Morgan 2014/9/17
         End If
         intCounter = intCounter + 1
         If IsNull(adoacc1k0.Fields("fa23").Value) = False Then
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa23").Value
            Else
            'END 2014/9/17
               m_Device.CurrentX = 500 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
               m_Device.Print adoacc1k0.Fields("fa23").Value
            End If 'Added by Morgan 2014/9/17
         End If
         If Option1.Value Then
            If MaskEdBox1.Text <> MsgText(29) And MaskEdBox2.Text <> MsgText(29) Then
               strExc(1) = IIf(Month(AFDate(CADate(FCDate(MaskEdBox1.Text)))) = 5, Format(AFDate(CADate(FCDate(MaskEdBox1.Text))), "mmm d, yyyy"), Format(AFDate(CADate(FCDate(MaskEdBox1.Text))), "mmm. d, yyyy")) & " - " & _
                            IIf(Month(AFDate(CADate(FCDate(MaskEdBox2.Text)))) = 5, Format(AFDate(CADate(FCDate(MaskEdBox2.Text))), "mmm d, yyyy"), Format(AFDate(CADate(FCDate(MaskEdBox2.Text))), "mmm. d, yyyy"))
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, 3, 3, strExc(1)
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 6500 + intInit
                  'Modify by Morgan 2011/7/27 新信紙
                  'm_Device.CurrentY = 3000
                  m_Device.CurrentY = 2400
                  m_Device.Print strExc(1)
               End If 'Added by Morgan 2014/9/17
            End If
         Else
            If "" & adoacc1k0.Fields("a1k02").Value <> "" Then
                '若為五月份
                If Month(AFDate(CADate(adoacc1k0.Fields("a1k02").Value))) = 5 Then
                    strExc(1) = Format(AFDate(CADate(adoacc1k0.Fields("a1k02").Value)), "mmm d, yyyy")
                '若非為五月份
                Else
                    strExc(1) = Format(AFDate(CADate(adoacc1k0.Fields("a1k02").Value)), "mmm. d, yyyy")
                End If
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, 3, 3, strExc(1)
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 6500 + intInit
                  'Modify by Morgan 2011/7/27 新信紙
                  'm_Device.CurrentY = 3000
                  m_Device.CurrentY = 2400
                  m_Device.Print strExc(1)
               End If 'Added by Morgan 2014/9/17
            End If
         End If
         intCounter = intCounter + 1
         If Option2.Value Then
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, strTitle, 2
               SetWordArray m_Item, intCounter, 3, "No." & Text3 & " - " & Right(Text4, 3), 2
               m_DocName = Text3 & "-" & Right(Text4, 3)
            Else
            'END 2014/9/17
               m_Device.CurrentX = 4000 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
               m_Device.Print strTitle & "                       No." & Text3 & " - " & Right(Text4, 3)
               'Modify By Sindy 2010/3/26
               'm_Device.Line (0 + intInit, 1500 + intCounter * 300 + 350)-(10200 + intInit, 1500 + intCounter * 300 + 350)
               m_Device.Line (0 + intInit, 1500 + intCounter * 300 + 250)-(10200 + intInit, 1500 + intCounter * 300 + 250)
               '2010/3/26 End
            End If 'Added by Morgan 2014/9/17
         End If
         intCounter = intCounter + 1
      Case "1"
         If IsNull(adoacc1k0.Fields("fa04").Value) = False Then
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa04").Value
            Else
            'END 2014/9/17
               m_Device.CurrentX = 500 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
               m_Device.Print adoacc1k0.Fields("fa04").Value
            End If 'Added by Morgan 2014/9/17
         End If
         intCounter = intCounter + 1
         If IsNull(adoacc1k0.Fields("fa18").Value) = False Then
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, adoacc1k0.Fields("fa18").Value
            Else
            'END 2014/9/17
               m_Device.CurrentX = 500 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
               m_Device.Print adoacc1k0.Fields("fa18").Value
            End If 'Added by Morgan 2014/9/17
         End If
         If Option1.Value Then
            If MaskEdBox1.Text <> MsgText(29) And MaskEdBox2.Text <> MsgText(29) Then
               strExc(1) = IIf(Month(AFDate(CADate(FCDate(MaskEdBox1.Text)))) = 5, Format(AFDate(CADate(FCDate(MaskEdBox1.Text))), "mmm d, yyyy"), Format(AFDate(CADate(FCDate(MaskEdBox1.Text))), "mmm. d, yyyy")) & " - " & _
                            IIf(Month(AFDate(CADate(FCDate(MaskEdBox2.Text)))) = 5, Format(AFDate(CADate(FCDate(MaskEdBox2.Text))), "mmm d, yyyy"), Format(AFDate(CADate(FCDate(MaskEdBox2.Text))), "mmm. d, yyyy"))
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, 3, 3, strExc(1)
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 6500 + intInit
                  'Modify by Morgan 2011/7/27 新信紙
                  'm_Device.CurrentY = 3000
                  m_Device.CurrentY = 2400
                  m_Device.Print strExc(1)
               End If 'Added by Morgan 2014/9/17
            End If
         Else
            If "" & adoacc1k0.Fields("a1k02").Value <> "" Then
               '若為五月份
               If Month(AFDate(CADate(adoacc1k0.Fields("a1k02").Value))) = 5 Then
                   strExc(1) = Format(AFDate(CADate(adoacc1k0.Fields("a1k02").Value)), "mmm d, yyyy")
               '若非為五月份
               Else
                   strExc(1) = Format(AFDate(CADate(adoacc1k0.Fields("a1k02").Value)), "mmm. d, yyyy")
               End If
               'Added by Morgan 2014/9/17
               If m_bolWord Then
                  SetWordArray m_Item, 3, 3, strExc(1)
               Else
               'END 2014/9/17
                  m_Device.CurrentX = 6500 + intInit
                  'Modify by Morgan 2011/7/27 新信紙
                  'm_Device.CurrentY = 3000
                  m_Device.CurrentY = 2400
                  m_Device.Print strExc(1)
               End If 'Added by Morgan 2014/9/17
            End If
         End If
         intCounter = intCounter + 1
         If Option2.Value Then
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               SetWordArray m_Item, intCounter, 2, strTitle, 2
               SetWordArray m_Item, intCounter, 3, "No." & Text3 & " - " & Right(Text4, 3), 2
               m_DocName = Text3 & "-" & Right(Text4, 3)
            Else
            'END 2014/9/17
               m_Device.CurrentX = 3000 + intInit
               m_Device.CurrentY = 1500 + intCounter * 300
               m_Device.Print strTitle & "                       No." & Text3 & " - " & Right(Text4, 3)
               'Modify By Sindy 2010/3/26
               'm_Device.Line (0 + intInit, 1500 + intCounter * 300 + 350)-(10200 + intInit, 1500 + intCounter * 300 + 350)
               m_Device.Line (0 + intInit, 1500 + intCounter * 300 + 250)-(10200 + intInit, 1500 + intCounter * 300 + 250)
               '2010/3/26 End
            End If 'Added by Morgan 2014/9/17
         End If
         intCounter = intCounter + 1
   End Select
End Sub

'*************************************************
' 合計位置
'
'*************************************************
Private Sub PrintSum()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim lngBoxX As Long, lngBoxY As Long 'Add by Morgan 2006/7/5
Dim DUsdRate As Double 'Added by Morgan 2014/9/23 外幣對美金匯率

   'Added by Morgan 2014/9/17
   If m_bolWord Then
      SetWordArray m_Item, intCounter, 2, "TOTAL", 4
   Else
   'END 2014/9/17
      m_Device.CurrentX = 5000 + 1200 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      m_Device.Print "TOTAL"
      m_Device.CurrentX = 7000 + 1200 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
   End If 'Added by Morgan 2014/9/17
   
   'Modified by Morgan 2013/1/2
'   Select Case strCurr
'      Case "U"
'         m_Device.Print "USD"
'        strAmount = PUB_ChgFormat("" & douAmount, True)
'      Case Else
'         m_Device.Print "NTD"
'        strAmount = PUB_ChgFormat("" & douAmount, True)
'   End Select
   
   'Added by Morgan 2014/9/19
   If m_bBatchRule Then douAmount = Trunc(douAmount)
   
   strAmount = PUB_ChgFormat("" & douAmount, True)
   Select Case m_iPrintCurrType
   Case 3, 4 '純外幣,外幣+美金合計
      'Added by Morgan 2014/9/17
      If m_bolWord Then
         SetWordArray m_Item, intCounter, 3, m_DNCurr, 4
      Else
      'END 2014/9/17
         m_Device.Print m_DNCurr
      End If 'Added by Morgan 2014/9/17
      
      '美金整數也要印 .00
      If m_DNCurr = "USD" Then
         strAmount = Format("" & douAmount, FDollar)
      End If
   Case Else
      'Added by Morgan 2014/9/17
      If m_bolWord Then
         SetWordArray m_Item, intCounter, 3, "NTD", 4
      Else
      'END 2014/9/17
         m_Device.Print "NTD"
      End If 'Added by Morgan 2014/9/17
   End Select
   'end 2013/1/2
   
   strAmount = Format(strAmount, FDollar) 'Added by Morgan 2013/1/3 金額都要印.00
   
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      SetWordArray m_Item, intCounter, 4, strAmount, 4
   Else
   'END 2014/9/17
      intLength = m_Device.TextWidth(strAmount)
      m_Device.CurrentX = 10200 - intLength + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      m_Device.Print strAmount
   End If 'Added by Morgan 2014/9/17
   
   intCounter = intCounter + 1
   strExc(1) = "( " & intRecords & " cases)"
   
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      SetWordArray m_Item, intCounter, 2, strExc(1), 4
   Else
   'END 2014/9/17
      m_Device.CurrentX = 5000 + 1200 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      m_Device.Print strExc(1)
   End If 'Added by Morgan 2014/9/17
   
   'Modified by Morgan 2013/1/2
'   Select Case strCurr
'      Case "N"
'      Case "U"
'      Case Else
   Select Case m_iPrintCurrType
   Case 2, 4 '2.台幣+外幣合計,4.外幣+美金合計
   'end 2013/1/2
      
      'Added by Morgan 2014/9/17
      If m_iPrintCurrType = "2" Then
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 3, m_DNCurr, 4
         Else
      'END 2014/9/23
      
            m_Device.CurrentX = 7000 + 1200 + intInit
            m_Device.CurrentY = 1500 + intCounter * 300
            m_Device.Print m_DNCurr
            
         'Added by Morgan 2014/9/23
         End If
         If m_bBatchRule Then
            strAmount = Trunc(douAmount / m_strOriExcRate)
            'Added by Lydia 2015/05/07
             Call UpdateBatchTrunc(Left(m_strA1K01, Len(m_strA1K01) - 1), Val(strAmount))
         Else
         'end 2014/9/23
         
            StrSQLa = "Select Sum(A1K08) From ACC1K0 Where A1K01 In (" & Left(m_strA1K01, Len(m_strA1K01) - 1) & ") "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               strAmount = PUB_ChgFormat(Val("" & rsA.Fields(0).Value), True)
            Else
               strAmount = "0"
            End If
         
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         
      'Added by Morgan 2014/9/23
         End If
      Else
         DUsdRate = PUB_GetDNRate(m_strA1K02, m_DNCurr)
         strAmount = Trunc(douAmount * DUsdRate)
         
         If m_bolWord Then
            SetWordArray m_Item, intCounter, 3, "USD", 4
         Else
            m_Device.CurrentX = 7000 + 1200 + intInit
            m_Device.CurrentY = 1500 + intCounter * 300
            m_Device.Print "USD"
         End If
         
      End If
      'end 2014/9/23
      
      strAmount = Format(strAmount, FDollar) 'Added by Morgan 2013/1/3 金額都要印.00
      'Added by Morgan 2014/9/17
      If m_bolWord Then
         SetWordArray m_Item, intCounter, 4, strAmount, 4
      Else
      'END 2014/9/17
         intLength = m_Device.TextWidth(strAmount)
         m_Device.CurrentX = 10200 - intLength + intInit
         m_Device.CurrentY = 1500 + intCounter * 300
         m_Device.Print strAmount
      End If 'Added by Morgan 2014/9/17
      intCounter = intCounter + 1
      
   End Select
   
   strExc(1) = String(17, "v")
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      SetWordArray m_Item, intCounter, 3, strExc(1), 4
   Else
   'END 2014/9/17
      m_Device.CurrentX = 7000 + 1200 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      m_Device.Print strExc(1)
   End If 'Added by Morgan 2014/9/17
   
   intCounter = intCounter + 2
   If (intCounter + 8) >= 45 Then
        'Modify by Morgan 2011/7/27 新信紙有信尾要上移
        'intCounter = 45
        intCounter = 43
      'Added by Morgan 2014/9/17
      If m_bolWord Then
         SetWordArray m_Item, intCounter, 2, "P." & m_intPage, 5
      Else
      'END 2014/9/17
         m_Device.CurrentX = 4500 + intInit
         m_Device.CurrentY = 1500 + intCounter * 300
         m_Device.Print "P." & m_intPage
      End If 'Added by Morgan 2014/9/17
        
        '若選擇加總明細首頁
        If Me.Text6.Text = "2" Or Me.Text6.Text = "4" Then
            'Added by Morgan 2014/9/17
            If m_bolWord Then
               runWord
            Else
            'end 2014/9/17
               If m_bPrinter = True Then
                  'Modified by Morgan 2013/1/3
                  'm_Device.NewPage
                  MyNewPage
               Else
                  SetPic m_intPage
               End If
            End If
            m_intPage = m_intPage + 1
        End If
        intCounter = 18
   End If
   
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      SetWordArray m_Item, intCounter, 1, ReportSum(71001), 5
   Else
   'END 2014/9/17
      m_Device.CurrentX = 0 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      m_Device.Print ReportSum(71001)
   End If 'Added by Morgan 2014/9/17
   
   intCounter = intCounter + 1
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      SetWordArray m_Item, intCounter, 1, ReportSum(73001), 5
   Else
   'END 2014/9/17
      m_Device.CurrentX = 0 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      m_Device.Print ReportSum(73001)
   End If 'Added by Morgan 2014/9/17
   
   'Add by Morgan 2006/7/5 加印 "Wire Transfer Preferred" -- 蘇副總
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      SetWordArray m_Item, intCounter, 2, "Wire", 5
      SetWordArray m_Item, intCounter + 1, 2, "Transfer", 5
      SetWordArray m_Item, intCounter + 2, 2, "Preferred", 5
   Else
   'END 2014/9/17
      lngBoxX = 500 + 7500
      lngBoxY = 1500 + intCounter * 300
      m_Device.DrawWidth = 5
      m_Device.Line (lngBoxX, lngBoxY - 100)-(lngBoxX + 1200, lngBoxY + 900), , B
      m_Device.DrawWidth = 1
      m_Device.CurrentX = lngBoxX + 150
      m_Device.CurrentY = lngBoxY
      m_Device.Print "Wire"
      m_Device.CurrentX = lngBoxX + 150
      m_Device.CurrentY = lngBoxY + 300
      m_Device.Print "Transfer"
      m_Device.CurrentX = lngBoxX + 150
      m_Device.CurrentY = lngBoxY + 600
      m_Device.Print "Preferred"
   End If 'Added by Morgan 2014/9/17
   'end 2006/7/5
   
   
   intCounter = intCounter + 1
   
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      SetWordArray m_Item, intCounter, 1, ReportSum(85), 5
   Else
   'END 2014/9/17
      m_Device.CurrentX = 0 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      m_Device.Print ReportSum(85)
   End If 'Added by Morgan 2014/9/17
   
   intCounter = intCounter + 1
   
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      SetWordArray m_Item, intCounter, 1, ReportSum(74), 5
   Else
   'END 2014/9/17
      m_Device.CurrentX = 0 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      m_Device.Print ReportSum(74)
   End If 'Added by Morgan 2014/9/17
   
'Removed by Morgan 2023/3/7 2013/12/17 已取消此帳戶
'   intCounter = intCounter + 1
'   'Added by Morgan 2014/9/17
'   If m_bolWord Then
'      SetWordArray m_Item, intCounter, 1, ReportSum(129), 5
'   Else
'   'END 2014/9/17
'      m_Device.CurrentX = 0 + intInit
'      m_Device.CurrentY = 1500 + intCounter * 300
'      m_Device.Print ReportSum(129)
'   End If 'Added by Morgan 2014/9/17
'end 2023/3/7
   
   intCounter = intCounter + 1
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      SetWordArray m_Item, intCounter, 1, ReportSum(121), 5
   Else
   'END 2014/9/17
      m_Device.CurrentX = 0 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      m_Device.Print ReportSum(121)
   End If 'Added by Morgan 2014/9/17
   
   intCounter = intCounter + 1
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      'Modified by Morgan 2016/12/6
      'SetWordArray m_Item, intCounter, 1, ReportSum(75) & Format(Val("" & m_strOriExcRate), FDollar), 5
      If m_DNCurr = "USD" Then
         SetWordArray m_Item, intCounter, 1, ReportSum(75) & Format(Val("" & m_strOriExcRate), FDollar), 5
      Else
         SetWordArray m_Item, intCounter, 1, Replace(ReportSum(75), "USD", m_DNCurr) & Format(Val("" & m_strOriExcRate), FDollar), 5
      End If
      'end 2016/12/6
   Else
   'END 2014/9/17
      m_Device.CurrentX = 0 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      'Modified by Morgan 2016/12/6
      'm_Device.Print ReportSum(75)
      If m_DNCurr = "USD" Then
         m_Device.Print ReportSum(75)
      Else
         m_Device.Print Replace(ReportSum(75), "USD", m_DNCurr)
      End If
      'end 2016/12/6
      m_Device.CurrentX = 3150 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      m_Device.Print Format(Val("" & m_strOriExcRate), FDollar)
   End If 'Added by Morgan 2014/9/17
   
   intCounter = intCounter + 1
   'Added by Morgan 2014/9/17
   If m_bolWord Then
      SetWordArray m_Item, intCounter, 1, ReportSum(86001), 5, True
   Else
   'END 2014/9/17
      m_Device.FontBold = True
      m_Device.CurrentX = 0 + intInit
      m_Device.CurrentY = 1500 + intCounter * 300
      m_Device.Print ReportSum(86001)
      m_Device.FontBold = False
   End If 'Added by Morgan 2014/9/17
   
End Sub

'*************************************************
' 選項一
'
'*************************************************
Public Sub FormEnabled1()
   MaskEdBox1.Enabled = True
   MaskEdBox2.Enabled = True
   Text3.Enabled = False
   Text4.Enabled = False
   Me.Text6.Text = "1"
End Sub

'*************************************************
' 選項二
'
'*************************************************
Public Sub FormEnabled2()
   MaskEdBox1.Enabled = False
   MaskEdBox2.Enabled = False
   Text3.Enabled = True
   Text4.Enabled = True
   Me.Text6.Text = "2"
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
    FormCheck = False
    '系統類別
    If Text2 = MsgText(601) Then
        MsgBox "請輸入系統類別!!!", vbExclamation + vbOKOnly
        Me.Text2.SetFocus
        Exit Function
    End If
    '選擇請款日期
    If Me.Option1.Value = True Then
        If MaskEdBox1.Text = MsgText(29) Then
            MsgBox "請輸入請款日期!!!", vbExclamation + vbOKOnly
            Me.MaskEdBox1.SetFocus
            Exit Function
        End If
        If MaskEdBox2.Text = MsgText(29) Then
            MsgBox "請輸入請款日期!!!", vbExclamation + vbOKOnly
            Me.MaskEdBox2.SetFocus
            Exit Function
        End If
    End If
    '選擇請款編號
    If Me.Option2.Value Then
        If Text3 = MsgText(601) Then
            MsgBox "請輸入請款編號!!!", vbExclamation + vbOKOnly
            Me.Text3.SetFocus
           Exit Function
        End If
        If Text4 = MsgText(601) Then
            MsgBox "請輸入請款編號!!!", vbExclamation + vbOKOnly
            Me.Text4.SetFocus
           Exit Function
        End If
    End If
    '請款對象
    If Text1 = MsgText(601) Then
        MsgBox "請輸入請款對象!!!", vbExclamation + vbOKOnly
        Me.Text1.SetFocus
        Exit Function
    End If
    '列印方式
    If Text6 = MsgText(601) Then
        MsgBox "請輸入列印方式!!!", vbExclamation + vbOKOnly
        Me.Text6.SetFocus
        Exit Function
    End If
    FormCheck = True
End Function

Private Sub Text5_GotFocus()
    TextInverse Me.Text5
   CloseIme
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/03/17
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
    'Add By Cheng 2003/03/17
    If Me.Text5.Text <> "" Then
        Me.Text5.Text = Left(Me.Text5.Text & "000000000", 9)
        'Modify by Morgan 2010/6/29
        'Me.lbl2.Caption = GetCustomerEngName(Me.Text5.Text)
        Lbl2.Caption = GetCustomerEngName(Text5.Text, Text2.Text, strExc(1))
        If Text2 <> "" And strExc(1) <> "" Then
            txtCopy = strExc(1)
        End If
        'end 2010/6/29
        
        If Me.Lbl2.Caption = "" Then
            MsgBox "請款對象輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            Text5_GotFocus
        End If
    Else
        Me.Lbl2.Caption = ""
    End If
End Sub

Private Sub Text6_GotFocus()
    'Add By Cheng 2003/03/21
    TextInverse Me.Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/03/17
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    'Modify By Cheng 2003/04/09
'    Case 49, 50, 8
    'Added by Lydia 2015/04/14 +5
    Case 49, 50, 51, 52, 53, 8
        '無動作
    Case Else
        KeyAscii = 0
    End Select
End Sub

'Add By Cheng 2003/03/17
'取得代理人名稱
'Modify by Morgan 2010/6/29 +strSYS:系統別,strCopys:請款單份數
Private Function GetFagentEngName(strFA0102 As String, Optional strSys As String, Optional strCopys As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
StrSQLa = "Select Decode(FA05,NULL,Nvl(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65),decode('" & strSys & "','FCT',fa90,fa89) CPY From Fagent Where FA01='" & Mid(strFA0102, 1, 8) & "' And FA02='" & Mid(strFA0102, 9, 1) & "' "
StrSQLa = StrSQLa & " UNION Select Decode(CU05,NULL,Nvl(CU04,CU06),CU05||' '||CU88||' '||CU89||' '||CU90),decode('" & strSys & "','FCT',cu136,cu135) CPY From Customer Where CU01='" & Mid(strFA0102, 1, 8) & "' And CU02='" & Mid(strFA0102, 9, 1) & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetFagentEngName = "" & rsA.Fields(0).Value
   If strSys <> "" Then
      strCopys = "" & rsA.Fields(1).Value
   Else
      strCopys = ""
   End If
Else
    GetFagentEngName = ""
    strCopys = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add By Cheng 2003/03/17
'取得申請人名稱
'Modify by Morgan 2010/6/29 +strSYS:系統別,strCopys:請款單份數
Private Function GetCustomerEngName(strCU0102 As String, Optional strSys As String, Optional strCopys As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
StrSQLa = "Select Decode(CU05,NULL,Nvl(CU04,CU06),CU05||' '||CU88||' '||CU89||' '||CU90),decode('" & strSys & "','FCT',cu136,cu135) CPY From Customer Where CU01='" & Mid(strCU0102, 1, 8) & "' And CU02='" & Mid(strCU0102, 9, 1) & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCustomerEngName = "" & rsA.Fields(0).Value
   If strSys <> "" Then
      strCopys = "" & rsA.Fields(1).Value
   Else
      strCopys = ""
   End If
Else
    GetCustomerEngName = ""
    strCopys = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add By Cheng 2003/03/13
'計算字數
Private Function CountLength(strWord As String) As Double
Dim ii As Integer
Dim strChr As String
    
CountLength = 0
If strWord <> "" Then
    For ii = 1 To Len(strWord)
        strChr = Mid(strWord, ii, 1)
        If Asc(strChr) >= 65 And Asc(strChr) <= 90 Then
            CountLength = CountLength + 1.5
        ElseIf Asc(strChr) >= 128 Then
            CountLength = CountLength + 2
        Else
            CountLength = CountLength + 1
        End If
    Next ii
End If
End Function

'Add By Cheng 2003/03/13
'文字折行
Private Sub PrintDropLine(strWord As String, strLineTitle As String, intRow As Integer, intLineChrs As Integer)
Dim ii As Integer
Dim jj As Integer
Dim kk As Integer
Dim dblChrCnt As Double
Dim strArr
Dim strWordPrint As String

kk = 0
dblChrCnt = 0
strWordPrint = ""
If Trim(strWord) <> "" Then
    strArr = Split(RTrim(strWord), " ")
    For ii = LBound(strArr) To UBound(strArr)
        'Modify By Cheng 2004/04/22
        '若陣列為空字串的設定為空白
'        If strArr(iI) <> "" Then
            If strArr(ii) = "" Then strArr(ii) = " "
            For jj = 1 To Len(strArr(ii))
                If Asc(Mid(strArr(ii), jj, 1)) >= 65 And Asc(Mid(strArr(ii), jj, 1)) <= 90 Then
                    dblChrCnt = dblChrCnt + 1.5
                ElseIf (Asc(Mid(strArr(ii), jj, 1)) >= 128 Or Asc(Mid(strArr(ii), jj, 1)) < 0) Then
                    dblChrCnt = dblChrCnt + 2
                Else
                    dblChrCnt = dblChrCnt + 1
                End If
            Next jj
            If dblChrCnt + 1 > intLineChrs Then
                If kk = 0 Then
                    m_Device.CurrentX = 0 + intInit
                    m_Device.CurrentY = 1500 + intRow * 300
                    m_Device.Print strLineTitle & strWordPrint
                Else
                    m_Device.CurrentX = 0 + intInit + m_Device.TextWidth(strLineTitle)
                    m_Device.CurrentY = 1500 + intRow * 300
                    m_Device.Print strWordPrint
                End If
                kk = kk + 1
                intRow = intRow + 1
                dblChrCnt = 0
                For jj = 1 To Len(strArr(ii))
                    If Asc(Mid(strArr(ii), jj, 1)) >= 65 And Asc(Mid(strArr(ii), jj, 1)) <= 90 Then
                        dblChrCnt = dblChrCnt + 1.5
                    ElseIf (Asc(Mid(strArr(ii), jj, 1)) >= 128 Or Asc(Mid(strArr(ii), jj, 1)) < 0) Then
                        dblChrCnt = dblChrCnt + 2
                    Else
                        dblChrCnt = dblChrCnt + 1
                    End If
                Next jj
                strWordPrint = ""
                strWordPrint = strWordPrint & strArr(ii) & IIf(strArr(ii) = " ", "", " ")
                dblChrCnt = dblChrCnt + 1
            Else
                strWordPrint = strWordPrint & strArr(ii) & IIf(strArr(ii) = " ", "", " ")
                dblChrCnt = dblChrCnt + 1
            End If
'        End If
    Next ii
    If strWordPrint <> "" Then
        m_Device.CurrentX = 0 + intInit + m_Device.TextWidth(strLineTitle)
        m_Device.CurrentY = 1500 + intRow * 300
        m_Device.Print strWordPrint
        kk = kk + 1
'        intRow = intRow + 1
        dblChrCnt = 0
        strWordPrint = ""
    End If
End If
End Sub

Private Sub txtAdd_GotFocus()
    'Add By Cheng 2003/03/19
    TextInverse Me.txtAdd
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/03/19
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 8, 89
        '無動作
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub txtCopy_GotFocus()
    'Add By Cheng 2003/03/18
    TextInverse Me.txtCopy
End Sub

Private Sub txtCopy_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/03/18
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case 8, 49, 50, 51
        '無動作
    Case Else
        KeyAscii = 51
    End Select
End Sub

'取得案件名稱
Private Function GetCaseName(strCaseNo As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetCaseName = ""
'Modify By Sindy 2015/7/8 tm05 ==> nvl(tm131,tm05)
StrSQLa = "Select nvl(tm131,tm05)||' '||TM06||' '||TM07 From TradeMark Where " & ChgTradeMark(Replace(strCaseNo, "-", ""))
'Add By Cheng 2003/03/31
'加服務業務基本檔
StrSQLa = StrSQLa & " union Select SP05||' '||SP06||' '||SP07 From ServicePractice Where " & ChgService(Replace(strCaseNo, "-", ""))
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCaseName = Trim("" & rsA.Fields(0).Value)
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'取得系統名稱
Private Function GetSystemName() As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetSystemName = ""

StrSQLa = " And a1k13='" & Me.Text2.Text & "' "
'選擇請款日期
If Me.Option1 Then
     StrSQLa = StrSQLa & " And a1k02>=" & Val(Replace(Me.MaskEdBox1.Text, "/", "")) & " "
     StrSQLa = StrSQLa & " And a1k02<=" & Val(Replace(Me.MaskEdBox2.Text, "/", "")) & " "
'選擇請款編號
Else
     StrSQLa = StrSQLa & " And a1k01>='" & Me.Text3.Text & "' "
     StrSQLa = StrSQLa & " And a1k01<='" & Me.Text4.Text & "' "
End If
StrSQLa = StrSQLa & " And a1k28='" & Me.Text1.Text & "' "
StrSQLa = "Select * From PatentTrademarkMap, Trademark ,acc1k0 Where PTM01='2' And PTM02=TM08 And TM01=a1k13 And TM02=a1k14 And TM03=a1k15 And TM04=a1k16 " & StrSQLa
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetSystemName = "" & rsA("PTM05").Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Sub GetRateAndUSD(strA1k01 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Modified by Morgan 2014/9/23 +a1k02
StrSQLa = "Select A1k10, A1k08,A1k02 From Acc1k0 Where A1k01='" & strA1k01 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    m_strOriExcRate = "" & rsA.Fields(0).Value
    m_strUSD = CDbl(m_strUSD) + rsA.Fields(1).Value
    m_strA1K02 = "" & rsA.Fields("a1k02").Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Sub

'Add By Cheng 2003/03/31
'取得請款項目目細
Private Function GetItemDetail(strR42501 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim blnPrtDetail As Boolean

GetItemDetail = ""
'Modify By Cheng 2003/04/08
''判斷系統類別
'Select Case Me.Text2.Text
'Case "FCT", "S"
'    '判斷請款細項
'    Select Case strR42501
'    Case "01", "02"
'        Exit Function
'    End Select
'End Select
blnPrtDetail = False
'Modify By Cheng 2003/04/09
'依金額由大到小排序
'strSQLA = "Select R42501, NVL(R42502,0), NVL(R42503,0), Count(*) From ACCRPT425 Where R42501='" & strR42501 & "' And ID='" & strUserNum & "' Group By R42501 , NVL(R42502,0), NVL(R42503,0) "
StrSQLa = "Select R42501, NVL(R42502,0), NVL(R42503,0), Count(*) From ACCRPT425 Where R42501='" & strR42501 & "' And ID='" & strUserNum & "' Group By R42501 , NVL(R42502,0), NVL(R42503,0) Order By NVL(R42502,0) Desc "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    '判斷系統類別
    If (Me.Text2.Text = "FCT" Or Me.Text2.Text = "S") Then
        '判斷請款科目代號
        If strR42501 = "01" Or strR42501 = "02" Then
            While Not rsA.EOF
                '若有折扣
                If rsA.Fields(2).Value > 0 Then GoTo PrtDetail
                rsA.MoveNext
            Wend
            GoTo NoPrtDetail
        Else
            GoTo PrtDetail
        End If
    End If
PrtDetail:
    rsA.MoveFirst
    GetItemDetail = GetItemDetail & " ( "
    While Not rsA.EOF
        '若每項超過一筆或有折扣或有多項時
        If rsA.Fields(3).Value > 1 Or rsA.Fields(2).Value > 0 Or rsA.RecordCount > 1 Then
            blnPrtDetail = True
            '金額
'            GetItemDetail = GetItemDetail & "NTD " & Format(rsA.Fields(1).Value, "#,##0.00")
            GetItemDetail = GetItemDetail & "NTD " & PUB_ChgFormat(rsA.Fields(1).Value, True)
            '折扣
            If Val("0" & rsA.Fields(2).Value) <> 0 Then
                'Modify By Cheng 2003/04/02
    '            GetItemDetail = GetItemDetail & " x " & ((100 - rsA.Fields(1).Value) / 100) & "% "
               'Modify by Morgan 2009/4/10 取整數
               'GetItemDetail = GetItemDetail & " x " & ((rsA.Fields(1).Value - rsA.Fields(2).Value) / rsA.Fields(1).Value * 100) & "% "
               GetItemDetail = GetItemDetail & " x " & Round((rsA.Fields(1).Value - rsA.Fields(2).Value) / rsA.Fields(1).Value * 100) & "% "
            End If
            '筆數
            If rsA.Fields(3).Value > 1 Then GetItemDetail = GetItemDetail & " x " & rsA.Fields(3).Value
        End If
        rsA.MoveNext
        If rsA.EOF = False Then
            GetItemDetail = GetItemDetail & " + "
        End If
    Wend
    GetItemDetail = GetItemDetail & " )"
End If
NoPrtDetail:
'edit by nickc 2007/02/08
'If rsA.State <> adstateclose Then rsA.Close
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
If blnPrtDetail = False Then GetItemDetail = ""
End Function

'Add By Cheng 2003/04/03
Private Function GetNationName(strKind As String) As String
'strKind : 2 英文
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset

GetNationName = ""
StrSqlB = ""
If Me.Option1.Value Then
    StrSqlB = StrSqlB & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
    StrSqlB = StrSqlB & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
Else
    StrSqlB = StrSqlB & " and a1k01 >= '" & Text3 & "'"
    StrSqlB = StrSqlB & " and a1k01 <= '" & Text4 & "'"
End If
StrSqlB = "Select a1k13, a1k14, a1k15, a1k16 From Acc1k0 Where a1k13='" & Me.Text2.Text & "' And a1k28='" & Me.Text1.Text & "' " & StrSqlB
rsB.CursorLocation = adUseClient
rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
If rsB.RecordCount > 0 Then
    StrSQLa = "Select NA03, NA04 From Nation, Patent Where NA01=PA09 And " & ChgPatent(rsB.Fields(0).Value & rsB.Fields(1).Value & rsB.Fields(2).Value & rsB.Fields(3).Value)
    StrSQLa = StrSQLa & " Union Select NA03, NA04 From Nation, Trademark Where NA01=TM10 And " & ChgTradeMark(rsB.Fields(0).Value & rsB.Fields(1).Value & rsB.Fields(2).Value & rsB.Fields(3).Value)
    StrSQLa = StrSQLa & " Union Select NA03, NA04 From Nation, Lawcase Where NA01=LC15 And " & ChgLawcase(rsB.Fields(0).Value & rsB.Fields(1).Value & rsB.Fields(2).Value & rsB.Fields(3).Value)
    StrSQLa = StrSQLa & " Union Select NA03, NA04 From Nation, Hirecase Where '000'=NA01 And " & ChgHirecase(rsB.Fields(0).Value & rsB.Fields(1).Value & rsB.Fields(2).Value & rsB.Fields(3).Value)
    StrSQLa = StrSQLa & " Union Select NA03, NA04 From Nation, Servicepractice Where NA01=SP09 And " & ChgService(rsB.Fields(0).Value & rsB.Fields(1).Value & rsB.Fields(2).Value & rsB.Fields(3).Value)
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        GetNationName = IIf(strKind = "2", "" & rsA.Fields(1).Value, "" & rsA.Fields(0).Value)
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End If
If rsB.State <> adStateClosed Then rsB.Close
Set rsB = Nothing

End Function

Private Sub txtOutMode_GotFocus()
   TextInverse txtOutMode
   CloseIme
End Sub

Private Sub txtOutMode_KeyPress(KeyAscii As Integer)
   If KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub LoadHeadPic()
   Dim strPicFileName As String, iNo As Integer
   strPicFileName = App.path & "\$TmpHead.jpg"
   If Dir(strPicFileName) = "" Then
      'Added by Morgan 2020/3/31
      If strSrvDate(1) >= 智慧所更名日 Then
         PUB_GetLetterPicID "2", , iNo, , , , , True
      Else
      'end 2020/3/31
         iNo = 6
      End If 'Added by Morgan 2020/3/31
      If PUB_ReadDB2File(strPicFileName, iNo) = False Then
         Exit Sub
      End If
   End If
   Set Picture1.Picture = LoadPicture(strPicFileName)
   Picture1.AutoSize = True
End Sub

Private Sub DelPic()
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_*.jpg"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
   'm_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
   LoadHeadPic
End Sub

Private Sub SetPic(idx As Integer)
   Dim strPicFileName1 As String
   Dim objImg As StdPicture
   Dim m_Image As New cImage
   Dim m_Jpeg  As New cJpeg
   Dim strFolder As String
   
   ChgCaseNo m_stCaseNo, strExc
   m_EFilePath = PUB_GetEFilePath(strExc(1))
   
   'Modified by Morgan 2013/1/2
   'm_EFilePath = PUB_GetEFilePath(strExc(1))
   'strFolder = m_EFilePath & "\" & m_stCaseNo
   'If Dir(strFolder, vbDirectory) = "" Then
   '   MkDir strFolder
   'End If
   'end 2013/1/2
   strFolder = m_EFilePath & "\" & strExc(1)
   If Dir(strFolder, vbDirectory) = "" Then
      MkDir strFolder
   End If
   
   strFolder = strFolder & "\" & Left(strExc(2), 3)
   If Dir(strFolder, vbDirectory) = "" Then
      MkDir strFolder
   End If
   
   strFolder = strFolder & "\" & m_stCaseNo
   If Dir(strFolder, vbDirectory) = "" Then
      MkDir strFolder
   End If
   'end 2013/1/2
         
   m_stFileName = strFolder & "\" & m_stCaseNo & "_DN" & m_strDN
   If idx = 0 Then
      strPicFileName1 = strFolder & "\" & m_stCaseNo & "_DN" & m_strDN & "S.jpg"
   Else
      strPicFileName1 = strFolder & "\" & m_stCaseNo & "_DN" & m_strDN & "S_P" & idx & ".jpg"
   End If
   
   RidFile strPicFileName1
   PUB_SavePic Picture1, strPicFileName1
   
   '要用覆蓋的否則會錯誤--VB Bug
   'Picture1.Cls
   'm_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
   LoadHeadPic
End Sub

Private Sub CopyImg()
   Dim arrCaseNo
   'Added by Morgan 2013/1/2
   Dim arrCaseCode(4) As String
   Dim strFolder As String
   Dim arrFolders
   
   If m_stCaseNos <> "" Then
      arrCaseNo = Split(m_stCaseNos, ",")
      arrFolders = Split(m_stCaseNos, ",") 'Added by Morgan 2013/1/2
      '建立目錄
      For intI = LBound(arrCaseNo) To UBound(arrCaseNo)
         If arrCaseNo(intI) <> "" Then
            'Modified by Morgan 2013/1/2
'            strExc(1) = m_EFilePath & "\" & arrCaseNo(intI)
'            If Dir(strExc(1), vbDirectory) = "" Then
'               MkDir strExc(1)
'            End If
            '上層目錄=系統別\本所號前3碼\
            ChgCaseNo arrCaseNo(intI), arrCaseCode
            strFolder = PUB_GetEFilePath(arrCaseCode(1)) & "\" & arrCaseCode(1)
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            
            strFolder = strFolder & "\" & Left(arrCaseCode(2), 3)
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            
            strFolder = strFolder & "\" & arrCaseNo(intI)
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            arrFolders(intI) = strFolder
            'end 2013/1/2
         End If
      Next
      'Modified by Morgan 2013/1/2
      'strExc(0) = m_EFilePath & "\" & m_stCaseNo & "\" & m_stCaseNo & "_DN" & m_strDN & "S*.jpg"
      If m_bPrint2Pdf = True Then
         strExc(0) = m_stFileName & ".pdf"
      Else
         strExc(0) = m_stFileName & "S*.jpg"
      End If
      'end 2013/1/2
      
      strExc(2) = Dir(strExc(0))
      Do
         If strExc(2) <> "" Then
            For intI = LBound(arrCaseNo) To UBound(arrCaseNo)
               If arrCaseNo(intI) <> "" Then
                  'Modified by Morgan 2013/1/2
                  'strExc(1) = Replace(strExc(2), m_stCaseNo, arrCaseNo(intI))
                  strExc(1) = arrFolders(intI) & "\" & Replace(strExc(2), m_stCaseNo, arrCaseNo(intI))
                  
                  'Modified by Morgan 2013/1/2
                  'strExc(3) = m_EFilePath & "\" & m_stCaseNo & "\" & strExc(2)
                  strExc(3) = Left(strExc(0), InStrRev(strExc(0), "\") - 1) & "\" & strExc(2)
                  FileCopy strExc(3), strExc(1)
               End If
            Next
         Else
            Exit Do
         End If
         strExc(2) = Dir
      Loop
   End If
End Sub

'Removed by Morgan 2009/12/30
'Private Function GetEMailSet(ByVal p_CuNo As String) As Boolean
'   If p_CuNo <> "" Then
'      If Left(p_CuNo, 1) = "X" Then
'         p_CuNo = Left(p_CuNo & "000", 9)
'         strExc(0) = "select cu124 from customer where cu01='" & Left(p_CuNo, 8) & "' and cu02='" & Mid(p_CuNo, 9) & "'"
'      Else
'         p_CuNo = Left(p_CuNo & "000", 9)
'         strExc(0) = "select FA86 from fagent where fa01='" & Left(p_CuNo, 8) & "' and fa02='" & Mid(p_CuNo, 9) & "'"
'      End If
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If Not IsNull(RsTemp(0)) Then
'            GetEMailSet = True
'         End If
'      End If
'   End If
'End Function

'Added by Morgan 2013/1/2
Private Sub MyNewPage(Optional bolIs1st As Boolean)
   Dim iPicNo As Integer, iPicNo2 As Integer
   Dim strFolder As String
   Dim strFileName As String
   
   If bolIs1st = False Then
      Printer.NewPage
   End If
   
   If m_bPrint2Pdf = True Then
      If bolIs1st Then
         ChgCaseNo m_stCaseNo, strExc
         '上層目錄=系統別\本所號前3碼\
         m_EFilePath = PUB_GetEFilePath(strExc(1))
         strFolder = m_EFilePath & "\" & strExc(1)
         If Dir(strFolder, vbDirectory) = "" Then
            MkDir strFolder
         End If
         
         strFolder = strFolder & "\" & Left(strExc(2), 3)
         If Dir(strFolder, vbDirectory) = "" Then
            MkDir strFolder
         End If
         
         strFolder = strFolder & "\" & m_stCaseNo
         If Dir(strFolder, vbDirectory) = "" Then
            MkDir strFolder
         End If
         
         m_stFileName = strFolder & "\" & m_stCaseNo & "_DN" & m_strDN & "S"
         
         Printer.Orientation = 1
         frmPDF.StartProcess strFolder, strFileName
      End If
      
      'If m_bWord2Pdf = False Then
         'Added by Morgan 2020/3/31
         If strSrvDate(1) >= 智慧所更名日 Then
            PUB_GetLetterPicID "2", Text2, iPicNo, iPicNo2, 2, True, Pub_StrUserSt03
         Else
         'end 2020/3/31
            iPicNo = 5
            iPicNo2 = 9
               
            If Left(Text1, 1) = "X" Then
               strExc(0) = "SELECT CU10 FROM CUSTOMER WHERE CU01='" & Left(Text1, 8) & "' and CU02='" & Mid(Text1, 9) & "'"
            Else
               strExc(0) = "SELECT FA10 FROM FAGENT WHERE FA01='" & Left(Text1, 8) & "' and FA02='" & Mid(Text1, 9) & "'"
            End If
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp(0) = "020" Then
                  iPicNo = 7
                  iPicNo2 = 0
               End If
            End If
            
         End If 'Added by Morgan 2020/3/31
         
         PrintPicture iPicNo, iPicNo2
      'End If
   End If
End Sub

Private Sub PrintPicture(iPicNo As Integer, Optional iPicNo2 As Integer)
   Dim tObj As New StdPicture, stFileName As String, pWidth As Long, pHeight As Long
   pWidth = 21 * 567
   If PUB_ReadDB2File(stFileName, iPicNo) = True Then
      Set tObj = pvGetStdPicture(stFileName)
      pHeight = tObj.Height * (pWidth / tObj.Width)
      Printer.PaintPicture tObj, 0, Int(0.5 * 567), pWidth, pHeight
   End If
   If iPicNo2 > 0 Then
      If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
         Set tObj = pvGetStdPicture(stFileName)
         pHeight = tObj.Height * (pWidth / tObj.Width)
         Printer.PaintPicture tObj, 0, Int(27 * 567), pWidth, pHeight
      End If
   End If
   Set tObj = Nothing
End Sub

Private Function PrinterIndex(Printername As String) As Long
   Dim i As Long
   PrinterIndex = -1
   For i = 0 To Printers.Count - 1
    If UCase(Printers(i).DeviceName) = UCase$(Printername) Then
     PrinterIndex = i
     Exit For
    End If
   Next i
End Function

'Add by Morgan 2014/9/9
Private Sub SetWordArray(pArray() As INVITEM, pRow As Integer, pIndex As Integer, pData As String, Optional pType As Single = 1, Optional pBold As Boolean = False, Optional pULine As Boolean = False, Optional pBottomLine As Boolean = False)
   If UBound(pArray) < pRow Then
      ReDim Preserve pArray(pRow)
   End If
   pArray(pRow).iType = pType
   
   If pIndex = 0 Then
      pArray(pRow).IBottomLine = pBottomLine
   End If
   
   Select Case pIndex
      Case 1
         pArray(pRow).IText1 = pArray(pRow).IText1 & pData
         pArray(pRow).IBold1 = pBold
         pArray(pRow).IULine1 = pULine
      Case 2
         pArray(pRow).IText2 = pArray(pRow).IText2 & pData
         pArray(pRow).IBold2 = pBold
         pArray(pRow).IULine2 = pULine
      Case 3
         pArray(pRow).IText3 = pArray(pRow).IText3 & pData
         pArray(pRow).IBold3 = pBold
         pArray(pRow).IULine3 = pULine
      Case 4
         pArray(pRow).IText4 = pArray(pRow).IText4 & pData
         pArray(pRow).IBold4 = pBold
         pArray(pRow).IULine4 = pULine
      Case 5
         pArray(pRow).IText5 = pArray(pRow).IText5 & pData
         pArray(pRow).IBold5 = pBold
         pArray(pRow).IULine5 = pULine
   End Select
End Sub


'Added by Morgan 2014/9/10
Private Sub runWord(Optional bolIsEndDoc As Boolean)

   Dim iRow As Integer, iCol As Integer, lngCellWidth As Long
   Dim strLastCountry As String, iFCol As Integer, iCols As Integer, bolPrintCountry As Boolean
   Dim strFontSize As String
   Dim iResumeCnt As Integer
   Dim bVisible As Boolean, stFileName As String, iHeadCount As Integer
   Dim strText As String
   Dim oShape
   Dim strLstStr As String
   Dim strLstNo As String
   Dim dblOffFeeSub As Double
   Dim dblAttFeeSub As Double
   Dim bolNewRow As Boolean
   Dim ii As Integer
   Dim sLastType As Single
   Dim sCnt As Single
   Dim sSubRowNo As Single
   Dim iPicNo As Integer, iPicNo2 As Integer
   
On Error GoTo ErrHnd

   strFontSize = 12
   
   If m_bolNewDoc = True Then
      If NewDoc() = False Then Exit Sub
   End If
   
   If UBound(m_Item) < 3 Then GoTo FlgAddPic
   
   With g_WordAp.Application
      If m_bolNewDoc = True Then
         '版面設定
         .Selection.PageSetup.Orientation = wdOrientPortrait
         .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
         .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.5)
         .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.3)
         .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1)
         .Selection.PageSetup.FooterDistance = .CentimetersToPoints(3)
         .Selection.PageSetup.CharsLine = 40
         .Selection.PageSetup.LinesPage = 38
         .Selection.Orientation = wdTextOrientationHorizontal
         .Selection.Font.Size = strFontSize
         '行距
         With .Selection.ParagraphFormat
           .SpaceBefore = 0
           .SpaceAfter = 0
           .LineSpacingRule = wdLineSpaceSingle
           .DisableLineHeightGrid = True
         End With
         m_bolNewDoc = False
      Else
         .Selection.EndKey Unit:=wdStory
         .Selection.InsertBreak wdPageBreak
      End If
      
      
      .Selection.TypeParagraph
      
      '新增表格(1*4)
      .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=4
      
      With .Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
      End With
    
      '設定表格高度欄寬
      .Selection.SelectRow
      .Selection.Cells.VerticalAlignment = wdAlignVerticalTop
      .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.9), RulerStyle:=wdAdjustProportional
      .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(8.7), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
      .Selection.Cells(3).VerticalAlignment = wdCellAlignVerticalTop
      
      .Selection.InsertRows UBound(m_Item) - 3
      .Selection.Collapse Direction:=wdCollapseStart
      
      For iRow = 3 To UBound(m_Item)
         If sLastType <> m_Item(iRow).iType Then
            If m_Item(iRow).iType <> 0 Then
               sLastType = m_Item(iRow).iType
               sSubRowNo = 0
            End If
            sCnt = 0
         End If
         sSubRowNo = sSubRowNo + 1
         Select Case m_Item(iRow).iType
         Case 0
            .Selection.MoveDown Unit:=wdLine, Count:=1
         Case 1 '表頭
            .Selection.Cells(1).SetHeight RowHeight:=12, HeightRule:=wdRowHeightAtLeast
            If m_Item(iRow).IText1 <> "" Then
               .Selection.TypeText Text:=m_Item(iRow).IText1
            End If
            
            If m_Item(iRow).IText2 <> "" Then
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               '第2欄欄位合併,否則若第3欄內容有折行時第2欄會隔空白
               If iRow > 1 Then
                  If m_Item(iRow - 1).IText2 <> "" Then
                     .Selection.MoveUp Unit:=wdLine, Count:=1, Extend:=wdExtend
                     .Selection.Cells.Merge
                     .Selection.EndKey
                     .Selection.TypeParagraph
                     sCnt = sCnt + 1
                  End If
               End If
               .Selection.TypeText Text:=m_Item(iRow).IText2
            Else
               .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
               .Selection.Cells.Merge
            End If
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            
            If sCnt > 0 Then
               .Selection.MoveDown Unit:=wdLine, Count:=sCnt
            End If
            
            If m_Item(iRow).IText3 <> "" Then
               .Selection.TypeText Text:=m_Item(iRow).IText3
            End If
            
            If m_Item(iRow).IText4 <> "" Then
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               .Selection.TypeText Text:=m_Item(iRow).IText4
            Else
               .Selection.MoveRight Unit:=wdCharacter, Count:=2, Extend:=wdExtend
               .Selection.Cells.Merge
            End If
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            
         Case 2 '標題
            .Selection.SelectRow
            If sSubRowNo = 1 Then
               '.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
               With .Selection.Cells
                  .VerticalAlignment = wdAlignVerticalBottom
                   With .Borders(wdBorderBottom)
                       .LineStyle = wdLineStyleSingle
                       .LineWidth = wdLineWidth100pt
                       .ColorIndex = wdAuto
                   End With
               End With
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
               .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(4.6), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
            Else
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.8), RulerStyle:=wdAdjustProportional
               .Selection.Collapse Direction:=wdCollapseStart
            End If
            If m_Item(iRow).IText1 <> "" Then
               .Selection.TypeText Text:=m_Item(iRow).IText1
            End If
            
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            If m_Item(iRow).IText2 <> "" Then
               .Selection.TypeText Text:=m_Item(iRow).IText2
            Else
               .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
               .Selection.Cells.Merge
            End If
            
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            If m_Item(iRow).IText3 <> "" Then
               .Selection.TypeText Text:=m_Item(iRow).IText3
            Else
               .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
               .Selection.Cells.Merge
            End If
            
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
               
         Case 3 '請款項目
            .Selection.SelectRow
            '加總首頁
            If m_strDetailKind = "" Then
               If sSubRowNo = 1 Then
                  .Selection.Font.Size = 16
               End If
               .Selection.Font.Italic = True
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(1.5), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(7), RulerStyle:=wdAdjustProportional
            '加總明細
            ElseIf m_strDetailKind = "1" Then
               If m_Item(iRow).IBottomLine = True Then
                  With .Selection.Cells
                     .VerticalAlignment = wdAlignVerticalBottom
                      With .Borders(wdBorderBottom)
                          .LineStyle = wdLineStyleSingle
                          .LineWidth = wdLineWidth100pt
                          .ColorIndex = wdAuto
                      End With
                  End With
               End If
               .Selection.Cells.Split NumRows:=1, NumColumns:=5, MergeBeforeSplit:=True
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(0.4), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(3.6), RulerStyle:=wdAdjustProportional
               'Modified by Morgan 2018/7/31 若彼所案號太長會折行而導致跳頁問題 --阿蓮 Ex:X10711284-292
               '.Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(4), RulerStyle:=wdAdjustProportional
               .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(3), RulerStyle:=wdAdjustProportional
               'end 2018/7/31
               .Selection.Cells(4).SetWidth ColumnWidth:=.CentimetersToPoints(4), RulerStyle:=wdAdjustProportional
            '加總明細項目
            Else
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(13), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(1.8), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).VerticalAlignment = wdCellAlignVerticalBottom 'Added by Morgan 2016/12/2
               .Selection.Cells(3).VerticalAlignment = wdCellAlignVerticalBottom 'Added by Morgan 2016/12/2
            End If
            
            .Selection.Collapse Direction:=wdCollapseStart
            If m_Item(iRow).IText1 <> "" Then
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               If m_Item(iRow).IBold1 = True Then .Selection.Font.Bold = True
               If m_Item(iRow).IULine1 = True Then .Selection.Font.Underline = True
               .Selection.TypeText Text:=m_Item(iRow).IText1
            End If
            
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            If m_Item(iRow).IText2 <> "" Then
               If m_Item(iRow).IBold2 = True Then .Selection.Font.Bold = True
               If m_Item(iRow).IULine2 = True Then .Selection.Font.Underline = True
               
               .Selection.TypeText Text:=m_Item(iRow).IText2
            End If
            
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            If m_Item(iRow).IText3 <> "" Then
               If m_Item(iRow).IBold3 = True Then .Selection.Font.Bold = True
               If m_Item(iRow).IULine3 = True Then .Selection.Font.Underline = True
               .Selection.TypeText Text:=m_Item(iRow).IText3
            Else
               .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
               .Selection.Cells.Merge
            End If
            
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            If m_Item(iRow).IText4 <> "" Then
               If m_Item(iRow).IBold4 = True Then .Selection.Font.Bold = True
               If m_Item(iRow).IULine4 = True Then .Selection.Font.Underline = True
               .Selection.TypeText Text:=m_Item(iRow).IText4
            Else
               .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
               .Selection.Cells.Merge
            End If
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            If m_Item(iRow).IText5 <> "" Then
               If m_Item(iRow).IBold5 = True Then .Selection.Font.Bold = True
               If m_Item(iRow).IULine5 = True Then .Selection.Font.Underline = True
               .Selection.TypeText Text:=m_Item(iRow).IText5
            Else
               .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
               .Selection.Cells.Merge
            End If
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            
         Case 4 '合計
            .Selection.SelectRow
            .Selection.Cells.Split NumRows:=1, NumColumns:=4, MergeBeforeSplit:=True
            '加總首頁
            If m_strDetailKind = "" Then
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
               .Selection.Font.Bold = True
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(4.5), RulerStyle:=wdAdjustProportional
            ElseIf m_strDetailKind = "1" Then
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(4), RulerStyle:=wdAdjustProportional
            Else
               .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight
               .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(9.6), RulerStyle:=wdAdjustProportional
               .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(3.4), RulerStyle:=wdAdjustProportional
               .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(1.8), RulerStyle:=wdAdjustProportional
            End If
            
            If m_strDetailKind <> "" Then
               If sSubRowNo = 1 Then
                  With .Selection.Cells
                      With .Borders(wdBorderTop)
                          .LineStyle = wdLineStyleSingle
                          .LineWidth = wdLineWidth100pt
                          .ColorIndex = wdAuto
                      End With
                  End With
                  If m_strDetailKind = "2" Then
                     .Selection.Cells(1).SetHeight RowHeight:=24, HeightRule:=wdRowHeightAtLeast
                     .Selection.Cells.VerticalAlignment = wdCellAlignVerticalBottom
                  End If
               End If
            End If
            .Selection.Collapse Direction:=wdCollapseStart
            If m_Item(iRow).IText1 <> "" Then
               .Selection.TypeText Text:=m_Item(iRow).IText1
            End If
            
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            If m_Item(iRow).IText2 <> "" Then
               If m_strDetailKind = "2" Then
                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
               End If
               .Selection.TypeText Text:=m_Item(iRow).IText2
            End If
            
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            If m_Item(iRow).IText3 <> "" Then
               .Selection.TypeText Text:=m_Item(iRow).IText3
            End If
            
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            If m_Item(iRow).IText4 <> "" Then
               .Selection.TypeText Text:=m_Item(iRow).IText4
            Else
               .Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
               .Selection.Cells.Merge
            End If
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            
         Case 5 '表尾
            .Selection.SelectRow
            .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=True
            
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(13), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(2.1), RulerStyle:=wdAdjustProportional
            
            .Selection.Collapse Direction:=wdCollapseStart
            If m_Item(iRow).IText1 <> "" Then
               If m_Item(iRow).IBold1 = True Then .Selection.Font.Bold = True
               .Selection.TypeText Text:=m_Item(iRow).IText1
            End If
            
            If m_Item(iRow).IText2 <> "" Then
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
               sCnt = sCnt + 1
               If m_Item(iRow).IBold2 = True Then .Selection.Font.Bold = True
               .Selection.TypeText Text:=m_Item(iRow).IText2
               If m_Item(iRow + 1).IText2 = "" Then
                  '.Selection.MoveUp Unit:=wdLine, Count:=sCnt, Extend:=wdExtend
                  .Selection.MoveUp Unit:=wdLine, Count:=sCnt - 1
                  .Selection.MoveDown Unit:=wdLine, Count:=sCnt - 1, Extend:=wdExtend
                  .Selection.Cells.Merge
                  With .Selection.Cells(1)
                      With .Borders(wdBorderLeft)
                          .LineStyle = wdLineStyleSingle
                          .LineWidth = wdLineWidth100pt
                          .ColorIndex = wdAuto
                      End With
                      With .Borders(wdBorderRight)
                          .LineStyle = wdLineStyleSingle
                          .LineWidth = wdLineWidth100pt
                          .ColorIndex = wdAuto
                      End With
                      With .Borders(wdBorderTop)
                          .LineStyle = wdLineStyleSingle
                          .LineWidth = wdLineWidth100pt
                          .ColorIndex = wdAuto
                      End With
                      With .Borders(wdBorderBottom)
                          .LineStyle = wdLineStyleSingle
                          .LineWidth = wdLineWidth100pt
                          .ColorIndex = wdAuto
                      End With
                  End With
                  .Selection.ParagraphFormat.LeftIndent = .CentimetersToPoints(0.2)
                  .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                  .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
                  If sCnt > 1 Then
                     .Selection.MoveDown Unit:=wdLine, Count:=sCnt - 1
                  End If
               Else
                  .Selection.MoveRight Unit:=wdCharacter, Count:=1
               End If
            ElseIf m_Item(iRow).IText2 & m_Item(iRow).IText3 = "" Then
               .Selection.MoveRight Unit:=wdCharacter, Count:=3, Extend:=wdExtend
               .Selection.Cells.Merge
            Else
               .Selection.MoveRight Unit:=wdCharacter, Count:=1
            End If
            
            If m_Item(iRow).IText3 <> "" Then
               .Selection.TypeText Text:=m_Item(iRow).IText3
            End If
            
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
         Case 6 '頁碼
            .Selection.SelectRow
            .Selection.Cells.Split NumRows:=1, NumColumns:=1, MergeBeforeSplit:=True
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Selection.Collapse Direction:=wdCollapseStart
            If m_Item(iRow).IText1 <> "" Then
               .Selection.TypeText Text:=m_Item(iRow).IText1
            End If
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            
         Case Else
            .Selection.MoveDown Unit:=wdLine, Count:=1
         End Select
      Next
      .Selection.WholeStory
      .Selection.Font.Name = "Times New Roman"
      .Selection.HomeKey Unit:=wdStory
   End With
   
FlgAddPic:

   With g_WordAp.Application
      If bolIsEndDoc Then
         .Selection.WholeStory
         .Selection.HomeKey Unit:=wdStory
         'Added by Morgan 2020/3/31
         If strSrvDate(1) >= 智慧所更名日 Then
            PUB_GetLetterPicID "2", Text2, iPicNo, iPicNo2, 2, True, Pub_StrUserSt03
         Else
         'end 2020/3/31
            iPicNo = 5
            iPicNo2 = 9
         End If 'Added by Morgan 2020/3/31
         
         If PUB_ReadDB2File(stFileName, iPicNo) Then
            For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
               Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
               oShape.ZOrder 4
               oShape.LockAnchor = True
               oShape.LockAspectRatio = -1
               oShape.Width = .CentimetersToPoints(21)
               oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
               oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
               oShape.Left = .CentimetersToPoints(0)
               oShape.Top = .CentimetersToPoints(0.5)
               oShape.WrapFormat.Type = wdWrapNone
               .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
            Next ii
            .Selection.HomeKey Unit:=wdStory
            
            If iPicNo2 > 0 Then
               If PUB_ReadDB2File(stFileName, iPicNo2) Then
                  For ii = 1 To .ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)
                     Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
                     oShape.ZOrder 4
                     oShape.LockAnchor = True
                     oShape.LockAspectRatio = -1
                     oShape.Width = .CentimetersToPoints(21)
                     oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                     oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
                     oShape.Left = .CentimetersToPoints(0)
                     'Added by Morgan 2020/3/31
                     If strSrvDate(1) >= 智慧所更名日 Then
                        oShape.Top = .CentimetersToPoints(27.6)
                     Else
                     'end 2020/3/31
                        oShape.Top = .CentimetersToPoints(27)
                     End If 'Added by Morgan 2020/3/31
                     oShape.WrapFormat.Type = wdWrapNone
                     .Selection.GoTo what:=wdGoToPage, which:=wdGoToNext, Count:=1
                  Next ii
               End If
               .Selection.HomeKey Unit:=wdStory
            End If
         End If
         
         'Added by Morgan 2016/10/11 外商人員列印時也要存電子檔(特殊請款單都要,因為財務處會要)--陳金蓮
         'Modified by Morgan 2020/11/2 外商改要印信頭(從上面移下來)--陳金蓮,徐湘�A
         If Pub_StrUserSt03 = "F12" And m_bPrinter Then
            '
            If m_bolEmail And Not m_bolPaper Then
               txtCopy.Text = 1
            End If
            
            'Word印表機切換為表單上的
            pub_OsPrinter = PUB_GetOsDefaultPrinter
            PUB_SetOsDefaultPrinter Combo1
            PUB_SetWordActivePrinter
            PUB_SetOsDefaultPrinter pub_OsPrinter
            .PrintOut FileName:="", Range:=wdPrintAllDocument, Item:=wdPrintDocumentContent, _
                        Copies:=IIf(Val("0" & Me.txtCopy.Text) = "0", 3, Val("0" & Me.txtCopy.Text)), Pages:="", PageType:=wdPrintAllPages, _
                        ManualDuplexPrint:=False, Collate:=True, Background:=False, PrintToFile:=False
         End If
         'end 2020/11/2
         'end 2016/10/11
         
         '存檔
         If Pub_StrUserSt03 = "F12" And Text6.Text = "2" And m_DocName <> "" Then
            'Modified by Morgan 2015/5/20
            'strExc(1) = "\\Typing2\國外部\外商\FCT Revised Debit Note\" & m_DocName & ".doc"
            strExc(1) = Pub_GetSpecMan("FCT特殊請款單存放路徑") & "\" & m_DocName & ".doc"
            'end 2015/5/20
            If Dir(strExc(1)) <> "" Then
               .Visible = False 'Added by Morgan 2015/5/25
               If MsgBox("檔案已存在，是否要覆蓋？" & vbCrLf & vbCrLf & "( " & strExc(1) & " )", vbYesNo + vbQuestion) = vbNo Then
                  strExc(1) = ""
               End If
            End If
            If strExc(1) <> "" Then
               RidFile strExc(1)
               .ActiveDocument.SaveAs strExc(1)
               .Visible = False 'Added by Morgan 2015/5/25
               MsgBox "電子檔已儲存！" & vbCrLf & vbCrLf & "( " & strExc(1) & " )", vbInformation
            End If
         End If
         
         'Modify By Sindy 2015/7/13 列印出來
         If (Text2 = "FCP" Or Text2 = "FG") And txtOutMode = "1" Then
            'Added by Morgan 2015/12/17 Word印表機切換為表單上的
            pub_OsPrinter = PUB_GetOsDefaultPrinter
            PUB_SetOsDefaultPrinter Combo1
            PUB_SetWordActivePrinter
            PUB_SetOsDefaultPrinter pub_OsPrinter
            'end 2015/12/17
            
            .ActiveDocument.SaveAs App.path & "\$$temp" & strSrvDate(1) & ServerTime & ".doc"
            .PrintOut FileName:="", Range:=wdPrintAllDocument, Item:=wdPrintDocumentContent, _
                        Copies:=IIf(Val("0" & Me.txtCopy.Text) = "0", 3, Val("0" & Me.txtCopy.Text)), Pages:="", PageType:=wdPrintAllPages, _
                        ManualDuplexPrint:=False, Collate:=True, Background:=False, PrintToFile:=False
            .ActiveDocument.Close
            .Quit
            
         'Added by Morgan 2016/10/11
         ElseIf m_bPrinter Then
            .ActiveDocument.Close
            .Quit
         'end 2016/10/11
         End If
         '2015/7/13 END
      End If
      
      If Not ((Text2 = "FCP" Or Text2 = "FG") And txtOutMode = "1") Then 'Add By Sindy 2015/7/13 +if
         If Not m_bPrinter Then  'Added by Morgan 2016/10/11
            .Visible = True 'Added by Morgan 2015/5/25
         End If
      End If
      
   End With
   If Not ((Text2 = "FCP" Or Text2 = "FG") And txtOutMode = "1") Then 'Add By Sindy 2015/7/13 +if
      If Not m_bPrinter Then 'Added by Morgan 2016/10/11
         g_WordAp.Activate
      End If
   End If
   
   Erase m_Item
   ReDim m_Item(1)
   Exit Sub
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox "錯誤 : " & Err.Description, vbCritical
   End If
End Sub

Private Function NewDoc() As Boolean

   Dim iResumeCnt As Integer
   
On Error GoTo ErrHnd
   
   If TypeName(g_WordAp) <> "Application" Then
      Set g_WordAp = New Word.Application
   End If
   g_WordAp.Documents.add
   '不顯示可能會有問題
   g_WordAp.Visible = True
   'g_WordAp.ActiveWindow.ActivePane.View.Zoom.Percentage = 100
   NewDoc = True
   Exit Function
   
ErrHnd:
   'Resume
   If Err.Number <> 0 Then
      If iResumeCnt > 3 Then
         MsgBox "錯誤 : " & Err.Description, vbCritical
      Else
         iResumeCnt = iResumeCnt + 1
         Select Case Err.Number
            Case 91:
               g_WordAp.Documents.add
               Resume Next
            Case 462:
               Set g_WordAp = New Word.Application
               Resume
            Case Else:
               MsgBox "錯誤 : " & Err.Description, vbCritical
         End Select
      End If
   End If
End Function

Private Function UpdateBatch(pNoList As String) As Boolean
   Dim stSQL As String, intR As Integer, bUpdateA1K08 As Boolean
   Dim arrNo() As String
   
   arrNo = Split(pNoList, ",")
   
   cnnConnection.BeginTrans
On Error GoTo ErrHnd
   'Modified by Lydia 2015/04/15 為了區別整批請款單,a1k32=C
   'stSQL = "update ACC1K0 set a1k32='Y' Where A1K01 In (" & pNoList & ")"
   stSQL = "update ACC1K0 set a1k32='C' Where A1K01 In (" & pNoList & ")"
   cnnConnection.Execute stSQL, intR
   
   stSQL = "delete ACC1T0  Where A1T01 In (" & pNoList & ")"
   cnnConnection.Execute stSQL, intR
   'Memo by Lydia 2015/05/08 供電腦中心查詢用
   stSQL = "insert into ACC1T0 (a1t01,a1t02,a1t03,a1t04,a1t05)  select a1k01," & arrNo(0) & ",'" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss') from acc1k0 where a1k01 In (" & pNoList & ")"
   cnnConnection.Execute stSQL, intR
   'Modified by Lydia 2015/05/07 改到UpdateBatchTrunc
'   'FCT,FMT,S,CFT,CFC 外幣金額改四捨五入到小數第二位
'   If m_bBatchRule Then
'      stSQL = "update acc1k0 set a1k08=round( a1k11/a1k10,2) where A1K01 In (" & pNoList & ")"
'      cnnConnection.Execute stSQL, intR
'   End If
   
   cnnConnection.CommitTrans
   UpdateBatch = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function
'Added by Morgan 2014/9/19
'是否適用整批列印規則
Private Function ChkBatchRule(pA1k01 As String) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset

   ChkBatchRule = False
   If InStr(",FCT,T,S,CFT,CFC,", "," & Text2 & ",") > 0 Then
      If Text2 = "T" Then
         stSQL = "select 1 from caseprogress where cp60='" & pA1k01 & "' and cp12 like 'F%'"
         intR = 1
         Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
         If intR = 1 Then
            ChkBatchRule = True
         End If
      Else
         ChkBatchRule = True
      End If
   End If
   Set rsQuery = Nothing
End Function
'Added by Lydia 2015/05/07 修正整批請款單小數點差異
Private Function UpdateBatchTrunc(iNoList As String, MaxDNamt As Double) As Boolean
   Dim stSQL As String, intR As Integer, bUpdateA1K08 As Boolean
   Dim rsQuery As ADODB.Recordset
   Dim MaxDNno As String '記錄最大金額請款單號
   Dim DetailDNamt As Double '記錄其他請款單外幣金額
   
On Error GoTo ErrHnd

    stSQL = "select a1k01,trunc(a1k11/a1k10,0) amt from acc1k0 where A1K01 In (" & iNoList & ") order by 2 desc "
    intR = 1
    Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
    If intR = 1 Then
       rsQuery.MoveFirst
       MaxDNno = rsQuery(0)
       Do While Not rsQuery.EOF
          If MaxDNno <> rsQuery(0) Then
             DetailDNamt = DetailDNamt + Val(rsQuery(1))
          End If
          rsQuery.MoveNext
       Loop
    End If
    
   cnnConnection.BeginTrans
   
   '除最大金額請款單外,其他請款單外幣金額小數點歸零
   stSQL = "update acc1k0 set a1k08=trunc(a1k11/a1k10,0) where A1K01 In (" & iNoList & ") and A1K01<>'" & MaxDNno & "'"
   cnnConnection.Execute stSQL, intR
   'FCT,FMT,S,CFT,CFC 整批請款單以最大金額請款單做整批金額的修正
   stSQL = "update acc1k0 set a1k08=" & MaxDNamt - DetailDNamt & " where A1K01='" & MaxDNno & "'"
   cnnConnection.Execute stSQL, intR
   
   cnnConnection.CommitTrans
   UpdateBatchTrunc = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function

