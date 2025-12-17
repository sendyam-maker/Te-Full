VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc2450 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "國外付款明細表"
   ClientHeight    =   4464
   ClientLeft      =   36
   ClientTop       =   252
   ClientWidth     =   5796
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4464
   ScaleWidth      =   5796
   Begin VB.TextBox txtSend 
      Height          =   315
      Left            =   4230
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   1950
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdPath2 
      Height          =   330
      Left            =   5340
      Picture         =   "Frmacc2450.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   2708
      Width           =   350
   End
   Begin VB.TextBox txtPath2 
      Height          =   315
      Left            =   2427
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2708
      Width           =   2925
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   1548
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2340
      Width           =   3804
   End
   Begin VB.CommandButton cmdPath 
      Height          =   330
      Left            =   5340
      Picture         =   "Frmacc2450.frx":0102
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   2340
      Width           =   350
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
      Left            =   1650
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   450
   End
   Begin VB.TextBox Text13 
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
      Height          =   315
      Left            =   2145
      TabIndex        =   20
      Top             =   120
      Width           =   3010
   End
   Begin VB.ComboBox cmbPrinter2 
      Height          =   300
      Left            =   1545
      Style           =   2  '單純下拉式
      TabIndex        =   17
      Top             =   3510
      Width           =   4152
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1545
      Style           =   2  '單純下拉式
      TabIndex        =   16
      Top             =   3105
      Width           =   4152
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
      Height          =   300
      Left            =   2085
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1560
      Width           =   450
   End
   Begin VB.TextBox Text5 
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
      Left            =   1650
      MaxLength       =   1
      TabIndex        =   1
      Top             =   480
      Width           =   450
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   435
      Left            =   5250
      ScaleHeight     =   384
      ScaleWidth      =   420
      TabIndex        =   13
      Top             =   900
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      ItemData        =   "Frmacc2450.frx":0204
      Left            =   600
      List            =   "Frmacc2450.frx":0206
      Style           =   2  '單純下拉式
      TabIndex        =   7
      Top             =   1956
      Width           =   2568
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
      Left            =   3570
      TabIndex        =   5
      Top             =   1200
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
      Left            =   1650
      TabIndex        =   4
      Top             =   1200
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1650
      TabIndex        =   2
      Top             =   840
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
      Caption         =   "執行(&P)"
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
      Left            =   600
      Style           =   1  '圖片外觀
      TabIndex        =   8
      Top             =   3870
      Width           =   4692
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3570
      TabIndex        =   3
      Top             =   840
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
   Begin VB.Label lblSend 
      BackStyle       =   0  '透明
      Caption         =   "寄送對象"
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
      Left            =   3270
      TabIndex        =   28
      Top             =   2010
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "特殊結匯帳單匯出路徑"
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
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   2250
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "帳單匯出路徑"
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
      Left            =   120
      TabIndex        =   24
      Top             =   2388
      Width           =   1368
   End
   Begin VB.Label Label9 
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
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   150
      Width           =   855
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印表機(留所)"
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
      Left            =   120
      TabIndex        =   19
      Top             =   3150
      Width           =   1305
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印表機(寄送)"
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
      Left            =   120
      TabIndex        =   18
      Top             =   3570
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "輸出選項          (1:列印 2.發EMail 3.單據匯出)"
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
      Left            =   600
      TabIndex        =   15
      Top             =   528
      Width           =   4632
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "是否存電子檔         (Y:是)"
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
      Left            =   600
      TabIndex        =   14
      Top             =   1605
      Width           =   2625
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
      Left            =   3330
      TabIndex        =   12
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1230
      Width           =   975
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
      Left            =   3330
      TabIndex        =   10
      Top             =   840
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "結匯日期"
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
      Left            =   600
      TabIndex        =   9
      Top             =   870
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc2450"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/2 日期欄已修改
Option Explicit
Public adoquery As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Public adoaccsum As New ADODB.Recordset
Public adoaccrpt219 As New ADODB.Recordset
Dim strConSql As String
Dim strNo As String
Dim m_a1b01 As String    '2012/2/23 add by sonia
Dim douAmount As Double
Dim strAmount As String
Dim intLength As Integer
Dim intCounter As Integer
Dim douUSDollar As Double
Dim strLanguage As String
Dim strCurrency As String
Dim intPage As Integer
Dim m_a1b06 As String
Dim intMargin As Integer
'Add by Morgan 2004/11/18 紀錄多種幣別與合計
Dim m_Cur() As String, m_Curs As String, m_Len As Integer, m_Idx As Integer
Dim m_Amount() As Double

'Add by Morgan 2006/12/11
Dim strLangTmp As String '暫存語文
Dim bolEmail As Boolean '是否發EMail
Dim strInform As String '是否寄發電匯通知
Dim strEMailBox As String 'EMail Box
Dim strEmailCC As String 'Added by Lydia 2024/09/18 財務副本信箱
Dim strPicLetter As String '暫存圖檔(信紙=信頭+信尾)路徑 'Memo by Lydia 2024/12/11 信頭/尾的暫存圖: 從strPicFileName=>strPicLetter
Dim strPicFileNames As String '暫存圖檔路徑組(*號分隔)
Dim strPayAmount As String '付款金額
Dim douExtRate As Double '字型位置縮放比
Dim lngX As Long, lngY As Long
'Add by Morgan 2008/4/16
Dim bol2File As Boolean '是否產生電子檔
Dim bol2Printer As Boolean '是否列印
Dim strSavePath As String '電子檔存放路徑
Dim m_strPayDate As String 'Add by Morgan 2010/6/7 結匯日期
Dim m_iNo As Integer, m_iNo2 As Integer '圖檔編號 Add by Morgan 2011/10/11
Dim intTop As Integer '上邊界調整值 Added by Morgan 2012/3/14
Dim intBottom As String '下邊界調整值 Added by Morgan 2012/3/14
Dim intPrtXFix As Integer '紙本上邊界調整值 Added by Morgan 2012/3/14
Dim strPrinter As String, strPrinter2 As String 'Added by Morgan 2012/10/11
Dim strCompName As String 'Add by Amy 2014/05/01 公司抬頭
'Add by Amy 2015/03/19  'modify by sonia 2016/3/1 +;81040閻副所長(因99047kelly留職停薪2016/09/01上班)
'Modify by Amy 2017/11/15 改通知楊雯芳;陳增廣
'Modify by Amy 2025/04/08 陳增廣留職停薪,故改為業拓公共信箱 原:99033;A4024
Const strMailTaie As String = "bd@taie.com.tw" '結匯完成寄信通知國外部
Dim stra1706 As String '代理人D/N or C/N
Dim bol2Email As Boolean 'Added by Lydia 2016/10/21 票匯是否可以Email
'Added by Lydia 2017/10/02 留所資料改成清單列印 (代理人編號  /代理人名稱  /本所案號  /代理人DN No.  /幣別  /金額)
'下列為ACCRPT2450的欄位
'ID  VARCHAR2(6) 建檔人
'R24500  NUMBER(4)   序號
'R24501  VARCHAR2(10)    代理人編號/小計
'R24502  VARCHAR2(150)   代理人名稱
'R24503  VARCHAR2(15)    本所案號
'R24504  VARCHAR2(100)   代理人DN No.
'R24505  VARCHAR2(3) 幣別
'R24506  NUMBER(15,2)    金額
'R24507  VARCHAR2(1) 是否下載Invoice(Acc152)
'R24508  VARCHAR2(15)    帳單編號
Dim bolList As Boolean '改成清單列印
Dim PLeft(0 To 7) As Integer '欄位邊界
Dim strTemp(0 To 5) As String
Dim mCnt As Integer  '序號
Dim bolPrintPage As Boolean '是否列印明細
Dim mPrtOrt As Integer  '原本預設印表機的列印方向
Private Const ciTitleFontSize = 14
Private Const ciFontSize = 10
Private Const ciStartX = 600, ciStartY = 500, ciColGap = 150
Dim iPrint As Integer, strPrtOrt As Integer
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long

'Added by Lydia 2018/09/03 一張匯票存成一個PDF檔
Dim bol2Pdf As Boolean
'Added by Lydia 2020/09/03 調整明細表的版面: 欄位的起始位置
Private Const ciXOur = 0 'OUR REF/Invoce Date
Private Const ciXYour = 2000 'YOUR REF
Private Const ciXDBnote = 5200 'YR DB NOTE
Private Const ciXCurr = 8400 '幣別
Private Const ciXAmt = 10300 '金額:右邊界
Private Const ciXA = 8700 'AMOUNT抬頭
Private Const ciDBwidth = 22 'YOUR REF / YR DB NOTE 的字數
Dim bolPdfStart As Boolean 'Added by Lydia 2023/11/15 是否開啟PDFCreator物件
'Added by Lydia 2024/12/11 Excel列印
Dim xRows As Integer, xRowE As Integer 'Excel列印資料的起始,終止位置
Dim xCols As Integer 'Excel列印資料的起始欄位Ascii值
Dim nRow As Integer '目前資料列位置
Private Const maxRows As Integer = 43 '頁面最大列數
Dim bolColTitle As Boolean '是否列印欄位抬頭
Dim strPrtPath As String, strPrtFile As String '列印Excel檔案路徑,名稱
Dim xlsRpt As New Excel.Application
Dim WksRpt1 As New Worksheet
Dim oShape
Dim oShape2

Private Sub cmdPath_Click()
   Dim fName As String, strStartFolder As String
   
   If Dir(txtPath & "\", vbDirectory) <> "" Then strStartFolder = txtPath
   
   fName = PUB_GetFolder(Me.hWnd, strStartFolder, "請選取資料夾:")
   If fName <> "" Then 'they did not hit cancel
      txtPath = fName
      SaveSetting "TAIE", "A", UCase(Me.Name) & "Dir", txtPath
   End If
End Sub

Private Sub Command2_Click()
   intTop = 0 'Added by Morgan 2012/3/14
   intBottom = 0 'Added by Morgan 2012/3/14
   
   Command2.Enabled = False    'add by sonia 2017/4/12 為防止列印二份
   If FormCheck = False Then
      Command2.Enabled = True  'add by sonia 2017/4/12
      Exit Sub
   End If
   
   
   strConSql = MsgText(601)
   If MaskEdBox1.Text <> MsgText(29) Then
      strConSql = strConSql & " and a1b03 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   'Added by Morgan 2024/10/30
   Else
      MsgBox "請輸入結匯日期起日！", vbExclamation
      Exit Sub
   'end 2024/10/30
   End If
   
   If MaskEdBox2.Text <> MsgText(29) Then
      strConSql = strConSql & " and a1b03 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   'Added by Morgan 2024/10/30
   Else
      MsgBox "請輸入結匯日期止日！", vbExclamation
      Exit Sub
   'end 2024/10/30
   End If
   
   If Text1 <> MsgText(601) Then
      strConSql = strConSql & " and a1b02 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strConSql = strConSql & " and a1b02 <= '" & Text2 & "'"
   End If
   
   Screen.MousePointer = vbHourglass
   
   'bolPdfStart = False 'Added by Lydia 2023/11/15 'Mark by Lydia 2024/12/11 改用EXCEL
   
'Added by Morgan 2018/1/19
'帳單/抵帳單電子檔匯出
If Text5 = "3" Then
   If txtPath = "" Then
      MsgBox "請點選匯出路徑！", vbExclamation
      GoTo EXITSUB 'Added by Lydia 2018/11/21
   ElseIf Dir(txtPath, vbDirectory) = "" Then
      MsgBox "匯出路徑已不存在，請重新點選！", vbExclamation
      GoTo EXITSUB 'Added by Lydia 2018/11/21
   ElseIf ExportFile() = True Then
      MsgBox "匯出完成！", vbExclamation
   End If
Else
'end 2018/1/19

   'Added by Lydia 2018/11/19 特殊結匯帳單匯出路徑
   If Text5 = "1" Then
        If txtPath2 = "" Then
           MsgBox "請點選特殊結匯帳單匯出路徑！", vbExclamation
           GoTo EXITSUB 'Modified by Lydia 2018/11/21 原本exit sub
        ElseIf Dir(txtPath2, vbDirectory) = "" Then
           MsgBox "特殊結匯帳單匯出路徑已不存在，請重新點選！", vbExclamation
           GoTo EXITSUB 'Modified by Lydia 2018/11/21 原本exit sub
        End If
   End If
   'end 2018/11/19
   
   'Add by Amy 2016/03/25 舜禹 Y52268 及 捷恩凱 Y53541 ,不需列印歸檔故排除-婉莘
   'modify by sonia 2017/10/12 再+Y54868迅達翻譯社
   'Mark by Lydia 2025/01/14 105年的規則已不適用---婉莘
   'If Text5 = "1" And Trim(Text1) = MsgText(601) And Trim(Text2) = MsgText(601) Then
   '   'Modified by Lydia 2017/10/16 改成共用變數
   '   'strConSql = strConSql & " and SubStr(a1b02,1,8) Not in('Y5226800','Y5354100','Y5486800')"
   '   strConSql = strConSql & " and SubStr(a1b02,1,8) Not in(" & GetAddStr(Replace(外翻Y編號, "000", "00")) & ")"
   'End If
   'end 2025/01/14
   
   'Add by Amy 2014/05/01 依公司別抓公司抬頭
   If Text3 = "J" Then
        strCompName = A0803Query(Text3)
   Else
        strCompName = A0803Query("2")
   End If
   'end 2014/05/01
   
   PUB_SetOsDefaultPrinter cmbPrinter 'Added by Lydia 2025/03/07 切換Word/Excel印表機
   PUB_RestorePrinter cmbPrinter 'Added by Morgan 2012/10/11
   bolPrintPage = False 'Added by Lydia 2017/10/05
   Select Case Combo2
      Case ComboItem(253)
         'Modify by Morgan 2006/12/14
         'Modified by Lydia 2024/12/11 改用EXCEL
         'PrintData
         PrintExcelMain
         'end 2024/12/11
         'Add by Amy 2015/04/08 結匯完成寄信通知國外部(kelly)
         'Modif by Amy 2017/11/15 改發其他人 '2017/08/11 取消發 mail-婉莘
         If Text5 = "2" Then
            MailToTaie
         End If
         'Added by Lydia 2017/10/05 列印清單
         'Remove by Lydia 2017/11/1 改成先印清單
         'PrintList
      Case ComboItem(252)
         ProcessDetail
         PrintData1
   End Select
   'Remove by Lydia 2016/08/03
  ' PUB_RestorePrinter strPrinter 'Added by Morgan 2012/10/11
  
End If 'Added by Morgan 2018/1/19

   FormClear 'Move by Lydia 2018/11/21 從Command2下方移上來
   
EXITSUB: 'Added by Lydia 2018/11/21
   Command2.Enabled = True  'add by sonia 2017/4/12
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   End If
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, Me.Width, Me.Height, strBackPicPath4
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   Combo2.Clear
   Combo2.AddItem ComboItem(253), 0
   Combo2.AddItem ComboItem(252), 1
   Combo2.ListIndex = 0
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter 'Added by Morgan 2012/10/11
   PUB_SetPrinter Me.Name, cmbPrinter2, strPrinter2 'Added by Morgan 2012/10/11
   pub_OsPrinter = strPrinter
   strPrtOrt = Printer.Orientation 'Added by Lydia 2017/10/05 預設印表機列印方向
   
   'Add by Amy 2014/05/01 +公司別
   Text3 = "1"
   Text13 = A0802Query(Text3)
   'end 2014/05/01

   txtPath = GetSetting("TAIE", "A", UCase(Me.Name) & "Dir", "") 'Added by Morgan 2018/1/23
   txtPath2 = GetSetting("TAIE", "B", UCase(Me.Name) & "Dir", "") 'Added by Lydia 2018/11/19
   
   'Added by Lydia 2024/12/11
   strPrtPath = App.path & "\" & strUserNum
   Call Pub_ChkExcelPath(strPrtPath)
   Call PUB_KillTempFile(strUserNum & "\$*.*")
   'end 2024/12/11

End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   'Added by Morgan 2012/10/11
   '若印表機變動, 則更新列印設定
   If cmbPrinter.Text <> cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   If cmbPrinter2.Text <> cmbPrinter2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter2.Name, "0", "0", Me.cmbPrinter2.Text
   End If
   PUB_SetOsDefaultPrinter pub_OsPrinter 'Added by Lydia 2025/03/07 切換Word/Excel印表機
   'Added by Lydia 2016/08/03 還原作業系統的預設印表機
   'Modified by Lydia 2017/10/03 +strPrtOrt
   PUB_RestorePrinter pub_OsPrinter, strPrtOrt
   
   Set Frmacc2450 = Nothing
End Sub

'add by sonia 2015/5/19
Private Sub MaskEdBox1_LostFocus()
   If MaskEdBox1.Text <> "___/__/__" And (MaskEdBox2.Text = "___/__/__" Or MaskEdBox2.Text = "") Then
      MaskEdBox2 = MaskEdBox1
   End If
End Sub
'2015/5/19 end

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Len(Text1) = 6 Then
      Text1 = AfterZero(Text1)
   End If
   '2009/6/2 ADD BY SONIA 預設尾碼999
   'Modify By Sindy 2014/8/11 999=>ZZZ
   'If Text1.Text <> "" Then Text2.Text = Left(Me.Text1.Text, 6) & "999"
   If Text1.Text <> "" Then Text2.Text = Left(Me.Text1.Text, 6) & "ZZZ"
End Sub

Private Sub Text2_GotFocus()
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
   Text5 = ""
   Text6 = ""
   Text3.SetFocus 'Modify by Amy 2014/04/22 原:Text5.SetFocus
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead()
   Dim intRow As Integer
   Dim StrSqlB As String
   Dim rsA As New ADODB.Recordset
   'Add by Amy 2018/10/31
   Dim strFA17 As String, strFA18 As String, strFA19 As String, strFA20 As String, strFA21 As String, strFA22 As String, strFA70 As String, strFA23 As String
   Dim strFA32 As String, strFA33 As String, strFA34 As String, strFA35 As String, strFA36 As String
   
      
   'Add by Morgan 2004/11/24
   'Modify by Morgan 2007/1/4 第一頁才要清
   If intPage = 0 Then
      Erase m_Cur
      Erase m_Amount
      m_Curs = ""
      m_Len = 0
   End If
   'end 2007/1/4
   
   'Modify by Morgan 2007/3/21 大陸地區改用中文--婧瑄
   'strLanguage = "2" 'Add by Morgan 2006/2/8 定稿語文固定用英文--婧瑄(避免上下語文不統一)
   If "" & adoquery("fa10") = "020" Then
      strLanguage = "1"
   Else
      strLanguage = "2"
   End If
   'end 2007/3/21
   
   bolPrintPage = True 'Added by Lydia 2017/10/05
   
   StrSqlB = " Select * From Fagent, (Select A2213 From ACC220, Fagent Where substr(A2201,1,8)=FA01 And substr(A2201,9,1)=FA02 And A2201='" & adoquery.Fields("A1b02").Value & "') A Where FA01=substr(A.A2213,1,8) And FA02=substr(A.A2213,9,1) "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
   'Modified by Lydia 2018/09/03 + bol2pdf
   If bol2Printer = True Or bol2Pdf = True Then
      Printer.FontSize = 12
      Printer.Font = "Times New Roman"
      Printer.CurrentX = 8700 + intMargin
      Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
      'Modify by Morgan 2010/6/7
      'Printer.Print Format(AFDate(strSrvDate(1)), "mmm. d, yyyy")
      Printer.Print Format(AFDate(m_strPayDate), "mmm. d, yyyy")
   End If
   If bol2File = True Then
      Picture1.FontSize = 12 * douExtRate
      Picture1.Font = "Times New Roman"
      Picture1.AutoRedraw = True
      Picture1.CurrentX = (8700 + intMargin) * douExtRate
      Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
      'Modify by Morgan 2010/6/7
      'Picture1.Print Format(AFDate(strSrvDate(1)), "mmm. d, yyyy")
      Picture1.Print Format(AFDate(m_strPayDate), "mmm. d, yyyy")
   End If
   
   intRow = intRow + 1
   'Add by Amy 2018/10/31 +地址有「竹曆退件」字樣不顯示地址
   If rsA.RecordCount > 0 Then
        strFA17 = "" & rsA.Fields("fa17").Value
        strFA18 = "" & rsA.Fields("fa18").Value: strFA19 = "" & rsA.Fields("fa19").Value: strFA20 = "" & rsA.Fields("fa20").Value
        strFA21 = "" & rsA.Fields("fa21").Value: strFA22 = "" & rsA.Fields("fa22").Value: strFA70 = "" & rsA.Fields("fa70").Value
        strFA23 = "" & rsA.Fields("fa23").Value
        strFA32 = "" & rsA.Fields("fa32").Value: strFA33 = "" & rsA.Fields("fa33").Value: strFA34 = "" & rsA.Fields("fa34").Value
        strFA35 = "" & rsA.Fields("fa35").Value: strFA36 = "" & rsA.Fields("fa36").Value
  '若未建付款明細列印對象
  Else
        strFA17 = "" & adoquery.Fields("fa17").Value
        strFA18 = "" & adoquery.Fields("fa18").Value: strFA19 = "" & adoquery.Fields("fa19").Value: strFA20 = "" & adoquery.Fields("fa20").Value
        strFA21 = "" & adoquery.Fields("fa21").Value: strFA22 = "" & adoquery.Fields("fa22").Value: strFA70 = "" & adoquery.Fields("fa70").Value
        strFA23 = "" & adoquery.Fields("fa23").Value
        strFA32 = "" & adoquery.Fields("fa32").Value: strFA33 = "" & adoquery.Fields("fa33").Value: strFA34 = "" & adoquery.Fields("fa34").Value
        strFA35 = "" & adoquery.Fields("fa35").Value: strFA36 = "" & adoquery.Fields("fa36").Value
  End If
  If InStr(strFA17, "竹曆退件") > 0 Then strFA17 = ""
  If InStr(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70, "竹曆退件") > 0 Then
        strFA18 = "": strFA19 = "": strFA20 = "": strFA21 = "": strFA22 = "": strFA70 = ""
  End If
  If InStr(strFA23, "竹曆退件") > 0 Then strFA23 = ""
  If InStr(strFA32 & strFA33 & strFA34 & strFA35 & strFA36, "竹曆退件") > 0 Then
        strFA32 = "": strFA33 = "": strFA34 = "": strFA35 = "": strFA36 = ""
  End If
   'end 2018/10/31
      
   '選擇語言
   Select Case strLanguage
      'Add by Morgan 2005/4/27
      Case "1" '中文
         '若有建付款明細列印對象
         If rsA.RecordCount > 0 Then
            '中文
            If IsNull(rsA.Fields("fa04").Value) = False Then
               If intPage = 0 Then
                  XPrint "" & rsA.Fields("fa04").Value, intRow
               End If
            '英文
            ElseIf IsNull(rsA.Fields("fa05").Value) = False Then
               If intPage = 0 Then
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print "" & rsA.Fields("fa05").Value
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print "" & rsA.Fields("fa05").Value
                  End If
               End If
               If IsNull(rsA.Fields("fa63").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa63").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa63").Value
                     End If
                  End If
               End If
               If IsNull(rsA.Fields("fa64").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa64").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa64").Value
                     End If
                  End If
               End If
               If IsNull(rsA.Fields("fa65").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa65").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa65").Value
                     End If
                  End If
               End If
               
            '日文
            ElseIf IsNull(rsA.Fields("fa06").Value) = False Then
               If intPage = 0 Then
                  XPrint "" & rsA.Fields("fa06").Value, intRow
               End If
            End If
         '若未建付款明細列印對象
         Else
            Set rsA = adoquery.Clone
            rsA.AbsolutePosition = adoquery.AbsolutePosition
            
            '中文名稱
            If "" & rsA.Fields("fa04").Value <> "" Then
               If intPage = 0 Then
                  XPrint "" & rsA.Fields("fa04").Value, intRow
               End If
            '英文名稱
            ElseIf IsNull(rsA.Fields("fa05").Value) = False Then
               If intPage = 0 Then
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print "" & rsA.Fields("fa05").Value
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print "" & rsA.Fields("fa05").Value
                  End If
               End If
               If IsNull(rsA.Fields("fa63").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa63").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa63").Value
                     End If
                  End If
               End If
               If IsNull(rsA.Fields("fa64").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa64").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa64").Value
                     End If
                  End If
               End If
               If IsNull(rsA.Fields("fa65").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa65").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa65").Value
                     End If
                  End If
               End If

            '日文名稱
            ElseIf IsNull(rsA.Fields("fa06").Value) = False Then
               If intPage = 0 Then
                  XPrint "" & rsA.Fields("fa06").Value, intRow
               End If
            End If
         End If
         intRow = intRow + 1
         If intPage = 0 Then
            '地址,順序：中文->POB->英文->日文
            '中文地址
            'Modify by Amy 2018/10/31 地址改為變數判斷 原:IsNull(rsA.Fields("faXX").Value) = False
            If strFA17 <> MsgText(601) Then
               intRow = intRow + 1
               XPrint strFA17, intRow
            'POB
            ElseIf strFA32 <> MsgText(601) Then
               'P0B1
               'Modified by Lydia 2018/09/03 + bol2pdf
               If bol2Printer = True Or bol2Pdf = True Then
                  Printer.CurrentX = 0 + intMargin
                  Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                  Printer.Print strFA32
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (0 + intMargin) * douExtRate
                  Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                  Picture1.Print strFA32
               End If
               'POB2
               If strFA33 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA33
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA33
                  End If
               End If
               'POB3
               If strFA34 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA34
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA34
                  End If
               End If
               'POB4
               If strFA35 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA35
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA35
                  End If
               End If
               'POB5
               If strFA36 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA36
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA36
                  End If
               End If
            
            '英文地址
            ElseIf strFA18 <> MsgText(601) Then
               '英文地址1
               'Modified by Lydia 2018/09/03 + bol2pdf
               If bol2Printer = True Or bol2Pdf = True Then
                  Printer.CurrentX = 0 + intMargin
                  Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                  Printer.Print strFA18
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (0 + intMargin) * douExtRate
                  Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                  Picture1.Print strFA18
               End If
               '英文地址2
               If strFA19 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA19
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA19
                  End If
               End If
               '英文地址3
               If strFA20 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA20
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA20
                  End If
               End If
               '英文地址4
               If strFA21 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA21
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA21
                  End If
               End If
               '英文地址5
               If strFA22 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA22
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA22
                  End If
               End If
               
               'Add by Morgan 2011/5/25
               '英文地址6
               If strFA70 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA70
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA70
                  End If
               End If
               
            '日文地址
            ElseIf strFA23 <> MsgText(601) Then
               intRow = intRow + 1
               XPrint strFA23, intRow
            End If
            'end 2018/10/31
         End If
         If intRow < 6 Then intRow = 6
      Case "2"
        'Modify By Cheng 2003/08/18
        '若有建付款明細列印對象
        'Modify by Amy 2018/10/31 地址改為變數判斷
        If rsA.RecordCount > 0 Then
            If IsNull(rsA.Fields("fa05").Value) = False Then
               If intPage = 0 Then
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print "" & rsA.Fields("fa05").Value
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print "" & rsA.Fields("fa05").Value
                  End If
               End If
            End If
            If IsNull(rsA.Fields("fa63").Value) = False Then
               If intPage = 0 Then
                  intRow = intRow + 1
                  intCounter = intCounter + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print "" & rsA.Fields("fa63").Value
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin)
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print "" & rsA.Fields("fa63").Value
                  End If
               End If
            End If
            If IsNull(rsA.Fields("fa64").Value) = False Then
               If intPage = 0 Then
                  intRow = intRow + 1
                  intCounter = intCounter + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print "" & rsA.Fields("fa64").Value
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print "" & rsA.Fields("fa64").Value
                  End If
               End If
            End If
            If IsNull(rsA.Fields("fa65").Value) = False Then
               If intPage = 0 Then
                  intRow = intRow + 1
                  intCounter = intCounter + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print "" & rsA.Fields("fa65").Value
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print "" & rsA.Fields("fa65").Value
                  End If
               End If
            End If

            intRow = intRow + 1
            '地址
            If intPage = 0 Then
               'If IsNull(rsA.Fields("fa32").Value) Then
               If strFA32 = MsgText(601) Then
                  If strFA18 <> MsgText(601) Then
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA18
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA18
                     End If
                  'Add By Cheng 2003/03/26
                  '若無英文地址時,  印中文地址
                  Else
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA17
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA17
                     End If
                  End If
               Else
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA32
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA32
                  End If
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               'If IsNull(rsA.Fields("fa32").Value) Then
               'Modified by Lydia 2024/12/06 strFA32->strFA33
               If strFA33 = MsgText(601) Then
                  If strFA19 <> MsgText(601) Then
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA19
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA19
                     End If
                  End If
               Else
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA33
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA33
                  End If
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               'If IsNull(rsA.Fields("fa32").Value) Then
               'Modified by Lydia 2024/12/06 strFA32->strFA34
               If strFA34 = MsgText(601) Then
                  If strFA20 <> MsgText(601) Then
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA20
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA20
                     End If
                  End If
               Else
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA34
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA34
                  End If
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               'If IsNull(rsA.Fields("fa32").Value) Then
               'Modified by Lydia 2024/12/06 strFA32->strFA35
               If strFA35 = MsgText(601) Then
                  If strFA21 <> MsgText(601) Then
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA21
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA21
                     End If
                  End If
               Else
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA35
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA35
                  End If
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               'If IsNull(rsA.Fields("fa32").Value) Then
               'Modified by Lydia 2024/12/06 strFA32->strFA36
               If strFA36 = MsgText(601) Then
                  If strFA22 <> MsgText(601) Then
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA22
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA22
                     End If
                  End If
                  
                  'Add by Morgan 2011/5/25
                  '英文地址6
                  If strFA70 <> MsgText(601) Then
                     intRow = intRow + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA70
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA70
                     End If
                  End If
                  
               Else
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA36
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA36
                  End If
               End If
            End If
            'end 2018/10/31
        '若未建付款明細列印對象
        Else
            '英文名稱
            If IsNull(adoquery.Fields("fa05").Value) = False Then
               If intPage = 0 Then
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print "" & adoquery.Fields("fa05").Value
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print "" & adoquery.Fields("fa05").Value
                  End If
               End If
                If IsNull(adoquery.Fields("fa63").Value) = False Then
                   If intPage = 0 Then
                      intRow = intRow + 1
                      intCounter = intCounter + 1
                      'Modified by Lydia 2018/09/03 + bol2pdf
                      If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & adoquery.Fields("fa63").Value
                      End If
                      If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & adoquery.Fields("fa63").Value
                      End If
                   End If
                End If
                If IsNull(adoquery.Fields("fa64").Value) = False Then
                   If intPage = 0 Then
                      intRow = intRow + 1
                      intCounter = intCounter + 1
                      'Modified by Lydia 2018/09/03 + bol2pdf
                      If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & adoquery.Fields("fa64").Value
                      End If
                      If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & adoquery.Fields("fa64").Value
                      End If
                   End If
                End If
                If IsNull(adoquery.Fields("fa65").Value) = False Then
                   If intPage = 0 Then
                      intRow = intRow + 1
                      intCounter = intCounter + 1
                      'Modified by Lydia 2018/09/03 + bol2pdf
                      If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & adoquery.Fields("fa65").Value
                      End If
                      If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & adoquery.Fields("fa65").Value
                      End If
                   End If
                End If
               
            '中文名稱
            ElseIf "" & adoquery.Fields("fa04").Value <> "" Then
               If intPage = 0 Then
                  XPrint "" & adoquery.Fields("fa04").Value, intRow
               End If
            '日文名稱
            ElseIf "" & adoquery.Fields("fa06").Value <> "" Then
               If intPage = 0 Then
                  XPrint "" & adoquery.Fields("fa06").Value, intRow
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               'If IsNull(adoquery.Fields("fa32").Value) Then
               If strFA32 = MsgText(601) Then
                  If strFA18 <> MsgText(601) Then
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA18
                     End If
                     'Add by Morgan 2008/2/15
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA18
                     End If
                  '若無英文地址時,  印中文地址
                  Else
                     XPrint strFA17, intRow
                  End If
               Else
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA32
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA32
                  End If
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               'If IsNull(adoquery.Fields("fa32").Value) Then
               If strFA32 = MsgText(601) Then
                  If strFA19 <> MsgText(601) Then
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA19
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA19
                     End If
                  End If
               Else
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA33
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA33
                  End If
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               'If IsNull(adoquery.Fields("fa32").Value) Then
               If strFA32 = MsgText(601) Then
                  If strFA20 <> MsgText(601) Then
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA20
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA20
                     End If
                  End If
               Else
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA34
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA34
                  End If
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               'If IsNull(adoquery.Fields("fa32").Value) Then
               If strFA32 = MsgText(601) Then
                  If strFA21 <> MsgText(601) Then
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA21
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA21
                     End If
                  End If
               Else
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA35
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA35
                  End If
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               'If IsNull(adoquery.Fields("fa32").Value) Then
               If strFA32 = MsgText(601) Then
                  If strFA22 <> MsgText(601) Then
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA22
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA22
                     End If
                  End If
                  
                  'Add by Morgan 2011/5/25
                  '英文地址6
                  If strFA70 <> MsgText(601) Then
                     intRow = intRow + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print strFA70
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print strFA70
                     End If
                  End If
               Else
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA36
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA36
                  End If
               End If
            End If
        End If
        'end 2018/10/31
        
      'Modify by Morgan 2006/1/4 改寫
      Case "3" '日文
      
        '若有建付款明細列印對象
         'Modify by Amy 2018/10/31 地址改為變數判斷
         If rsA.RecordCount > 0 Then
            '日文
            If IsNull(rsA.Fields("fa06").Value) = False Then
               If intPage = 0 Then
                  XPrint "" & rsA.Fields("fa06").Value, intRow
               End If
            '英文
            ElseIf IsNull(rsA.Fields("fa05").Value) = False Then
               If intPage = 0 Then
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print "" & rsA.Fields("fa05").Value
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print "" & rsA.Fields("fa05").Value
                  End If
               End If
               If IsNull(rsA.Fields("fa63").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa63").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa63").Value
                     End If
                  End If
               End If
               If IsNull(rsA.Fields("fa64").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa64").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa64").Value
                     End If
                  End If
               End If
               If IsNull(rsA.Fields("fa65").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa65").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa65").Value
                     End If
                  End If
               End If
            '中文
            ElseIf IsNull(rsA.Fields("fa04").Value) = False Then
               If intPage = 0 Then
                  XPrint "" & rsA.Fields("fa04").Value, intRow
               End If
            End If
         '若未建付款明細列印對象
         Else
            Set rsA = adoquery.Clone
            rsA.AbsolutePosition = adoquery.AbsolutePosition
            
            '日文名稱
            If IsNull(rsA.Fields("fa06").Value) = False Then
               If intPage = 0 Then
                  XPrint "" & rsA.Fields("fa06").Value, intRow
               End If
            '英文名稱
            ElseIf IsNull(rsA.Fields("fa05").Value) = False Then
               If intPage = 0 Then
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print "" & rsA.Fields("fa05").Value
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print "" & rsA.Fields("fa05").Value
                  End If
               End If
               If IsNull(rsA.Fields("fa63").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa63").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa63").Value
                     End If
                  End If
               End If
               If IsNull(rsA.Fields("fa64").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa64").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa64").Value
                     End If
                  End If
               End If
               If IsNull(rsA.Fields("fa65").Value) = False Then
                  If intPage = 0 Then
                     intRow = intRow + 1
                     intCounter = intCounter + 1
                     'Modified by Lydia 2018/09/03 + bol2pdf
                     If bol2Printer = True Or bol2Pdf = True Then
                        Printer.CurrentX = 0 + intMargin
                        Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                        Printer.Print "" & rsA.Fields("fa65").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (0 + intMargin) * douExtRate
                        Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                        Picture1.Print "" & rsA.Fields("fa65").Value
                     End If
                  End If
               End If

            '中文名稱
            ElseIf "" & rsA.Fields("fa04").Value <> "" Then
               If intPage = 0 Then
                  XPrint "" & rsA.Fields("fa04").Value, intRow
               End If
            
            End If
         End If
         intRow = intRow + 1
         If intPage = 0 Then
            '地址,順序：日文->POB->英文->中文
            '日文地址
            If strFA23 <> MsgText(601) Then
               intRow = intRow + 1
               XPrint strFA23, intRow
            'POB
            ElseIf strFA32 <> MsgText(601) Then
               'P0B1
               'Modified by Lydia 2018/09/03 + bol2pdf
               If bol2Printer = True Or bol2Pdf = True Then
                  Printer.CurrentX = 0 + intMargin
                  Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                  Printer.Print strFA32
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (0 + intMargin) * douExtRate
                  Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                  Picture1.Print strFA32
               End If
               'POB2
               If strFA33 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA33
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA33
                  End If
               End If
               'POB3
               If strFA34 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA34
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA34
                  End If
               End If
               'POB4
               If strFA35 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA35
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA35
                  End If
               End If
               'POB5
               If strFA36 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA36
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA36
                  End If
               End If
            '英文地址
            ElseIf strFA18 <> MsgText(601) Then
               '英文地址1
               'Modified by Lydia 2018/09/03 + bol2pdf
               If bol2Printer = True Or bol2Pdf = True Then
                  Printer.CurrentX = 0 + intMargin
                  Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                  Printer.Print strFA18
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (0 + intMargin) * douExtRate
                  Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                  Picture1.Print strFA18
               End If
               '英文地址2
               If strFA19 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA19
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA19
                  End If
               End If
               '英文地址3
               If strFA20 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA20
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA20
                  End If
               End If
               '英文地址4
               If strFA21 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA21
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA21
                  End If
               End If
               '英文地址5
               If strFA22 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA22
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA22
                  End If
               End If
               
               'Add by Morgan 2011/5/25
               '英文地址6
               If strFA70 <> MsgText(601) Then
                  intRow = intRow + 1
                  'Modified by Lydia 2018/09/03 + bol2pdf
                  If bol2Printer = True Or bol2Pdf = True Then
                     Printer.CurrentX = 0 + intMargin
                     Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 250
                     Printer.Print strFA70
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (0 + intMargin) * douExtRate
                     Picture1.CurrentY = (intTop + 2500 + intRow * 250) * douExtRate
                     Picture1.Print strFA70
                  End If
               End If
               
            '中文地址
            ElseIf strFA17 <> MsgText(601) Then
               intRow = intRow + 1
               XPrint strFA17, intRow
            End If
         End If
         'end 2018/10/31
         If intRow < 6 Then intRow = 6
   End Select
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
   intRow = intRow + 2
   m_a1b06 = adoquery.Fields("a1b06").Value
   Select Case adoquery.Fields("a1b06").Value
      Case "1" '票匯
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            Printer.FontSize = 15
         End If
         If bol2File = True Then
            Picture1.FontSize = 15 * douExtRate
         End If
         If intPage = 0 Then
            'Modified by Lydia 2018/09/03 + bol2pdf
            If bol2Printer = True Or bol2Pdf = True Then
               Printer.CurrentX = 0 + intMargin
               Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
               Printer.Print ReportSum(87)
            End If
            If bol2File = True Then
               Picture1.CurrentX = (0 + intMargin) * douExtRate
               Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
               Picture1.Print ReportSum(87)
            End If
         End If
         intRow = intRow + 1
         If intPage = 0 Then
            'Modified by Lydia 2018/09/03 + bol2pdf
            If bol2Printer = True Or bol2Pdf = True Then
               Printer.CurrentX = 500 + intMargin
               Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
               Printer.Print ReportSum(88)
            End If
            If bol2File = True Then
               Picture1.CurrentX = (500 + intMargin) * douExtRate
               Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
               Picture1.Print ReportSum(88)
            End If
         End If
         intRow = intRow + 1
         If intPage = 0 Then
            'Modified by Lydia 2018/09/03 + bol2pdf
            If bol2Printer = True Or bol2Pdf = True Then
               Printer.CurrentX = 0 + intMargin
               Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
               Printer.Print ReportSum(89)
            End If
            If bol2File = True Then
               Picture1.CurrentX = (0 + intMargin) * douExtRate
               Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
               Picture1.Print ReportSum(89)
            End If
         End If
         'Added by Lydia 2020/09/03 列印DRAFT NO
         If intPage = 0 Then
             intRow = intRow + 2
             intCounter = intCounter + 2
             If bol2Printer = True Or bol2Pdf = True Then
                Printer.CurrentX = ciXYour + intMargin
                Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
                Printer.Print "DRAFT NO. " & adoquery.Fields("a1b01").Value
             End If
             If bol2File = True Then
                Picture1.CurrentX = (ciXYour + intMargin) * douExtRate
                Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
                Picture1.Print "DRAFT NO. " & adoquery.Fields("a1b01").Value
             End If
             intRow = intRow + 1
         End If
         'end 2020/09/03
         
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            Printer.FontSize = 12
         End If
         If bol2File = True Then
            Picture1.FontSize = 12 * douExtRate
         End If
         If intPage >= 1 Then
            intRow = 0
         End If
         intRow = intRow + 2
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            'Modified by Lydia 2020/09/03 調整明細表的版面: 欄位的起始位置,拿掉DRAFT NO.
'            Printer.CurrentX = 0 + intMargin
'            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
'            Printer.Print "OUR REF"
'            Printer.CurrentX = 2000 + intMargin
'            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
'            Printer.Print "YOUR REF"
'            Printer.CurrentX = 4500 + intMargin
'            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
'            Printer.Print "YR DB NOTE"
'            Printer.CurrentX = 7000 + intMargin
'            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
'            Printer.Print "AMOUNT"
'            Printer.CurrentX = 8700 + intMargin
'            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
'            Printer.Print "DRAFT NO."
            Printer.CurrentX = ciXOur + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
            Printer.Print "OUR REF"
            Printer.CurrentX = ciXYour + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
            Printer.Print "YOUR REF"
            Printer.CurrentX = ciXDBnote + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
            Printer.Print "YR DB NOTE"
            Printer.CurrentX = ciXA + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
            Printer.Print "AMOUNT"
            'end 2020/09/03
         End If
         If bol2File = True Then
            'Modified by Lydia 2020/09/03 調整明細表的版面: 欄位的起始位置,拿掉DRAFT NO.
'            Picture1.CurrentX = (0 + intMargin) * douExtRate
'            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
'            Picture1.Print "OUR REF"
'            Picture1.CurrentX = (2000 + intMargin) * douExtRate
'            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
'            Picture1.Print "YOUR REF"
'            Picture1.CurrentX = (4500 + intMargin) * douExtRate
'            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
'            Picture1.Print "YR DB NOTE"
'            Picture1.CurrentX = (7000 + intMargin) * douExtRate
'            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
'            Picture1.Print "AMOUNT"
'            Picture1.CurrentX = (8700 + intMargin) * douExtRate
'            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
'            Picture1.Print "DRAFT NO."
            Picture1.CurrentX = (ciXOur + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
            Picture1.Print "OUR REF"
            Picture1.CurrentX = (ciXYour + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
            Picture1.Print "YOUR REF"
            Picture1.CurrentX = (ciXDBnote + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
            Picture1.Print "YR DB NOTE"
            Picture1.CurrentX = (ciXA + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
            Picture1.Print "AMOUNT"
            'end 2020/09/03
         End If
      Case Else '2 電匯
      
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            Printer.FontSize = 15
         End If
         If bol2File = True Then
            Picture1.FontSize = 15 * douExtRate
         End If
         If intPage = 0 Then
            'Modified by Lydia 2018/09/03 + bol2pdf
            If bol2Printer = True Or bol2Pdf = True Then
               Printer.CurrentX = 0 + intMargin
               Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
               Printer.Print ReportSum(87)
            End If
            If bol2File = True Then
               Picture1.CurrentX = (0 + intMargin) * douExtRate
               Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
               Picture1.Print ReportSum(87)
            End If
         End If
        'Add By Cheng 2003/09/02
        StrSqlB = "Select * From ACC220 Where A2201='" & adoquery("a1b02").Value & "' And A2202='" & adoquery("a1505").Value & "' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            intRow = intRow + 1
            If intPage = 0 Then
               'Modified by Lydia 2018/09/03 + bol2pdf
               If bol2Printer = True Or bol2Pdf = True Then
                  Printer.CurrentX = 500 + intMargin
                  Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
                  Printer.Print ReportSum(97001)
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (500 + intMargin) * douExtRate
                  Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
                  Picture1.Print ReportSum(97001)
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               'Modified by Lydia 2018/09/03 + bol2pdf
               If bol2Printer = True Or bol2Pdf = True Then
                  Printer.CurrentX = 0 + intMargin
                  Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
                  Printer.Print ReportSum(98001)
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (0 + intMargin) * douExtRate
                  Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
                  Picture1.Print ReportSum(98001)
               End If
            End If
            intRow = intRow + 2
            If intPage = 0 Then
               intCounter = intCounter + 2
               'Modified by Lydia 2018/09/03 + bol2pdf
               If bol2Printer = True Or bol2Pdf = True Then
                  Printer.CurrentX = 2000 + intMargin
                  Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
                  Printer.Print "" & rsA("a2208").Value
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (2000 + intMargin) * douExtRate
                  Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
                  Picture1.Print "" & rsA("a2208").Value
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               intCounter = intCounter + 1
               'Modified by Lydia 2018/09/03 + bol2pdf
               If bol2Printer = True Or bol2Pdf = True Then
                  Printer.CurrentX = 2000 + intMargin
                  Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
                  Printer.Print "Account #" & rsA("a2207").Value
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (2000 + intMargin) * douExtRate
                  Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
                  Picture1.Print "Account #" & rsA("a2207").Value
               End If
            End If
        Else
            intRow = intRow + 1
            If intPage = 0 Then
               'Modified by Lydia 2018/09/03 + bol2pdf
               If bol2Printer = True Or bol2Pdf = True Then
                  Printer.CurrentX = 500 + intMargin
                  Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
                  Printer.Print ReportSum(97)
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (500 + intMargin) * douExtRate
                  Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
                  Picture1.Print ReportSum(97)
               End If
            End If
            intRow = intRow + 1
            If intPage = 0 Then
               'Modified by Lydia 2018/09/03 + bol2pdf
               If bol2Printer = True Or bol2Pdf = True Then
                  Printer.CurrentX = 0 + intMargin
                  Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
                  Printer.Print ReportSum(98)
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (0 + intMargin) * douExtRate
                  Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
                  Picture1.Print ReportSum(98)
               End If
            End If
        End If
        
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            Printer.FontSize = 12
         End If
         If bol2File = True Then
            Picture1.FontSize = 12 * douExtRate
         End If
         If intPage >= 1 Then
            intRow = 0
         End If
         intRow = intRow + 2
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            'Modified by Lydia 2020/09/03 調整明細表的版面: 欄位的起始位置
'            Printer.CurrentX = 0 + intMargin
'            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
'            Printer.Print "OUR REF"
'            Printer.CurrentX = 2000 + intMargin
'            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
'            Printer.Print "YOUR REF"
'            Printer.CurrentX = 4500 + intMargin
'            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
'            Printer.Print "YR DB NOTE"
'            Printer.CurrentX = 7000 + intMargin
'            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
'            Printer.Print "AMOUNT"
            Printer.CurrentX = ciXOur + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
            Printer.Print "OUR REF"
            Printer.CurrentX = ciXYour + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
            Printer.Print "YOUR REF"
            Printer.CurrentX = ciXDBnote + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
            Printer.Print "YR DB NOTE"
            Printer.CurrentX = ciXA + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 2500 + intRow * 300
            Printer.Print "AMOUNT"
            'end 2020/09/03
         End If
         If bol2File = True Then
            'Modified by Lydia 2020/09/03 調整明細表的版面: 欄位的起始位置
'            Picture1.CurrentX = (0 + intMargin) * douExtRate
'            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
'            Picture1.Print "OUR REF"
'            Picture1.CurrentX = (2000 + intMargin) * douExtRate
'            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
'            Picture1.Print "YOUR REF"
'            Picture1.CurrentX = (4500 + intMargin) * douExtRate
'            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
'            Picture1.Print "YR DB NOTE"
'            Picture1.CurrentX = (7000 + intMargin) * douExtRate
'            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
'            Picture1.Print "AMOUNT"
            Picture1.CurrentX = (ciXOur + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
            Picture1.Print "OUR REF"
            Picture1.CurrentX = (ciXYour + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
            Picture1.Print "YOUR REF"
            Picture1.CurrentX = (ciXDBnote + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
            Picture1.Print "YR DB NOTE"
            Picture1.CurrentX = (ciXA + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 2500 + intRow * 300) * douExtRate
            Picture1.Print "AMOUNT"
            'end 2020/09/03
         End If
   End Select
   'Modified by Lydia 2018/09/03 + bol2pdf
   If bol2Printer = True Or bol2Pdf = True Then
      Printer.Line (0 + intMargin, intTop + intPrtXFix + 2500 + intRow * 300 + 350)-(10400 + intMargin, intTop + intPrtXFix + 2500 + intRow * 300 + 350)
   End If
   If bol2File = True Then
      lngX = (0 + intMargin) * douExtRate
      lngY = (intTop + 2500 + intRow * 300 + 350) * douExtRate
      'Modified by Morgan 2024/11/4
      'Picture1.Line (lngX, lngY)-(Picture1.Width - lngX, lngY)
      Picture1.Line (lngX, lngY)-((ciXAmt + intMargin) * douExtRate, lngY)
      'end 2024/11/4
   End If
   
   'Added by Morgan 2012/5/28 新版信紙列印紙本會壓到信尾改上移1行(欄位名與資料中間還有空間)
   If intPage = 0 Then 'Added by Morgan 2024/11/21 修正第2頁以後的欄位名稱下面的分隔線會壓到第1列資料問題
      intCounter = intCounter - 1
   End If
End Sub

'*************************************************
' 合計
'
'*************************************************
Private Sub PrintSum()
   'Modified by Lydia 2018/09/03 + bol2pdf
   If bol2Printer = True Or bol2Pdf = True Then
      Printer.Line (0 + intMargin, intTop + intPrtXFix + 6700 + intCounter * 300)-(10400 + intMargin, intTop + intPrtXFix + 6700 + intCounter * 300)
   End If
   'Add by Morgan 2006/12/13 加產生EMail用圖檔
   If bol2File = True Then
      lngX = (0 + intMargin) * douExtRate
      lngY = (intTop + 6700 + intCounter * 300) * douExtRate
      'Modified by Morgan 2024/11/4
      'Picture1.Line (lngX, lngY)-(Picture1.Width - lngX, lngY)
      Picture1.Line (lngX, lngY)-((ciXAmt + intMargin) * douExtRate, lngY)
      'end 2024/11/4
   End If
   intCounter = intCounter + 1
   
   If intCounter > 22 Then
      If intCounter = 23 Then
         intCounter = -10
      Else
         intCounter = 1
      End If
      intPage = intPage + 1
      'Modified by Lydia 2018/09/03 + bol2pdf
      If bol2Printer = True Or bol2Pdf = True Then
         PrintPageNo True
         Printer.NewPage
      End If
      If bol2File = True Then
         PrintPageNo
         PicNewPage
      End If
   End If
   'Modified by Lydia 2018/09/03 + bol2pdf
   If bol2Printer = True Or bol2Pdf = True Then
      Printer.CurrentX = 0 + intMargin
      Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
      Printer.Print "TOTAL"
      'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
      'Printer.CurrentX = 6500 + intMargin
      Printer.CurrentX = ciXCurr + intMargin
      Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
      Printer.Print m_Cur(1)
   End If
   If bol2File = True Then
      Picture1.CurrentX = (0 + intMargin) * douExtRate
      Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
      Picture1.Print "TOTAL"
      'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
      'Picture1.CurrentX = (6500 + intMargin) * douExtRate
      Picture1.CurrentX = (ciXCurr + intMargin) * douExtRate
      Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
      Picture1.Print m_Cur(1)
   End If
   strAmount = Format(m_Amount(1), FDollar)
   
   'Modified by Lydia 2018/09/03 + bol2pdf
   If bol2Printer = True Or bol2Pdf = True Then
      intLength = Printer.TextWidth(strAmount)
      'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
      'Printer.CurrentX = 8400 - intLength + intMargin
      Printer.CurrentX = ciXAmt - intLength + intMargin
      Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
      Printer.Print strAmount
      strPayAmount = m_Cur(1) & " " & strAmount 'Added by Lydia 2018/09/06
   End If
   If bol2File = True Then
      intLength = Picture1.TextWidth(strAmount)
      'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
      'Picture1.CurrentX = (8400 + intMargin) * douExtRate - intLength
      'Modified by Morgan 2024/11/4
      'Picture1.CurrentX = (ciXAmt + intMargin) * douExtRate - intLength
      'Modified by Morgan 2024/11/19 修正JPG格式的金額會被幣別蓋到問題
      'Picture1.CurrentX = (ciXAmt - intLength + intMargin) * douExtRate
      intLength = Picture1.TextWidth("9,999,999.00") - intLength
      If intLength < 0 Then
         intLength = Picture1.TextWidth("9")
      Else
         intLength = intLength + Picture1.TextWidth("9")
      End If
      Picture1.CurrentX = (ciXCurr + intMargin) * douExtRate + Picture1.TextWidth(m_Cur(1)) + intLength
      'end 2024/11/19
      'end 2024/11/4
      Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
      Picture1.Print strAmount
      strPayAmount = m_Cur(1) & " " & strAmount
   End If

   For m_Idx = 2 To UBound(m_Cur)
      If m_Idx > 1 Then intCounter = intCounter + 1
      If intCounter > 22 Then
         If intCounter = 23 Then
            intCounter = -10
         Else
            intCounter = 1
         End If
         intPage = intPage + 1
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            PrintPageNo True
            Printer.NewPage
         End If
         If bol2File = True Then
            PrintPageNo
            PicNewPage
         End If
      End If
      'Modified by Lydia 2018/09/03 + bol2pdf
      If bol2Printer = True Or bol2Pdf = True Then
         'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
         'Printer.CurrentX = 6500 + intMargin
         Printer.CurrentX = ciXCurr + intMargin
         Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
         Printer.Print m_Cur(m_Idx)
      End If
      If bol2File = True Then
         'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
         'Picture1.CurrentX = (6500 + intMargin) * douExtRate
         Picture1.CurrentX = (ciXCurr + intMargin) * douExtRate
         Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
         Picture1.Print m_Cur(m_Idx)
      End If
      strAmount = Format(m_Amount(m_Idx), FDollar)
      'Modified by Lydia 2018/09/03 + bol2pdf
      If bol2Printer = True Or bol2Pdf = True Then
         intLength = Printer.TextWidth(strAmount)
         'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
         'Printer.CurrentX = 8400 - intLength + intMargin
         Printer.CurrentX = ciXAmt - intLength + intMargin
         Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
         Printer.Print strAmount
      End If
      If bol2File = True Then
         intLength = Picture1.TextWidth(strAmount)
         'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
         'Picture1.CurrentX = (8400 + intMargin) * douExtRate - intLength
         'Modified by Morgan 2024/11/4
         'Picture1.CurrentX = (ciXAmt + intMargin) * douExtRate - intLength
         Picture1.CurrentX = (ciXAmt + intMargin - intLength) * douExtRate
         'end 2024/11/4
         Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
         Picture1.Print strAmount
      End If
   Next
   
   If intCounter > 18 Then
      intCounter = -14
      intPage = intPage + 1
      'Modified by Lydia 2018/09/03 + bol2pdf
      If bol2Printer = True Or bol2Pdf = True Then
         PrintPageNo True
         Printer.NewPage
      End If
      If bol2File = True Then
         PrintPageNo
         PicNewPage
      End If
   End If
   intCounter = intCounter + 2
   'Modified by Lydia 2018/09/03 + bol2pdf
   If bol2Printer = True Or bol2Pdf = True Then
      Printer.FontSize = 15
   End If
   If bol2File = True Then
      Picture1.FontSize = 15 * douExtRate
   End If
   Select Case m_a1b06
      Case "1"
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            Printer.CurrentX = 500 + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
            Printer.Print ReportSum(90)
         End If
         If bol2File = True Then
            Picture1.CurrentX = (500 + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
            Picture1.Print ReportSum(90)
         End If
         intCounter = intCounter + 1
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            Printer.CurrentX = 0 + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
            Printer.Print ReportSum(91)
         End If
         If bol2File = True Then
            Picture1.CurrentX = (0 + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
            Picture1.Print ReportSum(91)
         End If
         intCounter = intCounter + 1
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            Printer.CurrentX = 0 + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
            Printer.Print ReportSum(92)
         End If
         If bol2File = True Then
            Picture1.CurrentX = (0 + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
            Picture1.Print ReportSum(92)
         End If
      Case Else
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            Printer.CurrentX = 500 + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
            Printer.Print ReportSum(991)
         End If
         If bol2File = True Then
            Picture1.CurrentX = (500 + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
            Picture1.Print ReportSum(991)
         End If
         intCounter = intCounter + 1
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            Printer.CurrentX = 0 + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
            Printer.Print ReportSum(99101)
         End If
         If bol2File = True Then
            Picture1.CurrentX = (0 + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
            Picture1.Print ReportSum(99101)
         End If
   End Select
   
   intCounter = intCounter + 2
   'Modified by Lydia 2018/09/03 + bol2pdf
   If bol2Printer = True Or bol2Pdf = True Then
      Printer.FontSize = 15
      Printer.CurrentX = 500 + intMargin
      Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
      Printer.Print ReportSum(93)
   End If
   If bol2File = True Then
      Picture1.FontSize = 15 * douExtRate
      Picture1.CurrentX = (500 + intMargin) * douExtRate
      Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
      Picture1.Print ReportSum(93)
   End If
   intCounter = intCounter + 2
   'Modified by Lydia 2018/09/03 + bol2pdf
   If bol2Printer = True Or bol2Pdf = True Then
      Printer.CurrentX = 8000 + intMargin
      Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
      Printer.Print ReportSum(94)
   End If
   
   If bol2File = True Then
      Picture1.FontSize = 15 * douExtRate
      Picture1.CurrentX = (8000 + intMargin) * douExtRate
      Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
      Picture1.Print ReportSum(94)
   End If
   
   intCounter = intCounter + 2
   
   'Modify by Morgan 2009/4/15 註記列印改先●再＊
   'Modified by Lydia 2018/09/03 + bol2pdf
   If bol2Printer = True Or bol2Pdf = True Then
      Printer.FontSize = 15
      Printer.CurrentX = 0 + intMargin
      Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
      'Add by Morgan 2007/3/3 不必寄發電匯通知註記
      If strInform = "N" Then
         Printer.Print "●" & strNo
      'Modified by Lydia 2016/10/21 +票匯要Email => bol2Email = True
      ElseIf bolEmail = True Or bol2Email = True Then
         Printer.Print "＊" & strNo 'Memo by Lydia 2016/10/14 代理人編號前面加註記
      Else
         Printer.Print strNo
      End If
   End If
      
   If bol2File = True Then
      Picture1.FontSize = 15 * douExtRate
      Picture1.CurrentX = (0 + intMargin) * douExtRate
      Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
      'Add by Morgan 2007/3/3 不必寄發電匯通知註記
      If strInform = "N" Then
         Picture1.Print "●" & strNo
      'Modified by Lydia 2016/10/21 +票匯要Email => bol2Email = True
      ElseIf bolEmail = True Or bol2Email = True Then
         Picture1.Print "＊" & strNo
      Else
         Picture1.Print strNo
      End If
   End If
   
   'Modify by Morgan 2007/9/14 最後一頁頁次要累加，否則圖檔的檔名會有錯
   'If intCounter <= 16 Then
   '   If intPage > 0 Then
         intPage = intPage + 1
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            PrintPageNo True
         End If
         If bol2File = True Then
            PrintPageNo
         End If
   '   End If
   'End If
   
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Private Function FormCheck() As Boolean
   'Add by Amy 2014/05/01 +公司別
   If Text3 = "" Then
        MsgBox "公司別不可空白！"
        Text3.SetFocus
        Exit Function
   End If
   'end 2014/05/01
      
   'Add by Morgan 2010/6/10
   If Text5 = "" Then
      MsgBox "輸出選項不可空白！"
      Text5.SetFocus
      Exit Function
   ElseIf Text5 = "2" And Text6 <> "Y" Then
      If MsgBox("確定要發EMail？", vbYesNo + vbDefaultButton2) = vbNo Then
         Text5.SetFocus
         Exit Function
      End If
   'Added by Morgan 2018/1/22
   ElseIf Text5 = "3" Then
      If txtPath = "" Then
         MsgBox "請點選匯出路徑！", vbExclamation
         txtPath.SetFocus
         Exit Function
      ElseIf Dir(txtPath, vbDirectory) = "" Then
         MsgBox "匯出路徑已不存在，請重新點選！", vbExclamation
         txtPath.SetFocus
         Exit Function
      End If
   'end 2018/1/22
   End If
   'end 2010/6/10
   
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
   MsgBox MsgText(181), , MsgText(5)
End Function

'*************************************************
' 列印資料(匯票函件明細表)
'
'*************************************************
Private Sub PrintData1()
Dim strName As String

   strName = ""
   intCounter = 1
   Printer.FontSize = 12
   adoaccrpt219.CursorLocation = adUseClient
   adoaccrpt219.Open "select * from accrpt219 where r21901 = '" & strUserNum & "' order by r21908 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While adoaccrpt219.EOF = False
      If intCounter > 40 Then
         intCounter = 1
         Printer.NewPage
         PrintHead1
      End If
      If strName <> adoaccrpt219.Fields("r21902").Value Then
         If strName = "" Then
            PrintHead1
         End If
         strName = adoaccrpt219.Fields("r21902").Value
         Printer.CurrentX = 0
         Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
         If IsNull(adoaccrpt219.Fields("r21903").Value) Then
            Printer.Print ""
         Else
            Printer.Print Mid(adoaccrpt219.Fields("r21903").Value, 1, 17)
         End If
      End If
      Printer.CurrentX = 2500
      Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
      If IsNull(adoaccrpt219.Fields("r21904").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoaccrpt219.Fields("r21904").Value
      End If
      Printer.CurrentX = 7500
      Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
      If IsNull(adoaccrpt219.Fields("r21905").Value) Then
         Printer.Print ""
      Else
         Printer.Print adoaccrpt219.Fields("r21905").Value
      End If
      If IsNull(adoaccrpt219.Fields("r21906").Value) = True Or adoaccrpt219.Fields("r21906").Value = 0 Then
         strAmount = ""
      Else
         strAmount = Format(Val(adoaccrpt219.Fields("r21906").Value), FDollar)
      End If
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 9400 - intLength
      Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
      Printer.Print strAmount
      If IsNull(adoaccrpt219.Fields("r21907").Value) = True Or adoaccrpt219.Fields("r21907").Value = 0 Then
         strAmount = ""
      Else
         strAmount = Format(Val(adoaccrpt219.Fields("r21907").Value), FDollar)
      End If
      intLength = Printer.TextWidth(strAmount)
      Printer.CurrentX = 10700 - intLength
      Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
      Printer.Print strAmount
      intCounter = intCounter + 1
      adoaccrpt219.MoveNext
   Loop
   PrintSum1
   Printer.EndDoc
   adoaccrpt219.Close
End Sub

'*************************************************
'  抬頭列印(匯票函件明細表)
'
'*************************************************
Private Sub PrintHead1()
   intCounter = 0
   Printer.CurrentX = 4000
   Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
   Printer.Print ReportTitle(219)
   intCounter = intCounter + 2
   Printer.CurrentX = 9300
   Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
   Printer.Print Format(AFDate(CADate(ACDate(ServerDate))), "mmm. d, yyyy")
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
   Printer.Print "客戶名稱"
   Printer.CurrentX = 2500
   Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
   Printer.Print "本所案號"
   Printer.CurrentX = 7500
   Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
   Printer.Print "幣別"
   Printer.CurrentX = 8300
   Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
   Printer.Print "小計"
   Printer.CurrentX = 9500
   Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
   Printer.Print "合計"
   Printer.Line (0, intTop + intPrtXFix + 500 + intCounter * 300 + 350)-(10700, intTop + intPrtXFix + 500 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub

'*************************************************
'  列印小計
'
'*************************************************
Public Function PrintSum1() As Boolean
'Add By Cheng 2003/09/03
Dim dblCnt As Double '件數
Dim dblAmt As Double '金額合計
   
    dblCnt = 0: dblAmt = 0
   If intCounter > 38 Then
      intCounter = 1
      Printer.NewPage
   End If
   intCounter = intCounter + 1
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'Modify by Amy 2014/05/01 +公司別
   adoquery.Open "select count(distinct a1b02), count(distinct a1b01) from acc1b0, acc190 where a1b01 = a1908 and a1b06 = '1' And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & strConSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
         Printer.Print "代理人合計: " & adoquery.Fields(0).Value & " 家"
      End If
      intCounter = intCounter + 1
      If IsNull(adoquery.Fields(1).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
         Printer.Print "匯票合計: " & adoquery.Fields(1).Value & " 張"
      End If
      intCounter = intCounter + 1
   End If
   adoquery.Close
   adoquery.CursorLocation = adUseClient
  'Modify by Amy 2014/05/01 +公司別
   adoquery.Open "select count(distinct axf03) from acc1b0, acc190, acc151 where a1b01 = a1908 and a1902 = axf01 and a1b06 = '1' And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & strConSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
         Printer.Print "件數合計: " & adoquery.Fields(0).Value & " 件"
         intCounter = intCounter + 1
        dblCnt = dblCnt + Val("" & adoquery.Fields(0).Value)
      End If
   End If
   adoquery.Close
   adoquery.CursorLocation = adUseClient
   'Modify by Amy 2014/05/01 +公司別
   adoquery.Open "select sum(a1p08) from acc1b0, acc1p0 where a1b01||a1b02 = a1p04 and a1b06 = '1' And a1p01||''='" & Text3 & "'" & strConSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
         Printer.Print "台幣結匯金額合計: " & Format(adoquery.Fields(0).Value, FDollar)
        dblAmt = dblAmt + Val("" & adoquery.Fields(0).Value)
      End If
   End If
   adoquery.Close
    'Add By Cheng 2003/09/03
    '電匯資料
    intCounter = intCounter + 1
    intCounter = intCounter + 1
   If intCounter > 38 Then
      intCounter = 1
      Printer.NewPage
   End If
   intCounter = intCounter + 1
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   'Modify by Amy 2014/05/01 +公司別
   adoquery.Open "select count(distinct a1b02), count(distinct a1b01) from acc1b0, acc190 where a1b01 = a1908 and a1b06 = '2' And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & strConSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
         Printer.Print "電匯代理人合計: " & adoquery.Fields(0).Value & " 家"
      End If
      intCounter = intCounter + 1
      If IsNull(adoquery.Fields(1).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
         Printer.Print "電匯合計: " & adoquery.Fields(1).Value & " 張"
      End If
      intCounter = intCounter + 1
   End If
   adoquery.Close
   adoquery.CursorLocation = adUseClient
   'Modify by Amy 2014/05/01 +公司別
   adoquery.Open "select count(distinct axf03) from acc1b0, acc190, acc151 where a1b01 = a1908 and a1902 = axf01 and a1b06 = '2' And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & strConSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
         Printer.Print "電匯件數合計: " & adoquery.Fields(0).Value & " 件"
         intCounter = intCounter + 1
        dblCnt = dblCnt + Val("" & adoquery.Fields(0).Value)
      End If
   End If
   adoquery.Close
   adoquery.CursorLocation = adUseClient
   'Modify by Amy 2014/05/01 +公司別
   adoquery.Open "select sum(a1p08) from acc1b0, acc1p0 where a1b01||a1b02 = a1p04 and a1b06 = '2' And a1p01||''='" & Text3 & "'" & strConSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         Printer.CurrentX = 0
         Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
         Printer.Print "電匯台幣結匯金額合計: " & Format(adoquery.Fields(0).Value, FDollar)
        dblAmt = dblAmt + Val("" & adoquery.Fields(0).Value)
      End If
   End If
   adoquery.Close

    intCounter = intCounter + 1
    intCounter = intCounter + 1
   If intCounter > 38 Then
      intCounter = 1
      Printer.NewPage
   End If
   intCounter = intCounter + 1
     Printer.CurrentX = 0
     Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
     Printer.Print "件數總計: " & dblCnt & " 件"
   intCounter = intCounter + 1
     Printer.CurrentX = 0
     Printer.CurrentY = intTop + intPrtXFix + 500 + intCounter * 300
     Printer.Print "台幣金額總計: " & Format(dblAmt, FDollar)

End Function

'*************************************************
'  產生匯票函件明細表
'
'*************************************************
Public Function ProcessDetail() As Boolean
Dim intRow As Integer
Dim strName As String
Dim strCase As String
Dim douAmount As Double
Dim strTotalCase As String
Dim strFagentName As String

   intCounter = 1
   intRow = 1
   douAmount = 0
   adoTaie.Execute "delete from accrpt219 where r21901 = '" & strUserNum & "'"
   If adoquery.State = adStateOpen Then
      adoquery.Close
   End If
   adoquery.CursorLocation = adUseClient
   '2006/3/21 MODIFY BY SONIA
   'adoquery.Open "select * from acc190, acc180, acc1b0, fagent where a1901 = a1801 and a1908 = a1b01 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and a1b06 = '1'" & strConSql & " order by a1803 asc, a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2009/10/13 modify by sonia a1b01=1598重覆會抓到W000026748,所以加a1b02=a1803條件
   'adoquery.Open "select * from acc190, acc180, acc1b0, fagent,CUSTOMER where a1901 = a1801 and a1908 = a1b01 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and a1b06 = '1'" & strConSql & " order by a1803 asc, a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify by Amy 2014/05/01 +公司別
   adoquery.Open "select * from acc190, acc180, acc1b0, fagent,CUSTOMER where a1901 = a1801 and a1908 = a1b01 and a1b02=a1803 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and substr(a1803, 1, 8) = CU01 (+) and substr(a1803, 9, 1) = CU02 (+) and a1b06 = '1' And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & strConSql & " order by a1803 asc, a1902 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2009/10/13 end
   '2006/3/21 END
   Do While adoquery.EOF = False
      If IsNull(adoquery.Fields("a1803").Value) = False Then
         strName = adoquery.Fields("a1803").Value
      Else
         strName = ""
      End If
      adoaccsum.CursorLocation = adUseClient
      '2007/3/2 modify by sonia 加入cp87,cp88
      'adoaccsum.Open "select cp01||cp02 from caseprogress where (cp61 = '" & adoquery.Fields("a1902").Value & "' or cp62 = '" & adoquery.Fields("a1902").Value & "' or cp63 = '" & adoquery.Fields("a1902").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
      adoaccsum.Open "select cp01||cp02 from caseprogress where (cp61 = '" & adoquery.Fields("a1902").Value & "' or cp62 = '" & adoquery.Fields("a1902").Value & "' or cp63 = '" & adoquery.Fields("a1902").Value & "' or cp87 = '" & adoquery.Fields("a1902").Value & "' or cp88 = '" & adoquery.Fields("a1902").Value & "')", adoTaie, adOpenStatic, adLockReadOnly
      '2007/3/2 end
      Do While adoaccsum.EOF = False
         If intRow > 4 Then
            adoTaie.Execute "insert into accrpt219 values ('" & strUserNum & "', '" & adoquery.Fields("a1803").Value & "', '" & IIf(IsNull(adoquery.Fields("a1810").Value), adoquery.Fields("fa05").Value, adoquery.Fields("a1810").Value) & "', '" & strCase & "', null, 0, 0, " & intCounter & ")"
            strCase = ""
            intRow = 1
         End If
         If InStr(1, strTotalCase, adoaccsum.Fields(0).Value) > 0 Then
         Else
            strCase = strCase & " " & adoaccsum.Fields(0).Value
            strTotalCase = strTotalCase & " " & adoaccsum.Fields(0).Value
         End If
         intRow = intRow + 1
         adoaccsum.MoveNext
      Loop
      adoaccsum.Close
      douAmount = douAmount + Val(adoquery.Fields("a1904").Value)
      adoquery.MoveNext
      If adoquery.EOF = False Then
         adoquery.MovePrevious
         If IsNull(adoquery.Fields("a1810").Value) Then
            'Modify By Cheng 2003/12/17
'            strFagentName = Replace(adoquery.Fields("fa05").Value, "'", "''")
            If "" & adoquery.Fields("fa05").Value <> "" Then
                strFagentName = Replace(adoquery.Fields("fa05").Value, "'", "''")
            ElseIf "" & adoquery.Fields("fa04").Value <> "" Then
                strFagentName = PUB_StrToStr(Replace(adoquery.Fields("fa04").Value, "'", "''"), 20)
            ElseIf "" & adoquery.Fields("fa06").Value <> "" Then
                strFagentName = PUB_StrToStr(Replace(adoquery.Fields("fa06").Value, "'", "''"), 20)
            End If
         Else
            strFagentName = Replace(adoquery.Fields("a1810").Value, "'", "''")
         End If
         adoquery.MoveNext
         If strName <> adoquery.Fields("a1803").Value Then
            adoquery.MovePrevious
            'If strName <> "" Then
'               adoTaie.Execute "insert into accrpt219 values ('" & strUserNum & "', '" & adoquery.Fields("a1803").Value & "', '" & strFagentName & "', '" & strCase & "', '" & adoquery.Fields("a1903").Value & "', " & douAmount & ", " & douAmount & ", " & intCounter & ")"
               adoTaie.Execute "insert into accrpt219 values ('" & strUserNum & "', '" & ChgSQL("" & adoquery.Fields("a1803").Value) & "', '" & ChgSQL(strFagentName) & "', '" & strCase & "', '" & adoquery.Fields("a1903").Value & "', " & douAmount & ", " & douAmount & ", " & intCounter & ")"
               adoquery.MoveNext
               strCase = ""
               strTotalCase = ""
               douAmount = 0
               intRow = 1
            'End If
            'strName = adoquery.Fields("a1803").Value
         End If
      Else
         adoquery.MovePrevious
         If IsNull(adoquery.Fields("a1810").Value) Then
            'Modify By Cheng 2003/12/17
'            strFagentName = Replace(adoquery.Fields("fa05").Value, "'", "''")
            If "" & adoquery.Fields("fa05").Value <> "" Then
                strFagentName = Replace(adoquery.Fields("fa05").Value, "'", "''")
            ElseIf "" & adoquery.Fields("fa04").Value <> "" Then
                strFagentName = PUB_StrToStr(Replace(adoquery.Fields("fa04").Value, "'", "''"), 20)
            ElseIf "" & adoquery.Fields("fa06").Value <> "" Then
                strFagentName = PUB_StrToStr(Replace(adoquery.Fields("fa06").Value, "'", "''"), 20)
            ElseIf "" & adoquery.Fields("CU05").Value <> "" Then
                strFagentName = PUB_StrToStr(Replace(adoquery.Fields("CU05").Value, "'", "''"), 20)
            ElseIf "" & adoquery.Fields("CU04").Value <> "" Then
                strFagentName = PUB_StrToStr(Replace(adoquery.Fields("CU04").Value, "'", "''"), 20)
            ElseIf "" & adoquery.Fields("CU06").Value <> "" Then
                strFagentName = PUB_StrToStr(Replace(adoquery.Fields("CU06").Value, "'", "''"), 20)
            End If
         Else
            strFagentName = Replace(adoquery.Fields("a1810").Value, "'", "''")
         End If
         adoTaie.Execute "insert into accrpt219 values ('" & strUserNum & "', '" & adoquery.Fields("a1803").Value & "', '" & strFagentName & "', '" & strCase & "', '" & adoquery.Fields("a1903").Value & "', " & douAmount & ", " & douAmount & ", " & intCounter & ")"
         adoquery.MoveNext
      End If
      intCounter = intCounter + 1
   Loop
   adoquery.Close
End Function
'Add by Morgan 2005/5/18 折行列印
'p_stContent=列印內容,p_iRow=起始行數
Private Sub XPrint(ByVal p_stContent As String, ByRef p_iRow As Integer)
   Dim iPos As Integer, strTemp As String
   iPos = 1
   strTemp = Mid(p_stContent, iPos, 22)
   Do
      'Modified by Lydia 2018/09/03 + bol2pdf
      If bol2Printer = True Or bol2Pdf = True Then
         Printer.CurrentX = 0 + intMargin
         Printer.CurrentY = intTop + intPrtXFix + 2500 + p_iRow * 250
         Printer.Print strTemp
      End If
      'Add by Morgan 2008/2/15
      If bol2File = True Then
         Picture1.CurrentX = (0 + intMargin) * douExtRate
         Picture1.CurrentY = (intTop + 2500 + p_iRow * 250) * douExtRate
         Picture1.Print strTemp
      End If
      iPos = iPos + 22
      strTemp = Mid(p_stContent, iPos, 22)
      If strTemp <> "" Then
         p_iRow = p_iRow + 1
      Else
         Exit Do
      End If
   Loop
End Sub
'*************************************************
' 列印明細資料
'
'*************************************************
Private Sub PrintData()

   Dim strDescription As String
   Dim strName As String
   Dim StrSQLa As String
   Dim strAXF03 As String '本所案號
   Dim strA1501 As String '單據號碼
   Dim strMailFailList() As String 'Mail 失敗清單
   Dim iCopy As Integer
   Dim iRound As Integer '迴圈次數
   Dim bolChgPrinter As Boolean 'Added by Morgan 2012/10/11
   Dim tmpErr As String 'Added by Lydia 2020/09/10
   
   'douExtRate = Screen.TwipsPerPixelX / 15 'Remove by Morgan 2011/10/5
   
On Error GoTo ErrorHandle 'Added by Lydia 2019/05/28

   cnnConnection.Execute "delete from accrpt2450 where id='" & strUserNum & "' " 'Added by Lydia 2017/10/05 清除暫存檔"
   
   adoquery.CursorLocation = adUseClient
   'Modify by Morgan 2007/3/3 加fa84(cu120)
   'Modify by Morgan 2010/6/7 +a1b03
   'Modify by Morgan 2011/5/25 +fa70
   '2012/5/9 modify by sonia order by a1b02 asc, axf03 asc, a1501 asc 改為 order by a1b02 asc, a1b01, axf03 asc, a1501 asc (U10102708(CFP024278改與CFT開同一匯票)
   'Modify by Amy 2014/05/01 +公司別
   'Modify by Amy 2015/03/19 +a1901,a1902,fa29(for mail 給kelly) 和抓 acc170(B單號) 及 acc130(O單號)
   'modify by sonia 2017/7/10 付款對象固定改為另一家,應收a1b02抓FAGENT
   'strSQLa = "Select axf03, a1b02, cp45, a1504, a1505, axf04, a1b06, a1b01, a1502, cp01, cp02, cp03, cp04, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, a1c03, a1501, fa17, fa04, fa06, CP09,fa10,nvl(fa79,fa16) fa16,fa84,a1b03,fa70,a1917,a1901,a1902,fa29 " & _
                    "From acc1c0, acc1b0, acc151, acc150, caseprogress, fagent,acc190 Where a1c01 = a1b01 and a1c02 = a1b02 and a1c03 = a1501 (+) and a1c03 = axf01 (+) and axf02 = cp09 (+) and substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) And a1b01=a1908 And a1c03=a1902(+) " & _
                    "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " and axf01 is not null" & strConSql & " Union " & _
                    "Select axg03 as axf03, a1b02, cp45, a1604 as a1504, a1605 as a1505, axg04 as axf04, a1b06, a1b01, a1602, cp01, cp02, cp03, cp04, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, a1c03, a1601 as a1501, fa17, fa04, fa06, CP09,fa10,nvl(fa79,fa16) fa16,fa84,a1b03,fa70,a1917,a1901,a1902,fa29 " & _
                    "From acc1c0, acc1b0, acc161, acc160, caseprogress, fagent,acc190 Where a1c01 = a1b01 and a1c02 = a1b02 and a1c03 = a1601 (+) and a1c03 = axg01 (+) and axg02 = cp09 (+) and substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) And a1b01=a1908 And a1c03=a1902(+) " & _
                    "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " and axg01 is not null" & strConSql
   'Modified by Morgan 2018/2/26 +acc170.a1719
   'Modified by Lydia 2018/07/20 +fa105 (財務信箱CF), 優先抓FA105->FA79
   'Modified by Lydia 2018/11/19 +獨立水單(A1812),axf01
   'strSQLa = "Select axf03, a1b02, cp45, a1504, a1505, axf04, a1b06, a1b01, a1502, cp01, cp02, cp03, cp04, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, a1c03, a1501, fa17, fa04, fa06, CP09,fa10,nvl(fa105,nvl(fa79,fa16)) fa16,fa84,a1b03,fa70,a1917,a1901,a1902,fa29,a1719 " & _
                    "From acc1c0, acc1b0, acc151, acc150, caseprogress, fagent,acc190,acc170 Where a1c01 = a1b01 and a1c02 = a1b02 and a1c03 = a1501 (+) and a1c03 = axf01 (+) and axf02 = cp09 (+) and substr(a1b02, 1, 8) = fa01 (+) and substr(a1b02, 9, 1) = fa02 (+) And a1b01=a1908 And a1c03=a1902(+)  and a1702(+)=a1902 " & _
                    "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " and axf01 is not null" & strConSql & " Union " & _
                    "Select axg03 as axf03, a1b02, cp45, a1604 as a1504, a1605 as a1505, axg04 as axf04, a1b06, a1b01, a1602, cp01, cp02, cp03, cp04, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, a1c03, a1601 as a1501, fa17, fa04, fa06, CP09,fa10,nvl(fa105,nvl(fa79,fa16)) fa16,fa84,a1b03,fa70,a1917,a1901,a1902,fa29,a1719 " & _
                    "From acc1c0, acc1b0, acc161, acc160, caseprogress, fagent,acc190,acc170 Where a1c01 = a1b01 and a1c02 = a1b02 and a1c03 = a1601 (+) and a1c03 = axg01 (+) and axg02 = cp09 (+) and substr(a1b02, 1, 8) = fa01 (+) and substr(a1b02, 9, 1) = fa02 (+) And a1b01=a1908 And a1c03=a1902(+)  and a1702(+)=a1902 " & _
                    "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " and axg01 is not null" & strConSql
   'Modified by Lydia 2024/09/18 +財務副本信箱emailcc：寄財務信箱一併CC副本>>,decode(fa105||fa79,null,'',fa134) as emailcc
   StrSQLa = "Select axf03, a1b02, cp45, a1504, a1505, axf04, a1b06, a1b01, a1502, cp01, cp02, cp03, cp04, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, a1c03, a1501, fa17, fa04, fa06, CP09,fa10,nvl(fa105,nvl(fa79,fa16)) fa16,fa84,a1b03,fa70,a1917,a1901,a1902,fa29,a1719,a1812,axf01 " & _
                    ",decode(fa105||fa79,null,'',fa134) as emailcc From acc1c0, acc1b0, acc151, acc150, caseprogress, fagent,acc190,acc170,acc180 Where a1c01 = a1b01 and a1c02 = a1b02 and a1c03 = a1501 (+) and a1c03 = axf01 (+) and axf02 = cp09 (+) and substr(a1b02, 1, 8) = fa01 (+) and substr(a1b02, 9, 1) = fa02 (+) And a1b01=a1908 And a1c03=a1902(+)  and a1702(+)=a1902 " & _
                    "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " and a1801(+)=a1901 and axf01 is not null" & strConSql
   'Modified by Lydia 2024/09/18 +財務副本信箱emailcc：寄財務信箱一併CC副本>>,decode(fa105||fa79,null,'',fa134) as emailcc
   StrSQLa = StrSQLa & " Union Select axg03 as axf03, a1b02, cp45, a1604 as a1504, a1605 as a1505, axg04 as axf04, a1b06, a1b01, a1602, cp01, cp02, cp03, cp04, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, a1c03, a1601 as a1501, fa17, fa04, fa06, CP09,fa10,nvl(fa105,nvl(fa79,fa16)) fa16,fa84,a1b03,fa70,a1917,a1901,a1902,fa29,a1719,'' as a1812,'' as axf01 " & _
                    ",decode(fa105||fa79,null,'',fa134) as emailcc From acc1c0, acc1b0, acc161, acc160, caseprogress, fagent,acc190,acc170 Where a1c01 = a1b01 and a1c02 = a1b02 and a1c03 = a1601 (+) and a1c03 = axg01 (+) and axg02 = cp09 (+) and substr(a1b02, 1, 8) = fa01 (+) and substr(a1b02, 9, 1) = fa02 (+) And a1b01=a1908 And a1c03=a1902(+)  and a1702(+)=a1902 " & _
                    "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " and axg01 is not null" & strConSql
    'Mark by Amy 2015/04/08 婉莘:只寄給kelly的才需要抓 acc170(B單號),所以改至最後獨立做
'    strSQLa = strSQLa & " Union " & _
'                    "Select a1707 as axf03, a1b02, '' as cp45, a1706 as a1504, a1703 as a1505, a1704 as axf04, a1b06, a1b01, a1708, '' as cp01, '' as cp02, '' as cp03, '' as cp04, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, a1c03, a1702 as a1501, fa17, fa04, fa06, '' as CP09,fa10,nvl(fa79,fa16) fa16,fa84,a1b03,fa70,a1917,a1901,a1902,fa29 " & _
'                    "From acc1c0, acc1b0,acc170, fagent,acc190 Where a1c01 = a1b01 and a1c02 = a1b02 and a1c03 = a1702 (+) and substr(a1705, 1, 8) = fa01 (+) and substr(a1705, 9, 1) = fa02 (+) And a1b01=a1908 And a1c03=a1902(+) And SubStr(a1c03,1,1)= 'B' " & _
'                    "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " and a1702 is not null" & strConSql
    'Mark by Amy 2015/04/07 婉莘:O單號另外做不需出現於此 acc130(O單號)
'    strSQLa = strSQLa & " Union " & _
'                    " Select '' as axf03, a1b02, '' as cp45, '' as a1504, a1306 as a1505, a1307 as axf04, a1b06, a1b01, a1302, '' as cp01, '' as cp02, '' as cp03, '' as cp04, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, a1c03, a1301 as a1501, fa17, fa04, fa06, '' as CP09,fa10,nvl(fa79,fa16) fa16,fa84,a1b03,fa70,a1917,a1901,a1902,fa29 " & _
'                    "From acc1c0, acc1b0,acc130, fagent,acc190 Where a1c01 = a1b01 and a1c02 = a1b02 and a1c03 = a1301 (+) and substr(a1304, 1, 8) = fa01 (+) and substr(a1304, 9, 1) = fa02 (+) And a1b01=a1908 And a1c03=a1902(+) And SubStr(a1c03,1,1)= 'O' " & _
'                    "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " and a1301 is not null" & strConSql
    StrSQLa = StrSQLa & " Order by a1b02 asc, a1b01, axf03 asc, a1501 asc"
    'end 2015/03/19
    
   adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount = 0 Then
      adoquery.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   'Add by Morgan 2008/4/16 存電子檔時不印
   If Text6 = "Y" Then
      strSavePath = PUB_Getdesktop
      bol2Printer = False
   Else
      strSavePath = App.path
      If Text5 = "1" Then
         bol2Printer = True
      Else
         bol2Printer = False
      End If
   End If
   
   
   '刪除舊的暫存圖檔
   strExc(1) = App.path & "\$*.jpg"
   If Dir(strExc(1)) <> "" Then Kill strExc(1)
   'Added by Lydia 2018/09/03 刪除舊的PDF檔
   strExc(1) = App.path & "\$*.pdf"
   If Dir(strExc(1)) <> "" Then Kill strExc(1)
   'end 2018/09/03

'Modify by Morgan 2008/11/20 改複寫為單張列印，不發Mail的另外多印一份
'Modify by Morgan 2010/6/10 要列印且非存電子檔的才要兩次
'For iCopy = 1 To 2
'Modified by Lydia 2016/10/21 因為有可能要email票匯,並且同時印2份
'If Text5 = "1" And Text6 <> "Y" Then
'Modified by Lydia 2016/11/16 重新界定需求:選列印時票匯印2份電匯印1份,選email則都要發mail不做列印
'If Text6 <> "Y" Then
If Text5 = "1" And Text6 <> "Y" Then
   iRound = 2
Else
   iRound = 1
End If
'Added by Lydia 2020/10/08
If Text5 = "2" Then
   txtSend.Visible = True: lblSend.Visible = True
End If
'end 2020/10/08

For iCopy = 1 To iRound
'end 2010/6/10
   
   'Added by Lydia 2017/11/01  改成先印清單 (第2次列印之前)
   If iCopy = 2 Then
      'Memo by Lydia 2018/11/19 已與婉莘確認,若輸入代理人編號也不用輸出Invoice
      If Trim(Text1.Text & Text2.Text) = "" Then 'Added by Lydia 2017/11/13 若輸入代理人編號, 則不需要列印清單
         PrintList
      End If   'end 2017/11/13
   End If
   'end 2017/11/01
   
   'Added by Morgan 2012/10/11
   If iCopy = 2 And bolChgPrinter Then
      PUB_RestorePrinter cmbPrinter2
      'Added by Lydia 2017/11/08 因為是同一台印表機的不同紙匣,要多設定讓driver改到第2次列印的印表機
      Printer.PaperSize = 9
      Printer.EndDoc
      'end 2017/11/08
   End If

   Erase strMailFailList
   ReDim strMailFailList(0)
   intLength = 0
   douAmount = 0
   douUSDollar = 0
   intPage = 0
   strNo = ""
   strPicLetter = ""
   strPicFileNames = ""
   mCnt = 0 'Added by Lydia 2017/10/05
   adoquery.MoveFirst
   Do While adoquery.EOF = False
    'Add by Morgan 2010/6/7 改印結匯日期
    If Not IsNull(adoquery.Fields("a1b03")) Then
       m_strPayDate = DBDATE(adoquery.Fields("a1b03").Value)
    Else
       m_strPayDate = strSrvDate(1)
    End If
      
    strAXF03 = "" & adoquery.Fields("axf03").Value
    strA1501 = "" & adoquery.Fields("a1501").Value
    If IsNull(adoquery.Fields("axf03").Value) = False Then
       strName = adoquery.Fields("axf03").Value
    Else
       strName = ""
    End If
      
    'Added by Lydia 2017/10/05 列印:留所資料改成清單方式產出
    bolList = False
    If iCopy = 1 And bol2Printer = True And Text5.Text = "1" And Text6.Text <> "Y" Then
       bolList = True
    Else
       strName = strName
    End If
    'end 2017/10/05
    
      '2012/2/23 modify by sonia 同一代理人匯票號碼不同也要分開印 (Y28215於100/8/2)
      'If strNo <> adoquery.Fields("a1b02").Value Then
      If strNo <> adoquery.Fields("a1b02").Value Or m_a1b01 <> adoquery.Fields("a1b01").Value Then
         If douAmount <> 0 Then
            'Added by Lydia 2017/10/05 暫存清單資料
            If bolList = True Then
                mCnt = mCnt + 1
                strExc(0) = "insert into accrpt2450 (id,r24500,r24501,r24502,r24503,r24504,r24505,r24506) values (" & CNULL(strUserNum) & _
                            ", " & mCnt & ", NULL, NULL,'小　計' ,NULL,NULL," & douAmount & ") "
                cnnConnection.Execute strExc(0)
                mCnt = mCnt + 1
            Else
            'end 2017/10/05
                PrintSum
                If bol2Printer = True Then
                   Printer.NewPage
                'Added by Lydia 2018/09/03 一張匯票存成一個PDF檔
                ElseIf bol2Pdf = True Then
                      Printer.EndDoc
                      frmPDF.EndtProcess
                      bolPdfStart = False 'Added by Lydia 2023/11/15
                      'Added by Lydia 2019/08/21 判斷檔案是否存在, 超過時間就繼續
                      'Modified by Lydia 2020/02/15 +改成共用
                      'If ChkFileStatus(strSavePath & "\$" & strNo & m_a1b01 & ".PDF" & "*") = False Then
                      'Modified by Lydia 2020/09/10 超過時間，改不發email直接出發信失敗清單
                      'If PUB_ChkFileStatus(strSavePath & "\$" & strNo & m_a1b01 & ".PDF" & "*") = False Then
                      tmpErr = ""
                      If PUB_ChkFileStatus(strSavePath & "\$" & strNo & m_a1b01 & ".PDF" & "*", False, tmpErr) = False Then
                          If tmpErr <> "" Then
                              If strMailFailList(0) <> "" Then
                                 ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                              End If
                              strMailFailList(UBound(strMailFailList)) = strNo & "：" & Mid(tmpErr, 2) & IIf((bolEmail = True Or bol2Email = True) And Text5.Text = "2" And Text6.Text <> "Y", "請重新Email給該代理人", "")
                          End If
                      'end 2020/09/10
                      End If
                      'end 2019/08/21
                      Unload frmPDF
                      If tmpErr = "" Then 'Added by Lydia 2020/09/10 增加判斷
                           strPicFileNames = strPicFileNames & strSavePath & "\$" & strNo & m_a1b01 & ".PDF" & "*"
                      End If
                      '開新檔案
                      frmPDF.Show
                      frmPDF.StartProcess strSavePath, "$" & adoquery.Fields("a1b02").Value & adoquery.Fields("a1b01").Value & ".PDF"
                      bolPdfStart = True 'Added by Lydia 2023/11/15
                      '列印信頭
                      Printer.PaintPicture Picture1, 0, 0, Picture1.Width, Picture1.Height
                'end 2018/09/03
                End If
            End If 'end 2017/10/05
            
            'Modified by Lydia 2018/09/03
            'If bol2File = True Then
            '   PicNewPage
            If bol2File = True Or bol2Pdf = True Then
               If bol2File = True Then
                  PicNewPage '存檔後,開新檔
                  
                  'Added by Morgan 2024/10/30 因PDFCreator常出現無法預期的錯誤,改用Word將JPG轉PDF
                  strExc(1) = strSavePath & "\$" & strNo & m_a1b01 & ".PDF"
                  'Modified by Morgan 2024/9/4 中文要加印收據章
                  If PUB_JPG2PDF(strPicFileNames, strExc(1)) = True Then
                     strPicFileNames = strExc(1)
                  End If
                  'end 2024/10/30
               End If
            'end 2018/09/03
               
               'Modify by Morgan 2010/6/10 選發EMail且未設要存電子檔才寄送
               'If bolEmail = True And bol2Printer = True Then
               'Modified by Lydia 2016/10/21 票匯增加可Email的功能
               'If bolEmail = True And Text5.Text = "2" And Text6.Text <> "Y" Then
               'Modified by Lydia 2016/11/16 重新界定需求:選列印時票匯印2份電匯印1份,選email則都要發mail不做列印
               'If (bolEmail = True Or (bol2Email = True And iCopy = 2)) And Text5.Text = "2" And Text6.Text <> "Y" Then
               'Modified by Lydia 2018/09/03 判斷有檔案才寄信
               'If (bolEmail = True Or bol2Email = True) And Text5.Text = "2" And Text6.Text <> "Y" Then
               If (bolEmail = True Or bol2Email = True) And Text5.Text = "2" And Text6.Text <> "Y" And strPicFileNames <> "" Then
                  bolMailFailNoAlert = True
                  bolMailSendOk = False
                  txtSend.Text = strNo   'Added by Lydia 2020/10/08
                  'Modify by Morgan 2011/4/22 改以ipdept@taie.com.tw 寄但回覆還是給寄件人(70004)
                  'PUB_SendMail strUserNum, strEMailBox, "", "ADVICE OF WIRE PAYMENT", GetMailContent, , strPicFileNames, True, True, True
                  'Modified by Morgan 2011/10/12 改用 account@taie.com.tw 寄
                  'PUB_SendMail strUserNum, strEMailBox, "", "ADVICE OF WIRE PAYMENT", GetMailContent, , strPicFileNames, True, True, True, , "ipdept@taie.com.tw", "TAI E INTERNATIONAL PATENT & LAW OFFICE", strUserNum
                  'Modify by Amy 2014/05/01 改公司抬頭
                  'Modified by Morgan 2014/8/27 改回覆到財務信箱 -- 婧瑄
                  'PUB_SendMail strUserNum, strEMailBox, "", "ADVICE OF WIRE PAYMENT", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, strCompName, strUserNum
                  'Modified by Lydia 2016/10/19 "ADVICE OF WIRE PAYMENT" => "ADVICE OF WIRE/CHECK PAYMENT"
                  'Modified by Lydia 2020/09/22 +代理人編號 ADVICE OF WIRE/CHECK PAYMENT=> ADVICE OF WIRE/CHECK PAYMENT (Yxxxxx)
                  'Modified by Lydia 2024/09/18 +財務副本信箱strEmailCC
                  PUB_SendMail strUserNum, strEMailBox, "", "ADVICE OF WIRE/CHECK PAYMENT (" & strNo & ")", GetMailContent, , strPicFileNames, True, True, True, strEmailCC, strAccMailBox, strCompName, strAccMailBox
                  bolMailFailNoAlert = False
                  If bolMailSendOk = False Then
                     If strMailFailList(0) <> "" Then
                        ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                     End If
                     strMailFailList(UBound(strMailFailList)) = strNo & "：" & strEMailBox
                  End If
                  strPicFileNames = ""
               End If
               
               strPicFileNames = "" 'Added by Lydia 2024/12/09 已存電子檔,清空JPG圖片路徑
            End If
            
            douAmount = 0
            douUSDollar = 0
            intPage = 0
         End If
         intCounter = 1
         
         '是否寄明細
         strInform = "" & adoquery("fa84")
         '電子信箱
         strEMailBox = "" & adoquery("fa16")
         strEmailCC = "" & adoquery("emailcc")  'Added by Lydia 2024/09/18
         '檢查非大陸案有Email的代理人
         'Modify by Morgan 2007/3/3
         'Modify by Morgan 2008/10/30 大陸的也要了
         'If "" & adoquery("fa10") <> "020" And strEMailBox <> "" And UCase(strEMailBox) <> "NO" And "" & adoquery("a1b06") = "2" Then
         If strEMailBox <> "" And UCase(strEMailBox) <> "NO" And "" & adoquery("a1b06") = "2" Then
            'Remove by Morgan 2006/12/21 開始使用
            'strEMailBox = "jasjaswu@gmail.com" 'Add by Morgan 2006/12/14 測試用
            bolEmail = True
            bol2File = True
         Else
            bolEmail = False
            bol2File = False
         End If
         
         'Add by Morgan 2008/4/17 放在迴圈內是因為要印"是否EMail通知","不寄明細"的註記
         '產生電子檔
         If Text6.Text = "Y" Then
            bol2File = True
         '不發Mail 或 不寄收據 時只列印不產生電子檔
         'Modify by Morgan 2010/6/10
         'ElseIf Text5.Text = "N" Or strInform = "N" Then
         ElseIf Text5.Text = "1" Or strInform = "N" Then
         'end 2010/6/10
            bol2File = False
         End If
         
         'Added by Lydia 2016/10/21 票匯增加可Email的功能(不存電子檔),並且要列印
         'Modified by Lydia 2016/11/16 重新界定需求:選列印時票匯印2份電匯印1份,選email則都要發mail不做列印
         'bol2Email = False
         'If strEMailBox <> "" And UCase(strEMailBox) <> "NO" And "" & adoquery("a1b06") = "1" And Text5.Text = "2" And Text6 <> "Y" Then
         '   bol2Email = True
         '   bol2File = True
         '   bol2Printer = True
         'End If
         'end 2016/10/21
         'Remove by Lydia 2018/09/03 取消票匯的email通知
'         If strEMailBox <> "" And UCase(strEMailBox) <> "NO" And "" & adoquery("a1b06") = "1" Then
'            bol2Email = True
'            If Text5.Text = "2" And Text6 <> "Y" Then bol2File = True
'         Else
'            bol2Email = False
'         End If
'         'end 2016/11/16
         'end 2018/09/03
         
         'Add by Morgan 2008/11/20
         If iCopy = 2 Then
            'Added by Lydia 2016/10/21 排除email的票匯
           ' If bol2Email = False Then 'Remove by Lydia 2016/11/16 debug
                '不寄收據 或 EMail 或 存電子檔的不要印第二份
                If strInform = "N" Or bolEmail = True Or bol2File = True Then
                   bol2Printer = False
                Else
                   bol2Printer = True
                End If
                bol2File = False
                bolEmail = False
           ' End If  'end 2016/10/21
            
         'Added by Morgan 2012/10/11
         Else
            'Modified by Lydia 2018/09/03 + bol2pdf
            If Not (strInform = "N" Or bolEmail = True Or bol2File = True Or bol2Pdf = True) Then bolChgPrinter = True
            
         End If
         
         bol2Pdf = False 'Added by Lydia 2019/10/14 未還原判斷,導致票匯不需Email也會產生pdf且會加到下一個Email的附件內
         'Added by Lydia 2018/09/03 一張匯票存成一個PDF檔(發email 和存電子檔)
         If bol2File = True And (Text6 = "Y" Or (Text5 = "2" And Text6 <> "Y")) Then
            bol2Printer = False
            
            'Modified by Morgan 2024/10/30 因PDFCreator常出現無法預期的錯誤,改用Word將JPG轉PDF
            'bol2File = False
            'bol2Pdf = True
            bol2File = True
            'end 2024/10/30
            
            If bol2Pdf Then
               '預設PDF印表機
               'Modified by Lydia 2023/11/15
               'If strNo = "" Then
               If bolPdfStart = False Then
                   frmPDF.Show
                   frmPDF.StartProcess strSavePath, "$" & adoquery.Fields("a1b02").Value & adoquery.Fields("a1b01").Value & ".PDF"
               End If
            End If
         End If
         'end 2018/09/03
         
         strNo = adoquery.Fields("a1b02").Value
         m_a1b01 = adoquery.Fields("a1b01").Value   '2012/2/23 ADD BY SONIA
         m_a1b06 = adoquery.Fields("a1b06").Value 'Added by Lydia 2024/12/11
         
         'Added by Lydia 2017/10/05 代理人編號和名稱
         strTemp(0) = strNo
         strTemp(1) = Trim(adoquery.Fields("fa05") & " " & adoquery.Fields("fa63") & " " & adoquery.Fields("fa64") & " " & adoquery.Fields("fa65"))
         'end 2017/10/05
         
         '加信頭-存電子檔
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2File = True Or bol2Pdf = True Then
         
            'Added by Morgan 2020/3/30
            If strSrvDate(1) >= 智慧所更名日 Then
               m_iNo2 = m_iNo
               PUB_GetLetterPicID Text3, , m_iNo, , , , , True
               If m_iNo <> m_iNo2 Then
                  strPicLetter = ""
               End If
            Else
            'end 2020/3/30
               
               'Modify by Morgan 2011/10/11
               'm_iNo = 6
               '巨京Y52269000的付款明細請改用 台一國際專利商標事務所
               'Modify by Amy 2014/07/08 巨京Y52269000的付款明細為J公司用智權,其他仍用專利商標(2011/10/11)
               If strNo = "Y52269000" And Text3 <> "J" Then
                  If m_iNo <> 8 Then
                     m_iNo = 8
                     strPicLetter = ""
                  End If
               'Add by Amy 2014/05/01 +J公司使用智權公司抬頭
               ElseIf Text3 = "J" Then
                   If m_iNo <> 25 Then
                     m_iNo = 25
                     strPicLetter = ""
                  End If
               'end 2014/05/01
               Else
                  If m_iNo <> 6 Then
                     m_iNo = 6
                     strPicLetter = ""
                  End If
               End If
               'end 2011/10/11
               
            End If 'Added by Morgan 2020/3/30
            
            If strPicLetter = "" Then
               strPicLetter = App.path & "\$Tmp.jpg"
               If PUB_ReadDB2File(strPicLetter, m_iNo) = True Then
                  Set Picture1.Picture = LoadPicture(strPicLetter)
                  Picture1.AutoSize = True
                  douExtRate = Picture1.Height / 16836 'Add by Morgan 2011/8/10
                  'Added by Lydia 2018/09/03 PDF
                  If bol2Pdf = True Then
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
                       Printer.PaintPicture Picture1, 0, 0, Picture1.Width, Picture1.Height
                  End If
                  'end 2018/09/03
               End If
            End If
         End If
         
         'Added by Morgan 2012/3/14
         'Modify by Amy 2014/07/08 巨京Y52269000的付款明細為J公司用智權,其他仍用專利商標(2011/10/11)
         If strNo = "Y52269000" And Text3 <> "J" Then
            intTop = 250
            intBottom = 0
         'Added by Lydia 2018/09/03
         ElseIf bol2Pdf = True Then
            intTop = -100
            intPrtXFix = 0
            intBottom = 0
         'end 2018/09/03
         Else
            intTop = -200
            intPrtXFix = -400
            intBottom = 0
         End If
         
         'Modified by Lydia 2017/10/05 非清單才列印
         'PrintHead
         If bolList = False Then
            PrintHead
         End If
         'end 2017/10/05
      End If
      
      'Modified by Lydia 2017/10/05 非清單才列印
      'If intCounter > 20 Then
      If intCounter > 20 And bolList = False Then
         If intPage >= 1 Then
            If intCounter > 20 Then
               intPage = intPage + 1
               'Modified by Lydia 2018/09/03 + bol2pdf
               If bol2Printer = True Or bol2Pdf = True Then
                  PrintPageNo True
                  Printer.NewPage
                  'Added by Lydia 2018/09/03 列印信頭
                  If bol2Pdf = True Then
                      Printer.PaintPicture Picture1, 0, 0, Picture1.Width, Picture1.Height
                  End If
                  'end 2018/09/03
               End If
               If bol2File = True Then
                  PrintPageNo
                  PicNewPage
               End If
               intCounter = -10
               PrintHead
            Else
               intCounter = intCounter + 1
            End If
         Else
            intPage = intPage + 1
            'Modified by Lydia 2018/09/03 + bol2pdf
            If bol2Printer = True Or bol2Pdf = True Then
               PrintPageNo True
               Printer.NewPage
               'Added by Lydia 2018/09/03 列印信頭
               If bol2Pdf = True Then
                   Printer.PaintPicture Picture1, 0, 0, Picture1.Width, Picture1.Height
               End If
               'end 2018/09/03
            End If
            If bol2File = True Then
               PrintPageNo
               PicNewPage
            End If
            intCounter = -10
            PrintHead
         End If
      End If
      If IsNull(adoquery.Fields("axf03").Value) = False Then
         If Mid(adoquery.Fields("axf03").Value, Len(adoquery.Fields("axf03").Value) - 2, 3) = "000" Then
            strExc(0) = Mid(adoquery.Fields("axf03").Value, 1, Len(adoquery.Fields("axf03").Value) - 9) & "-" & Mid(adoquery.Fields("axf03").Value, Len(adoquery.Fields("axf03").Value) - 8, 6)
         Else
            strExc(0) = Mid(adoquery.Fields("axf03").Value, 1, Len(adoquery.Fields("axf03").Value) - 9) & "-" & Mid(adoquery.Fields("axf03").Value, Len(adoquery.Fields("axf03").Value) - 8, 6) & "-" & Mid(adoquery.Fields("axf03").Value, Len(adoquery.Fields("axf03").Value) - 2, 1) & "-" & Mid(adoquery.Fields("axf03").Value, Len(adoquery.Fields("axf03").Value) - 1, 2)
         End If
         'Added by Lydia 2017/10/05 本所案號
         If bolList = True Then
            strTemp(2) = "" & strExc(0)
         Else
         'end 2017/10/05
            'Modified by Lydia 2018/09/03 + bol2pdf
            If bol2Printer = True Or bol2Pdf = True Then
               'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
               'Printer.CurrentX = 0 + intMargin
               Printer.CurrentX = ciXOur + intMargin
               Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
               Printer.Print strExc(0)
            End If
            If bol2File = True Then
               'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
               'Picture1.CurrentX = (0 + intMargin) * douExtRate
               Picture1.CurrentX = (ciXOur + intMargin) * douExtRate
               Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
               Picture1.Print strExc(0)
            End If
         End If
         'end 2017/10/05
      End If
      
      strCurrency = "" & adoquery.Fields("a1505").Value
      'Added by Lydia 2017/10/05 DB.note 和幣別
      If bolList = True Then
         strTemp(3) = "" & adoquery.Fields("a1504").Value
         strTemp(4) = strCurrency
      Else
      'end 2017/10/05
         'YOUR REF和DB NOTE
         'Modified by Lydia 2018/09/03 + bol2pdf
         If bol2Printer = True Or bol2Pdf = True Then
            'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數,超出寬度折行
            'Printer.CurrentX = 2000 + intMargin
            'Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
            'Printer.Print "" & adoquery.Fields("cp45").Value
            'Printer.CurrentX = 4500 + intMargin
            'Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
            'Printer.Print "" & adoquery.Fields("a1504").Value
            'Printer.CurrentX = 6500 + intMargin
            'Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
            'Printer.Print strCurrency
            Printer.CurrentX = ciXYour + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
            If GetTextLength("" & adoquery.Fields("cp45").Value) > ciDBwidth Then
                strExc(1) = convForm("" & adoquery.Fields("cp45").Value, ciDBwidth)
                Printer.Print strExc(1)
                Printer.CurrentX = ciXYour + intMargin
                Printer.CurrentY = intTop + intPrtXFix + 6700 + (intCounter + 1) * 300 '折行
                Printer.Print convForm(Replace("" & adoquery.Fields("cp45").Value, strExc(1), ""), ciDBwidth)
            Else
                Printer.Print "" & adoquery.Fields("cp45").Value
            End If
            
            Printer.CurrentX = ciXDBnote + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
            If GetTextLength("" & adoquery.Fields("a1504").Value) > ciDBwidth Then
                strExc(1) = convForm("" & adoquery.Fields("a1504").Value, ciDBwidth)
                Printer.Print strExc(1)
                Printer.CurrentX = ciXDBnote + intMargin
                Printer.CurrentY = intTop + intPrtXFix + 6700 + (intCounter + 1) * 300 '折行
                Printer.Print convForm(Replace("" & adoquery.Fields("a1504").Value, strExc(1), ""), ciDBwidth)
            Else
                Printer.Print "" & adoquery.Fields("a1504").Value
            End If
            Printer.CurrentX = ciXCurr + intMargin
            Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
            Printer.Print strCurrency
            'end 2020/09/03
         End If
         If bol2File = True Then
            'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數,超出寬度折行
            'Picture1.CurrentX = (2000 + intMargin) * douExtRate
            'Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
            'Picture1.Print "" & adoquery.Fields("cp45").Value
            'Picture1.CurrentX = (4500 + intMargin) * douExtRate
            'Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
            'Picture1.Print "" & adoquery.Fields("a1504").Value
            'Picture1.CurrentX = (6500 + intMargin) * douExtRate
            'Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
            'Picture1.Print strCurrency
            Picture1.CurrentX = (ciXYour + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
            If GetTextLength("" & adoquery.Fields("cp45").Value) > ciDBwidth Then
                strExc(1) = convForm("" & adoquery.Fields("cp45").Value, ciDBwidth)
                Picture1.Print strExc(1)
                Picture1.CurrentX = (ciXYour + intMargin) * douExtRate
                Picture1.CurrentY = (intTop + 6700 + (intCounter + 1) * 300) * douExtRate '折行
                Picture1.Print convForm(Replace("" & adoquery.Fields("cp45").Value, strExc(1), ""), ciDBwidth)
            Else
                Picture1.Print "" & adoquery.Fields("cp45").Value
            End If
            Picture1.CurrentX = (ciXDBnote + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
            If GetTextLength("" & adoquery.Fields("a1504").Value) > ciDBwidth Then
                strExc(1) = convForm("" & adoquery.Fields("a1504").Value, ciDBwidth)
                Picture1.Print strExc(1)
                Picture1.CurrentX = (ciXYour + intMargin) * douExtRate
                Picture1.CurrentY = (intTop + 6700 + (intCounter + 1) * 300) * douExtRate '折行
                Picture1.Print convForm(Replace("" & adoquery.Fields("a1504").Value, strExc(1), ""), ciDBwidth)
            Else
                Picture1.Print "" & adoquery.Fields("a1504").Value
            End If
            Picture1.CurrentX = (ciXCurr + intMargin) * douExtRate
            Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
            Picture1.Print strCurrency
            'end 2020/09/03
         End If
      End If 'If bolList = True Then
      
      If IsNull(adoquery.Fields("axf04").Value) = False Then
         If Mid(adoquery.Fields("a1c03").Value, 1, 1) = "U" Then
            '若帳單編號及本所案號相同則金額合併
            strAmount = Val("" & adoquery.Fields("axf04").Value)
            adoquery.MoveNext
            If adoquery.EOF = True Then
                adoquery.MovePrevious
            Else
                Do While adoquery.Fields("axf03").Value = strAXF03 And adoquery.Fields("a1501").Value = strA1501
                    'Modify by Amy 2015/03/25 B與O單號皆為正數,只有V為負數
                    'strAmount = Val(strAmount) + Val(IIf(Mid(adoquery.Fields("a1c03").Value, 1, 1) = "U", "1", "-1")) * Val(adoquery.Fields("axf04").Value)
                    strAmount = Val(strAmount) + Val(IIf(Mid(adoquery.Fields("a1c03").Value, 1, 1) = "V", "-1", "1")) * Val(adoquery.Fields("axf04").Value)
                    adoquery.MoveNext
                    If adoquery.EOF = True Then Exit Do
                Loop
                adoquery.MovePrevious
            End If
            strAmount = Format(Val(strAmount), FDollar)
            'intLength = Printer.TextWidth(strAmount) 'Removed by Morgan 2024/11/4 移到下面依物件設定
            'Added by Lydia 2017/10/05 金額
            If bolList = True Then
                strTemp(5) = Val(CDbl(strAmount))
            Else
            'end 2017/10/05
                'Modified by Lydia 2018/09/03 + bol2pdf
                If bol2Printer = True Or bol2Pdf = True Then
                  intLength = Printer.TextWidth(strAmount)
                   'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
                   'Printer.CurrentX = 8400 - intLength + intMargin
                   Printer.CurrentX = ciXAmt - intLength + intMargin
                   Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
                   Printer.Print strAmount
                End If
                If bol2File = True Then
                   intLength = Picture1.TextWidth(strAmount)
                   'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
                   'Picture1.CurrentX = (8400 - intLength + intMargin) * douExtRate
                   'Modified by Morgan 2024/11/19 修正JPG格式的金額會被幣別蓋到問題
                   'Picture1.CurrentX = (ciXAmt - intLength + intMargin) * douExtRate
                   intLength = Picture1.TextWidth("9,999,999.00") - intLength
                   If intLength < 0 Then
                      intLength = Picture1.TextWidth("9")
                   Else
                      intLength = intLength + Picture1.TextWidth("9")
                   End If
                   Picture1.CurrentX = (ciXCurr + intMargin) * douExtRate + Picture1.TextWidth(strCurrency) + intLength
                   'end 2024/11/19
                   Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
                   Picture1.Print strAmount
                End If
            End If 'end 2017/10/05
            douAmount = douAmount + Val(CDbl(strAmount))
         Else
            '若帳單編號及本所案號相同則金額合併
            'Modify by Amy 2015/03/25 B與O單號皆為正數,只有V為負數
            'strAmount = "-" & Val("" & adoquery.Fields("axf04").Value)
            strAmount = Val(IIf(Mid(adoquery.Fields("a1c03").Value, 1, 1) = "V", "-1", "1")) * Val("" & adoquery.Fields("axf04").Value)
            adoquery.MoveNext
            If adoquery.EOF = True Then
                adoquery.MovePrevious
            Else
                Do While adoquery.Fields("axf03").Value = strAXF03 And adoquery.Fields("a1501").Value = strA1501
                    'Modify by Amy 2015/03/25 B與O單號皆為正數,只有V為負數
                    'strAmount = Val(strAmount) + Val(IIf(Mid(adoquery.Fields("a1c03").Value, 1, 1) = "U", "1", "-1")) * Val(adoquery.Fields("axf04").Value)
                    strAmount = Val(strAmount) + Val(IIf(Mid(adoquery.Fields("a1c03").Value, 1, 1) = "V", "-1", "1")) * Val(adoquery.Fields("axf04").Value)
                    adoquery.MoveNext
                    If adoquery.EOF = True Then Exit Do
                Loop
                adoquery.MovePrevious
            End If
            strAmount = Format(Val(strAmount), FDollar)
            'Added by Lydia 2017/10/05 金額
            If bolList = True Then
                strTemp(5) = Val(CDbl(strAmount))
            Else
            'end 2017/10/05
                'Modified by Lydia 2018/09/03 + bol2pdf
                If bol2Printer = True Or bol2Pdf = True Then
                   intLength = Printer.TextWidth(strAmount)
                   'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
                   'Printer.CurrentX = 8400 - intLength + intMargin
                   Printer.CurrentX = ciXAmt - intLength + intMargin
                   Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
                   Printer.Print strAmount
                End If
                If bol2File = True Then
                   intLength = Picture1.TextWidth(strAmount)
                   'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
                   'Picture1.CurrentX = (8400 + intMargin) * douExtRate - intLength
                   'Modified by Morgan 2024/11/4
                   'Picture1.CurrentX = (ciXAmt + intMargin) * douExtRate - intLength
                   Picture1.CurrentX = (ciXAmt + intMargin - intLength) * douExtRate
                   'end 2024/11/4
                   Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
                   Picture1.Print strAmount
                End If
            End If 'end 2017/10/05
            douAmount = douAmount + Val(CDbl(strAmount))
         End If
      End If
      'Remove by Lydia 2020/09/03 調整明細表的版面:Draft NO. 改到上方
      'Select Case adoquery.Fields("a1b06").Value
      '   Case "1"
      '      'Added by Lydia 2017/10/05
      '      If bolList = True Then
      '      Else
      '      'end 2017/10/05
      '          'Modified by Lydia 2018/09/03 + bol2pdf
      '          If bol2Printer = True Or bol2Pdf = True Then
      '             Printer.CurrentX = 8700 + intMargin
      '             Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
      '             Printer.Print "" & adoquery.Fields("a1b01").Value
      '          End If
      '          If bol2File = True Then
      '             Picture1.CurrentX = (8700 + intMargin) * douExtRate
      '             Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
      '             Picture1.Print "" & adoquery.Fields("a1b01").Value
      '          End If
      '      End If 'end 2017/10/05
      'End Select
      'end 2020/09/03
      intCounter = intCounter + 1
      If Not IsNull(adoquery.Fields("a1502").Value) Then
         strExc(0) = Format(AFDate(CADate(adoquery.Fields("a1502").Value)), "mmm. d, yyyy")
         'Added by Lydia 2017/10/05
         If bolList = True Then
         Else
         'end 2017/10/05
            'Modified by Lydia 2018/09/03 + bol2pdf
            If bol2Printer = True Or bol2Pdf = True Then
               'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
               'Printer.CurrentX = 4500 + intMargin
               Printer.CurrentX = ciXOur + intMargin
               Printer.CurrentY = intTop + intPrtXFix + 6700 + intCounter * 300
               Printer.Print strExc(0)
            End If
            If bol2File = True Then
               'Modified by Lydia 2020/09/03 調整明細表的版面:改用變數
               'Picture1.CurrentX = (4500 + intMargin) * douExtRate
               Picture1.CurrentX = (ciXOur + intMargin) * douExtRate
               Picture1.CurrentY = (intTop + 6700 + intCounter * 300) * douExtRate
               Picture1.Print strExc(0)
            End If
         End If 'end 2017/10/05
      End If
      
      intCounter = intCounter + 1
      If InStr(m_Curs, strCurrency) = 0 Then
         m_Len = m_Len + 1
         ReDim Preserve m_Cur(m_Len)
         ReDim Preserve m_Amount(m_Len)
         m_Curs = m_Curs & "," & strCurrency
         m_Cur(m_Len) = strCurrency
         m_Amount(m_Len) = Val(CDbl(strAmount))
      Else
         For m_Idx = 1 To UBound(m_Cur)
            If m_Cur(m_Idx) = strCurrency Then
               m_Amount(m_Idx) = m_Amount(m_Idx) + Val(CDbl(strAmount))
               Exit For
            End If
         Next
      End If
      
      'Added by Lydia 2017/10/05 暫存清單資料
      If bolList = True Then
         mCnt = mCnt + 1 'Added by Lydia 2017/10/05
         'Modified by Morgan 2018/2/26 +adoquery.Fields("a1719")
         'Modified by Lydia 2018/11/19 +是否下載Invoice(R24507),帳單編號(R24508)
         'strExc(0) = "insert into accrpt2450 (id,r24500,r24501,r24502,r24503,r24504,r24505,r24506) values (" & CNULL(strUserNum) & _
                     ", " & mCnt & ", " & CNULL(strTemp(0)) & ", " & CNULL(ChgSQL(PUB_StrToStr(strTemp(1), 150))) & " ," & CNULL(strTemp(2)) & "," & CNULL(PUB_StrToStr(IIf("" & adoquery.Fields("a1719") = "Y", "＊", "") & strTemp(3), 100)) & " ," & CNULL(strTemp(4)) & " ," & strTemp(5) & ") "
         strExc(0) = "insert into accrpt2450 (id,r24500,r24501,r24502,r24503,r24504,r24505,r24506,r24507,r24508) values (" & CNULL(strUserNum) & _
                     ", " & mCnt & ", " & CNULL(strTemp(0)) & ", " & CNULL(ChgSQL(PUB_StrToStr(strTemp(1), 150))) & " ," & CNULL(strTemp(2)) & "," & CNULL(PUB_StrToStr(IIf("" & adoquery.Fields("a1719") = "Y", "＊", "") & strTemp(3), 100)) & " ," & CNULL(strTemp(4)) & " ," & strTemp(5) & _
                     ",'" & IIf("" & adoquery.Fields("a1812") = "Y", "Y", "") & "', " & CNULL("" & adoquery.Fields("axf01")) & ") "
         cnnConnection.Execute strExc(0)
      End If
      'end 2017/10/05
NextSkip:
      adoquery.MoveNext
   Loop
'   If adoquery.RecordCount <> 0 Then
'      adoquery.MoveLast
'   End If
   'Modified by Lydia 2017/10/05
   'PrintSum
   'Printer.EndDoc
   If bolList = True Then
        mCnt = mCnt + 1
        strExc(0) = "insert into accrpt2450 (id,r24500,r24501,r24502,r24503,r24504,r24505,r24506) values (" & CNULL(strUserNum) & _
                    ", " & mCnt & ", NULL, NULL ,'小　計',NULL,NULL," & douAmount & ") "
        cnnConnection.Execute strExc(0)
   Else
       PrintSum
   End If
   If bolPrintPage = True Then Printer.EndDoc
   'end 2017/10/05
   
   'Modified by Lydia 2017/10/05 非清單才印
   'If bol2File = True Then
   'Modified by Lydia 2018/09/03 + bol2pdf
   If bol2File = True Or bol2Pdf = True Then
      'Modified by Morgan 2014/7/21 +Y51895 103/7/15 檔案產生錯誤
      'strExc(1) = strSavePath & "\$" & strNo & IIf(intPage > 0, "_" & Format(intPage, "00"), "") & ".jpg"
      If bol2File = True Then 'Added by Lydia 2018/09/03
          strExc(1) = strSavePath & "\$" & strNo & m_a1b01 & IIf(intPage > 0, "_" & Format(intPage, "00"), "") & ".jpg"
          PUB_SavePic Picture1, strExc(1)
      'Added by Lydia 2018/09/03
      ElseIf bol2Pdf = True Then
          strExc(1) = strSavePath & "\$" & strNo & m_a1b01 & ".PDF"
          frmPDF.EndtProcess
          bolPdfStart = False 'Added by Lydia 2023/11/15
          'Added by Lydia 2019/08/21 判斷檔案是否存在, 超過時間就繼續
          'Modified by Lydia 2020/02/15 +改成共用
          'If ChkFileStatus(strExc(1)) = False Then
          'Modified by Lydia 2020/09/10 超過時間，改不發email直接出發信失敗清單
          'If PUB_ChkFileStatus(strExc(1)) = False Then
          tmpErr = ""
          If PUB_ChkFileStatus(strExc(1), False, tmpErr) = False Then
              If tmpErr <> "" Then
                  If strMailFailList(0) <> "" Then
                     ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                  End If
                  strMailFailList(UBound(strMailFailList)) = strNo & "：" & Mid(tmpErr, 2) & IIf((bolEmail = True Or bol2Email = True) And Text5.Text = "2" And Text6.Text <> "Y", "請重新Email給該代理人", "")
              End If
          'end 2020/09/10
          End If
          'end 2019/08/21
          Unload frmPDF
      End If
      'end 2018/09/03
      
      If tmpErr = "" Then 'Added by Lydia 2020/09/10 增加判斷
         strPicFileNames = strPicFileNames & strExc(1) & "*"

         'Added by Morgan 2024/10/30 因PDFCreator常出現無法預期的錯誤,改用Word將JPG轉PDF
         If bol2File = True Then
            strExc(1) = strSavePath & "\$" & strNo & m_a1b01 & ".PDF"
            'Modified by Morgan 2024/9/4 中文要加印收據章
            If PUB_JPG2PDF(strPicFileNames, strExc(1)) = True Then
               strPicFileNames = strExc(1)
            End If
         End If
         'end 2024/10/30
      End If
              
      'Modify by Morgan 2010/6/10 選發EMail且未設要存電子檔才寄送
      'If bolEmail = True And bol2Printer = True Then
      'Modified by Lydia 2016/10/21 票匯增加可Email的功能
      'If bolEmail = True And Text5.Text = "2" And Text6.Text <> "Y" Then
      'Modified by Lydia 2016/11/16 重新界定需求:選列印時票匯印2份電匯印1份,選email則都要發mail不做列印
      'If (bolEmail = True Or (bol2Email = True And iCopy = 2)) And Text5.Text = "2" And Text6.Text <> "Y" Then
      'Modified by Lydia 2020/09/10 判斷有檔案才寄信
      'If (bolEmail = True Or bol2Email = True) And Text5.Text = "2" And Text6.Text <> "Y" Then
      If (bolEmail = True Or bol2Email = True) And Text5.Text = "2" And Text6.Text <> "Y" And strPicFileNames <> "" Then
         bolMailFailNoAlert = True
         bolMailSendOk = False
         txtSend.Text = strNo   'Added by Lydia 2020/10/08
         'Modify by Morgan 2011/4/22 改以ipdept@taie.com.tw 寄但回覆還是給寄件人(70004)
         'PUB_SendMail strUserNum, strEMailBox, "", "ADVICE OF WIRE PAYMENT", GetMailContent, , strPicFileNames, True, True, True
         'Modified by Morgan 2011/10/12 改用 account@taie.com.tw 寄
         'PUB_SendMail strUserNum, strEMailBox, "", "ADVICE OF WIRE PAYMENT", GetMailContent, , strPicFileNames, True, True, True, , "ipdept@taie.com.tw", "TAI E INTERNATIONAL PATENT & LAW OFFICE", strUserNum
         'Modify by Amy 2014/05/01 改公司抬頭
         'Modified by Morgan 2014/8/27 改回覆到財務信箱 -- 婧瑄
         'PUB_SendMail strUserNum, strEMailBox, "", "ADVICE OF WIRE PAYMENT", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, strCompName, strUserNum
         'Modified by Lydia 2016/10/19 "ADVICE OF WIRE PAYMENT" => "ADVICE OF WIRE/CHECK PAYMENT"
         'Moified by Lydia 2020/09/22 +代理人編號 ADVICE OF WIRE/CHECK PAYMENT=> ADVICE OF WIRE/CHECK PAYMENT (Yxxxxx)
         'Modified by Lydia 2024/09/18 +財務副本信箱strEmailCC
         PUB_SendMail strUserNum, strEMailBox, "", "ADVICE OF WIRE/CHECK PAYMENT (" & strNo & ")", GetMailContent, , strPicFileNames, True, True, True, strEmailCC, strAccMailBox, strCompName, strAccMailBox
         bolMailFailNoAlert = False
         If bolMailSendOk = False Then
            If strMailFailList(0) <> "" Then
               ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
            End If
            strMailFailList(UBound(strMailFailList)) = strNo & "：" & strEMailBox
         End If
      Else
         If strPicFileNames <> "" Then
            MsgBox "電子檔已存桌面！"
         End If
      End If
      '刪除舊的暫存圖檔
      strExc(1) = App.path & "\$*.jpg"
      If Dir(strExc(1)) <> "" Then Kill strExc(1)
      'Added by Lydia 2018/09/03 刪除舊的PDF檔
      strExc(1) = App.path & "\$*.pdf"
      If Dir(strExc(1)) <> "" Then Kill strExc(1)
      'end 2018/09/03
   End If
   
   'Add by Morgan 2007/1/24
   If strMailFailList(0) <> "" Then
      strExc(0) = "E-Mail失敗清單：" & vbCrLf & vbCrLf
      For intI = 0 To UBound(strMailFailList)
         strExc(0) = strExc(0) & strMailFailList(intI) & vbCrLf
      Next
      'Modified by Lydia 2020/09/10 改成彈訊息，並且直接寄email給使用者
      'If MsgBox(strExc(0) & vbCrLf & "是否要列印？" & vbCrLf, vbYesNo + vbDefaultButton1) = vbYes Then
      '   Printer.Print strExc(0)
      '   Printer.EndDoc
      'End If
      PUB_SendMail strUserNum, strUserNum, "", Me.Caption & "-" & "E-Mail失敗清單", vbCrLf & strExc(0)
      MsgBox strExc(0) & vbCrLf & "請參考！", vbInformation, Me.Caption & "-" & "E-Mail失敗清單"
      'end 2020/09/10
   End If
   'end 2007/1/24
         
Next

   adoquery.Close
   
   txtSend.Visible = False: lblSend.Visible = False  'Added by Lydia 2020/10/08
'Added by Lydia 2019/05/28
    Exit Sub
    
ErrorHandle:
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Resume Next
    End If
End Sub

'Modify by Morgan 2008/3/10
Private Function GetMailContent() As String
   Dim StrMailContent As String
   'Modified by Morgan 2024/4/10 對外統一用 Dear Colleagues --林總
   'StrMailContent = "Dear Sirs," & vbCrLf
   StrMailContent = "Dear Colleagues," & vbCrLf
   'end 2024/4/10
   'Modified by Morgan 2015/10/16 內容調整
   'StrMailContent = StrMailContent & vbCrLf & "A payment of " & strPayAmount & " has been sent to your account in settlement of your debit notes as the attachment." & vbCrLf
   'Modified by Morgan 2017/8/23 票匯調整內容
   If m_a1b06 = "1" Then
      StrMailContent = StrMailContent & vbCrLf & "A check of " & strPayAmount & " has been sent to you by registered mail today in settlement of your invoices as below." & vbCrLf
   Else
      StrMailContent = StrMailContent & vbCrLf & "A payment of " & strPayAmount & " has been sent to your account in settlement of your invoices as below." & vbCrLf
      'Added by Lydia 2018/09/06 改變內文
      If bol2Pdf = True Then
          StrMailContent = Replace(StrMailContent, "your invoices as below.", "your invoices as attachment.")
      End If
      'end 2018/09/06
   End If
   'end 2017/8/23
   StrMailContent = StrMailContent & vbCrLf & "Please be advised that this e-mail address be reserved for account matters only."
   StrMailContent = StrMailContent & vbCrLf & "Please direct all case matter to ipdept@taie.com.tw to ensure the quickest reply." & vbCrLf
   StrMailContent = StrMailContent & vbCrLf & "Thank you for your services and do not hesitate to contact us regarding accounts matters." & vbCrLf
   'Modify by Amy 2014/05/01 改抓acc080.a0803
   'StrMailContent = StrMailContent & vbCrLf & "TAI E INTERNATIONAL PATENT & LAW OFFICE"
   StrMailContent = StrMailContent & vbCrLf & strCompName
   'end 2014/05/01
   StrMailContent = StrMailContent & vbCrLf & "Accounting Department"
   GetMailContent = StrMailContent & vbCrLf & vbCrLf & vbCrLf & vbCrLf
End Function
'end 2008/3/10

Private Sub PrintPageNo(Optional bPrinter As Boolean)
   Dim strPages As String
   
   strPages = "**" & intPage & "**"
   If bPrinter Then
      Printer.CurrentX = 5000
      Printer.CurrentY = Printer.Height - (1700 + intBottom)
      Printer.Print strPages
   Else
      Picture1.CurrentX = Picture1.Width / 2 - Picture1.TextWidth(strPages) / 2
      Picture1.CurrentY = Picture1.Height - (1700 + intBottom) * douExtRate
      Picture1.Print strPages
   End If
End Sub

Private Sub PicNewPage()
   'Modified by Morgan 2014/7/21 +Y51895 103/7/15 檔案產生錯誤
   'strExc(1) = strSavePath & "\$" & strNo & IIf(intPage > 0, "_" & Format(intPage, "00"), "") & ".jpg"
   strExc(1) = strSavePath & "\$" & strNo & m_a1b01 & IIf(intPage > 0, "_" & Format(intPage, "00"), "") & ".jpg"
   PUB_SavePic Picture1, strExc(1)
   strPicFileNames = strPicFileNames & strExc(1) & "*"
   Set Picture1.Picture = LoadPicture(strPicLetter)
   Picture1.AutoSize = True
   douExtRate = Picture1.Height / 16836 'Add by Morgan 2011/8/10
End Sub

'Add by Amy 2014/05/01
Private Sub Text3_Change()
    If Text3 = MsgText(601) Then
        Text13 = ""
        Exit Sub
    End If
    If Text3 = "1" Or Text3 = "J" Then
        Text13 = A0802Query(Text3)
    End If
End Sub

Private Sub Text3_GotFocus()
    TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    If Text3 = "" Then Exit Sub
    If Text3 <> "1" And Text3 <> "J" Then
        Text13 = ""
        MsgBox "公司別輸入錯誤請確認 ！"
        Cancel = True
        Exit Sub
    End If
End Sub
'end 2014/05/01

'Add by Morgan 2007/1/24
Private Sub Text5_GotFocus()
   CloseIme
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2018/1/19
   'If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
   '   KeyAscii = 0
   If KeyAscii = Asc("3") Then
      Combo2.ListIndex = -1
   ElseIf KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   ElseIf Combo2.ListIndex = -1 Then
      Combo2.ListIndex = 0
   'end 2018/1/19
   End If
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
   End If
End Sub

'Add by Amy 2014/05/01 +抓公司名稱(英)
Private Function A0803Query(ByVal InputNo As String) As String
    Dim adoacc080 As New ADODB.Recordset
   
   adoacc080.CursorLocation = adUseClient
   adoacc080.Open "Select * From acc080 Where a0801 = '" & InputNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc080.RecordCount <> 0 Then
      If IsNull(adoacc080.Fields("a0803").Value) Then
         A0803Query = MsgText(601)
      Else
         A0803Query = adoacc080.Fields("a0803").Value
      End If
   Else
      A0803Query = MsgText(601)
   End If
   adoacc080.Close
End Function
'end 2014/05/01

'Add by Amy 2015/04/08 結匯完成寄信通知國外部-原2015/03/19 程式改至此
Private Sub MailToTaie()
    Dim adoQ As New ADODB.Recordset
    Dim strQ As String
    Dim StrMailContent As String, strSubject As String 'Mail 內容,主旨
    
    strNo = ""
    '抓取B單號相關資料
    'Modify by Amy 2016/03/17 婉莘通知科目調整 原:6120
    'Modified by Morgan 2024/11/8 +612003 國外廣告-國際會議相關,612004 國外廣告 -其他 --斯閔
    'strQ = "Select a1b02 as 代理人,a1b03 as 結匯日期,a1p14 as 摘要,a1706 as Invoice,a2207 as 帳戶,a2208||' '||a2209 as 帳號,a1b01 as 匯票號碼 " & _
                "From acc1c0, acc1b0,acc170, fagent,acc190,acc1p0,acc220 " & _
                "Where a1c01 = a1b01 And a1c02 = a1b02 And a1c03 = a1702 (+) And a1b01=a1908 And a1c03=a1902(+) And SubStr(a1c03,1,1)= 'B' " & _
                "And SubStr(a1705, 1, 8) = fa01 (+) and SubStr(a1705, 9, 1) = fa02 (+) And InStr(fa29,'財務處結匯用')>0 And (fa84<>'N' or fa84 is null) " & _
                "And a1p04=a1c01||a1c02 And (''||a1p05='612001' Or ''||a1p05='612002') And SubStr(a1p14,InStr(a1p14,'單號:')+3,9)=a1c03 And a1705=a2201(+) And a1703=a2202(+) " & _
                "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " And a1702 is not null" & strConSql
    strQ = "Select a1b02 as 代理人,a1b03 as 結匯日期,a1p14 as 摘要,a1706 as Invoice,a2207 as 帳戶,a2208||' '||a2209 as 帳號,a1b01 as 匯票號碼 " & _
                "From acc1c0, acc1b0,acc170, fagent,acc190,acc1p0,acc220 " & _
                "Where a1c01 = a1b01 And a1c02 = a1b02 And a1c03 = a1702 (+) And a1b01=a1908 And a1c03=a1902(+) And SubStr(a1c03,1,1)= 'B' " & _
                "And SubStr(a1705, 1, 8) = fa01 (+) and SubStr(a1705, 9, 1) = fa02 (+) And InStr(fa29,'財務處結匯用')>0 And (fa84<>'N' or fa84 is null) " & _
                "And a1p04=a1c01||a1c02 And a1p05 in ('612001','612002','612003','612004') And SubStr(a1p14,InStr(a1p14,'單號:')+3,9)=a1c03 And a1705=a2201(+) And a1703=a2202(+) " & _
                "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " And a1702 is not null" & strConSql
    strQ = strQ & " Order by a1b02 asc, a1b01"
    
    adoQ.CursorLocation = adUseClient
    adoQ.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoQ.RecordCount <> 0 Then
        With adoQ
            Do While Not .EOF
                strSubject = "" & .Fields("摘要")
                StrMailContent = "您好," & vbCrLf 'Modify by Amy 2017/11/21 原:Kelly
                StrMailContent = StrMailContent & vbCrLf & ChangeTStringToTDateString("" & .Fields("結匯日期")) & " 匯款支付 "
                StrMailContent = StrMailContent & vbCrLf & strSubject
                StrMailContent = StrMailContent & vbCrLf & "Invoice no. : " & "" & .Fields("Invoice") & vbCrLf
                StrMailContent = StrMailContent & vbCrLf & "匯款帳號資料如下:"
                StrMailContent = StrMailContent & vbCrLf & "帳戶:" & .Fields("帳戶")
                StrMailContent = StrMailContent & vbCrLf & "帳號:" & .Fields("帳號") & vbCrLf
                StrMailContent = StrMailContent & vbCrLf & "特此告知"
                StrMailContent = StrMailContent & vbCrLf & vbCrLf
                
                PUB_SendMail strUserNum, strMailTaie, "", strSubject, StrMailContent, , , , , , , strAccMailBox
               
                .MoveNext
            Loop
        End With
        MsgBox "結匯完成已發Mail 給相關人員!" 'Modify by Amy 2017/11/21 原:Kelly
    End If
End Sub

'Added by Lydia 2017/10/05 列印清單
Private Sub PrintList()
Dim RsQ As New ADODB.Recordset
Dim intQ As Integer
Dim strGrp As String

   strGrp = "select * from accrpt2450 where id = '" & strUserNum & "' order by r24500 "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strGrp)
   If intQ = 1 Then
      SettingPrtSet
      With RsQ
          .MoveFirst
          PrintListHead
          Do While Not .EOF
             If strGrp <> "" & .Fields("R24501") Then
                '代理人編號
                strTemp(0) = "" & .Fields("R24501")
                '代理人名稱
                strTemp(1) = convForm("" & .Fields("R24502"), 26)
             Else
                strTemp(0) = ""
                strTemp(1) = ""
             End If
             
             '本所案號
             strTemp(2) = "" & .Fields("R24503")
             '代理人DN No.
             strTemp(3) = convForm("" & .Fields("R24504"), 16)
             '幣別
             strTemp(4) = "" & .Fields("R24505")
             '金額
             strTemp(5) = Format("" & .Fields("R24506"), FDollar)

             If Trim(strTemp(2)) = "小　計" Then
                Printer.Line (PLeft(0), iPrint)-(PLeft(7), iPrint)
                iPrint = iPrint + 150
                Printer.Font.Bold = True
             Else
                Printer.Font.Bold = False
             End If
             
             For intQ = 0 To UBound(PLeft) - 2
                If intQ = 5 Then
                   Printer.CurrentX = PLeft(intQ + 1) - ciColGap - Printer.TextWidth(strTemp(intQ))
                Else
                   Printer.CurrentX = PLeft(intQ)
                End If
                Printer.CurrentY = iPrint
                Printer.Print strTemp(intQ)
             Next
             PrintNewLine
             strGrp = "" & .Fields("R24501")
             .MoveNext
          Loop
          
          Printer.EndDoc
          'Remove by Lydia 2017/11/01
          'MsgBox "清單列印完成!"
      End With
   End If
   Set RsQ = Nothing
   
   'Added by Lydia 2018/11/19 匯出特殊結匯帳單Invoice
   If ExportFile("2") Then
        MsgBox "匯出完成！", vbExclamation
   End If
   
End Sub

'清單-換行
Private Sub PrintNewLine(Optional ByVal mRate As Single = 1, Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight * mRate
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      Printer.NewPage
      PrintListHead
   End If
End Sub

'清單-表頭
Private Sub PrintListHead()
Dim iPos As Integer

iPrint = ciStartY

strExc(0) = "國外付款明細表整批清單"
Printer.Font.Size = ciTitleFontSize
Printer.Font.Bold = True
Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strExc(0))) / 2
Printer.CurrentY = iPrint
Printer.Print strExc(0)

Printer.Font.Size = ciFontSize
Printer.Font.Bold = True
PrintNewLine
PrintNewLine

'列印條件
strExc(1) = ""
If Text3.Text <> "" Then
   strExc(1) = "公司別：" & Text3.Text & " " & Text13.Text & " "
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(strExc(1))) / 2
   Printer.CurrentY = iPrint
   Printer.Print strExc(1)
   PrintNewLine
End If
strExc(1) = ""
If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
   strExc(1) = "結匯日期：" & FCDate(MaskEdBox1.Text) & " ~ "
End If
If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
   strExc(1) = strExc(1) & IIf(InStr(strExc(1), "結匯日期：") = 0, "結匯日期： ~ ", "") & FCDate(MaskEdBox2.Text)
End If
If strExc(1) <> "" Then
   iPos = (Printer.ScaleWidth - Printer.TextWidth(strExc(1))) / 2
   Printer.CurrentX = iPos
   Printer.CurrentY = iPrint
   Printer.Print strExc(1)
   PrintNewLine
End If

strExc(1) = ""
If Text1.Text <> "" Then
   strExc(1) = "代理人：" & Text1.Text & " ~ "
End If
If Text2.Text <> "" Then
   strExc(1) = strExc(1) & IIf(InStr(strExc(1), "代理人：") = 0, "代理人： ~ ", "") & Text2.Text
End If
If strExc(1) <> "" Then
   If iPos = 0 Then iPos = (Printer.ScaleWidth - Printer.TextWidth(strExc(1))) / 2
   Printer.CurrentX = iPos
   Printer.CurrentY = iPrint
   Printer.Print strExc(1)
   PrintNewLine
End If

Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'Added by Morgan 2018/2/26 --婉莘
strExc(1) = "有＊者為ｅ化帳單！"
iPos = (Printer.ScaleWidth - Printer.TextWidth(strExc(1))) / 2
Printer.CurrentX = iPos
Printer.CurrentY = iPrint
Printer.Print strExc(1)
'end 2018/2/26
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "頁　　數：" & Printer.Page

PrintNewLine
For intI = 0 To UBound(PLeft) - 2
   Select Case intI
     Case 0: strExc(1) = "代理人編號"
     Case 1: strExc(1) = "代理人名稱"
     Case 2: strExc(1) = "本所案號"
     Case 3: strExc(1) = "代理人DN No."
     Case 4: strExc(1) = "幣別"
     Case 5: strExc(1) = "金　　額"
   End Select
   If intI = 5 Then
      Printer.CurrentX = PLeft(intI + 1) - ciColGap - Printer.TextWidth(strExc(1))
   Else
      Printer.CurrentX = PLeft(intI)
   End If
   Printer.CurrentY = iPrint
   Printer.Print strExc(1)
Next intI
PrintNewLine

Printer.Font.Bold = False
Printer.Line (PLeft(0), iPrint)-(PLeft(7), iPrint)
iPrint = iPrint + 150

End Sub

'清單-設定印表機
Private Sub SettingPrtSet()
Dim inX As Integer
Dim tmpArr As Variant, tmpArr2 As Variant

    '設定印表機
    PUB_RestorePrinter cmbPrinter
    Printer.EndDoc
    Printer.PaperSize = 9  'A4
    Printer.Orientation = 1 '1.直印
    
    lngPageHeight = Printer.ScaleHeight
    lngPageWidth = Printer.ScaleWidth
    lngLineHeight = 300
    Printer.Font.Name = "新細明體"
    Printer.Font.Size = ciFontSize
    Erase PLeft
    '代理人編號  /代理人名稱  /本所案號  /代理人DN No.  /幣別  /金額
    PLeft(0) = ciStartX
    PLeft(1) = PLeft(0) + Printer.TextWidth(String(5, "　")) + ciColGap     '+代理人編號
    PLeft(2) = PLeft(1) + Printer.TextWidth(String(15, "　")) + ciColGap    '+代理人名稱
    PLeft(3) = PLeft(2) + Printer.TextWidth(String(7, "　")) + ciColGap     '+本所案號
    PLeft(4) = PLeft(3) + Printer.TextWidth(String(8, "　")) + ciColGap     '+代理人DN No.
    PLeft(5) = PLeft(4) + Printer.TextWidth(String(2, "　")) + ciColGap     '+幣別
    PLeft(6) = PLeft(5) + Printer.TextWidth(String(7, "　")) + ciColGap     '+金額
    PLeft(7) = PLeft(6) + ciColGap
End Sub

'Added by Morgan 2018/1/19
'匯出帳單/抵帳單電子檔
'Modified by Lydia 2018/11/19 +特殊結匯帳單
'Private Function ExportFile() As Boolean
Private Function ExportFile(Optional ByVal iKind As String = "1") As Boolean
Dim stSavePath As String, stFileName As String
Dim strGrp As String 'Added by Lydia 2018/11/19
   
   If iKind = "1" Then 'Added by Lydia 2018/11/19 匯出帳單/抵帳單電子檔
        strExc(0) = "Select distinct ayf01,ayf02,a1b02,a1b03" & _
           " From acc1b0, acc1c0, acc190, acc152" & _
           " Where a1c01(+)=a1b01 and a1c02(+)=a1b02 and a1908(+)=a1c01 And a1902(+)=a1c03 " & _
           " And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & strConSql & _
           " and ayf01(+)=a1902 and ayf02 is not null"
   'Added by Lydia 2018/11/19 特殊結匯帳單
   ElseIf iKind = "2" Then
        strExc(0) = "select a.*,ayf01,ayf02 from accrpt2450 a, acc152 b " & _
                          "where id='" & strUserNum & "' and r24507='Y' and r24508 is not null and r24508=ayf01(+) and ayf02 is not null "
   End If
   'end 2018/11/19
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         If iKind = "1" Then 'Added by Lydia 2018/11/19 匯出帳單/抵帳單電子檔
                stFileName = .Fields("a1b02") & "_" & .Fields("ayf01") & "_" & .Fields("ayf02")
                stSavePath = txtPath & "\" & .Fields("a1b03")
                If PUB_GetAttachFile_Invoice(.Fields("ayf01"), .Fields("ayf02"), stSavePath, stFileName) = False Then
                   MsgBox "無法匯出檔案[ " & stFileName & " ]！"
                   Exit Function
                End If
         'Added by Lydia 2018/11/19 特殊結匯帳單
         Else
                '案號 幣別 金額(預設無小數點) 帳單號碼
                stFileName = Replace(.Fields("R24503"), "-", "") & " " & .Fields("R24505") & " " & .Fields("R24506") & " " & .Fields("R24508") & ".pdf"
                If strGrp = stFileName Then
                    stFileName = Replace(.Fields("R24503"), "-", "") & " " & Mid(.Fields("ayf02"), 1, InStrRev(.Fields("ayf02"), ".") - 1) & " " & .Fields("R24505") & " " & .Fields("R24506") & " " & .Fields("R24508") & ".pdf"
                End If
                stSavePath = txtPath2
                If PUB_GetAttachFile_Invoice(.Fields("ayf01"), .Fields("ayf02"), stSavePath, stFileName) = False Then
                   MsgBox "無法匯出檔案[ " & stFileName & " ]！"
                   Exit Function
                End If
                strGrp = stFileName
         End If
         'end 2018/11/19
         .MoveNext
      Loop
      End With
      If iKind = "1" Then 'Added by Lydia 2018/11/19
          ExportFile = True
      'Added by Lydia 2018/11/19
      ElseIf iKind = "2" And strGrp <> "" Then
          ExportFile = True
      End If
      'end 2018/11/19
   Else
      MsgBox "無電子檔可匯出！", vbInformation
   End If
End Function

'Added by Lydia 2018/11/19
Private Sub cmdPath2_Click()
   Dim fName As String, strStartFolder As String
   
   If Dir(txtPath2 & "\", vbDirectory) <> "" Then strStartFolder = txtPath2
   
   fName = PUB_GetFolder(Me.hWnd, strStartFolder, "請選取資料夾:")
   If fName <> "" Then 'they did not hit cancel
      txtPath2 = fName
      SaveSetting "TAIE", "B", UCase(Me.Name) & "Dir", txtPath2
   End If
End Sub

'Added by Lydia 2024/12/11 Excel列印-匯票信頭、信尾
Private Function PrintExcel_BFile(ByVal bolOpenFile As Boolean, ByVal iPicNo1 As Integer, Optional ByVal iPicNo2 As Integer) As Boolean
Dim strPic01 As String, strPic02 As String '信頭、信尾下載檔案路徑

   If bolOpenFile = True Then
      strPrtFile = strPrtPath & "\$" & Me.Caption & "-匯票" & MsgText(43)
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
      WksRpt1.PageSetup.HeaderMargin = Excel.Application.InchesToPoints(0.3) '頁首
      WksRpt1.PageSetup.FooterMargin = Excel.Application.InchesToPoints(0.3) '頁尾
      WksRpt1.PageSetup.TopMargin = xlsRpt.InchesToPoints(0.2) '上
      WksRpt1.PageSetup.BottomMargin = xlsRpt.InchesToPoints(0.2) '下
      WksRpt1.PageSetup.LeftMargin = xlsRpt.InchesToPoints(0.1) '左邊界
      WksRpt1.PageSetup.RightMargin = xlsRpt.InchesToPoints(0.1) '右邊界
      xlsRpt.Visible = False
   Else
      If intPage = 0 Then  '刪除前一張匯票的內容
         WksRpt1.Shapes.SelectAll
         xlsRpt.Selection.Delete  '刪除所有圖片
         WksRpt1.Range("A:G").Select
         xlsRpt.Selection.Delete  '刪除文字
      Else
         '跨頁不清除
      End If
   End If
'-------------------欄寬和列高-----------------------------
   If intPage = 0 Then
      For intI = 0 To 6
         WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Name = "Times New Roman"
         WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Size = 12
         Select Case intI
            Case 0, 6  'A,G
               WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 7.5
            Case 1 'B
               WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 14
            Case 2, 3 'C,D
               WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 24
            Case 4 'E
               WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 6
            Case 5 'F
               WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 14
         End Select
         If intI <> 5 Then
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).NumberFormatLocal = "@"
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).HorizontalAlignment = xlLeft
         End If
      Next intI
      bolColTitle = True
   End If
   For intI = 1 To maxRows
      If intI = 1 Then
         WksRpt1.Range(intI + intPage * maxRows & ":" & intI + intPage * maxRows).RowHeight = 110
      ElseIf intI = maxRows Then
         WksRpt1.Range(intI + intPage * maxRows & ":" & intI + intPage * maxRows).RowHeight = 35
      Else
         WksRpt1.Range(intI + intPage * maxRows & ":" & intI + intPage * maxRows).RowHeight = 17
      End If
   Next intI

'-------------------欄寬和列高-----------------------------
   'Excel列印資料的起始,終止位置
   xRows = (intPage * maxRows) + 2
   xRowE = ((intPage + 1) * maxRows) - 2
   nRow = xRows '目前
   xCols = 66  '因為信頭JPG範圍包含上+左右邊界的空白，所以從B欄開始放入資料
   
   WksRpt1.Range("D" & xRowE + 1).Value = "**" & intPage + 1 & "**"
   If iPicNo1 > 0 Then  '信頭
      strPic01 = strPrtPath & "\$Tmp01.jpg"
      If intPage = 0 Then
         If PUB_ReadDB2File(strPic01, iPicNo1) = True Then
         End If
      Else
         strExc(0) = Dir(strPic01)
         If strExc(0) = "" Then
            If PUB_ReadDB2File(strPic01, iPicNo1) = True Then
            End If
         End If
      End If
      Set oShape = WksRpt1.Shapes.AddPicture(strPic01, True, True, 0, WksRpt1.Cells((intPage * maxRows) + 1, "A").Top, xlsRpt.CentimetersToPoints(19.5), xlsRpt.CentimetersToPoints(3.66))
   End If
   
   If iPicNo2 > 0 Then  '信尾
      strPic02 = strPrtPath & "\$Tmp02.jpg"
      If intPage = 0 Then
         If PUB_ReadDB2File(strPic02, iPicNo2) = True Then
         End If
      Else
         strExc(0) = Dir(strPic02)
         If strExc(0) = "" Then
            If PUB_ReadDB2File(strPic02, iPicNo2) = True Then
            End If
         End If
      End If
      Set oShape2 = WksRpt1.Shapes.AddPicture(strPic02, True, True, 0, WksRpt1.Cells(((intPage + 1) * maxRows), "A").Top + 2, xlsRpt.CentimetersToPoints(19.5), xlsRpt.CentimetersToPoints(0.91))
   End If
   
   PrintExcel_BFile = True
   Exit Function
End Function

'Added by Lydia 2024/12/11 改用EXCEL：列印匯票、清單
Private Sub PrintExcelMain()
Dim strDescription As String
Dim strName As String
Dim StrSQLa As String
Dim strAXF03 As String '本所案號
Dim strA1501 As String '單據號碼
Dim strMailFailList() As String 'Mail 失敗清單
Dim iCopy As Integer
Dim iRound As Integer '迴圈次數
Dim bolChgPrinter As Boolean
Dim tmpErr As String
Dim bolOpenXls As Boolean
  
On Error GoTo ErrorHandle

   cnnConnection.Execute "delete from accrpt2450 where id='" & strUserNum & "' "
   
   adoquery.CursorLocation = adUseClient

   StrSQLa = "Select axf03, a1b02, cp45, a1504, a1505, axf04, a1b06, a1b01, a1502, cp01, cp02, cp03, cp04, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, a1c03, a1501, fa17, fa04, fa06, CP09,fa10,nvl(fa105,nvl(fa79,fa16)) fa16,fa84,a1b03,fa70,a1917,a1901,a1902,fa29,a1719,a1812,axf01 " & _
                    ",decode(fa105||fa79,null,'',fa134) as emailcc From acc1c0, acc1b0, acc151, acc150, caseprogress, fagent,acc190,acc170,acc180 Where a1c01 = a1b01 and a1c02 = a1b02 and a1c03 = a1501 (+) and a1c03 = axf01 (+) and axf02 = cp09 (+) and substr(a1b02, 1, 8) = fa01 (+) and substr(a1b02, 9, 1) = fa02 (+) And a1b01=a1908 And a1c03=a1902(+)  and a1702(+)=a1902 " & _
                    "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " and a1801(+)=a1901 and axf01 is not null" & strConSql

   StrSQLa = StrSQLa & " Union Select axg03 as axf03, a1b02, cp45, a1604 as a1504, a1605 as a1505, axg04 as axf04, a1b06, a1b01, a1602, cp01, cp02, cp03, cp04, fa05, fa63, fa64, fa65, fa32, fa18, fa33, fa19, fa34, fa20, fa35, fa21, fa36, fa22, fa06, fa23, a1c03, a1601 as a1501, fa17, fa04, fa06, CP09,fa10,nvl(fa105,nvl(fa79,fa16)) fa16,fa84,a1b03,fa70,a1917,a1901,a1902,fa29,a1719,'' as a1812,'' as axf01 " & _
                    ",decode(fa105||fa79,null,'',fa134) as emailcc From acc1c0, acc1b0, acc161, acc160, caseprogress, fagent,acc190,acc170 Where a1c01 = a1b01 and a1c02 = a1b02 and a1c03 = a1601 (+) and a1c03 = axg01 (+) and axg02 = cp09 (+) and substr(a1b02, 1, 8) = fa01 (+) and substr(a1b02, 9, 1) = fa02 (+) And a1b01=a1908 And a1c03=a1902(+)  and a1702(+)=a1902 " & _
                    "And a1917" & IIf(Text3 = "J", "='" & Text3 & "'", "<>'J'") & " and axg01 is not null" & strConSql
   StrSQLa = StrSQLa & " Order by a1b02 asc, a1b01, axf03 asc, a1501 asc"
    
   adoquery.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount = 0 Then
      adoquery.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   '存電子檔時不印
   If Text6 = "Y" Then
      strSavePath = PUB_Getdesktop
      bol2Printer = False
   Else
      strSavePath = App.path
      If Text5 = "1" Then
         bol2Printer = True
      Else
         bol2Printer = False
      End If
   End If
   
   Call Pub_ChkExcelPath(strPrtPath)
   Call PUB_KillTempFile(strUserNum & "\$*.*")

'Modified by Lydia 2016/11/16 重新界定需求:選列印時票匯印2份(其中1份是清單)電匯印1份,選email則都要發mail不做列印
If Text5 = "1" And Text6 <> "Y" Then
   iRound = 2
Else
   iRound = 1
End If

If Text5 = "2" Then  '觀察寄件狀態
   txtSend.Visible = True: lblSend.Visible = True
End If

For iCopy = 1 To iRound

   '改成先印清單 (第2次列印之前)
   If iCopy = 2 Then
      'Memo by Lydia 2018/11/19 已與婉莘確認,若輸入代理人編號也不用輸出Invoice
      If Trim(Text1.Text & Text2.Text) = "" Then 'Added by Lydia 2017/11/13 若輸入代理人編號, 則不需要列印清單
         PrintList
      End If
      GoTo JumpToListEnd
   End If

   Erase strMailFailList
   ReDim strMailFailList(0)
   intLength = 0
   douAmount = 0
   douUSDollar = 0
   intPage = 0
   strNo = ""
   strPicLetter = ""
   strPicFileNames = ""
   mCnt = 0
   adoquery.MoveFirst
   Do While adoquery.EOF = False
      '改印結匯日期
      If Not IsNull(adoquery.Fields("a1b03")) Then
         m_strPayDate = DBDATE(adoquery.Fields("a1b03").Value)
      Else
         m_strPayDate = strSrvDate(1)
      End If
      
      strAXF03 = "" & adoquery.Fields("axf03").Value
      strA1501 = "" & adoquery.Fields("a1501").Value
      If IsNull(adoquery.Fields("axf03").Value) = False Then
         strName = adoquery.Fields("axf03").Value
      Else
         strName = ""
      End If
      
      '留所資料改成清單方式產出
      bolList = False
      If iCopy = 1 And bol2Printer = True And Text5.Text = "1" And Text6.Text <> "Y" Then
         bolList = True
      Else
         strName = strName
      End If

      If iCopy = 1 And adoquery.AbsolutePosition = 1 Then
         bolOpenXls = True
      Else
         bolOpenXls = False
      End If

      '2012/2/23 modify by sonia 同一代理人匯票號碼不同也要分開印 (Y28215於100/8/2)
      If strNo <> adoquery.Fields("a1b02").Value Or m_a1b01 <> adoquery.Fields("a1b01").Value Then
         If douAmount <> 0 Then
            '暫存清單資料
            If bolList = True Then
                mCnt = mCnt + 1
                strExc(0) = "insert into accrpt2450 (id,r24500,r24501,r24502,r24503,r24504,r24505,r24506) values (" & CNULL(strUserNum) & _
                            ", " & mCnt & ", NULL, NULL,'小　計' ,NULL,NULL," & douAmount & ") "
                cnnConnection.Execute strExc(0)
                mCnt = mCnt + 1
            Else
                Call PrintExcel_BSum
            End If
            
            If bol2File = True And bolList = False Then
               '先存PDF檔(另存新檔)放在桌面，不關EXCEL後面再處理信頭、信尾>>PrintExcel_BFile
               If PUB_PrintExcel2File(xlsRpt, strSavePath, "$" & strNo & m_a1b01 & ".PDF", strExc(1), False) = True Then
                  strPicFileNames = strSavePath & "\" & strExc(1)
               End If
               
               If (bolEmail = True Or bol2Email = True) And Text5.Text = "2" And Text6.Text <> "Y" And strPicFileNames <> "" Then
                  bolMailFailNoAlert = True
                  bolMailSendOk = False
                  txtSend.Text = strNo
                  PUB_SendMail strUserNum, strEMailBox, "", "ADVICE OF WIRE/CHECK PAYMENT (" & strNo & ")", GetMailContent, , strPicFileNames, True, True, True, strEmailCC, strAccMailBox, strCompName, strAccMailBox
                  bolMailFailNoAlert = False
                  If bolMailSendOk = False Then
                     If strMailFailList(0) <> "" Then
                        ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                     End If
                     strMailFailList(UBound(strMailFailList)) = strNo & "：" & strEMailBox
                  End If
                  strPicFileNames = ""
               End If
               strPicFileNames = "" '已存電子檔,清空路徑
            End If
            If bol2Printer = True And bolList = False Then
               WksRpt1.PrintOut Copies:=1, Collate:=True '列印
            End If
            
            douAmount = 0
            douUSDollar = 0
            intPage = 0
         End If
         intCounter = 0
         
         '是否寄明細
         strInform = "" & adoquery("fa84")
         '電子信箱
         strEMailBox = "" & adoquery("fa16")
         strEmailCC = "" & adoquery("emailcc")
         If strEMailBox <> "" And UCase(strEMailBox) <> "NO" And "" & adoquery("a1b06") = "2" Then
            bolEmail = True
            bol2File = True
         Else
            bolEmail = False
            bol2File = False
         End If
         
         'Add by Morgan 2008/4/17 放在迴圈內是因為要印"是否EMail通知","不寄明細"的註記
         '產生電子檔
         If Text6.Text = "Y" Then
            bol2File = True
         '不發Mail 或 不寄收據 時只列印不產生電子檔
         ElseIf Text5.Text = "1" Or strInform = "N" Then
            bol2File = False
         End If
         
         'Add by Morgan 2008/11/20
         If iCopy = 2 Then
            '不寄收據 或 EMail 或 存電子檔的不要印第二份
            If strInform = "N" Or bolEmail = True Or bol2File = True Then
               bol2Printer = False
            Else
               bol2Printer = True
            End If
            bol2File = False
            bolEmail = False
         Else
            If Not (strInform = "N" Or bolEmail = True Or bol2File = True) Then bolChgPrinter = True
         End If

         If bol2File = True And (Text6 = "Y" Or (Text5 = "2" And Text6 <> "Y")) Then
            bol2Printer = False
            bol2File = True
         End If
         
         strNo = adoquery.Fields("a1b02").Value
         m_a1b01 = adoquery.Fields("a1b01").Value
         m_a1b06 = adoquery.Fields("a1b06").Value
         
         '代理人編號和名稱
         strTemp(0) = strNo
         strTemp(1) = Trim(adoquery.Fields("fa05") & " " & adoquery.Fields("fa63") & " " & adoquery.Fields("fa64") & " " & adoquery.Fields("fa65"))

         
         '清空信頭變數
         strPicLetter = ""
         m_iNo = 0
         m_iNo2 = 0

         '加信頭-存電子檔
         If bol2File = True Then
            'Added by Morgan 2020/3/30
            If strSrvDate(1) >= 智慧所更名日 Then
               m_iNo2 = m_iNo
               '改用EXCEL：信頭、信尾分開來
               PUB_GetLetterPicID Text3, , m_iNo, m_iNo2, , , "HALF"
            Else
               'Modify by Amy 2014/07/08 巨京Y52269000的付款明細為J公司用智權,其他仍用專利商標(2011/10/11)
               If strNo = "Y52269000" And Text3 <> "J" Then
                   m_iNo = 8
               'Add by Amy 2014/05/01 +J公司使用智權公司抬頭
               ElseIf Text3 = "J" Then
                   m_iNo = 25
               'end 2014/05/01
               Else
                   m_iNo = 6
               End If
               'end 2011/10/11
            End If
         End If
         
         '非清單才列印匯票=帳單
         If bolList = False Then
            If PrintExcel_BFile(bolOpenXls, m_iNo, m_iNo2) = False Then
               GoTo ErrorHandle
            End If
         End If
         'end 2017/10/05
         '非清單才列印
         If bolList = False Then
            Call PrintExcel_BHead '匯票=帳單抬頭列印
         End If
      End If

      intCounter = 0 '列印欄位Column
      If IsNull(adoquery.Fields("axf03").Value) = False Then
         If Mid(adoquery.Fields("axf03").Value, Len(adoquery.Fields("axf03").Value) - 2, 3) = "000" Then
            strExc(0) = Mid(adoquery.Fields("axf03").Value, 1, Len(adoquery.Fields("axf03").Value) - 9) & "-" & Mid(adoquery.Fields("axf03").Value, Len(adoquery.Fields("axf03").Value) - 8, 6)
         Else
            strExc(0) = Mid(adoquery.Fields("axf03").Value, 1, Len(adoquery.Fields("axf03").Value) - 9) & "-" & Mid(adoquery.Fields("axf03").Value, Len(adoquery.Fields("axf03").Value) - 8, 6) & "-" & Mid(adoquery.Fields("axf03").Value, Len(adoquery.Fields("axf03").Value) - 2, 1) & "-" & Mid(adoquery.Fields("axf03").Value, Len(adoquery.Fields("axf03").Value) - 1, 2)
         End If
         '清單：本所案號
         If bolList = True Then
            strTemp(2) = "" & strExc(0)
         Else
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strExc(0)
         End If
      End If
      
      intCounter = intCounter + 1
      strCurrency = "" & adoquery.Fields("a1505").Value
      '清單：DB.note 和幣別
      If bolList = True Then
         strTemp(3) = "" & adoquery.Fields("a1504").Value
         strTemp(4) = strCurrency
      Else
         'YOUR REF和DB NOTE
         If GetTextLength("" & adoquery.Fields("cp45").Value) > ciDBwidth Then
            strExc(1) = convForm("" & adoquery.Fields("cp45").Value, ciDBwidth)
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strExc(1)
            WksRpt1.Range(Chr(xCols + intCounter) & nRow + 1).Value = convForm(Replace("" & adoquery.Fields("cp45").Value, strExc(1), ""), ciDBwidth)
         Else
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "" & adoquery.Fields("cp45").Value
         End If
         intCounter = intCounter + 1
         If GetTextLength("" & adoquery.Fields("a1504").Value) > ciDBwidth Then
            strExc(1) = convForm("" & adoquery.Fields("a1504").Value, ciDBwidth)
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strExc(1)
            WksRpt1.Range(Chr(xCols + intCounter) & nRow + 1).Value = convForm(Replace("" & adoquery.Fields("a1504").Value, strExc(1), ""), ciDBwidth)
         Else
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "" & adoquery.Fields("a1504").Value
         End If
      End If  'If bolList = True Then
      
      intCounter = intCounter + 1
      '幣別
      If bolList = False Then
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strCurrency
      End If
      '金額
      intCounter = intCounter + 1
      If IsNull(adoquery.Fields("axf04").Value) = False Then
         If Mid(adoquery.Fields("a1c03").Value, 1, 1) = "U" Then
            '若帳單編號及本所案號相同則金額合併
            strAmount = Val("" & adoquery.Fields("axf04").Value)
            adoquery.MoveNext
            If adoquery.EOF = True Then
                adoquery.MovePrevious
            Else
                Do While adoquery.Fields("axf03").Value = strAXF03 And adoquery.Fields("a1501").Value = strA1501
                    strAmount = Val(strAmount) + Val(IIf(Mid(adoquery.Fields("a1c03").Value, 1, 1) = "V", "-1", "1")) * Val(adoquery.Fields("axf04").Value)
                    adoquery.MoveNext
                    If adoquery.EOF = True Then Exit Do
                Loop
                adoquery.MovePrevious
            End If
            strAmount = Format(Val(strAmount), FDollar)
            '清單：金額
            If bolList = True Then
                strTemp(5) = Val(CDbl(strAmount))
            Else
                WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strAmount
                WksRpt1.Range(Chr(xCols + intCounter) & nRow).NumberFormatLocal = FDollar
            End If
            douAmount = douAmount + Val(CDbl(strAmount))
         Else
            strAmount = Val(IIf(Mid(adoquery.Fields("a1c03").Value, 1, 1) = "V", "-1", "1")) * Val("" & adoquery.Fields("axf04").Value)
            adoquery.MoveNext
            If adoquery.EOF = True Then
                adoquery.MovePrevious
            Else
                Do While adoquery.Fields("axf03").Value = strAXF03 And adoquery.Fields("a1501").Value = strA1501
                    strAmount = Val(strAmount) + Val(IIf(Mid(adoquery.Fields("a1c03").Value, 1, 1) = "V", "-1", "1")) * Val(adoquery.Fields("axf04").Value)
                    adoquery.MoveNext
                    If adoquery.EOF = True Then Exit Do
                Loop
                adoquery.MovePrevious
            End If
            strAmount = Format(Val(strAmount), FDollar)
            '清單：金額
            If bolList = True Then
                strTemp(5) = Val(CDbl(strAmount))
            Else
                WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strAmount
                WksRpt1.Range(Chr(xCols + intCounter) & nRow).NumberFormatLocal = FDollar
            End If
            douAmount = douAmount + Val(CDbl(strAmount))
         End If
      End If

      '資料第二行
      nRow = nRow + 1
      intCounter = 0
      If Not IsNull(adoquery.Fields("a1502").Value) Then
         '匯票=帳單日期
         strExc(0) = Format(AFDate(CADate(adoquery.Fields("a1502").Value)), "mmm. d, yyyy")
         If bolList = True Then
         Else
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strExc(0)
         End If
      End If

      If InStr(m_Curs, strCurrency) = 0 Then
         m_Len = m_Len + 1
         ReDim Preserve m_Cur(m_Len)
         ReDim Preserve m_Amount(m_Len)
         m_Curs = m_Curs & "," & strCurrency
         m_Cur(m_Len) = strCurrency
         m_Amount(m_Len) = Val(CDbl(strAmount))
      Else
         For m_Idx = 1 To UBound(m_Cur)
            If m_Cur(m_Idx) = strCurrency Then
               m_Amount(m_Idx) = m_Amount(m_Idx) + Val(CDbl(strAmount))
               Exit For
            End If
         Next
      End If
      
      '暫存清單資料
      If bolList = True Then
         mCnt = mCnt + 1
         strExc(0) = "insert into accrpt2450 (id,r24500,r24501,r24502,r24503,r24504,r24505,r24506,r24507,r24508) values (" & CNULL(strUserNum) & _
                     ", " & mCnt & ", " & CNULL(strTemp(0)) & ", " & CNULL(ChgSQL(PUB_StrToStr(strTemp(1), 150))) & " ," & CNULL(strTemp(2)) & "," & CNULL(PUB_StrToStr(IIf("" & adoquery.Fields("a1719") = "Y", "＊", "") & strTemp(3), 100)) & " ," & CNULL(strTemp(4)) & " ," & strTemp(5) & _
                     ",'" & IIf("" & adoquery.Fields("a1812") = "Y", "Y", "") & "', " & CNULL("" & adoquery.Fields("axf01")) & ") "
         cnnConnection.Execute strExc(0)
      Else
         Call PrintExcel_BPage
      End If

      adoquery.MoveNext
   Loop

   If bolList = True Then
        mCnt = mCnt + 1
        strExc(0) = "insert into accrpt2450 (id,r24500,r24501,r24502,r24503,r24504,r24505,r24506) values (" & CNULL(strUserNum) & _
                    ", " & mCnt & ", NULL, NULL ,'小　計',NULL,NULL," & douAmount & ") "
        cnnConnection.Execute strExc(0)
   Else
       Call PrintExcel_BSum
   End If

   If bol2File = True And bolList = False Then
      If bol2File = True Then
          'PDF檔放在桌面
          If PUB_PrintExcel2File(xlsRpt, strSavePath, "$" & strNo & m_a1b01 & ".PDF", strExc(1), False) = True Then
             strPicFileNames = strPicFileNames & strSavePath & "\" & strExc(1) & "*"
          End If
      End If
      If (bolEmail = True Or bol2Email = True) And Text5.Text = "2" And Text6.Text <> "Y" And strPicFileNames <> "" Then
         bolMailFailNoAlert = True
         bolMailSendOk = False
         txtSend.Text = strNo
         PUB_SendMail strUserNum, strEMailBox, "", "ADVICE OF WIRE/CHECK PAYMENT (" & strNo & ")", GetMailContent, , strPicFileNames, True, True, True, strEmailCC, strAccMailBox, strCompName, strAccMailBox
         bolMailFailNoAlert = False
         If bolMailSendOk = False Then
            If strMailFailList(0) <> "" Then
               ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
            End If
            strMailFailList(UBound(strMailFailList)) = strNo & "：" & strEMailBox
         End If
      Else
         If strPicFileNames <> "" Then
            MsgBox "電子檔已存桌面！"
         End If
      End If
   End If

   If bolList = False Then
      xlsRpt.Workbooks(1).Save
      If bol2Printer = True Then
         WksRpt1.PrintOut Copies:=1, Collate:=True '列印
      End If
      xlsRpt.Workbooks.Close
      xlsRpt.Quit
      Set xlsRpt = Nothing
      Set WksRpt1 = Nothing
   End If
   
   '刪除舊的暫存檔
   Call PUB_KillTempFile(strUserNum & "\$*.*")
      
   If strMailFailList(0) <> "" Then
      strExc(0) = "E-Mail失敗清單：" & vbCrLf & vbCrLf
      For intI = 0 To UBound(strMailFailList)
         strExc(0) = strExc(0) & strMailFailList(intI) & vbCrLf
      Next

      PUB_SendMail strUserNum, strUserNum, "", Me.Caption & "-" & "E-Mail失敗清單", vbCrLf & strExc(0)
      MsgBox strExc(0) & vbCrLf & "請參考！", vbInformation, Me.Caption & "-" & "E-Mail失敗清單"

   End If
         
Next

JumpToListEnd:

   adoquery.Close
   
   txtSend.Visible = False: lblSend.Visible = False

   Exit Sub
    
ErrorHandle:
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Resume Next
    End If
End Sub

'Adeed by Lydia 2024/12/11 改用EXCEL：匯票抬頭列印
Private Sub PrintExcel_BHead()
Dim StrSqlB As String
Dim rsA As New ADODB.Recordset

Dim strFA17 As String, strFA18 As String, strFA19 As String, strFA20 As String, strFA21 As String, strFA22 As String, strFA70 As String, strFA23 As String
Dim strFA32 As String, strFA33 As String, strFA34 As String, strFA35 As String, strFA36 As String

   If intPage = 0 Then
      Erase m_Cur
      Erase m_Amount
      m_Curs = ""
      m_Len = 0
   End If
   
   StrSqlB = " Select '1' ord1, f1.* From Fagent f1, (Select A2213 From ACC220, Fagent Where substr(A2201,1,8)=FA01 And substr(A2201,9,1)=FA02 And A2201='" & strNo & "') A Where fa01 = substr(a.a2213, 1, 8) And fa02 = substr(a.a2213, 9, 1) " & _
             " union select '2' ord1,f2.* from fagent f2 where fa01='" & Mid(strNo, 1, 8) & "' and fa02='" & Mid(strNo, 9, 1) & "' " & _
             " order by ord1"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
   '跨頁發生在頁尾內文，重新抓資料
   If "" & rsA.Fields("fa10") = "020" Then
      strLanguage = "1"
   Else
      strLanguage = "2"
   End If
   
   '信頭印結匯日期
   WksRpt1.Range("F" & nRow).Value = Format(AFDate(m_strPayDate), "mmm. d, yyyy")
   nRow = nRow + 1
   
   If intPage > 0 Then '跨頁
      '最後會跳行
   Else
'--------------------------------------------------
      '地址有「竹曆退件」字樣不顯示地址
      strFA17 = "" & rsA.Fields("fa17").Value
      strFA18 = "" & rsA.Fields("fa18").Value: strFA19 = "" & rsA.Fields("fa19").Value: strFA20 = "" & rsA.Fields("fa20").Value
      strFA21 = "" & rsA.Fields("fa21").Value: strFA22 = "" & rsA.Fields("fa22").Value: strFA70 = "" & rsA.Fields("fa70").Value
      strFA23 = "" & rsA.Fields("fa23").Value
      strFA32 = "" & rsA.Fields("fa32").Value: strFA33 = "" & rsA.Fields("fa33").Value: strFA34 = "" & rsA.Fields("fa34").Value
      strFA35 = "" & rsA.Fields("fa35").Value: strFA36 = "" & rsA.Fields("fa36").Value

     If InStr(strFA17, "竹曆退件") > 0 Then strFA17 = ""
     If InStr(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70, "竹曆退件") > 0 Then
           strFA18 = "": strFA19 = "": strFA20 = "": strFA21 = "": strFA22 = "": strFA70 = ""
     End If
     If InStr(strFA23, "竹曆退件") > 0 Then strFA23 = ""
     If InStr(strFA32 & strFA33 & strFA34 & strFA35 & strFA36, "竹曆退件") > 0 Then
           strFA32 = "": strFA33 = "": strFA34 = "": strFA35 = "": strFA36 = ""
     End If
         
      '選擇語言
      Select Case strLanguage
         Case "1" '中文
            If rsA.RecordCount > 0 Then
               If intPage = 0 Then
                  '中文
                  If IsNull(rsA.Fields("fa04").Value) = False Then
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa04").Value
                  '英文
                  ElseIf IsNull(rsA.Fields("fa05").Value) = False Then
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa05").Value
                     If IsNull(rsA.Fields("fa63").Value) = False Then
                        nRow = nRow + 1
                        WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa63").Value
                     End If
                     If IsNull(rsA.Fields("fa64").Value) = False Then
                        nRow = nRow + 1
                        WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa64").Value
                     End If
                     If IsNull(rsA.Fields("fa65").Value) = False Then
                        nRow = nRow + 1
                        WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa65").Value
                     End If
                  '日文
                  ElseIf IsNull(rsA.Fields("fa06").Value) = False Then
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa06").Value
                  End If
               End If
            End If
            nRow = nRow + 1
            If intPage = 0 Then
               '地址,順序：中文->POB->英文->日文
               '中文地址
               If strFA17 <> MsgText(601) Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA17
               'POB
               ElseIf strFA32 <> MsgText(601) Then
                  'P0B1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA32
                  'POB2
                  If strFA33 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA33
                  End If
                  'POB3
                  If strFA34 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA34
                  End If
                  'POB4
                  If strFA35 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA35
                  End If
                  'POB5
                  If strFA36 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA36
                  End If
               
               '英文地址
               ElseIf strFA18 <> MsgText(601) Then
                  '英文地址1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA18
                  '英文地址2
                  If strFA19 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA19
                  End If
                  '英文地址3
                  If strFA20 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA20
                     'Modified by Lydia 2018/09/03 + bol2pdf
                  End If
                  '英文地址4
                  If strFA21 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA21
                  End If
                  '英文地址5
                  If strFA22 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA22
                  End If
                  '英文地址6
                  If strFA70 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA70
                  End If
               '日文地址
               ElseIf strFA23 <> MsgText(601) Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA23
               End If
               'end 2018/10/31
            End If
            If nRow < 6 Then nRow = 6
         Case "2"  '英文
            If rsA.RecordCount > 0 Then
               If intPage = 0 Then
                  If IsNull(rsA.Fields("fa05").Value) = False Then
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa05").Value
                  End If
                  If IsNull(rsA.Fields("fa63").Value) = False Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa63").Value
                  End If
                  If IsNull(rsA.Fields("fa64").Value) = False Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa64").Value
                  End If
                  If IsNull(rsA.Fields("fa65").Value) = False Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa65").Value
                  End If
               End If
            End If
            nRow = nRow + 1
            '地址,順序：POB->英文->中文----Memo by Lydia 2025/09/25
            '地址
            If intPage = 0 Then
               'Modified by Lydia 2025/09/25 將寫法調整與中文方式相似
                'POB
                If strFA32 <> MsgText(601) Then
                   'P0B1
                   WksRpt1.Range(Chr(xCols) & nRow).Value = strFA32
                   'POB2
                   If strFA33 <> MsgText(601) Then
                      nRow = nRow + 1
                      WksRpt1.Range(Chr(xCols) & nRow).Value = strFA33
                   End If
                   'POB3
                   If strFA34 <> MsgText(601) Then
                      nRow = nRow + 1
                      WksRpt1.Range(Chr(xCols) & nRow).Value = strFA34
                   End If
                   'POB4
                   If strFA35 <> MsgText(601) Then
                      nRow = nRow + 1
                      WksRpt1.Range(Chr(xCols) & nRow).Value = strFA35
                   End If
                   'POB5
                   If strFA36 <> MsgText(601) Then
                      nRow = nRow + 1
                      WksRpt1.Range(Chr(xCols) & nRow).Value = strFA36
                   End If
                
                '英文地址
                ElseIf strFA18 <> MsgText(601) Then
                   '英文地址1
                   WksRpt1.Range(Chr(xCols) & nRow).Value = strFA18
                   '英文地址2
                   If strFA19 <> MsgText(601) Then
                      nRow = nRow + 1
                      WksRpt1.Range(Chr(xCols) & nRow).Value = strFA19
                   End If
                   '英文地址3
                   If strFA20 <> MsgText(601) Then
                      nRow = nRow + 1
                      WksRpt1.Range(Chr(xCols) & nRow).Value = strFA20
                   End If
                   '英文地址4
                   If strFA21 <> MsgText(601) Then
                      nRow = nRow + 1
                      WksRpt1.Range(Chr(xCols) & nRow).Value = strFA21
                   End If
                   '英文地址5
                   If strFA22 <> MsgText(601) Then
                      nRow = nRow + 1
                      WksRpt1.Range(Chr(xCols) & nRow).Value = strFA22
                   End If
                   '英文地址6
                   If strFA70 <> MsgText(601) Then
                      nRow = nRow + 1
                      WksRpt1.Range(Chr(xCols) & nRow).Value = strFA70
                   End If
                '中文地址
                ElseIf strFA17 <> MsgText(601) Then
                   nRow = nRow + 1
                   WksRpt1.Range(Chr(xCols) & nRow).Value = strFA17
                End If
            End If  'If intPage = 0 Then
            If nRow < 6 Then nRow = 6

         Case "3" '日文
            If rsA.RecordCount > 0 Then
               If intPage = 0 Then
                  '日文
                  If IsNull(rsA.Fields("fa06").Value) = False Then
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa06").Value
                  '英文
                  ElseIf IsNull(rsA.Fields("fa05").Value) = False Then
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa05").Value
                     If IsNull(rsA.Fields("fa63").Value) = False Then
                        nRow = nRow + 1
                        WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa63").Value
                     End If
                     If IsNull(rsA.Fields("fa64").Value) = False Then
                        nRow = nRow + 1
                        WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa64").Value
                     End If
                     If IsNull(rsA.Fields("fa65").Value) = False Then
                        nRow = nRow + 1
                        WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa65").Value
                     End If
                  '中文
                  ElseIf IsNull(rsA.Fields("fa04").Value) = False Then
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & rsA.Fields("fa04").Value
                  End If
               End If
            End If
            'Added by Lydia 2025/09/25 檢查後，發現有缺少列印地址
            nRow = nRow + 1
            '地址,順序：日文->POB->英文->中文
            If intPage = 0 Then
               '日文地址
               If strFA23 <> MsgText(601) Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA23
               'POB
               ElseIf strFA32 <> MsgText(601) Then
                  'P0B1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA32
                  'POB2
                  If strFA33 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA33
                  End If
                  'POB3
                  If strFA34 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA34
                  End If
                  'POB4
                  If strFA35 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA35
                  End If
                  'POB5
                  If strFA36 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA36
                  End If
               
               '英文地址
               ElseIf strFA18 <> MsgText(601) Then
                  '英文地址1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA18
                  '英文地址2
                  If strFA19 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA19
                  End If
                  '英文地址3
                  If strFA20 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA20
                  End If
                  '英文地址4
                  If strFA21 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA21
                  End If
                  '英文地址5
                  If strFA22 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA22
                  End If
                  '英文地址6
                  If strFA70 <> MsgText(601) Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = strFA70
                  End If
               '中文地址
               ElseIf strFA17 <> MsgText(601) Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA17
               End If
            End If  'If intPage = 0 Then
            'end 2025/09/25
            If nRow < 6 Then nRow = 6
      End Select
      
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      nRow = nRow + 2
      Select Case m_a1b06
         Case "1" '票匯
            If intPage = 0 Then
               WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
               WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(87) 'Gentlemen:
            End If
            nRow = nRow + 1
            If intPage = 0 Then
               WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
               WksRpt1.Range(Chr(xCols) & nRow).Value = String(6, " ") & ReportSum(88)  'We are sending ...
            End If
            nRow = nRow + 1
            If intPage = 0 Then
               WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
               WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(89)  'detailed hereunder.
               '列印DRAFT NO
               nRow = nRow + 2
               WksRpt1.Range(Chr(xCols + 1) & nRow).Value = "DRAFT NO. " & m_a1b01
               nRow = nRow + 1
            End If
            
         Case Else '2 電匯
            If intPage = 0 Then
               WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
               WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(87) 'Gentlemen:
            End If
            StrSqlB = "Select * From ACC220 Where A2201='" & strNo & "' And A2202='" & strCurrency & "' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               nRow = nRow + 1
               If intPage = 0 Then
                  WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
                  WksRpt1.Range(Chr(xCols) & nRow).Value = String(6, " ") & ReportSum(97001)   'We inform ...
               End If
               nRow = nRow + 1
               If intPage = 0 Then
                  WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
                  WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(98001) 'bank account, i.,e.
               End If
               nRow = nRow + 2
               If intPage = 0 Then  '受款銀行名稱
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Font.Size = 14
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = "" & rsA("a2208").Value
               End If
               nRow = nRow + 1
               If intPage = 0 Then  '受款人帳號
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Font.Size = 14
                  WksRpt1.Range(Chr(xCols + 1) & nRow).Value = "Account #" & rsA("a2207").Value
               End If
            Else
                nRow = nRow + 1
                If intPage = 0 Then
                   WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
                   WksRpt1.Range(Chr(xCols) & nRow).Value = String(6, " ") & ReportSum(97)   'A remittance ...
                End If
                nRow = nRow + 1
                If intPage = 0 Then
                   WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
                   WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(98) 'notes(invoices) as follows
                End If
            End If
      End Select
   End If  '-----If intPage > 0 Then '跨頁
   
   intPage = intPage + 1
   nRow = nRow + 2
   '欄位抬頭
   If bolColTitle = True Then
      WksRpt1.Range(Chr(xCols) & nRow).Value = "OUR REF"
      WksRpt1.Range(Chr(xCols + 1) & nRow).Value = "YOUR REF"
      WksRpt1.Range(Chr(xCols + 2) & nRow).Value = "YR DB NOTE"
      WksRpt1.Range(Chr(xCols + 3) & nRow).Value = "    AMOUNT"
      
      WksRpt1.Range(Chr(xCols) & nRow & ":" & Chr(xCols + 4) & nRow).Borders(xlEdgeBottom).LineStyle = xlContinuous  '儲存格底線
      nRow = nRow + 1
   End If
   
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Sub

'Adeed by Lydia 2024/12/11 改用EXCEL：匯票換行
Private Sub PrintExcel_BPage(Optional ByVal pAddLine As Integer = 1)
   nRow = nRow + pAddLine
   If nRow >= xRowE Then
      Call PrintExcel_BFile(False, m_iNo, m_iNo2)
      Call PrintExcel_BHead
   End If
End Sub

'Added by Lydia 2024/12/11 改用EXCEL：匯票TOTAL列印
Private Sub PrintExcel_BSum()
Dim strTmp As String

   WksRpt1.Range(Chr(xCols) & nRow & ":" & Chr(xCols + 4) & nRow).Borders(xlEdgeTop).LineStyle = xlContinuous  '儲存格上邊界框線
   Call PrintExcel_BPage
   
   WksRpt1.Range(Chr(xCols) & nRow).Value = "TOTAL"
   WksRpt1.Range(Chr(xCols + 3) & nRow).Value = m_Cur(1)  '幣別
   strAmount = Format(m_Amount(1), FDollar)
   WksRpt1.Range(Chr(xCols + 4) & nRow).Value = strAmount '金額
   WksRpt1.Range(Chr(xCols + 4) & nRow).NumberFormatLocal = FDollar
   strPayAmount = m_Cur(1) & " " & strAmount
   Call PrintExcel_BPage
   
   '同一張匯票有不同幣別的帳單
   For m_Idx = 2 To UBound(m_Cur)
      Call PrintExcel_BPage(False)
      WksRpt1.Range(Chr(xCols + 3) & nRow).Value = m_Cur(m_Idx)  '幣別
      strAmount = Format(m_Amount(1), FDollar)
      WksRpt1.Range(Chr(xCols + 4) & nRow).Value = Format(m_Amount(m_Idx), FDollar) '金額
      WksRpt1.Range(Chr(xCols + 4) & nRow).NumberFormatLocal = FDollar
      strPayAmount = m_Cur(1) & " " & strAmount
   Next
   
   bolColTitle = False
   
   Call PrintExcel_BPage(2)

   Select Case m_a1b06 '匯票方式
      Case "1"  '票匯
         WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
         WksRpt1.Range(Chr(xCols) & nRow).Value = String(6, " ") & ReportSum(90)   'Please acknowledge safe receipt ...
         Call PrintExcel_BPage
         WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
         WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(91) 'be appreciated if you ...
         Call PrintExcel_BPage
         WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
         WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(92) 'debit notes or statements.
         
      Case Else
         WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
         WksRpt1.Range(Chr(xCols) & nRow).Value = String(6, " ") & ReportSum(991)   'If your account information ...
         Call PrintExcel_BPage
         WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
         WksRpt1.Range(Chr(xCols) & nRow).Value = ReportSum(99101) 'and we shall update our records.
         
   End Select
   
   Call PrintExcel_BPage(2)
   WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
   WksRpt1.Range(Chr(xCols) & nRow).Value = String(6, " ") & ReportSum(93)  'With best regards.
   Call PrintExcel_BPage(2)
   WksRpt1.Range(Chr(xCols + 3) & nRow).Font.Size = 14
   WksRpt1.Range(Chr(xCols + 3) & nRow).Value = ReportSum(94) 'Sincerely yours,
   
   Call PrintExcel_BPage(2)
   'Modify by Morgan 2009/4/15 註記列印改先●再＊
   'Add by Morgan 2007/3/3 不必寄發電匯通知註記
   If strInform = "N" Then
      strTmp = "●" & strNo
   ElseIf bolEmail = True Or bol2Email = True Then
      strTmp = "＊" & strNo
   Else
      strTmp = strNo
   End If
   WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
   WksRpt1.Range(Chr(xCols) & nRow).Value = strTmp  '代理人編號前面加註記
   
End Sub

