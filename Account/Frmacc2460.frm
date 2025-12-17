VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc2460 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "FC收據列印"
   ClientHeight    =   4260
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5472
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5472
   Begin VB.CheckBox Check1 
      Caption         =   "產生特殊收據清單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3300
      TabIndex        =   29
      Top             =   2430
      Width           =   2000
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
      Height          =   315
      Left            =   2685
      MaxLength       =   15
      TabIndex        =   8
      Top             =   2040
      Width           =   852
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
      Left            =   1485
      MaxLength       =   15
      TabIndex        =   7
      Top             =   2040
      Width           =   852
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   1665
      Style           =   2  '單純下拉式
      TabIndex        =   23
      Top             =   2790
      Width           =   3270
   End
   Begin VB.TextBox txtReceiver 
      Height          =   285
      Left            =   1170
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3810
      Width           =   3750
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
      Left            =   1755
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2430
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
      Left            =   1485
      MaxLength       =   1
      TabIndex        =   0
      Top             =   360
      Width           =   405
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   225
      Left            =   5025
      ScaleHeight     =   180
      ScaleWidth      =   324
      TabIndex        =   17
      Top             =   2820
      Visible         =   0   'False
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
      Height          =   300
      Left            =   3405
      MaxLength       =   15
      TabIndex        =   6
      Top             =   1650
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
      Height          =   300
      Left            =   1485
      MaxLength       =   15
      TabIndex        =   5
      Top             =   1650
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
      Height          =   300
      Left            =   1485
      TabIndex        =   3
      Top             =   1260
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
      Height          =   300
      Left            =   3405
      TabIndex        =   4
      Top             =   1260
      Width           =   1572
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
      Left            =   270
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   3180
      Width           =   4965
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1485
      TabIndex        =   1
      Top             =   900
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
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   3405
      TabIndex        =   2
      Top             =   900
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
      AutoSize        =   -1  'True
      Caption         =   "強制列印會印出符合條件之所有要通知的收據"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1485
      TabIndex        =   28
      Top             =   690
      Width           =   3600
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "大陸一定要輸！"
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
      Left            =   3645
      TabIndex        =   27
      Top             =   2040
      Width           =   1680
   End
   Begin VB.Label Label13 
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
      Left            =   2445
      TabIndex        =   26
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  '透明
      Caption         =   "國籍"
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
      Left            =   285
      TabIndex        =   25
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "印表機"
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
      Left            =   285
      TabIndex        =   24
      Top             =   2820
      Width           =   675
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "代理人編號前符號：""●""  不必寄發收 據；""＊""  要發EMail。"
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
      Height          =   195
      Left            =   225
      TabIndex        =   22
      Top             =   60
      Width           =   5055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "收件人"
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
      Left            =   270
      TabIndex        =   21
      Top             =   3810
      Width           =   855
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
      Left            =   285
      TabIndex        =   19
      Top             =   2430
      Width           =   2625
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "輸出選項           (1:列印 2.發EMail 3:強制列印)"
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
      Left            =   285
      TabIndex        =   18
      Top             =   390
      Width           =   4755
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "收款單號"
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
      Left            =   285
      TabIndex        =   16
      Top             =   1650
      Width           =   1215
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
      Left            =   3165
      TabIndex        =   15
      Top             =   1650
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
      Left            =   3165
      TabIndex        =   14
      Top             =   1260
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "代理人編號"
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
      Left            =   285
      TabIndex        =   13
      Top             =   1260
      Width           =   1215
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
      Height          =   255
      Left            =   3165
      TabIndex        =   12
      Top             =   900
      Width           =   255
   End
   Begin VB.Label Label1 
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
      Index           =   0
      Left            =   285
      TabIndex        =   11
      Top             =   900
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc2460"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/2 日期欄已修改
Option Explicit
Public adoacc1k0 As New ADODB.Recordset
Public adoquery As New ADODB.Recordset
Public adocheck As New ADODB.Recordset
Private Const intLeft As Integer = 500
'Modified by Morgan 2012/3/14 新信紙上移 1 cm
'Private Const intTop As Integer = 300
Private Const intTop As Integer = 0
Dim strSql As String
Dim strNo As String
Dim douAmount As Double
Dim strAmount As String
Dim intLength As Integer
Dim intCounter As Integer
Dim douUSDollar As Double
Dim strLanguage As String
Dim strCurrency As String
Dim strFNo As String
'2005/9/12 ADD BY SONIA 有暫收款
Dim strA0Y10 As String
'Add by Morgan 2006/12/11
Dim strLangTmp As String '暫存語文
Dim bolEmail As Boolean '是否發EMail
Dim strInform As String '是否寄發收據
Dim strEMailBox As String 'EMail Box
Dim strEmailCC As String 'Added by Lydia 2024/09/18 財務副本信箱
Dim strPicLetter As String '暫存圖檔路徑  'Memo by Lydia 2024/12/19 信頭/尾的暫存圖: 從strPicFileName=>strPicLetter
Dim strPicFileNames As String '暫存圖檔路徑組(*號分隔)
Dim iPageNo As Integer '頁數
Dim strRecDate As String '收款日期
Dim strRecAmount As String '收款金額
Dim douExtRate As Double '字型位置縮放比
Dim lngX As Long, lngY As Long
'Add by Morgan 2008/4/16
Dim bol2File As Boolean '是否產生電子檔
Dim bol2Printer As Boolean '是否列印
Dim strSavePath As String '電子檔存放路徑
Dim bolChinese As Boolean '本文是否印中文
Dim iCopy As Integer
Dim strA1K28 As String '請款對象 Added by Morgan 2012/7/11
Dim strPrinter As String, strPrinter2 As String 'Added by Morgan 2012/10/11
Dim bolChina As Boolean '是否大陸 Add By Sindy 2014/3/10
Dim strNoList As String 'Added by Morgan 2015/12/29 請款單號
Dim strBCC As String 'Added by Morgan 2016/3/11 密件副本(T案要寄給程序)
Dim strContent As String 'Added by Morgan 2016/4/12
'Added by Lydia 2024/12/19 Excel列印
Dim xRows As Integer, xRowE As Integer 'Excel列印資料的起始,終止位置
Dim xCols As Integer, xColE As Integer 'Excel列印資料的起始/終止欄位Ascii值
Dim nRow As Integer '目前資料列位置
Private Const maxRows As Integer = 42 '頁面最大列數
Dim bolColTitle As Boolean '是否列印欄位抬頭
Dim m_iNo As Integer, m_iNo2 As Integer '圖檔編號
Dim strPrtPath As String, strPrtFile As String '列印Excel檔案路徑,名稱
Dim xlsRpt As New Excel.Application
Dim WksRpt1 As New Worksheet
Dim oShape
Dim oShape2
Dim oShape3

Private Sub Command2_Click()
   If FormCheck = False Then
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   'Add by Amy 2020/10/06 清暫存檔
   cnnConnection.Execute "Delete From Accrpt2460 Where ID='" & strUserNum & "' "
   PUB_SetOsDefaultPrinter cmbPrinter 'Added by Lydia 2025/03/07 切換Word/Excel印表機
   'Modify by Morgan 2006/12/14
   PUB_RestorePrinter cmbPrinter 'Added by Morgan 2012/10/11
   'Modified by Lydia 2024/12/19 改用EXCEL
   'PrintData
   PrintExcelMain
   
   PUB_SetOsDefaultPrinter strPrinter  'Added by Lydia 2025/03/07 切換Word/Excel印表機
   PUB_RestorePrinter strPrinter 'Added by Morgan 2012/10/11
   'Add by Amy 2020/10/06 +特殊收據清單
   If Check1.Value = vbChecked Then
        If SpecReceiptList = False Then
            MsgBox "無特殊收據清單產生！"
        End If
   End If
   'end 2020/10/06
   FormClear
   Screen.MousePointer = vbDefault
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
End Sub

'Add by Amy 2020/10/06  產生特殊收據清單
Private Function SpecReceiptList() As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim intQ As Integer
    Dim strQ As String, strText As String
    
    SpecReceiptList = False
    '抓取有特殊收據備註資料
    strQ = "Select fa01||fa02 AgNo,Decode(fa05,null,Nvl(fa04,fa06),fa05||' '||fa63||' '||fa64||' '||fa65) FName,fa118 Memo From Accrpt2460,FAgent " & _
                "Where ID='" & strUserNum & "' And SubStr(R001,1,8)=fa01(+) And SubStr(R001,9,1)=fa02(+) And fa01 is not null And fa125='Y' " & _
    "Union Select cu01||cu02 AgNo,Decode(cu05,null,Nvl(cu04,cu06),cu05||' '||cu88||' '||cu89||' '||cu90) FName,cu159 Memo From Accrpt2460,Customer " & _
                "Where ID='" & strUserNum & "' And SubStr(R001,1,8)=cu01(+) And SubStr(R001,9,1)=cu02(+) And cu01 is not null And cu184='Y' "
     intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        Do While RsQ.EOF = False
            strText = strText & "" & RsQ.Fields("AgNo") & "　" & PUB_StrToStr(RsQ.Fields("FName"), 15) & "　" & RsQ.Fields("Memo") & vbCrLf
            RsQ.MoveNext
        Loop
        strText = "特殊收據備註如下：" & vbCrLf & vbCrLf & strText
        PUB_SendMail strUserNum, strUserNum, "", "特殊收據備註清單", strText
        SpecReceiptList = True
    End If
    RsQ.Close
End Function

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
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(100)
   
   PUB_SetPrinter Me.Name, cmbPrinter, strPrinter 'Added by Morgan 2012/10/11
   'Remove by Morgan 2014/5/14 電子化不再留卷
   'PUB_SetPrinter Me.Name, cmbPrinter2, strPrinter2 'Added by Morgan 2012/10/15
   
   'Added by Lydia 2024/12/19
   strPrtPath = App.path & "\" & strUserNum
   Call Pub_ChkExcelPath(strPrtPath)
   Call PUB_KillTempFile(strUserNum & "\$*.*")
   'end 2024/12/19
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
   
   'Remove by Morgan 2014/5/14 電子化不再留卷
   'If cmbPrinter2.Text <> cmbPrinter2.Tag Then
   '    PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter2.Name, "0", "0", Me.cmbPrinter2.Text
   'End If
   
   Set Frmacc2460 = Nothing
End Sub

'add by sonia 2015/5/19
Private Sub MaskEdBox1_LostFocus()
   If MaskEdBox1.Text <> "___/__/__" And (MaskEdBox2.Text = "___/__/__" Or MaskEdBox2.Text = "") Then
      MaskEdBox2 = MaskEdBox1
   End If
End Sub
'2015/5/19 end

'Add by Amy 2020/12/15 有輸收款日期且輸出選項="2"(發email),產生特殊收據清單預設勾選-莘
Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    Check1.Value = 0
    If (MaskEdBox1.Text <> "___/__/__" And MaskEdBox1.Text <> MsgText(601)) _
          Or (MaskEdBox2.Text <> "___/__/__" And MaskEdBox2.Text <> MsgText(601)) Then
        If Text5 = "2" Then
            Check1.Value = 1
        End If
    End If
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
    Check1.Value = 0
    If (MaskEdBox1.Text <> "___/__/__" And MaskEdBox1.Text <> MsgText(601)) _
          Or (MaskEdBox2.Text <> "___/__/__" And MaskEdBox2.Text <> MsgText(601)) Then
        If Text5 = "2" Then
            Check1.Value = 1
        End If
    End If
End Sub
'end 2020/12/15

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
   Text3 = ""
   Text4 = ""
   Text5 = ""
   Text6 = ""
   'Add by Amy 20130703
   Text7 = ""
   Text8 = ""
   'end 20130703
   Text5.SetFocus
   Check1.Value = 0 'Add by Amy 2020/10/06
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead(Optional bolPrint As Boolean = True)
   Dim intRow As Integer
   Dim iW0 As Integer, iH0 As Integer, iW1 As Integer, iH1 As Integer
   Dim iW0x As Integer, iH0x As Integer, iW1x As Integer, iH1x As Integer
   'Add by Amy 2018/10/31
   Dim strFA17 As String, strFA18 As String, strFA19 As String, strFA20 As String, strFA21 As String, strFA22 As String, strFA70 As String, strFA23 As String
   Dim strFA32 As String, strFA33 As String, strFA34 As String, strFA35 As String, strFA36 As String
   
   iPageNo = iPageNo + 1
   
   If bol2Printer = True Then
      Printer.FontSize = 14
      Printer.Font = "Times New Roman"
   End If
   If bol2File = True Then
      Picture1.FontSize = 14 * douExtRate
      Picture1.Font = "Times New Roman"
      Picture1.AutoRedraw = True
   End If
   
   intRow = 5
   If bolPrint Then
   
      'Modify by Morgan 2006/10/25 沒設定時預設英文
      'strLanguage = "2"
      strLanguage = strLangTmp
      If strLanguage <> "1" Then strLanguage = "2"
      'end 2006/10/25
      
      'Add by Morgan 2005/3/18 代理人Y20412要求不寄紙本故加註"收據不印"字樣--婧瑄
      If Left("" & adoacc1k0.Fields("fa01").Value, 5) = "Y20412" Then
         If bol2Printer = True Then
            Printer.FontSize = 20
            Printer.CurrentX = intLeft + 4580
            Printer.CurrentY = intTop + 2200 + intRow * 300
            Printer.Print "收據不寄"
         End If
         If bol2File = True Then
            Picture1.FontSize = 20 * douExtRate
            Picture1.CurrentX = (intLeft + 4580) * douExtRate
            Picture1.CurrentY = (intTop + 2200 + intRow * 300) * douExtRate
            Picture1.Print "收據不寄"
         End If
      End If
            
      'Add by Morgan 2009/4/8 中文加印 RECEIPT
      strExc(1) = "": iW0 = 0: iW0x = 0: iH0 = 0: iH0x = 0: iW1 = 0: iW1x = 0: iH1 = 0: iH1x = 0
      
      'Modify by Morgan 2006/10/26 加中文格式
      If strLanguage = "1" Then
         '2009/3/31 MODIFY BY SONIA 葉經理提出,經婧瑄確認
          'strExc(0) = "收款確認通知書"
         'Modified by Morgan 2024/9/5
         'strExc(0) = "　　收　據　　"
         'strExc(1) = "RECEIPT"
         strExc(0) = "　發　票　"
         'end 2024/9/5
         lngX = 4100
         
         If bol2Printer = True Then
            Printer.FontBold = True
            'Removed by Morgan 2024/9/5
            'Printer.FontSize = 12
            'iW1 = Printer.TextWidth(strExc(1))
            'iH1 = Printer.TextHeight(strExc(1))
            'end 2024/9/5
            Printer.FontSize = 20
            iW0 = Printer.TextWidth(strExc(0))
            iH0 = Printer.TextHeight(strExc(0))
         End If
         If bol2File = True Then
            Picture1.FontBold = True
            'Removed by Morgan 2024/9/5
            'Picture1.FontSize = 12 * douExtRate
            'iW1x = Picture1.TextWidth(strExc(1))
            'iH1x = Picture1.TextHeight(strExc(1))
            'end 2024/9/5
            Picture1.FontSize = 24 * douExtRate
            iW0x = Picture1.TextWidth(strExc(0))
            iH0x = Picture1.TextHeight(strExc(0))
         End If
      Else
         strExc(0) = "RECEIPT"
         lngX = 4650
         
         If bol2Printer = True Then
            Printer.FontSize = 26
            iW0 = Printer.TextWidth(strExc(0))
            iH0 = Printer.TextHeight(strExc(0))
         End If
         If bol2File = True Then
            Picture1.FontSize = 26 * douExtRate
            iW0x = Picture1.TextWidth(strExc(0))
            iH0x = Picture1.TextHeight(strExc(0))
         End If
      End If
      
      'Modify by Morgan 2011/8/17
      'lngY = 1400 + intRow * 300
      lngY = intTop + 900 + intRow * 300
      
      'Add by Morgan 2009/4/8
      If strExc(1) <> "" Then
         lngY = lngY - 250
      End If
      
      If bol2Printer = True Then
         Printer.DrawWidth = 5 'Add by Morgan 2010/5/11 框加粗
         
         '要先印框再印字
         Printer.Line (lngX - 60, lngY - 60)-(lngX + iW0 + 60, lngY - 60)
         Printer.Line (lngX - 60, lngY - 60)-(lngX - 60, lngY + iH0 + iH1 + 60)
         Printer.CurrentX = lngX
         Printer.CurrentY = lngY
         Printer.Print strExc(0)
         
         'Add by Morgan 2009/4/8 中文加印 RECEIPT
         If strExc(1) <> "" Then
            Printer.CurrentX = lngX + (iW0 - iW1) / 2
            Printer.CurrentY = lngY + iH0
            Printer.FontSize = 12
            Printer.Print strExc(1)
         End If
         
         Printer.Line (lngX + iW0 + 60, lngY - 60)-(lngX + iW0 + 60, lngY + iH0 + iH1 + 60)
         Printer.Line (lngX - 60, lngY + iH0 + iH1 + 60)-(lngX + iW0 + 60, lngY + iH0 + iH1 + 60)
         
         Printer.DrawWidth = 1
         
         '收款日期起日
         Printer.FontBold = False
         Printer.FontSize = 12
         Printer.CurrentX = intLeft + 8500
         Printer.CurrentY = intTop + 1400 + intRow * 300
      End If
      
      If bol2File = True Then
         lngX = lngX * douExtRate
         lngY = lngY * douExtRate
         Picture1.Line (lngX - 60, lngY - 60)-(lngX + iW0x + 60, lngY - 60)
         Picture1.Line (lngX - 60, lngY - 60)-(lngX - 60, lngY + iH0x + iH1x + 60)
         Picture1.CurrentX = lngX
         Picture1.CurrentY = lngY
         Picture1.Print strExc(0)
         
         'Add by Morgan 2009/4/8 中文加印 RECEIPT
         If strExc(1) <> "" Then
            'Modified by Morgan 2014/5/26
            'Picture1.CurrentX = ((lngX + (iW0 - iW1) / 2) + 150) * douExtRate '(lngX + (iW0x - iW1x) / 2) * douExtRate Modify By Sindy 2014/3/24
            'Picture1.CurrentY = ((lngY + iH0) + 60) * douExtRate '(lngY + iH0x) * douExtRate Modify By Sindy 2014/3/24
            Picture1.CurrentX = lngX + (iW0x - iW1x) / 2
            Picture1.CurrentY = lngY + iH0x
            'end 2014/5/26
            Picture1.FontSize = 12 * douExtRate
            Picture1.Print strExc(1)
         End If
         
         Picture1.Line (lngX + iW0x + 60, lngY - 60)-(lngX + iW0x + 60, lngY + iH0x + iH1x + 60)
         Picture1.Line (lngX - 60, lngY + iH0x + iH1x + 60)-(lngX + iW0x + 60, lngY + iH0x + iH1x + 60)
         
         Picture1.FontBold = False
         Picture1.FontSize = 12 * douExtRate
         Picture1.CurrentX = (intLeft + 8500) * douExtRate
         Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
      End If
      'Modify by Morgan 2006/10/26 加中文格式
      If strLanguage = "1" Then
         If Me.MaskEdBox1.Text = "___/__/__" Then
            If bol2Printer = True Then
               Printer.Print Format(strSrvDate(1), "#### 年 ## 月 ## 日")
            End If
            If bol2File = True Then
               Picture1.Print Format(strSrvDate(1), "#### 年 ## 月 ## 日")
            End If
         Else
            If bol2Printer = True Then
               Printer.Print Format(DBDATE(Me.MaskEdBox1.Text), "#### 年 ## 月 ## 日")
            End If
            If bol2File = True Then
               Picture1.Print Format(DBDATE(Me.MaskEdBox1.Text), "#### 年 ## 月 ## 日")
            End If
         End If
      Else
         'edit by nick 2004/11/26 有輸日期起時才帶，沒有將用當天
         If Me.MaskEdBox1.Text = "___/__/__" Then
            If bol2Printer = True Then
               Printer.Print Format(AFDate(strSrvDate(1)), "mmm. d, yyyy")
            End If
            If bol2File = True Then
               Picture1.Print Format(AFDate(strSrvDate(1)), "mmm. d, yyyy")
            End If
         Else
            If bol2Printer = True Then
               Printer.Print Format(AFDate(ChangeTStringToWString(Replace(Me.MaskEdBox1.Text, "/", ""))), "mmm. d, yyyy")
            End If
            If bol2File = True Then
               Picture1.Print Format(AFDate(ChangeTStringToWString(Replace(Me.MaskEdBox1.Text, "/", ""))), "mmm. d, yyyy")
            End If
         End If
      End If
      intRow = intRow + 2
      
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
      'end 2018/10/31
            
      'Modify by Amy 2018/10/31 地址改為變數判斷
      Select Case strLanguage
         Case "1" '中文(中-->英-->日)
            '代理人名稱
            'Modify By Sindy 2013/4/29
            If IsNull(adoacc1k0.Fields("a0y19").Value) = False Then
               XPrint "" & adoacc1k0.Fields("a0y19").Value, intRow
               intRow = intRow + 1
            Else
            '2013/4/29 End
               If IsNull(adoacc1k0.Fields("fa04").Value) = False Then
                  XPrint "" & adoacc1k0.Fields("fa04").Value, intRow
                  intRow = intRow + 1
               ElseIf IsNull(adoacc1k0.Fields("fa05").Value) = False Then
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print "" & adoacc1k0.Fields("fa05").Value
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print "" & adoacc1k0.Fields("fa05").Value
                  End If
                  If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
                     intRow = intRow + 1
                     If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print "" & adoacc1k0.Fields("fa63").Value
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (intLeft) * douExtRate
                        Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                        Picture1.Print "" & adoacc1k0.Fields("fa63").Value
                     End If
                  End If
                  If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
                     intRow = intRow + 1
                     If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print "" & adoacc1k0.Fields("fa64").Value
                     End If
                     If bol2File = True Then
                         Picture1.CurrentX = (intLeft) * douExtRate
                         Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                         Picture1.Print "" & adoacc1k0.Fields("fa64").Value
                     End If
                  End If
                  If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
                     intRow = intRow + 1
                     If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print "" & adoacc1k0.Fields("fa65").Value
                     End If
                     If bol2File = True Then
                         Picture1.CurrentX = (intLeft) * douExtRate
                         Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                         Picture1.Print "" & adoacc1k0.Fields("fa65").Value
                     End If
                  End If
               ElseIf IsNull(adoacc1k0.Fields("fa06").Value) = False Then
                   '日文名稱
                   XPrint "" & adoacc1k0.Fields("fa06").Value, intRow
                   intRow = intRow + 1
               End If
            End If
            '代理人地址
            If strFA17 <> MsgText(601) Then  'If IsNull(adoacc1k0.Fields("fa17").Value) = False Then
               '中文地址
               'Removed by Morgan 2024/2/23
               'If bol2Printer = True Then
               '   Printer.CurrentX = intLeft
               '   Printer.CurrentY = intTop + 1400 + intRow * 250
               'End If
               'If bol2File = True Then
               '   Picture1.CurrentX = (intLeft) * douExtRate
               '   Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
               'End If
               'end 2024/2/23
               If LenB(strFA17) > 44 Then
                  'Modified by Morgan 2024/2/23
                  'If bol2Printer = True Then
                  '   Printer.Print MidB(strFA17, 1, 44)
                  'End If
                  'If bol2File = True Then
                  '   Picture1.Print MidB(strFA17, 1, 44)
                  'End If
                  strExc(0) = PUB_StrToStr(strFA17, 42)
                  strFA17 = Mid(strFA17, Len(strExc(0)) + 1)
                  XPrint strExc(0), intRow
                  'end 2024/2/23
               Else
                  'Modified by Morgan 2024/2/23
                  'If bol2Printer = True Then
                  '   Printer.Print strFA17
                  'End If
                  'If bol2File = True Then
                  '   Picture1.Print strFA17
                  'End If
                  XPrint strFA17, intRow
                  'end 2024/2/23
                  strFA17 = "" 'Added by Morgan 2024/9/6
               End If
               intRow = intRow + 1
               'Modified by Morgan 2024/2/23
               'If LenB(strFA17) > 44 Then
               '   If bol2Printer = True Then
               '      Printer.CurrentX = intLeft
               '      Printer.CurrentY = intTop + 1400 + intRow * 250
               '      Printer.Print MidB(strFA17, 45, LenB(strFA17) - 44)
               '   End If
               '   If bol2File = True Then
               '      Picture1.CurrentX = (intLeft) * douExtRate
               '      Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
               '      Picture1.Print MidB(strFA17, 45, LenB(strFA17) - 44)
               '   End If
               'End If
               If strFA17 <> "" Then
                  XPrint strFA17, intRow
               End If
               intRow = intRow + 4
               
            'ElseIf IsNull(adoacc1k0.Fields("fa32").Value) = False Or IsNull(adoacc1k0.Fields("fa18").Value) = False Then
            ElseIf strFA32 <> MsgText(601) Or strFA18 <> MsgText(601) Then
               'If IsNull(adoacc1k0.Fields("fa32").Value) Then
               If strFA32 = MsgText(601) Then
                  If strFA18 <> MsgText(601) Then
                     '英文地址1
                     intRow = intRow + 1
                     If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print strFA18
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (intLeft) * douExtRate
                        Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                        Picture1.Print strFA18
                     End If
                  End If
               Else
                  'POB1
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA32
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA32
                  End If
               End If
               'If IsNull(adoacc1k0.Fields("fa32").Value) Then
               If strFA32 = MsgText(601) Then
                  If strFA19 <> MsgText(601) Then
                     '英文地址2
                     intRow = intRow + 1
                     If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print strFA19
                     End If
                     If bol2File = True Then
                       Picture1.CurrentX = (intLeft) * douExtRate
                       Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                       Picture1.Print strFA19
                     End If
                  End If
               Else
                  'POB2
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA33
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA33
                  End If
               End If
               'If IsNull(adoacc1k0.Fields("fa32").Value) Then
               If strFA32 = MsgText(601) Then
                  If strFA20 <> MsgText(601) Then
                     '英文地址3
                     intRow = intRow + 1
                     If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print strFA20
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (intLeft) * douExtRate
                        Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                        Picture1.Print strFA20
                     End If
                  End If
               Else
                  'POB3
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA34
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA34
                  End If
               End If
               'If IsNull(adoacc1k0.Fields("fa32").Value) Then
               If strFA32 = MsgText(601) Then
                  '英文地址4
                  If strFA21 <> MsgText(601) Then
                     intRow = intRow + 1
                     If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print strFA21
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (intLeft) * douExtRate
                        Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                        Picture1.Print strFA21
                     End If
                  End If
               Else
                  'POB1
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA35
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA35
                  End If
               End If
               'If IsNull(adoacc1k0.Fields("fa32").Value) Then
               If strFA32 = MsgText(601) Then
                  '英文地址5
                  If strFA22 <> MsgText(601) Then
                     intRow = intRow + 1
                     If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print strFA22
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (intLeft) * douExtRate
                        Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                        Picture1.Print strFA22
                     End If
                  End If
                  
                  'Add by Morgan 2011/5/25
                  '英文地址6
                  If strFA70 <> MsgText(601) Then
                     intRow = intRow + 1
                     If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print strFA70
                     End If
                     If bol2File = True Then
                        Picture1.CurrentX = (intLeft) * douExtRate
                        Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                        Picture1.Print strFA70
                     End If
                  End If
                  
               Else
                  'POB5
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA36
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA36
                  End If
               End If
            ElseIf strFA23 <> MsgText(601) Then
               '日文地址
               'Modified by Morgan 2024/2/23
               'If bol2Printer = True Then
               '   Printer.CurrentX = intLeft
               '   Printer.CurrentY = intTop + 1400 + intRow * 250
               '   Printer.Print strFA23
               'End If
               'If bol2File = True Then
               '   Picture1.CurrentX = (intLeft) * douExtRate
               '   Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
               '   Picture1.Print strFA23
               'End If
               XPrint strFA23, intRow
               'end 2024/2/23
               intRow = intRow + 4
            End If
            
         Case "2" '英文(英-->中-->日)
            '代理人名稱
            If IsNull(adoacc1k0.Fields("fa05").Value) = False Then
               If bol2Printer = True Then
                  Printer.CurrentX = intLeft
                  Printer.CurrentY = intTop + 1400 + intRow * 250
                  Printer.Print "" & adoacc1k0.Fields("fa05").Value
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (intLeft) * douExtRate
                  Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                  Picture1.Print "" & adoacc1k0.Fields("fa05").Value
               End If
               If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print "" & adoacc1k0.Fields("fa63").Value
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print "" & adoacc1k0.Fields("fa63").Value
                  End If
               End If
                
                If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
                   intRow = intRow + 1
                   If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print "" & adoacc1k0.Fields("fa64").Value
                   End If
                   If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print "" & adoacc1k0.Fields("fa64").Value
                   End If
                End If
                If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
                   intRow = intRow + 1
                   If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print "" & adoacc1k0.Fields("fa65").Value
                   End If
                   If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print "" & adoacc1k0.Fields("fa65").Value
                   End If
                End If
            ElseIf IsNull(adoacc1k0.Fields("fa04").Value) = False Then
               XPrint "" & adoacc1k0.Fields("fa04").Value, intRow
            ElseIf IsNull(adoacc1k0.Fields("fa06").Value) = False Then
               XPrint "" & adoacc1k0.Fields("fa06").Value, intRow
            End If
            '代理人地址
            'If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If strFA32 = MsgText(601) Then
                If strFA18 <> MsgText(601) Then
                    intRow = intRow + 1
                    If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print strFA18
                    End If
                    If bol2File = True Then
                        Picture1.CurrentX = (intLeft) * douExtRate
                        Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                        Picture1.Print strFA18
                    End If
                ElseIf strFA17 <> MsgText(601) Then
                    intRow = intRow + 1
                    XPrint strFA17, intRow
                ElseIf strFA23 <> MsgText(601) Then
                    intRow = intRow + 1
                    XPrint strFA23, intRow
                End If
            Else
               intRow = intRow + 1
               If bol2Printer = True Then
                  Printer.CurrentX = intLeft
                  Printer.CurrentY = intTop + 1400 + intRow * 250
                  Printer.Print strFA32
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (intLeft) * douExtRate
                  Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                  Picture1.Print strFA32
               End If
            End If
            'If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If strFA32 = MsgText(601) Then
               If strFA19 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA19
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA19
                  End If
               End If
            Else
               intRow = intRow + 1
               If bol2Printer = True Then
                  Printer.CurrentX = intLeft
                  Printer.CurrentY = intTop + 1400 + intRow * 250
                  Printer.Print strFA33
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (intLeft) * douExtRate
                  Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                  Picture1.Print strFA33
               End If
            End If
            'If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If strFA32 = MsgText(601) Then
               If strFA20 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA20
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA20
                  End If
               End If
            Else
               intRow = intRow + 1
               If bol2Printer = True Then
                  Printer.CurrentX = intLeft
                  Printer.CurrentY = intTop + 1400 + intRow * 250
                  Printer.Print strFA34
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (intLeft) * douExtRate
                  Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                  Picture1.Print strFA34
               End If
            End If
            'If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If strFA32 = MsgText(601) Then
               If strFA21 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA21
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA21
                  End If
               End If
            Else
               intRow = intRow + 1
               If bol2Printer = True Then
                  Printer.CurrentX = intLeft
                  Printer.CurrentY = intTop + 1400 + intRow * 250
                  Printer.Print strFA35
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (intLeft) * douExtRate
                  Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                  Picture1.Print strFA35
               End If
            End If
            'If IsNull(adoacc1k0.Fields("fa32").Value) Then
            If strFA32 = MsgText(601) Then
               If strFA22 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA22
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA22
                  End If
               End If
              'Add by Morgan 2011/5/25
              '英文地址6
               If strFA70 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA70
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA70
                  End If
               End If
               
            Else
               intRow = intRow + 1
               If bol2Printer = True Then
                  Printer.CurrentX = intLeft
                  Printer.CurrentY = intTop + 1400 + intRow * 250
                  Printer.Print strFA36
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (intLeft) * douExtRate
                  Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                  Picture1.Print strFA36
               End If
            End If
            
         Case "3" '日文(日-->英-->中)
            '代理人名稱
            If IsNull(adoacc1k0.Fields("fa06").Value) = False Then
                XPrint "" & adoacc1k0.Fields("fa06").Value, intRow
                intRow = intRow + 1
            ElseIf IsNull(adoacc1k0.Fields("fa05").Value) = False Then
                If bol2Printer = True Then
                  Printer.CurrentX = intLeft
                  Printer.CurrentY = intTop + 1400 + intRow * 250
                  Printer.Print "" & adoacc1k0.Fields("fa05").Value
                End If
                If bol2File = True Then
                  Picture1.CurrentX = (intLeft) * douExtRate
                  Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                  Picture1.Print "" & adoacc1k0.Fields("fa05").Value
                End If
                If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
                    intRow = intRow + 1
                    If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print "" & adoacc1k0.Fields("fa63").Value
                    End If
                    If bol2File = True Then
                        Picture1.CurrentX = (intLeft) * douExtRate
                        Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                        Picture1.Print "" & adoacc1k0.Fields("fa63").Value
                    End If
                End If
                If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
                    intRow = intRow + 1
                    If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print "" & adoacc1k0.Fields("fa64").Value
                    End If
                    If bol2File = True Then
                        Picture1.CurrentX = (intLeft) * douExtRate
                        Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                        Picture1.Print "" & adoacc1k0.Fields("fa64").Value
                    End If
                End If
                If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
                    intRow = intRow + 1
                    If bol2Printer = True Then
                        Printer.CurrentX = intLeft
                        Printer.CurrentY = intTop + 1400 + intRow * 250
                        Printer.Print "" & adoacc1k0.Fields("fa65").Value
                    End If
                    If bol2File = True Then
                        Picture1.CurrentX = (intLeft) * douExtRate
                        Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                        Picture1.Print "" & adoacc1k0.Fields("fa65").Value
                    End If
                End If
            ElseIf IsNull(adoacc1k0.Fields("fa04").Value) = False Then
                XPrint "" & adoacc1k0.Fields("fa04").Value, intRow
                intRow = intRow + 1
            End If

            '代理人地址
            '順序：日文->POB->英文->中文
            '日文
            If strFA23 <> MsgText(601) Then
                XPrint strFA23, intRow
                intRow = intRow + 5
            'POB
            ElseIf strFA32 <> MsgText(601) Then
               'P0B1
               intRow = intRow + 1
               If bol2Printer = True Then
                  Printer.CurrentX = intLeft
                  Printer.CurrentY = intTop + 1400 + intRow * 250
                  Printer.Print strFA32
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (intLeft) * douExtRate
                  Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                  Picture1.Print strFA32
               End If
               'P0B2
               If strFA33 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA33
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA33
                  End If
               End If
               'P0B3
               If strFA34 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA34
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA34
                  End If
               End If
               'P0B4
               If strFA35 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA35
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA35
                  End If
               End If
               'P0B5
               If strFA36 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA36
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA36
                  End If
               End If
            '英文地址
            'ElseIf IsNull(adoacc1k0.Fields("fa18").Value) = False Then
            ElseIf strFA18 <> MsgText(601) Then
               '英文地址1
               intRow = intRow + 1
               If bol2Printer = True Then
                  Printer.CurrentX = intLeft
                  Printer.CurrentY = intTop + 1400 + intRow * 250
                  Printer.Print strFA18
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (intLeft) * douExtRate
                  Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                  Picture1.Print strFA18
               End If
               '英文地址2
               If strFA19 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA19
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA19
                  End If
               End If
               '英文地址3
               If strFA20 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA20
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA20
                  End If
               End If
               '英文地址4
               If strFA21 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA21
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA21
                  End If
               End If
               '英文地址5
               If strFA22 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA22
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA22
                  End If
               End If
              'Add by Morgan 2011/5/25
              '英文地址6
               If strFA70 <> MsgText(601) Then
                  intRow = intRow + 1
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft
                     Printer.CurrentY = intTop + 1400 + intRow * 250
                     Printer.Print strFA70
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft) * douExtRate
                     Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                     Picture1.Print strFA70
                  End If
               End If
               
            '中文
            ElseIf strFA17 <> MsgText(601) Then
               If bol2Printer = True Then
                  Printer.CurrentX = intLeft
                  Printer.CurrentY = intTop + 1400 + intRow * 250
                  Printer.Print strFA17
               End If
               If bol2File = True Then
                  Picture1.CurrentX = (intLeft) * douExtRate
                  Picture1.CurrentY = (intTop + 1400 + intRow * 250) * douExtRate
                  Picture1.Print strFA17
               End If
               intRow = intRow + 5
            End If
      End Select
      'end 2018/10/31
      'Modify by Morgan 2006/10/26 加中文格式
      If strLanguage = "1" Then
         strExc(1) = ""
         strExc(0) = "select SUM(A0Z04) FROM ACC0Z0 WHERE A0Z01='" & adoacc1k0("A0Y01") & "'"
         intI = 1
         'edit by nickc 2007/02/07 不用 dll 了
         'Set AdoRecordSet3 = objLawDll.ReadRstMsg(intI, strExc(0))
         Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "" & adoacc1k0("a1k18") & " " & Format(AdoRecordSet3.Fields(0) + Val("" & adoacc1k0("A0Y06")), FDollar)
         End If
         
         intRow = 14
         strExc(0) = "茲通知收到　貴方的付款，金額為 " & strExc(1) & "，已支付下列帳單。"
         If bol2Printer = True Then
            Printer.FontSize = 14
            Printer.CurrentX = intLeft + 500
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print strExc(0)
         End If
         If bol2File = True Then
            Picture1.FontSize = 14 * douExtRate
            Picture1.CurrentX = (intLeft + 500) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print strExc(0)
         End If
      Else
         intRow = 14
         If bol2Printer = True Then
            Printer.FontSize = 15
            Printer.CurrentX = intLeft + 500
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print ReportSum(100)
         End If
         If bol2File = True Then
            Picture1.FontSize = 15 * douExtRate
            Picture1.CurrentX = (intLeft + 500) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print ReportSum(100)
         End If
      End If
   End If
   If bol2Printer = True Then
      Printer.FontSize = 12
   End If
   If bol2File = True Then
      Picture1.FontSize = 12 * douExtRate
   End If
   intRow = intRow + 2
   'Modify by Morgan 2006/10/26 加中文格式
   If strLanguage = "1" Then
      'Add By Sindy 2014/3/12
      If bolChina = True Then
         If bol2Printer = True Then
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "我方文號"
            Printer.CurrentX = intLeft + 1400
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "貴方文號"
            Printer.CurrentX = intLeft + 4100
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "帳單編號"
            Printer.CurrentX = intLeft + 6000
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "金額"
            Printer.CurrentX = intLeft + 7300
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "案件名稱"
            Printer.CurrentX = intLeft + 9400
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "申請人"
         End If
         If bol2File = True Then
            Picture1.CurrentX = (intLeft) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "我方文號"
            Picture1.CurrentX = (intLeft + 1400) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "貴方文號"
            Picture1.CurrentX = (intLeft + 4100) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "帳單編號"
            Picture1.CurrentX = (intLeft + 6000) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "金額"
            Picture1.CurrentX = (intLeft + 7300) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "案件名稱"
            Picture1.CurrentX = (intLeft + 9500) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "申請人"
         End If
      Else
      '2014/3/12 END
         If bol2Printer = True Then
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "我方文號"
            Printer.CurrentX = intLeft + 2500
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "貴方文號"
            Printer.CurrentX = intLeft + 5500
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "帳單編號"
            Printer.CurrentX = intLeft + 8500
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "金額"
         End If
         If bol2File = True Then
            Picture1.CurrentX = (intLeft) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "我方文號"
            Picture1.CurrentX = (intLeft + 2500) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "貴方文號"
            Picture1.CurrentX = (intLeft + 5500) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "帳單編號"
            Picture1.CurrentX = (intLeft + 8500) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "金額"
         End If
      End If
   Else
      'Modify By Sindy 2014/7/3
      If bolChina = True Then
         If bol2Printer = True Then
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "OUR REF"
            Printer.CurrentX = intLeft + 1400
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "YOUR REF"
            Printer.CurrentX = intLeft + 4100
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "DEBIT NOTE."
            Printer.CurrentX = intLeft + 6000
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "AMOUNT"
            Printer.CurrentX = intLeft + 7300
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "TITLE"
            Printer.CurrentX = intLeft + 9400
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "APPLICANT"
         End If
         If bol2File = True Then
            Picture1.CurrentX = (intLeft) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "OUR REF"
            Picture1.CurrentX = (intLeft + 1400) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "YOUR REF"
            Picture1.CurrentX = (intLeft + 4100) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "DEBIT NOTE."
            Picture1.CurrentX = (intLeft + 6000) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "AMOUNT"
            Picture1.CurrentX = (intLeft + 7300) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "TITLE"
            Picture1.CurrentX = (intLeft + 9400) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "APPLICANT"
         End If
      Else
         If bol2Printer = True Then
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "OUR REF"
            Printer.CurrentX = intLeft + 2500
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "YOUR REF"
            Printer.CurrentX = intLeft + 5500
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "DEBIT NOTE."
            Printer.CurrentX = intLeft + 8500
            Printer.CurrentY = intTop + 1400 + intRow * 300
            Printer.Print "AMOUNT"
         End If
         If bol2File = True Then
            Picture1.CurrentX = (intLeft) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "OUR REF"
            Picture1.CurrentX = (intLeft + 2500) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "YOUR REF"
            Picture1.CurrentX = (intLeft + 5500) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "DEBIT NOTE."
            Picture1.CurrentX = (intLeft + 8500) * douExtRate
            Picture1.CurrentY = (intTop + 1400 + intRow * 300) * douExtRate
            Picture1.Print "AMOUNT"
         End If
      End If
      '2014/7/3 END
   End If
   If bol2Printer = True Then
      Printer.Line (intLeft, intTop + 1400 + intRow * 300 + 350)-(11000 + intLeft, intTop + 1400 + intRow * 300 + 350)
   End If
   If bol2File = True Then
      lngX = (intLeft) * douExtRate
      lngY = (intTop + 1400 + intRow * 300 + 350) * douExtRate
      Picture1.Line (lngX, lngY)-(Picture1.Width - lngX, lngY)
   End If
   
End Sub

'*************************************************
' 合計位置 -- 7(11)列
'
'*************************************************
Private Sub PrintSum()
   
   Dim strDollar As String
   Dim bNewPage As Boolean
   'Modify by Morgan 2006/4/13 表尾控制須包含合計資料
   'Printer.Line (0 + intLeft, 5700 + intCounter * 300 - 200 - intTop)-(10500 + intLeft, 5700 + intCounter * 300 - 200 - intTop)
   Dim iResLine As Integer
   If strA0Y10 <> "" Then
      iResLine = 11
   Else
      iResLine = 7
   End If
   
   'Modify by Morgan 2009/8/26 修正跳頁控制
   bNewPage = False
   If bol2Printer = True Then
      If Printer.CurrentY + iResLine * Printer.TextHeight("X") + 1400 > Printer.ScaleHeight Then
         Printer.NewPage
         bNewPage = True
      End If
   End If
   If bol2File = True Then
      If Picture1.CurrentY + iResLine * Picture1.TextHeight("X") + 1400 * douExtRate > Picture1.ScaleHeight Then
         PicNewPage
         bNewPage = True
      End If
   End If
      
   If bNewPage = True Then
      intCounter = -4
      PrintHead False
   Else
      If bol2Printer = True Then
         Printer.Line (intLeft, intTop + 5100 + intCounter * 300 - 200)-(intLeft + 11000, intTop + 5100 + intCounter * 300 - 200)
      End If
      If bol2File = True Then
         lngX = (intLeft) * douExtRate
         lngY = (intTop + 5100 + intCounter * 300 - 200) * douExtRate
         Picture1.Line (lngX, lngY)-(Picture1.Width - lngX, lngY)
      End If
   End If
   '2006/4/13 end
   
   'Add by Morgan 2006/10/26
   If strLanguage = "1" Then
      strExc(0) = "金額合計"
      If bol2Printer = True Then
         'Add By Sindy 2014/3/12
         If bolChina = True Then
            Printer.CurrentX = intLeft + 5400 - Printer.TextWidth(strExc(0)) - 100
         Else
         '2014/3/12 END
            Printer.CurrentX = intLeft + 7400 - Printer.TextWidth(strExc(0)) - 100
         End If
         Printer.CurrentY = intTop + 5100 + intCounter * 300
         Printer.Print strExc(0)
      End If
      If bol2File = True Then
         'Add By Sindy 2014/3/12
         If bolChina = True Then
            Picture1.CurrentX = (intLeft + 5400 - 100) * douExtRate - Picture1.TextWidth(strExc(0))
         Else
         '2014/3/12 END
            Picture1.CurrentX = (intLeft + 7400 - 100) * douExtRate - Picture1.TextWidth(strExc(0))
         End If
         Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
         Picture1.Print strExc(0)
      End If
   End If
   
   strAmount = Format(douAmount, FDollar)
   
   If bol2Printer = True Then
      'Add By Sindy 2014/3/12
      If bolChina = True Then
         Printer.CurrentX = intLeft + 5400
      Else
      '2014/3/12 END
         Printer.CurrentX = intLeft + 7400
      End If
      Printer.CurrentY = intTop + 5100 + intCounter * 300
      Printer.Print strCurrency
      intLength = Printer.TextWidth(strAmount)
      'Add By Sindy 2014/3/12
      If bolChina = True Then
         Printer.CurrentX = intLeft + 7100 - intLength
      Else
      '2014/3/12 END
         Printer.CurrentX = intLeft + 9900 - intLength
      End If
      Printer.CurrentY = intTop + 5100 + intCounter * 300
      Printer.Print strAmount
   End If
   If bol2File = True Then
      'Add By Sindy 2014/3/12
      If bolChina = True Then
         Picture1.CurrentX = (intLeft + 5400) * douExtRate
      Else
      '2014/3/12 END
         Picture1.CurrentX = (intLeft + 7400) * douExtRate
      End If
      Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
      Picture1.Print strCurrency
      intLength = Picture1.TextWidth(strAmount)
      'Add By Sindy 2014/3/12
      If bolChina = True Then
         Picture1.CurrentX = (intLeft + 7100) * douExtRate - intLength
      Else
      '2014/3/12 END
         Picture1.CurrentX = (intLeft + 9900) * douExtRate - intLength
      End If
      Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
      Picture1.Print strAmount
      strRecAmount = strCurrency & " " & strAmount
   End If
   
   intCounter = intCounter + 1
   If bol2Printer = True Then
      'Add By Sindy 2014/3/12
      If bolChina = True Then
         Printer.CurrentX = intLeft + 5400
         Printer.CurrentY = intTop + 5100 + intCounter * 300
         Printer.Print String(15, "v")
      Else
      '2014/3/12 END
         Printer.CurrentX = intLeft + 7400
         Printer.CurrentY = intTop + 5100 + intCounter * 300
         Printer.Print String(23, "v")
      End If
   End If
   If bol2File = True Then
      'Add By Sindy 2014/3/12
      If bolChina = True Then
         Picture1.CurrentX = (intLeft + 5400) * douExtRate
         Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
         Picture1.Print String(15, "v")
      Else
      '2014/3/12 END
         Picture1.CurrentX = (intLeft + 7400) * douExtRate
         Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
         Picture1.Print String(23, "v")
      End If
   End If
   
   'Modify by Morgan 2006/10/26 加中文格式
   If strLanguage = "1" Then
      intCounter = intCounter + 4
      'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
      'strExc(0) = "台一國際專利商標事務所"
      strExc(0) = PUB_GetCompName2("1")
      'end 2020/3/30
      If bol2Printer = True Then
         Printer.FontSize = 14
         Printer.CurrentX = intLeft + 7200
         Printer.CurrentY = intTop + 5100 + intCounter * 300
         'Printer.Print strExc(0)   'cancel by sonia 2020/5/27
      End If
      If bol2File = True Then
         Picture1.FontSize = 14 * douExtRate
         Picture1.CurrentX = (intLeft + 7200) * douExtRate
         Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
         'Picture1.Print strExc(0)     'cancel by sonia 2020/5/27
      End If
      
      intCounter = intCounter + 1
      If bol2Printer = True Then
         Printer.CurrentX = intLeft + 7200 + Printer.TextWidth("台一國際")
         Printer.CurrentY = intTop + 5200 + intCounter * 300
         'Printer.Print "財務處"      'cancel by sonia 2020/5/27
      End If
      If bol2File = True Then
         Picture1.CurrentX = (intLeft + 7200) * douExtRate + Picture1.TextWidth("台一國際")
         Picture1.CurrentY = (intTop + 5200 + intCounter * 300) * douExtRate
         'Picture1.Print "財務處"     'cancel by sonia 2020/5/27
      End If
      
      'Added by Morgan 2024/9/5 加印智慧所收據章,固定印在第15列 (不支援png,但jpg又會遮到字,先取消)
      'If bol2File = True Then
      '   Call PrintPicture("2", 16, (intLeft + 7200) * douExtRate, (intTop + 5100 + 17 * 300) * douExtRate, douExtRate)
      'End If
      'end 2024/9/5
      
   Else
      intCounter = intCounter + 4
      If bol2Printer = True Then
         Printer.FontSize = 15
         Printer.CurrentX = intLeft + 8000
         Printer.CurrentY = intTop + 5100 + intCounter * 300
         'Printer.Print ReportSum(94)      'cancel by sonia 2020/5/27
      End If
      If bol2File = True Then
         Picture1.FontSize = 15 * douExtRate
      End If
   End If
   'Modify by Morgan 2009/4/15 註記列印改先●再＊
   intCounter = intCounter + 2
   If bol2Printer = True Then
      Printer.CurrentX = intLeft
      Printer.CurrentY = intTop + 5100 + intCounter * 300
      'Add by Morgan 2007/3/3 不必寄發收據註記
      If strInform = "N" Then
         '不必寄發收據
         Printer.Print "●" & strFNo
      ElseIf bolEmail = True Then
         '要發Mail
         Printer.Print "＊" & strFNo
      Else
         Printer.Print strFNo
      End If
   End If
   
   If bol2File = True Then
      Picture1.FontSize = 15 * douExtRate
      Picture1.CurrentX = (intLeft) * douExtRate
      Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
      'Add by Morgan 2007/3/3 不必寄發收據註記
      If strInform = "N" Then
         Picture1.Print "●" & strFNo
      ElseIf bolEmail = True Then
         Picture1.Print "＊" & strFNo
      Else
         Picture1.Print strFNo
      End If
   End If
   
    'Added by Lydia 2016/01/29 應代理人要求,大陸地區加印收據圖章
    If bolChina = True Then
       'Remove by Lydia 2016/02/04 因為印出來是黑色的,所以紙本不加圖章
       'If bol2Printer = True Then
       '    Call PrintPicture("1", 46, (intLeft + 7200) + Picture1.TextWidth("台一國際"), (intTop + 5200 + intCounter * 300), douExtRate)
       'End If
       'Removed by Morgan 2020/3/30 取消--婉莘
       'If bol2File = True Then
       '    Call PrintPicture("2", 46, (intLeft + 7200) * douExtRate + Picture1.TextWidth("台一"), (intTop + 5100 + intCounter * 300) * douExtRate, douExtRate)
       'End If
       'end 2020/3/30
    End If
    'end 2016/01/29
    
   '2005/9/12 ADD BY SONIA
   If strA0Y10 <> "" Then
      intCounter = intCounter + 2
      If bol2Printer = True Then
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 5100 + intCounter * 300
         'Modify By Sindy 2014/10/9
         'Printer.Print "P.S. : The credit shown in this receipt is an overpayment from your payment, which"
         Printer.Print "P.S. : The credit amount shown in this receipt derives from your payment, which "
      End If
      If bol2File = True Then
         Picture1.CurrentX = (intLeft) * douExtRate
         Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
         'Modify By Sindy 2014/10/9
         'Picture1.Print "P.S. : The credit shown in this receipt is an overpayment from your payment, which"
         Picture1.Print "P.S. : The credit amount shown in this receipt derives from your payment, which "
      End If
      intCounter = intCounter + 1
      If bol2Printer = True Then
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 5100 + intCounter * 300
         'Modify By Sindy 2014/10/9
         'Printer.Print "          can be deducted from your next payment if requested. Howerer, if you want"
         Printer.Print "          can be deducted from your next payment if requested. However, if you prefer"
      End If
      If bol2File = True Then
         Picture1.CurrentX = (intLeft) * douExtRate
         Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
         'Modify By Sindy 2014/10/9
         'Picture1.Print "          can be deducted from your next payment if requested. Howerer, if you want"
         Picture1.Print "          can be deducted from your next payment if requested. However, if you prefer"
      End If
      intCounter = intCounter + 1
      If bol2Printer = True Then
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 5100 + intCounter * 300
         'Modify By Sindy 2014/10/9
         'Printer.Print "          us to return this overpayment by a check, please advise us immediately."
         Printer.Print "          to return this credit, please advise us immediately."
      End If
      If bol2File = True Then
         Picture1.CurrentX = (intLeft) * douExtRate
         Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
         'Modify By Sindy 2014/10/9
         'Picture1.Print "          us to return this overpayment by a check, please advise us immediately."
         Picture1.Print "          to return this credit, please advise us immediately."
      End If
   End If
   '2005/9/12 END
   
   PrintCoverPage 'Added by Morgan 2012/7/11
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Private Function FormCheck() As Boolean
   If Text5 = "" Then
      MsgBox "輸出選項不可空白！"
      Text5.SetFocus
      Exit Function
   ElseIf Text5 = "2" And Text6 <> "Y" Then
      If MsgBox("確定要發EMail？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
         Text5.SetFocus
         Exit Function
      End If
   End If
   
   'Modify by Morgan 2010/6/10 只要有輸收款單號或日期就可執行
   If Text3 <> "" And Text4 <> "" Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox1.Text & MaskEdBox2 <> "" And MaskEdBox1.Text <> MsgText(29) And MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
   MsgBox MsgText(181), , MsgText(5)
End Function

Private Sub Text3_GotFocus()
TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
If Text3 = "" Then Exit Sub
If Mid(Text3, 1, 1) <> "M" Then
    MsgBox "收款單號應該是 M 開頭！", , "錯誤！"
    Text3.SetFocus
    Cancel = True
    Exit Sub
End If
If Not nickChgRan(Text3, Text4, "收款單號") Then
    Cancel = True
End If
End Sub

Private Sub Text4_GotFocus()
TextInverse Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Text4 = "" Then Exit Sub
If Mid(Text4, 1, 1) <> "M" Then
    MsgBox "收款單號應該是 M 開頭！", , "錯誤！"
    Text4.SetFocus
    Cancel = True
    Exit Sub
End If
If Not nickChgRan(Text3, Text4, "收款單號") Then
    Cancel = True
End If
End Sub
'Add by Morgan 2005/5/18 折行列印
'p_stContent=列印內容,p_iRow=起始行數
Private Sub XPrint(ByVal p_stContent As String, ByRef p_iRow As Integer)
   Dim iPos As Integer, strTemp As String
   iPos = 1
   'Modified by Morgan 2024/2/3
   'strTemp = Mid(p_stContent, iPos, 22)
   strTemp = PUB_StrToStr(p_stContent, 44)
   'end 2024/2/3
   Do
      If bol2Printer = True Then
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 1400 + p_iRow * 250
         'Modified by Morgan 2024/2/23
         'Printer.Print strTemp
         PUB_PrintUnicodeText strTemp, Printer.CurrentX, Printer.CurrentY, 0
         'end 2024/2/23
      End If
      If bol2File = True Then
         Picture1.CurrentX = (intLeft) * douExtRate
         Picture1.CurrentY = (intTop + 1400 + p_iRow * 250) * douExtRate
         'Modified by Morgan 2024/2/23
         'Picture1.Print strTemp
         PUB_PrintUnicodeText strTemp, Picture1.CurrentX, Picture1.CurrentY, 0, Picture1
         'end 2024/2/23
      End If
      
      'Modified by Morgan 2024/2/23
      'iPos = iPos + 22
      'strTemp = Mid(p_stContent, iPos, 22)
      p_stContent = Mid(p_stContent, Len(strTemp) + 1)
      strTemp = ""
      If p_stContent <> "" Then
         strTemp = PUB_StrToStr(p_stContent, 44)
      End If
      'end 2024/2/23
      If strTemp <> "" Then
         p_iRow = p_iRow + 1
      Else
         Exit Do
      End If
   Loop
End Sub

'Modify by Morgan 2006/12/1 加存圖檔功能
Private Sub PrintData()
   Dim strKeyNo As String
   Dim strDocNo As String
   Dim strSQL1 As String
   Dim strMailFailList() As String 'Mail 失敗清單
   Dim bNewPage As Boolean
   Dim iRound As Integer '迴圈次數
   Dim bolChgPrinter As Boolean 'Added by Morgan 2012/10/11
   Dim strNation As String 'Add by Amy 2013/07/03
   Dim strData As String 'Add By Sindy 2014/3/12
   Dim strPA77 As String 'Add by Amy 2016/03/31 彼所案號抓取欄位
   Dim dblAmount As Double 'Add by Morgan 2016/5/31
   Dim iPicNo As Integer 'Added by Morgan 2020/3/30
   Dim StrSQLa As String 'Add by Amy 2020/10/06
   
   'douExtRate = Screen.TwipsPerPixelX / 15 'Remove by Morgan 2011/10/5
   
   strSql = ""
   strSQL1 = ""
   '收款日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0y02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSQL1 = strSQL1 & " and a0y02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'Remove by Morgan 2010/6/10
'   'Add by Morgan 2009/9/28
'   Else
'      MsgBox "收款日期起日不可空白!!"
'      Exit Sub
   End If

   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0y02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSQL1 = strSQL1 & " and a0y02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'Remove by Morgan 2010/6/10
'   'Add by Morgan 2009/9/28
'   Else
'      MsgBox "收款日期迄日不可空白!!"
'      Exit Sub
   End If
   
   '代理人編號
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and decode(a0y18, '1', a0y07, '2', a0y08, a0y09) >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and decode(a0y18, '1', a0y07, '2', a0y08, a0y09) <= '" & Text2 & "'"
   End If
   '收款單號
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0y01>= '" & Text3 & "'"
   End If
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and a0y01 <= '" & Text4 & "'"
   End If
   
   'Add by Amy 2013/07/03 +國藉判斷
   If Text7 = "020" Then
        strNation = " Where fa10||''='020' "
   Else
        strNation = " Where fa10||''<>'020' "
   End If
   'end 2013/07/03
   
   'Added by Morgan 2015/12/29
   '發Email改只以內文說明不再夾帶收據圖檔(大陸地區除外)
   bolChinese = False 'Added by Morgan 2016/2/2
   strBCC = "" 'Added by Morgan 2016/3/11
   strContent = "" 'Added by Morgan 2016/4/2
   If Text5 = "2" And Text7 <> "020" And Text6 <> "Y" Then
      'Modified by Morgan 2016/3/11 +a1k13(商標案要BCC給程序)
      'Modified by Lydia 2024/09/18 +財務副本信箱emailcc：寄財務信箱一併CC副本>>decode(fa105||fa79,null,decode(cu115,null,'',cu200),fa134) as emailcc
      strExc(0) = "select * from (select a0y01,a0y02,AgNo,a0z02,nvl(fa79,nvl(fa16,nvl(cu115,cu20))) fa16, nvl(fa83,cu119) fa83,a1k29,nvl(fa10,cu10) fa10 " & _
         ",a1k13,decode(fa105||fa79,null,decode(cu115,null,'',cu200),fa134) as emailcc from (select a0y01,a0y02,a0z02,decode(a0y18, '1', a0y07, '2', a0y08, a0y09) AgNo,a1k29" & _
         ",a1k13 From acc0y0,acc0z0,acc1k0 where a0z01(+)=a0y01 and (length(a0z02)=9 or substr(a0z02,-2)='00') and a1k01(+)=a0z02" & strSql & _
         "),fagent,customer where fa01(+)=substr(AgNo,1,8) and fa02(+)=substr(AgNo,9) and cu01(+)=substr(AgNo,1,8) and cu02(+)=substr(AgNo,9)) X" & strNation & " order by 1,2,3"
      intI = 1
      Set adoacc1k0 = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With adoacc1k0
         Erase strMailFailList
         ReDim strMailFailList(0)
         Do While Not .EOF
            strBCC = "" 'Added by Morgan 2016/3/11
            bolEmail = False
            strKeyNo = .Fields("a0y01")
            strInform = "" & adoacc1k0("fa83") '是否寄收據
            strRecDate = .Fields("a0y02") '收款日
            strFNo = "" & .Fields("AgNo")
            If strInform <> "N" Then
               strEMailBox = "" & adoacc1k0("fa16") '電子信箱
               strEmailCC = "" & adoacc1k0.Fields("emailcc") 'Added by Lydia 2024/09/18 財務副本信箱
               If strEMailBox <> "" And UCase(strEMailBox) <> "NO" Then
                  If txtReceiver <> "" Then
                     strEMailBox = txtReceiver
                     strEmailCC = "" 'Added by Lydia 2024/09/18
                  End If
                  bolEmail = True
               End If
            End If
            
            '請款單號
            strNoList = .Fields("a0z02")
            If .Fields("a1k29") <> "Y" Then bolEmail = False '部分收款不發Mail
            'Remove by Lydia 2016/07/01 程序改在每日批次(StrMenu77)
            'If .Fields("a1k13") = "T" And strBcc = "" Then strBcc = GetTReceiver 'Added by Morgan 2016/3/11 T案要BCC給程序
            intI = 1
            .MoveNext
            Do While Not .EOF
               If strKeyNo <> .Fields("a0y01") Then Exit Do
               strNoList = strNoList & ", " & .Fields("a0z02")
               intI = intI + 1
               If .Fields("a1k29") <> "Y" Then bolEmail = False '部分收款不發Mail
               'Remove by Lydia 2016/07/01 程序改在每日批次(StrMenu77)
               'If .Fields("a1k13") = "T" And strBcc = "" Then strBcc = GetTReceiver 'Added by Morgan 2016/3/11 T案要BCC給程序
               .MoveNext
            Loop
            If intI > 1 Then
               '最後一個 ", " 換成 " and "
               strNoList = Left(strNoList, InStrRev(strNoList, ", ") - 1) & " and " & Mid(strNoList, InStrRev(strNoList, ", ") + 2)
            End If
            
            If bolEmail Then
               strRecAmount = ""
               'Modified by Morgan 2016/2/17 若為暫收款要抓原貸方的付款方式
               'Modified by Morgan 2016/9/19 修正
               'Modified by Morgan 2018/10/17 暫收款若沒有收款資料時不要帶 Ex:M10704930,N10700159 --婉莘
               'strExc(0) = "select a.a1p19,a.a1p21,a.a1p23,nvl(a.a1p24,b.a1p24) a1p24,a.a1p05 from acc1p0 a,acc1p0 b where a.a1p04='" & strKeyNo & "' and a.a1p07>0 and b.a1p30(+)=a.a1p30 and b.a1p05(+)=a.a1p05 and b.a1p05(+)='2401' and b.a1p08(+)>0 and nvl(a.a1p24,b.a1p24) is not null"
               'Modified by Morgan 2023/6/30 暫收款金額合併且不再帶日期 --斯閔
               'strExc(0) = "select a.a1p19,a.a1p21,c.a1p23,c.a1p24,a.a1p05,1 srt,a0y02 from acc1p0 a,acc1p0 b,acc1p0 c, acc0y0 where a.a1p04='" & strKeyNo & "' and a.a1p07>0 and a.a1p05='2401' and b.a1p30(+)=a.a1p30 and b.a1p05(+)=a.a1p05 and b.a1p08>0 and c.a1p04(+)=b.a1p04 and c.a1p07>0 and c.a1p24 is not null and a0y01(+)=c.a1p04 and a0y01 is not null" & _
                  " union select a.a1p19,a.a1p21,a.a1p23,a.a1p24,a.a1p05,2 srt,0 from acc1p0 a where a.a1p04='" & strKeyNo & "' and a.a1p07>0 and a.a1p05<>'2401' and a.a1p24 is not null order by srt"
               strExc(0) = "select a.a1p19,sum(a.a1p21) a1p21,'' a1p23,'' a1p24,'' a1p05,1 srt from acc1p0 a where a.a1p04='" & strKeyNo & "' and a.a1p07>0 and a.a1p05='2401' group by a.a1p19" & _
                  " union select a.a1p19,a.a1p21,a.a1p23,a.a1p24,a.a1p05,3 srt from acc1p0 a where a.a1p04='" & strKeyNo & "' and a.a1p07>0 and a.a1p05<>'2401' and a1p05<>'611301' and a.a1p24 is not null" & _
                  " order by srt"
               'end 2016/9/19
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(1) = "" 'Added by Morgan 2023/7/22
                  strExc(2) = "" 'Added by Morgan 2023/7/22
                  strExc(3) = "" 'Added by Morgan 2023/7/4
                  'Added by Morgan 2016/9/19
                  If RsTemp("srt") = 1 Then
                     'Modified by Morgan 2023/7/4
                     'strRecDate = "" & RsTemp("a0y02")
                     strRecDate = ""
                     If RsTemp.RecordCount = 1 Then
                        'Modified by Morgan 2024/7/22
                        'strExc(1) = "credit"
                        strExc(1) = "credit in the amount of " & RsTemp("a1p19") & Format(RsTemp("a1p21"), FDollar)
                        'end 2024/7/22
                     Else
                        strExc(3) = " and included credit in the amount of " & RsTemp("a1p19") & Format(RsTemp("a1p21"), FDollar)
                        RsTemp.MoveNext
                     End If
                     'end 2023/7/4
                  End If
                  'end 2016/9/19
                  
                  dblAmount = RsTemp("a1p21")
                  
                  'Added by Morgan 2016/5/31
                  '台幣(科目 110204)收款時, 手續費 (科目 611301) 的金額與收款金額相加後, 其金額再列示在收據內(收台幣的時候不會有多筆收款的情況)
                  If RsTemp("a1p05") = "110204" Then
                     strExc(0) = "select a1p21 from acc1p0 where a1p04='" & strKeyNo & "' and a1p05='611301' and a1p07>0"
                     intI = 1
                     Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        dblAmount = dblAmount + adoRecordset(0)
                     End If
                  End If
                  'end 2016/5/31
                  
                  If "" & RsTemp("srt") <> 1 Then 'Added by Morgan 2023/7/5
                  
                     'CB
                     If RsTemp("a1p24") = "2" Or RsTemp("a1p24") = "3" Then
                        strExc(1) = "check #" & RsTemp("a1p23")
                        strExc(2) = " in the amount of " & RsTemp("a1p19") & Format(dblAmount, FDollar)
                     'IR
                     Else
                        'Modified by Morgan 2023/7/5
                        'strExc(1) = "remittance"
                        'strExc(2) = " in the amount of " & RsTemp("a1p19") & Format(dblAmount, FDollar)
                        strExc(1) = "your remittance"
                        strExc(2) = " of " & RsTemp("a1p19") & Format(dblAmount, FDollar)
                        'end 2023/7/5
                     End If
                  
                  End If 'Added by Morgan 2023/7/5
                  
                  'strExc(2) = " in the amount of " & RsTemp("a1p19") & Format(dblAmount, FDollar) 'Removed by Morgan 2023/7/5 移到上面
                  
                  RsTemp.MoveNext
                  Do While Not RsTemp.EOF
                     intI = intI + 1
                     If RsTemp("a1p24") = "2" Then
                        strExc(1) = strExc(1) & ", #" & RsTemp("a1p23")
                     End If
                     strExc(2) = strExc(2) & ", " & RsTemp("a1p19") & Format(RsTemp("a1p21"), FDollar)
                     RsTemp.MoveNext
                  Loop
                  If intI > 2 Then
                     If InStrRev(strExc(1), ", ") > 0 Then
                        strExc(1) = Left(strExc(1), InStrRev(strExc(1), ", ") - 1) & Replace(strExc(1), ", ", " and ", InStrRev(strExc(1), ", "))
                     End If
                     strExc(2) = Left(strExc(2), InStrRev(strExc(2), ", ") - 1) & Replace(strExc(2), ", ", " and ", InStrRev(strExc(2), ", "))
                  ElseIf intI = 2 Then
                     strExc(1) = Replace(strExc(1), ", ", " and ")
                     strExc(2) = Replace(strExc(2), ", ", " and ")
                  End If
                  
                  'Modified by Morgan 2023/7/4 +strExc(3)
                  strRecAmount = strExc(1) & strExc(2) & strExc(3)
                  
                  'Modified by Morgan 2016/2/16 主旨+代理人編號
                  'Modified by Morgan 2016/3/11 +BCC
                  'modify by sonia 2016/10/24 婉莘Y52431,Y45848主旨加Ticket # [N-297175]
                  'PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , , True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBcc
                  If strFNo = "Y52431000" Or strFNo = "Y45848000" Then
                     'Modified by Lydia 2024/09/18 財務副本信箱strEmailCC
                     PUB_SendMail strUserNum, strEMailBox, "", "Ticket # [N-297175] Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , , True, True, True, strEmailCC, strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
                  Else
                     'Modified by Lydia 2024/09/18 財務副本信箱strEmailCC
                     PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , , True, True, True, strEmailCC, strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
                  End If
                  'end 2016/10/24
                  'Add by Amy 2020/10/06 有寄mail的產生特殊收據清單
                  If Check1.Value = vbChecked Then
                        StrSQLa = "Insert Into Accrpt2460 (ID,R001) Values ('" & strUserNum & "','" & strFNo & "')"
                        cnnConnection.Execute StrSQLa
                  End If
                  'end 2020/10/06
                  
                  bolMailFailNoAlert = False
                  If bolMailSendOk = False Then
                     If strMailFailList(0) <> "" Then
                        ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                     End If
                     strMailFailList(UBound(strMailFailList)) = strFNo & " : " & strEMailBox
                  End If
               Else
                  If strMailFailList(0) <> "" Then
                     ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                  End If
                  strMailFailList(UBound(strMailFailList)) = strFNo & " : " & strEMailBox & " => " & strKeyNo & "無法確認收款內容"
               End If
            End If
         Loop
         End With
         
         If strMailFailList(0) <> "" Then
            strExc(0) = "E-Mail失敗清單：" & vbCrLf & vbCrLf
            For intI = 0 To UBound(strMailFailList)
               strExc(0) = strExc(0) & strMailFailList(intI) & vbCrLf
            Next
            If MsgBox(strExc(0) & vbCrLf & "是否要列印？" & vbCrLf, vbYesNo + vbDefaultButton1) = vbYes Then
               Printer.Print strExc(0)
               Printer.EndDoc
            End If
         End If
   
      Else
         MsgBox MsgText(28), , MsgText(5)
      End If
      adoacc1k0.Close
      Exit Sub
   End If
   'end 2015/12/29
      
   adoacc1k0.CursorLocation = adUseClient
   'Add by Morgan 2007/3/3 加fa83(cu119)
   'Modify by Morgan 2011/5/25 +fa70
   'Modified by Morgan2012/7/11 +a1k28
   'Modify By Sindy 2013/1/11 +a0y19
   'Modify by Amy 2013/07/03 +國藉判斷
   'Modified by Lydia 2024/09/18 +財務副本信箱emailcc：寄財務信箱一併CC副本
   'strSql = "select * from (select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, fa05, fa63, fa64, fa65, fa32, fa33, fa34, fa35, fa36, fa06, fa23, fa18, fa19, fa20, fa21, fa22, fa04, fa17, fa01, fa02, a1k02, FA31, substr(a1k01, 1, 8) as DocNo, a1k08 as Amount, a0y02,fa10,nvl(fa79,fa16) fa16,fa83,fa70,a1k28,a0y19 from acc1k0, acc0z0, acc0y0, fagent where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 " & strSql & " and length(a1k01) = 9" & _
            " union select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu66 as fa33, cu67 as fa34, cu68 as fa35, cu69 as fa36, cu06 as fa06, cu29 as fa23, cu24 as fa18, cu25 as fa19, cu26 as fa20, cu27 as fa21, cu28 as fa22, cu04 as fa04, NVL(CU31,cu23) as fa17, cu01 as fa01, cu02 as fa02, a1k02, cu64 as FA31, substr(a1k01, 1, 8) as DocNo, a1k08 as Amount, a0y02,cu10 as fa10,nvl(cu115,cu20) as fa16, cu119 as fa83,cu102 as fa70,a1k28,a0y19 from acc1k0, acc0z0, acc0y0, customer where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 " & strSql & " and length(a1k01) = 9" & _
            ") new union select * from (select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, fa05, fa63, fa64, fa65, fa32, fa33, fa34, fa35, fa36, fa06, fa23, fa18, fa19, fa20, fa21, fa22, fa04, fa17, fa01, fa02, a1k02, fa31 from acc1k0, acc0z0, acc0y0, fagent where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 " & strSql & " and substr(a1k01, 9, 2) = '00'" & _
            " union select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu66 as fa33, cu67 as fa34, cu68 as fa35, cu69 as fa36, cu06 as fa06, cu29 as fa23, cu24 as fa18, cu25 as fa19, cu26 as fa20, cu27 as fa21, cu28 as fa22, cu04 as fa04, NVL(CU31,cu23) as fa17, cu01 as fa01, cu02 as fa02, a1k02, cu64 as fa31 from acc1k0, acc0z0, acc0y0, customer where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 " & strSql & " and substr(a1k01, 9, 2) = '00'" & _
            ") new,(select substr(a1k01, 1, 8) as DocNo, sum(a1k08) as Amount,max(a0y02) as a0y02,max(fa10) as fa10,max(nvl(fa79,fa16)) as fa16,max(fa83) fa83,max(fa70) fa70,max(a1k28) a1k28,max(a0y19) a0y19 from acc1k0, acc0z0, acc0y0, fagent where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 " & strSql & " and length(a1k01) = 10 group by substr(a1k01, 1, 8)" & _
            " union select substr(a1k01, 1, 8) as DocNo, sum(a1k08) as Amount,max(a0y02) as a0y02,max(cu10) as fa10,max(nvl(cu115,cu20)) as fa16,max(cu119) fa83,max(cu102) fa70,max(a1k28) a1k28,max(a0y19) a0y19 from acc1k0, acc0z0, acc0y0, customer where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 " & strSql & " and length(a1k01) = 10 group by substr(a1k01, 1, 8)" & _
            ") old where substr(a1k01, 1, 8) = DocNo "
   'strSql = "select * from (" & strSql & ") X " & strNation & "order by a0y01, decode(a0y18, '1', a0y07, '2', a0y08, a0y09) asc, a1k13 asc, a1k14 asc, a1k15 asc, a1k16 asc, a1k01 asc"
   strExc(0) = "select * from (select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, fa05, fa63, fa64, fa65, fa32, fa33, fa34, fa35, fa36, fa06, fa23, fa18, fa19, fa20, fa21, fa22, fa04, fa17, fa01, fa02, a1k02, FA31, substr(a1k01, 1, 8) as DocNo, a1k08 as Amount, a0y02,fa10,nvl(fa79,fa16) fa16,fa83,fa70,a1k28,a0y19,decode(fa105||fa79,null,'',fa134) as emailcc from acc1k0, acc0z0, acc0y0, fagent where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 " & strSql & " and length(a1k01) = 9" & _
            " union select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu66 as fa33, cu67 as fa34, cu68 as fa35, cu69 as fa36, cu06 as fa06, cu29 as fa23, cu24 as fa18, cu25 as fa19, cu26 as fa20, cu27 as fa21, cu28 as fa22, cu04 as fa04, NVL(CU31,cu23) as fa17, cu01 as fa01, cu02 as fa02, a1k02, cu64 as FA31, substr(a1k01, 1, 8) as DocNo, a1k08 as Amount, a0y02,cu10 as fa10,nvl(cu115,cu20) as fa16, cu119 as fa83,cu102 as fa70,a1k28,a0y19,decode(cu115,null,'',cu200) as emailcc from acc1k0, acc0z0, acc0y0, customer where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 " & strSql & " and length(a1k01) = 9" & _
            ") new "
   strExc(0) = strExc(0) & " union select  a0y01,a0y18,a0y07,a0y08,a0y09,a0y10,a1k18,a0y06,a1k13,a1k14,a1k15,a1k16,a1k01,fa05,fa63,fa64,fa65,fa32,fa33,fa34,fa35,fa36 " & _
            ",fa06,fa23,fa18,fa19,fa20,fa21,fa22,fa04,fa17,fa01,fa02,a1k02,fa31,docno,amount,a0y02,fa10,fa16,fa83,fa70,a1k28,a0y19,emailcc " & _
            " from (select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, fa05, fa63, fa64, fa65, fa32, fa33, fa34, fa35, fa36, fa06, fa23, fa18, fa19, fa20, fa21, fa22, fa04, fa17, fa01, fa02, a1k02, fa31,decode(fa105||fa79,null,'',fa134) as emailcc from acc1k0, acc0z0, acc0y0, fagent where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 " & strSql & " and substr(a1k01, 9, 2) = '00'" & _
            " union select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu66 as fa33, cu67 as fa34, cu68 as fa35, cu69 as fa36, cu06 as fa06, cu29 as fa23, cu24 as fa18, cu25 as fa19, cu26 as fa20, cu27 as fa21, cu28 as fa22, cu04 as fa04, NVL(CU31,cu23) as fa17, cu01 as fa01, cu02 as fa02, a1k02, cu64 as fa31,decode(cu115,null,'',cu200) as emailcc from acc1k0, acc0z0, acc0y0, customer where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 " & strSql & " and substr(a1k01, 9, 2) = '00'" & _
            ") new,(select substr(a1k01, 1, 8) as DocNo, sum(a1k08) as Amount,max(a0y02) as a0y02,max(fa10) as fa10,max(nvl(fa79,fa16)) as fa16,max(fa83) fa83,max(fa70) fa70,max(a1k28) a1k28,max(a0y19) a0y19 from acc1k0, acc0z0, acc0y0, fagent where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 " & strSql & " and length(a1k01) = 10 group by substr(a1k01, 1, 8)" & _
            " union select substr(a1k01, 1, 8) as DocNo, sum(a1k08) as Amount,max(a0y02) as a0y02,max(cu10) as fa10,max(nvl(cu115,cu20)) as fa16,max(cu119) fa83,max(cu102) fa70,max(a1k28) a1k28,max(a0y19) a0y19 from acc1k0, acc0z0, acc0y0, customer where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 " & strSql & " and length(a1k01) = 10 group by substr(a1k01, 1, 8)" & _
            ") old where substr(a1k01, 1, 8) = DocNo "
   strSql = "select * from (" & strExc(0) & ") X " & strNation & "order by a0y01, decode(a0y18, '1', a0y07, '2', a0y08, a0y09) asc, a1k13 asc, a1k14 asc, a1k15 asc, a1k16 asc, a1k01 asc"
   'end 2024/09/18
   adoacc1k0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc1k0.RecordCount = 0 Then
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   'Add by Morgan 2008/4/16 存電子檔時不印
   If Text6 = "Y" Then
      strSavePath = PUB_Getdesktop
      bol2Printer = False
   Else
      strSavePath = App.path
      'Modify by Morgan 2010/6/10
      'bol2Printer = True
      'Modified by Morgan 2015/1/21
      'If Text5 = "1" Then
      If Text5 <> "2" Then
         bol2Printer = True
      Else
         bol2Printer = False
      End If
      'end 2010/6/10
   End If
   
   '刪除舊的暫存圖檔
   'Modified by Morgan 2023/11/9 pdf也要刪
   'strExc(1) = App.path & "\$*.jpg"
   strExc(1) = App.path & "\$*.*"
   If Dir(strExc(1)) <> "" Then Kill strExc(1)

'Modify by Morgan 2008/11/20 改複寫為單張列印，不發Mail的另外多印一份
'Modify by Morgan 2010/6/10 要列印且非存電子檔的才要兩次
'For iCopy = 1 To 2

'Modified by Morgan 2014/5/14 電子化不再留卷
'If Text5 = "1" And Text6 <> "Y" Then
'   iRound = 2
'Else
'   iRound = 1
'End If
iRound = 1
'end 2014/5/14
For iCopy = 1 To iRound
'end 2010/6/10
   
   'Added by Morgan 2012/10/11
   'Remove by Morgan 2014/5/14 電子化不再留卷
   'If iCopy = 2 And bolChgPrinter Then
   '   PUB_RestorePrinter cmbPrinter2
   'End If
   
   Erase strMailFailList
   ReDim strMailFailList(0)
   strPicLetter = ""
   strPicFileNames = ""
   strNo = ""
   strDocNo = ""
   douAmount = 0
   intLength = 0
   iPageNo = 0
   strBCC = "" 'Added by Morgan 2016/3/11
   adoacc1k0.MoveFirst
   
   'Add By Sindy 2014/3/10 加控制,收據要增加案件名稱,申請人
   If Text7 = "020" Then
      bolChina = True
   Else
      bolChina = False
   End If
   '2014/3/10 END
   
   Do While adoacc1k0.EOF = False
      
      '定稿語文
      strLangTmp = "" & adoacc1k0("fa31")
      
      '若有下國籍條件 Amy 2014/11/28加
      If Text7 <> MsgText(601) Then
         If adoacc1k0.Fields("fa10").Value < Text7 Then
            GoTo NextSkip
         End If
      End If
      If Text8 <> MsgText(601) Then
         If adoacc1k0.Fields("fa10").Value > Text8 & "z" Then
            GoTo NextSkip
         End If
      End If
      '舊系統請款單
      If Len(adoacc1k0.Fields("a1k01").Value) = 10 Then
         If strDocNo = Mid(adoacc1k0.Fields("a1k01").Value, 1, 8) Then
            GoTo NextSkip
         Else
            strDocNo = Mid(adoacc1k0.Fields("a1k01").Value, 1, 8)
         End If
      End If
      
      '收款單號
      If IsNull(adoacc1k0.Fields("a0y01").Value) Then
         strKeyNo = ""
      Else
         strKeyNo = adoacc1k0.Fields("a0y01").Value
      End If
      
      If strNo <> strKeyNo Then
         If douAmount <> 0 Then
            intCounter = intCounter + 1
            PrintSum
            douAmount = 0
            If bol2Printer = True Then
               Printer.NewPage
               'Added by Morgan 2016/4/12
               'T案印紙本的也要通知程序
               'Remove by Lydia 2016/07/01 程序改在每日批次(StrMenu77)
               'If strBcc <> "" Then
               '   PUB_SendMail strUserNum, strBcc, "", "FC收據列印通知 (" & strFNo & ")", "收款單號：" & strNo & vbCrLf & vbCrLf & "請款單號：" & strNoList
               'End If
               ''end 2016/4/12
            End If
            If bol2File = True Then
               PicNewPage
               
               'Added by Morgan 2023/11/9 jpg轉pdf
               strExc(1) = strSavePath & "\$" & strNo & ".pdf"
               'Modified by Morgan 2024/9/4 中文要加印收據章
               If PUB_JPG2PDF(strPicFileNames, strExc(1), , IIf(strLanguage = "1", "16", "")) = True Then
                  strPicFileNames = strExc(1)
               End If
               'end 2023/11/9
               
               '當附件
               'Modify by Morgan 2010/6/10 選發EMail且未設要存電子檔才寄送
               'If bolEmail = True And bol2Printer = True Then
               If bolEmail = True And Text5.Text = "2" And Text6.Text <> "Y" Then
                  bolMailFailNoAlert = True
                  bolMailSendOk = False
                  
                  'Modify by Morgan 2011/4/22 改以ipdept@taie.com.tw 寄但回覆還是給寄件人(70004)
                  'PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement", GetMailContent, , strPicFileNames, True, True, True
                  'Modified by Morgan 2011/10/12 改用 account@taie.com.tw 寄
                  'PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement", GetMailContent, , strPicFileNames, True, True, True, , "ipdept@taie.com.tw", "TAI E INTERNATIONAL PATENT & LAW OFFICE", strUserNum
                  'Modified by Morgan 2014/8/27 改回覆到財務信箱 -- 婧瑄
                  'PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strUserNum
                  'Modified by Morgan 2016/2/16 主旨+代理人編號
                  'Modified by Morgan 2016/3/11 +BCC
                  'modify by sonia 2016/10/24 婉莘Y52431,Y45848主旨加Ticket # [N-297175]
                  'PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBcc
                  If strFNo = "Y52431000" Or strFNo = "Y45848000" Then
                     PUB_SendMail strUserNum, strEMailBox, "", "Ticket # [N-297175] Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
                  Else
                     PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
                  End If
                  'end 2016/10/24
                  'Add by Amy 2020/10/06 有寄mail的產生特殊收據清單
                  If Check1.Value = vbChecked Then
                        StrSQLa = "Insert Into Accrpt2460 (ID,R001) Values ('" & strUserNum & "','" & strFNo & "')"
                        cnnConnection.Execute StrSQLa
                  End If
                  'end 2020/10/06
                  
                  bolMailFailNoAlert = False
                  If bolMailSendOk = False Then
                     If strMailFailList(0) <> "" Then
                        ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                     End If
                     strMailFailList(UBound(strMailFailList)) = strFNo & " : " & strEMailBox
                  End If
                  strPicFileNames = ""
               End If
               
               strPicFileNames = "" 'Added by Morgan 2023/11/9
               
            End If
            iPageNo = 0
         End If
         
         strBCC = "" 'Added by Morgan 2016/3/11
         strNoList = "" 'Added by Morgan 2016/4/12
         
         intCounter = 5
         
         '是否寄收據
         strInform = "" & adoacc1k0("fa83")
         '電子信箱
         strEMailBox = "" & adoacc1k0("fa16")
         strEmailCC = "" & adoacc1k0.Fields("emailcc") 'Added by Lydia 2024/09/18 財務副本信箱
         'Modify by Morgan 2007/3/3
         '檢查非大陸案有Email的代理人
         'Modify by Morgan 2008/10/30 大陸的也要了
         'If "" & adoacc1k0("fa10") <> "020" And strEMailBox <> "" And UCase(strEMailBox) <> "NO" Then
         If strEMailBox <> "" And UCase(strEMailBox) <> "NO" Then
            'Remove by Morgan 2006/12/21 開始使用
            'strEMailBox = "jasjaswu@gmail.com" 'Add by Morgan 2006/12/14 測試用
            
            'Add by Morgan 2009/2/23 測試用
            If txtReceiver <> "" Then
               strEMailBox = txtReceiver
               strEmailCC = "" 'Added by Lydia 2024/09/18
            End If
            
            bolEmail = True
            bol2File = True
            If "" & adoacc1k0("fa10") = "020" Then
               bolChinese = True
            Else
               bolChinese = False
            End If
         Else
            bolEmail = False
            bol2File = False
         End If
         
         'Add by Morgan 2008/4/16 放在迴圈內是因為要印是否"EMail通知","不寄收據"的註記
         '產生電子檔
         If Text6.Text = "Y" Then
            bol2File = True
         '不發Mail 或 不寄收據 時只列印不產生電子檔
         'Modify by Morgan 2010/6/10
         'ElseIf Text5.Text = "N" Or strInform = "N" Then
         'Modified by Morgan 2015/1/21
         'ElseIf Text5.Text = "1" Or strInform = "N" Then
         ElseIf Text5.Text <> "2" Or strInform = "N" Then
         'end 2015/1/21
         'end 2010/6/10
            bol2File = False
         End If
         
         'Add by Morgan 2008/11/20
         'Modified by Morgan 2014/5/19 電子化不再留卷
         'If iCopy = 2 Then
         '   '不寄收據 或 EMail 或 存電子檔的不要印第二份
         '   If strInform = "N" Or bolEmail = True Or bol2File = True Then
         '      bol2Printer = False
         '   Else
         '      bol2Printer = True
         '   End If
         '   bol2File = False
         '   bolEmail = False
         'Added by Morgan 2012/10/11
         'ElseIf Not (strInform = "N" Or bolEmail = True Or bol2File = True) Then
         '   bolChgPrinter = True
         
         If Text5.Text = "1" Then
            If bol2File = True Then
               bol2Printer = False
            'Added by Morgan 2015/6/12 +E+印的也要印
            ElseIf strInform = "B" Then
               bol2Printer = True
            'end 2015/6/12
            ElseIf (strInform = "N" Or bolEmail = True) Then
               bol2Printer = False
            Else
               bol2Printer = True
            End If
         'end 2014/5/19
         
         'Added by Morgan 2015/1/21
         '強制列印
         ElseIf Text5.Text = "3" Then
            If (strInform = "N" Or bol2File = True) Then
               bol2Printer = False
            Else
               bol2Printer = True
            End If
         'end 2015/1/21
         
         End If
          
         '是否存電子檔
         If bol2File = True Then
            If strPicLetter = "" Then
               strPicLetter = App.path & "\$Tmp.jpg"
               'Modified by Morgan 2020/3/30
               'If PUB_ReadDB2File(strPicLetter, 6) = True Then
               If strSrvDate(1) >= 智慧所更名日 Then
                  PUB_GetLetterPicID "2", , iPicNo, , , , , True
               Else
                  iPicNo = 6
               End If
               If PUB_ReadDB2File(strPicLetter, iPicNo) = True Then
               'end 2020/3/30
                  Set Picture1.Picture = LoadPicture(strPicLetter)
                  Picture1.AutoSize = True
                  douExtRate = Picture1.Height / 16836 'Add by Morgan 2011/8/10
               End If
            End If
         End If
         
         PrintHead
         
         strA0Y10 = ""
         If IsNull(adoacc1k0.Fields("a0y10").Value) = False Then
            If bol2Printer = True Then
               Printer.CurrentX = intLeft
               Printer.CurrentY = intTop + 5100 + intCounter * 300
               Printer.Print "CREDIT"
'               Printer.CurrentX = intLeft + 7400
'               Printer.CurrentY = intTop + 5100 + intCounter * 300
               'Add By Sindy 2014/3/12
               If bolChina = True Then
                  Printer.CurrentX = intLeft + 5400
               Else
               '2014/3/12 END
                  Printer.CurrentX = intLeft + 7400
               End If
               Printer.CurrentY = intTop + 5100 + intCounter * 300
               Printer.Print "" & adoacc1k0.Fields("a1k18").Value
            End If
            If bol2File = True Then
               Picture1.CurrentX = (intLeft) * douExtRate
               Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
               Picture1.Print "CREDIT"
'               Picture1.CurrentX = (intLeft + 7400) * douExtRate
'               Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
               'Add By Sindy 2014/3/12
               If bolChina = True Then
                  Picture1.CurrentX = (intLeft + 5400) * douExtRate
               Else
               '2014/3/12 END
                  Picture1.CurrentX = (intLeft + 7400) * douExtRate
               End If
               Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
               Picture1.Print "" & adoacc1k0.Fields("a1k18").Value
            End If
            
            If IsNull(adoacc1k0.Fields("a0y06").Value) = False Then
               strAmount = Format(Val(adoacc1k0.Fields("a0y06").Value), FDollar)
               If bol2Printer = True Then
'                  intLength = Printer.TextWidth(strAmount)
'                  Printer.CurrentX = intLeft + 9900 - intLength
'                  Printer.CurrentY = intTop + 5100 + intCounter * 300
                  intLength = Printer.TextWidth(strAmount)
                  'Add By Sindy 2014/3/12
                  If bolChina = True Then
                     Printer.CurrentX = intLeft + 7100 - intLength
                  Else
                  '2014/3/12 END
                     Printer.CurrentX = intLeft + 9900 - intLength
                  End If
                  Printer.CurrentY = intTop + 5100 + intCounter * 300
                  Printer.Print strAmount
               End If
               If bol2File = True Then
'                  intLength = Picture1.TextWidth(strAmount)
'                  Picture1.CurrentX = (intLeft + 9900) * douExtRate - intLength
'                  Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
                  intLength = Picture1.TextWidth(strAmount)
                  'Add By Sindy 2014/3/12
                  If bolChina = True Then
                     Picture1.CurrentX = (intLeft + 7100) * douExtRate - intLength
                  Else
                  '2014/3/12 END
                     Picture1.CurrentX = (intLeft + 9900) * douExtRate - intLength
                  End If
                  Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
                  Picture1.Print strAmount
               End If
               douAmount = douAmount + Val(adoacc1k0.Fields("a0y06").Value)
            End If
            intCounter = intCounter + 1
         End If
         strNo = strKeyNo
         strRecDate = "" & adoacc1k0("a0y02")
         strA1K28 = "" & adoacc1k0("a1k28")
      End If
      'Remove by Lydia 2016/07/01 程序改在每日批次(StrMenu77)
      'If adoacc1k0.Fields("a1k13") = "T" And strBcc = "" Then strBcc = GetTReceiver 'Added by Morgan 2016/3/11 T案要BCC給程序
      strNoList = strNoList & adoacc1k0.Fields("a1k01").Value & " (" & adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & IIf(adoacc1k0.Fields("a1k15") & adoacc1k0.Fields("a1k16") = "000", "", "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value & ")") & ")" & vbCrLf & String(5, "　") 'Added by Morgan 2016/4/12
      
'      'Modify by Sindy 2013/1/11
'      If Not IsNull(adoacc1k0.Fields("a0y19").Value) Then
'         strFNo = adoacc1k0.Fields("a0y19").Value
'      Else
'      '2013/1/11 End
         Select Case adoacc1k0.Fields("a0y18").Value
            Case "1"
               If IsNull(adoacc1k0.Fields("a0y07").Value) Then
                  strFNo = ""
               Else
                  strFNo = adoacc1k0.Fields("a0y07").Value
               End If
            Case "2"
               If IsNull(adoacc1k0.Fields("a0y08").Value) Then
                  strFNo = ""
               Else
                  strFNo = adoacc1k0.Fields("a0y08").Value
               End If
            Case "3"
               If IsNull(adoacc1k0.Fields("a0y09").Value) Then
                  strFNo = ""
               Else
                  strFNo = adoacc1k0.Fields("a0y09").Value
               End If
            Case Else
               If IsNull(adoacc1k0.Fields("a0y09").Value) Then
                  strFNo = ""
               Else
                  strFNo = adoacc1k0.Fields("a0y09").Value
               End If
         End Select
'      End If
      
      'Modify by Morgan 2009/8/26 修正跳頁控制
      bNewPage = False
      If bol2Printer = True Then
         If Printer.CurrentY + 1400 > Printer.ScaleHeight Then
            Printer.NewPage
            bNewPage = True
         End If
      End If
      If bol2File = True Then
         If Picture1.CurrentY + 1400 * douExtRate > Picture1.ScaleHeight Then
            PicNewPage
            bNewPage = True
         End If
      End If
      
      If bNewPage = True Then
         intCounter = -4
         PrintHead False
      End If
      'end 2009/8/26
      
      adoquery.CursorLocation = adUseClient
      If Len(adoacc1k0.Fields("a1k01").Value) = 10 Then
         adoquery.Open "select sum(a0z04) from acc0z0, acc0y0 where a0z01 = a0y01 and substr(a0z02, 1, 8) = '" & Mid(adoacc1k0.Fields("a1k01").Value, 1, 8) & "' and A0y01='" & strKeyNo & "' " & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
      Else
         adoquery.Open "select sum(a0z04) from acc0z0, acc0y0 where a0z01 = a0y01 and a0z02 = '" & adoacc1k0.Fields("a1k01").Value & "' and A0y01='" & strKeyNo & "' " & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
      End If
      
      If adoquery.RecordCount <> 0 Then
         If adoquery.Fields(0).Value <> 0 Then
            If Not IsNull(adoacc1k0.Fields("A0Y10").Value) Then
               strA0Y10 = adoacc1k0.Fields("A0Y10").Value
            End If
            
            '我方文號
            If bol2Printer = True Then
               Printer.CurrentX = intLeft
               Printer.CurrentY = intTop + 5100 + intCounter * 300
            End If
            If bol2File = True Then
               Picture1.CurrentX = (intLeft) * douExtRate
               Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
            End If
            If adoacc1k0.Fields("a1k15").Value = "0" And adoacc1k0.Fields("a1k16").Value = "00" Then
               If bol2Printer = True Then
                  Printer.Print adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value
               End If
               If bol2File = True Then
                  Picture1.Print adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value
               End If
            Else
               If bol2Printer = True Then
                  Printer.Print adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value
               End If
               If bol2File = True Then
                  Picture1.Print adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value
               End If
            End If
                        
            '貴方文號
            
            'Add by Amy 2016/03/31 +巨京沒彼所案號抓分所案號
'            adocheck.CursorLocation = adUseClient
'            adocheck.Open "select " & strField(0) & " as pa77 from patent where pa01 = '" & adoacc1k0.Fields("a1k13").Value & "' and pa02 = '" & adoacc1k0.Fields("a1k14").Value & "' and pa03 = '" & adoacc1k0.Fields("a1k15").Value & "' and pa04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
'                          "select " & strField(1) & " as pa77 from trademark where tm01 = '" & adoacc1k0.Fields("a1k13").Value & "' and tm02 = '" & adoacc1k0.Fields("a1k14").Value & "' and tm03 = '" & adoacc1k0.Fields("a1k15").Value & "' and tm04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
'                          "select " & strField(2) & " as pa77 from lawcase where lc01 = '" & adoacc1k0.Fields("a1k13").Value & "' and lc02 = '" & adoacc1k0.Fields("a1k14").Value & "' and lc03 = '" & adoacc1k0.Fields("a1k15").Value & "' and lc04 = '" & adoacc1k0.Fields("a1k16").Value & "' union " & _
'                          "select " & strField(3) & " as pa77 from servicepractice where sp01 = '" & adoacc1k0.Fields("a1k13").Value & "' and sp02 = '" & adoacc1k0.Fields("a1k14").Value & "' and sp03 = '" & adoacc1k0.Fields("a1k15").Value & "' and sp04 = '" & adoacc1k0.Fields("a1k16").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
            
            
            strPA77 = GetYourRefNo1(adoacc1k0.Fields("a1k13").Value, adoacc1k0.Fields("a1k14").Value, adoacc1k0.Fields("a1k15").Value, adoacc1k0.Fields("a1k16").Value, _
                            IIf(Left(strFNo, 6) = "Y52269", True, False))
'            If adocheck.RecordCount <> 0 Then
             'Added by Morgan 2023/10/16
             If strPA77 = MsgText(601) Then
               strPA77 = GetYourRefNo2(adoacc1k0.Fields("a1k13").Value, adoacc1k0.Fields("a1k14").Value, adoacc1k0.Fields("a1k15").Value, adoacc1k0.Fields("a1k16").Value)
             End If
             'end 2023/10/16
             If strPA77 <> MsgText(601) Then
               If bol2Printer = True Then
                  'Add By Sindy 2014/3/12
                  If bolChina = True Then
                     Printer.CurrentX = intLeft + 1400
                  Else
                  '2014/3/12 END
                     Printer.CurrentX = intLeft + 2500
                  End If
                  Printer.CurrentY = intTop + 5100 + intCounter * 300
                  'Printer.Print "" & Left(adocheck.Fields("pa77").Value, 18) 'Modify by Amy 2014/07/11 只取18個字-婉莘
                  Printer.Print "" & Left(strPA77, 18)
               End If
               
               If bol2File = True Then
                  'Add By Sindy 2014/3/12
                  If bolChina = True Then
                     Picture1.CurrentX = (intLeft + 1400) * douExtRate
                  Else
                  '2014/3/12 END
                     Picture1.CurrentX = (intLeft + 2500) * douExtRate
                  End If
                  Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
                  'Picture1.Print "" & Left(adocheck.Fields("pa77").Value, 18) 'Modify by Amy 2014/07/11 只取18個字-婉莘
                  Picture1.Print "" & Left(strPA77, 18)
               End If
            End If
'            adocheck.Close
            'end 2016/03/31
            
            '帳單編號
            If bol2Printer = True Then
               'Add By Sindy 2014/3/12
               If bolChina = True Then
                  Printer.CurrentX = intLeft + 4100
               Else
               '2014/3/12 END
                  Printer.CurrentX = intLeft + 5500
               End If
               Printer.CurrentY = intTop + 5100 + intCounter * 300
            End If
            If bol2File = True Then
               'Add By Sindy 2014/3/12
               If bolChina = True Then
                  Picture1.CurrentX = (intLeft + 4100) * douExtRate
               Else
               '2014/3/12 END
                  Picture1.CurrentX = (intLeft + 5500) * douExtRate
               End If
               Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
            End If
            If Not IsNull(adoacc1k0.Fields("a1k01").Value) Then
               If Len(adoacc1k0.Fields("a1k01").Value) = 10 Then
                  If bol2Printer = True Then
                     Printer.Print Mid(adoacc1k0.Fields("a1k01").Value, 3, 6)
                  End If
                  If bol2File = True Then
                     Picture1.Print Mid(adoacc1k0.Fields("a1k01").Value, 3, 6)
                  End If
               Else
                  If bol2Printer = True Then
                     Printer.Print adoacc1k0.Fields("a1k01").Value
                  End If
                  If bol2File = True Then
                     Picture1.Print adoacc1k0.Fields("a1k01").Value
                  End If
               End If
            End If
            
            '幣別
            adocheck.CursorLocation = adUseClient
            adocheck.Open "select distinct a0y03 from acc0z0, acc0y0 where a0z01 = a0y01 and a0z02 = '" & adoacc1k0.Fields("a1k01").Value & "' and A0y01='" & strKeyNo & "' ", adoTaie, adOpenStatic, adLockReadOnly
            If adocheck.RecordCount <> 0 Then
               If IsNull(adocheck.Fields("a0y03").Value) Then
                  strCurrency = "USD"
               Else
                  strCurrency = adocheck.Fields("a0y03").Value
               End If
            Else
               strCurrency = "USD"
            End If
            adocheck.Close
            If bol2Printer = True Then
               'Add By Sindy 2014/3/12
               If bolChina = True Then
                  Printer.CurrentX = intLeft + 5400
               Else
               '2014/3/12 END
                  Printer.CurrentX = intLeft + 7400
               End If
               Printer.CurrentY = intTop + 5100 + intCounter * 300
               Printer.Print strCurrency
            End If
            If bol2File = True Then
               'Add By Sindy 2014/3/12
               If bolChina = True Then
                  Picture1.CurrentX = (intLeft + 5400) * douExtRate
               Else
               '2014/3/12 END
                  Picture1.CurrentX = (intLeft + 7400) * douExtRate
               End If
               Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
               Picture1.Print strCurrency
            End If
            
            '金額
            If IsNull(adoquery.Fields(0).Value) = False Then
               strAmount = Format(Val(adoquery.Fields(0).Value), FDollar)
               
               If bol2Printer = True Then
                  intLength = Printer.TextWidth(strAmount)
                  'Add By Sindy 2014/3/12
                  If bolChina = True Then
                     Printer.CurrentX = intLeft + 7100 - intLength
                  Else
                  '2014/3/12 END
                     Printer.CurrentX = intLeft + 9900 - intLength
                  End If
                  Printer.CurrentY = intTop + 5100 + intCounter * 300
                  Printer.Print strAmount
               End If
               If bol2File = True Then
                  intLength = Picture1.TextWidth(strAmount)
                  'Add By Sindy 2014/3/12
                  If bolChina = True Then
                     Picture1.CurrentX = (intLeft + 7100) * douExtRate - intLength
                  Else
                  '2014/3/12 END
                     Picture1.CurrentX = (intLeft + 9900) * douExtRate - intLength
                  End If
                  Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
                  Picture1.Print strAmount
               End If
               douAmount = douAmount + Val(adoquery.Fields(0).Value)
            End If
            
            'Add By Sindy 2014/3/12
            If bolChina = True Then
               '案件名稱
               strData = GetPrjName("" & adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value)
               If strData <> "" Then
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft + 7300
                     Printer.CurrentY = intTop + 5100 + intCounter * 300
                     Printer.Print convForm(CheckStr(strData), 16)
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft + 7300) * douExtRate
                     Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
                     Picture1.Print convForm(CheckStr(strData), 16)
                  End If
               End If
               '申請人
               strData = GetPrjPeopleNum1("" & adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value)
               If strData <> "" Then
                  strData = GetPrjPeople1(strData)
                  If bol2Printer = True Then
                     Printer.CurrentX = intLeft + 9400
                     Printer.CurrentY = intTop + 5100 + intCounter * 300
                     Printer.Print convForm(CheckStr(strData), 12)
                  End If
                  If bol2File = True Then
                     Picture1.CurrentX = (intLeft + 9400) * douExtRate
                     Picture1.CurrentY = (intTop + 5100 + intCounter * 300) * douExtRate
                     Picture1.Print convForm(CheckStr(strData), 12)
                  End If
               End If
            End If
            '2014/3/12 END
            
            intCounter = intCounter + 1
         End If
      End If
      adoquery.Close
NextSkip:
      adoacc1k0.MoveNext
   Loop
   intCounter = intCounter + 1
'   If adoacc1k0.RecordCount <> 0 Then
'      adoacc1k0.MoveLast
'   End If
   PrintSum
   
   If bol2Printer = True Then
      Printer.EndDoc
      'Added by Morgan 2016/4/12
      'T案印紙本的也要通知程序
      'Remove by Lydia 2016/07/01 程序改在每日批次(StrMenu77)
      'If strBcc <> "" Then
      '   PUB_SendMail strUserNum, strBcc, "", "FC收據列印通知 (" & strFNo & ")", "收款單號：" & strNo & vbCrLf & vbCrLf & "請款單號：" & strNoList
      'End If
      ''end 2016/4/12
   End If
   
   If bol2File = True Then
      PripagenoNo
      strExc(1) = strSavePath & "\$" & strNo & IIf(iPageNo > 0, "_" & Format(iPageNo, "00"), "") & ".jpg"
      PUB_SavePic Picture1, strExc(1)
      strPicFileNames = strPicFileNames & strExc(1) & "*"
      
      'Added by Morgan 2023/11/9 jpg轉pdf
      strExc(0) = strSavePath & "\$" & strNo & ".pdf"
      'Modified by Morgan 2024/9/4 中文要加印收據章
      If PUB_JPG2PDF(strPicFileNames, strExc(0), , IIf(strLanguage = "1", "16", "")) = True Then
         strPicFileNames = strExc(0)
      End If
      'end 2023/11/9
         
      '當附件
      'Modify by Morgan 2010/6/10 選發EMail且未設要存電子檔才寄送
      'If bolEmail = True And bol2Printer = True Then
      If bolEmail = True And Text5.Text = "2" And Text6.Text <> "Y" Then
         bolMailFailNoAlert = True
         bolMailSendOk = False
         
         'Modify by Morgan 2011/4/22 改以ipdept@taie.com.tw 寄但回覆還是給寄件人(70004)
         'PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement", GetMailContent, , strPicFileNames, True, True, True
         'Modified by Morgan 2011/10/12 改用 account@taie.com.tw 寄
         'PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement", GetMailContent, , strPicFileNames, True, True, True, , "ipdept@taie.com.tw", "TAI E INTERNATIONAL PATENT & LAW OFFICE", strUserNum
         'Modified by Morgan 2014/8/27 改回覆到財務信箱 -- 婧瑄
         'PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strUserNum
         'Modified by Morgan 2016/2/16 主旨+代理人編號
         'Modified by Morgan 2016/3/11 +BCC
         'modify by sonia 2016/10/24 婉莘Y52431,Y45848主旨加Ticket # [N-297175]
         'PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBcc
         If strFNo = "Y52431000" Or strFNo = "Y45848000" Then
            PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
         Else
            PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
         End If
         'end 2016/10/24
         'Add by Amy 2020/10/06 有寄mail的產生特殊收據清單
         If Check1.Value = vbChecked Then
                StrSQLa = "Insert Into Accrpt2460 (ID,R001) Values ('" & strUserNum & "','" & strFNo & "')"
                cnnConnection.Execute StrSQLa
         End If
         'end 2020/10/06
         
         bolMailFailNoAlert = False
         If bolMailSendOk = False Then
            If strMailFailList(0) <> "" Then
               ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
            End If
            strMailFailList(UBound(strMailFailList)) = strFNo & " : " & strEMailBox
         End If
      Else
         If strPicFileNames <> "" Then
            MsgBox "電子檔已存桌面！"
         End If
      End If
      'Modified by Morgan 2023/11/9 pdf也要刪
      'strExc(1) = App.path & "\$*.jpg"
      strExc(1) = App.path & "\$*.*"
      If Dir(strExc(1)) <> "" Then Kill strExc(1)
   End If
   
   'Add by Morgan 2007/1/24
   If strMailFailList(0) <> "" Then
      strExc(0) = "E-Mail失敗清單：" & vbCrLf & vbCrLf
      For intI = 0 To UBound(strMailFailList)
         strExc(0) = strExc(0) & strMailFailList(intI) & vbCrLf
      Next
      If MsgBox(strExc(0) & vbCrLf & "是否要列印？" & vbCrLf, vbYesNo + vbDefaultButton1) = vbYes Then
         Printer.Print strExc(0)
         Printer.EndDoc
      End If
   End If
   'end 2007/1/24
Next

   adoacc1k0.Close
End Sub

Private Sub PripagenoNo()
   Picture1.CurrentX = Picture1.ScaleWidth / 2 - 50
   Picture1.CurrentY = Picture1.ScaleHeight - 500
   Picture1.Print Format(iPageNo, "#")
End Sub
'Modify by Morgan 2008/3/11 調整內容--婧瑄
Private Function GetMailContent() As String
   Dim StrMailContent As String
      
   If bolChinese = True Then
                                StrMailContent = "敬啟者："
      StrMailContent = StrMailContent & vbCrLf
      StrMailContent = StrMailContent & vbCrLf & "　　貴公司" & TranslateKeyWord(incCNV_CHINESE_CUN1, DBDATE(strRecDate), "") & "付來的款項 " & strRecAmount & " 已收悉，現附上此款項所支付的本所帳款明細，請核對。"
      StrMailContent = StrMailContent & vbCrLf
      StrMailContent = StrMailContent & vbCrLf & "　　非常感謝您的付款，若有疑問請隨時與本所聯繫。"
      StrMailContent = StrMailContent & vbCrLf
      'Modified by Morgan 2020/3/30 事務所名稱改用函數抓
      'StrMailContent = StrMailContent & vbCrLf & "台一國際專利法律事務所"
      StrMailContent = StrMailContent & vbCrLf & CompNameQuery("2")
      'end 2020/3/30
      StrMailContent = StrMailContent & vbCrLf & "財務處"
      StrMailContent = StrMailContent & vbCrLf
      'Modified by Morgan 2024/6/14 taie@seed.net.com-->ipdept@taie.com.tw
      StrMailContent = StrMailContent & vbCrLf & "PS.此信箱只限於與財務處有關事項聯絡時使用，如與案件有關事項的聯繫請使用 ipdept@taie.com.tw 信箱"
      StrMailContent = StrMailContent & vbCrLf
      
   Else
   
'Modified by Morgan 2015/12/29
'      StrMailContent = "Dear Sirs," & vbCrLf
'      '2010/3/12 MODIFY BY SONIA 取消 received on  ChgEngDate(DBDATE(strRecDate))
'      'StrMailContent = StrMailContent & vbCrLf & "This is to acknowledge safe receipt of your remittance in the amount of " & strRecAmount & " received on " & ChgEngDate(DBDATE(strRecDate)) & " "
'      StrMailContent = StrMailContent & vbCrLf & "This is to acknowledge safe receipt of your remittance in the amount of " & strRecAmount & " "
'      '2010/3/12 END
'      StrMailContent = StrMailContent & vbCrLf & "for our debit notes as the attachment." & vbCrLf
'      StrMailContent = StrMailContent & vbCrLf & "Please be advised that this e-mail address be reserved for account matters only."
'      StrMailContent = StrMailContent & vbCrLf & "Please direct all case matter to ipdept@taie.com.tw to ensure the quickest reply." & vbCrLf
'      StrMailContent = StrMailContent & vbCrLf & "Thank you for your remittance and do not hesitate to contact us regarding accounts matters." & vbCrLf
            
      StrMailContent = "Attn: Accounting Dept." & vbCrLf
      'Modified by Morgan 2024/4/10 對外統一用 Dear Colleagues --林總
      'StrMailContent = StrMailContent & vbCrLf & "Dear Sirs," & vbCrLf
      StrMailContent = StrMailContent & vbCrLf & "Dear Colleagues," & vbCrLf
      'end 2024/4/10
      'Modified by Morgan 2023/7/4
      'StrMailContent = StrMailContent & vbCrLf & "We confirm with thanks safe receipt of " & strRecAmount & " on " & ChgEngDate(DBDATE(strRecDate)) & " for the payment of invoice No." & strNoList & "." & vbCrLf
      StrMailContent = StrMailContent & vbCrLf & "We confirm with thanks safe receipt of " & strRecAmount & IIf(strRecDate <> "", " on " & ChgEngDate(DBDATE(strRecDate)), "") & " for the payment of invoice No." & strNoList & "." & vbCrLf
      'end 2023/7/4
      StrMailContent = StrMailContent & vbCrLf & "Please feel free to contact us, if we can be of any further assistance." & vbCrLf
      StrMailContent = StrMailContent & vbCrLf & "Best regards," & vbCrLf
'end 2015/12/29
      StrMailContent = StrMailContent & vbCrLf & "TAI E INTERNATIONAL PATENT & LAW OFFICE"
      StrMailContent = StrMailContent & vbCrLf & "Accounting department" & vbCrLf
       
   End If
   GetMailContent = StrMailContent & vbCrLf & vbCrLf & vbCrLf & "&nbsp;" 'Modified by Morgan 2015/12/31 +最後加&nbsp;以忽略轉換，否則會發生無法自動折行情形
End Function

Private Sub PicNewPage()
   PripagenoNo
   strExc(1) = strSavePath & "\$" & strNo & IIf(iPageNo > 0, "_" & Format(iPageNo, "00"), "") & ".jpg"
   PUB_SavePic Picture1, strExc(1)
   strPicFileNames = strPicFileNames & strExc(1) & "*"
   Set Picture1.Picture = LoadPicture(strPicLetter)
   Picture1.AutoSize = True
   douExtRate = Picture1.Height / 16836 'Add by Morgan 2011/8/10
End Sub

'Add by Morgan 2007/1/24
Private Sub Text5_GotFocus()
   CloseIme
   TextInverse Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2015/1/21 +3
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
      KeyAscii = 0
   End If
End Sub

'Add by Amy 2020/12/15 輸出選項="2"(發email)且有輸收款日期,產生特殊收據清單預設勾選-莘
Private Sub Text5_Validate(Cancel As Boolean)
    Check1.Value = 0
    If Text5 = "2" Then
        If (MaskEdBox1.Text <> "___/__/__" And MaskEdBox1.Text <> MsgText(601)) _
          Or (MaskEdBox2.Text <> "___/__/__" And MaskEdBox2.Text <> MsgText(601)) Then
          Check1.Value = 1
        End If
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
'Added by Morgan 2012/7/11
Private Sub PrintCoverPage()
   Dim stSQL As String, intR As Integer, intRow As Integer
   Dim adoTmp As ADODB.Recordset
   'Add by Amy 2018/10/31
   Dim strFA18 As String, strFA19 As String, strFA20 As String, strFA21 As String, strFA22 As String, strFA70 As String
   Dim strFA32 As String, strFA33 As String, strFA34 As String, strFA35 As String, strFA36 As String
   
   If bol2Printer = False Then Exit Sub
   If strA1K28 <> "Y51371000" Then Exit Sub
   'If iCopy <> 2 Then Exit Sub 'Remove by Morgan 2014/6/30 電子化不再留卷
   
   stSQL = "select * from fagent where fa01='" & Left(strA1K28, 8) & "' and fa02='" & Mid(strA1K28, 9) & "'"
   intR = 1
   Set adoTmp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With adoTmp
      
      Printer.NewPage
      Printer.Font = "Times New Roman"
      Printer.FontSize = 12
      
      intRow = 6
      '收款日期
      Printer.CurrentX = intLeft + 8500
      Printer.CurrentY = intTop + 1400 + intRow * 250
      If Me.MaskEdBox1.Text = "___/__/__" Then
         Printer.Print Format(AFDate(strSrvDate(1)), "mmm. d, yyyy")
      Else
         Printer.Print Format(AFDate(ChangeTStringToWString(Replace(Me.MaskEdBox1.Text, "/", ""))), "mmm. d, yyyy")
      End If
      
      '代理人名稱
      intRow = intRow + 1
      Printer.CurrentX = intLeft
      Printer.CurrentY = intTop + 1400 + intRow * 250
      Printer.Print Trim(.Fields("fa05") & " " & .Fields("fa63"))
      If Not IsNull(.Fields("fa64")) Then
         intRow = intRow + 1
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print Trim(.Fields("fa64") & " " & .Fields("fa65"))
      End If
      
      '代理人地址
      'Add by Amy 2018/10/31 +地址有「竹曆退件」字樣不顯示地址
      strFA18 = "" & .Fields("fa18").Value: strFA19 = "" & .Fields("fa19").Value: strFA20 = "" & .Fields("fa20").Value
      strFA21 = "" & .Fields("fa21").Value: strFA22 = "" & .Fields("fa22").Value: strFA70 = "" & .Fields("fa70").Value
     
      strFA32 = "" & .Fields("fa32").Value: strFA33 = "" & .Fields("fa33").Value: strFA34 = "" & .Fields("fa34").Value
      strFA35 = "" & .Fields("fa35").Value: strFA36 = "" & .Fields("fa36").Value
    
      If InStr(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70, "竹曆退件") > 0 Then
        strFA18 = "": strFA19 = "": strFA20 = "": strFA21 = "": strFA22 = "": strFA70 = ""
      End If
      
      If InStr(strFA32 & strFA33 & strFA34 & strFA35 & strFA36, "竹曆退件") > 0 Then
        strFA32 = "": strFA33 = "": strFA34 = "": strFA35 = "": strFA36 = ""
      End If
      'end 2018/10/31
      'POB
      'Modify by Amy 2018/10/31 地址改為變數判斷
      'If Not IsNull(.Fields("fa32")) Then
      If strFA32 <> MsgText(601) Then
         intRow = intRow + 1
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print strFA32
         If strFA33 <> MsgText(601) Then
            intRow = intRow + 1
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 250
            Printer.Print strFA33
         End If
         If strFA34 <> MsgText(601) Then
            intRow = intRow + 1
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 250
            Printer.Print strFA34
         End If
         If strFA35 <> MsgText(601) Then
            intRow = intRow + 1
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 250
            Printer.Print strFA35
         End If
         If strFA36 <> MsgText(601) Then
            intRow = intRow + 1
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 250
            Printer.Print strFA36
         End If
      '地址
      Else
         intRow = intRow + 1
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print strFA18
         If strFA19 <> MsgText(601) Then
            intRow = intRow + 1
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 250
            Printer.Print strFA19
         End If
         If strFA20 <> MsgText(601) Then
            intRow = intRow + 1
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 250
            Printer.Print strFA20
         End If
         If strFA21 <> MsgText(601) Then
            intRow = intRow + 1
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 250
            Printer.Print strFA21
         End If
         If strFA22 <> MsgText(601) Then
            intRow = intRow + 1
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 250
            Printer.Print strFA22
         End If
         If strFA70 <> MsgText(601) Then
            intRow = intRow + 1
            Printer.CurrentX = intLeft
            Printer.CurrentY = intTop + 1400 + intRow * 250
            Printer.Print strFA70
         End If
         'end 2018/10/31
         
         intRow = 15
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print "Dear Sirs,"
         
         intRow = intRow + 2
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print "Thank you for the payment in the amount of " & strCurrency & " " & strAmount & " on " & ChgEngDate(DBDATE(strRecDate)) & "."
         
         intRow = intRow + 2
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print "Enclosed herewith please find the receipt, which is in the name of your client."
         
         intRow = intRow + 2
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print "Best regards,"
         
         intRow = intRow + 5
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print "Jasmine Wu"
         Printer.CurrentX = intLeft + 5000
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print "Fred C.T. Yen"
         
         intRow = intRow + 1
         Printer.CurrentX = intLeft
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print "Accounting  Department"
         Printer.CurrentX = intLeft + 5000
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print "Patent Attorney"
         
         intRow = intRow + 1
         Printer.CurrentX = intLeft + 5000
         Printer.CurrentY = intTop + 1400 + intRow * 250
         Printer.Print "Managing Partner"
      End If
      End With
   End If
   Set adoTmp = Nothing
   
End Sub

Private Sub Text7_GotFocus()
    TextInverse Text7
End Sub

Private Sub Text8_GotFocus()
    TextInverse Text8
End Sub

'Added by Lydia 2016/01/29
Private Sub PrintPicture(Typ As String, iPicNo As Integer, inX1 As Double, inY1 As Double, iRate As Double)
   Dim tObj As New StdPicture, stFileName As String, pWidth As Long, pHeight As Long
   'Modified by Morgn 2024/9/5 新的收據章改在M31且比例也要調整
   If PUB_ReadDB2File(stFileName, iPicNo, "M31") Then '讀取圖檔
      Set tObj = pvGetStdPicture(stFileName)
      'pHeight = tObj.Height * 0.48
      'pWidth = tObj.Width * 0.48
      pHeight = tObj.Height * 0.88
      pWidth = tObj.Width * 0.88
      If Typ = "1" Then
         Printer.PaintPicture tObj, inX1, inY1, pWidth, pHeight
      Else
         pHeight = pHeight * iRate
         pWidth = pWidth * iRate
         Picture1.PaintPicture tObj, inX1, inY1, pWidth, pHeight
      End If
   End If
   Set tObj = Nothing
End Sub
'end 2016/01/29

'Added by Morgan 2016/3/11
'T案Email收件人
Private Function GetTReceiver()
'Modified by Morgan 2016/4/12 改抓"每日批次商標結案通知郵件收件人"
'   Dim stSQL As String, intQ As Integer
'   Dim rsQuery  As ADODB.Recordset
'   Dim stReturn As String
'
'   stSQL = "select st01 from staff where st05='93' and st04='1'"
'   intQ = 1
'   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
'   If intQ = 1 Then
'      stReturn = rsQuery.GetString(, , , ";")
'   End If
'   GetTReceiver = stReturn
'   Set rsQuery = Nothing
GetTReceiver = Pub_GetSpecMan("P")
'end 2016/4/12
End Function

'Added by Morgan 2023/10/16
'案件無FC代理人時彼號改帶客戶案件案號 Ex:P-132092
Private Function GetYourRefNo2(pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset

   stSQL = "Select PA48 From Patent Where pa01='" & pCP01 & "' and pa02='" & pCP02 & "' and pa03='" & pCP03 & "' and pa04='" & pCP04 & "' and pa75 is null"
   stSQL = stSQL & " Union Select TM35 From Trademark Where tm01='" & pCP01 & "' and tm02='" & pCP02 & "' and tm03='" & pCP03 & "' and tm04='" & pCP04 & "' and tm44 is null"
   stSQL = stSQL & " Union Select LC17 From Lawcase Where lc01='" & pCP01 & "' and lc02='" & pCP02 & "' and lc03='" & pCP03 & "' and lc04='" & pCP04 & "' and lc22 is null"
   stSQL = stSQL & " Union Select SP29 From Servicepractice Where SP01='" & pCP01 & "' and sp02='" & pCP02 & "' and sp03='" & pCP03 & "' and sp04='" & pCP04 & "' and sp26 is null"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      GetYourRefNo2 = "" & rsQuery(0)
   End If
   Set rsQuery = Nothing

End Function

'Added by Lydia 2024/12/19 Excel列印-收據信頭、信尾
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
      WksRpt1.PageSetup.HeaderMargin = Excel.Application.InchesToPoints(0.3) '頁首
      WksRpt1.PageSetup.FooterMargin = Excel.Application.InchesToPoints(0.3) '頁尾
      WksRpt1.PageSetup.TopMargin = xlsRpt.InchesToPoints(0.2) '上
      WksRpt1.PageSetup.BottomMargin = xlsRpt.InchesToPoints(0.2) '下
      WksRpt1.PageSetup.LeftMargin = xlsRpt.InchesToPoints(0.1) '左邊界
      WksRpt1.PageSetup.RightMargin = xlsRpt.InchesToPoints(0.1) '右邊界
      xlsRpt.Visible = False
   Else
      If iPageNo = 0 Then  '刪除前一張收據的內容
         WksRpt1.Shapes.SelectAll
         xlsRpt.Selection.Delete  '刪除所有圖片
         WksRpt1.Range(Chr(xCols - 1) & ":" & Chr(xColE + 1)).Select
         xlsRpt.Selection.Delete  '刪除文字
      Else
         '跨頁不清除
      End If
   End If
'-------------------欄寬和列高-----------------------------
   If iPageNo = 0 Then
      If bolChina = True Then  '大陸代理人
         xColE = 65 + 12
         For intI = 0 To 13
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Name = "Times New Roman"
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Size = 12
            Select Case intI
               Case 0, 13  'A,N
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 3.9
               Case 1, 8 'B,I
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 12
               Case 3 'D
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 17
               Case 7 'H
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 5
               Case 5, 10, 12 'F,K,M
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 13
               Case 2, 4, 6, 9, 11 '欄位間距
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 0.3
            End Select
         Next intI
      Else
         xColE = 65 + 8
         For intI = 0 To 9
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Name = "Times New Roman"
            WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).Font.Size = 12
            Select Case intI
               Case 0, 7, 9 'A,H,J
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 6
               Case 1 'B
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 17
               Case 3 'D
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 30.5
               Case 5 'F
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 16
               Case 8 'I
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 14
               Case 2, 4, 6  '欄位間距
                  WksRpt1.Range(Chr(65 + intI) & ":" & Chr(65 + intI)).ColumnWidth = 0.3
            End Select
         Next intI
      End If
      bolColTitle = True
   End If
   For intI = 1 To maxRows
      If intI = 1 Then  '信頭
         WksRpt1.Range(intI + iPageNo * maxRows & ":" & intI + iPageNo * maxRows).RowHeight = 110 '列高=3.87CM
      ElseIf intI = 2 Then  '發票/INVOICE抬頭
         WksRpt1.Range(intI + iPageNo * maxRows & ":" & intI + iPageNo * maxRows).RowHeight = 37  '列高=1.29CM
      ElseIf intI = maxRows Then '信尾
         WksRpt1.Range(intI + iPageNo * maxRows & ":" & intI + iPageNo * maxRows).RowHeight = 35  '列高=1.23CM
      Else
         WksRpt1.Range(intI + iPageNo * maxRows & ":" & intI + iPageNo * maxRows).RowHeight = 17  '列高=0.59CM
      End If
   Next intI
    
'-------------------欄寬和列高-----------------------------
   'Excel列印資料的起始,終止位置
   xRows = (iPageNo * maxRows) + 2
   xRowE = ((iPageNo + 1) * maxRows) - 2
   nRow = xRows '目前
   xCols = 66  '因為信頭JPG範圍包含上+左右邊界的空白，所以從B欄開始放入資料
   
   If bolChina = True Then
       WksRpt1.Range("F" & xRowE + 1).Value = "**" & iPageNo + 1 & "**"
       WksRpt1.Range("F" & xRowE + 1).HorizontalAlignment = xlRight
   Else
       WksRpt1.Range("D" & xRowE + 1).Value = "**" & iPageNo + 1 & "**"
       WksRpt1.Range("D" & xRowE + 1).HorizontalAlignment = xlRight
   End If

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

   PrintExcel_BFile = True
   Exit Function
End Function

'Added by Lydia 2024/12/19 改用EXCEL：列印收據
Private Sub PrintExcelMain()
Dim strKeyNo As String
Dim strDocNo As String
Dim strSQL1 As String
Dim strMailFailList() As String 'Mail 失敗清單
Dim bNewPage As Boolean
Dim iRound As Integer '迴圈次數
Dim bolChgPrinter As Boolean
Dim strNation As String
Dim strData As String
Dim strPA77 As String
Dim dblAmount As Double
Dim iPicNo As Integer
Dim StrSQLa As String
Dim bolOpenXls As Boolean

   strSql = ""
   strSQL1 = ""
   '收款日期
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strSql = strSql & " and a0y02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
      strSQL1 = strSQL1 & " and a0y02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If

   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strSql = strSql & " and a0y02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
      strSQL1 = strSQL1 & " and a0y02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   
   '代理人編號
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and decode(a0y18, '1', a0y07, '2', a0y08, a0y09) >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and decode(a0y18, '1', a0y07, '2', a0y08, a0y09) <= '" & Text2 & "'"
   End If
   '收款單號
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0y01>= '" & Text3 & "'"
   End If
   If Text4 <> MsgText(601) Then
      strSql = strSql & " and a0y01 <= '" & Text4 & "'"
   End If
   
   '國藉判斷
   If Text7 = "020" Then
        strNation = " Where fa10||''='020' "
   Else
        strNation = " Where fa10||''<>'020' "
   End If
   
   '發Email改只以內文說明不再夾帶收據圖檔(大陸地區除外)
   bolChinese = False
   strBCC = ""
   strContent = ""
   If Text5 = "2" And Text7 <> "020" And Text6 <> "Y" Then
      'Modified by Lydia 2024/09/18 +財務副本信箱emailcc：寄財務信箱一併CC副本
      strExc(0) = "select * from (select a0y01,a0y02,AgNo,a0z02,nvl(fa79,nvl(fa16,nvl(cu115,cu20))) fa16, nvl(fa83,cu119) fa83,a1k29,nvl(fa10,cu10) fa10 " & _
         ",a1k13,decode(fa105||fa79,null,decode(cu115,null,'',cu200),fa134) as emailcc from (select a0y01,a0y02,a0z02,decode(a0y18, '1', a0y07, '2', a0y08, a0y09) AgNo,a1k29" & _
         ",a1k13 From acc0y0,acc0z0,acc1k0 where a0z01(+)=a0y01 and (length(a0z02)=9 or substr(a0z02,-2)='00') and a1k01(+)=a0z02" & strSql & _
         "),fagent,customer where fa01(+)=substr(AgNo,1,8) and fa02(+)=substr(AgNo,9) and cu01(+)=substr(AgNo,1,8) and cu02(+)=substr(AgNo,9)) X" & strNation & " order by 1,2,3"
      intI = 1
      Set adoacc1k0 = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With adoacc1k0
         Erase strMailFailList
         ReDim strMailFailList(0)
         Do While Not .EOF
            strBCC = ""
            bolEmail = False
            strKeyNo = .Fields("a0y01")
            strInform = "" & adoacc1k0("fa83") '是否寄收據
            strRecDate = .Fields("a0y02") '收款日
            strFNo = "" & .Fields("AgNo")
            If strInform <> "N" Then
               strEMailBox = "" & adoacc1k0("fa16") '電子信箱
               strEmailCC = "" & adoacc1k0.Fields("emailcc") '財務副本信箱
               If strEMailBox <> "" And UCase(strEMailBox) <> "NO" Then
                  If txtReceiver <> "" Then
                     strEMailBox = txtReceiver
                     strEmailCC = ""
                  End If
                  bolEmail = True
               End If
            End If
            
            '請款單號
            strNoList = .Fields("a0z02")
            If .Fields("a1k29") <> "Y" Then bolEmail = False '部分收款不發Mail
            intI = 1
            .MoveNext
            Do While Not .EOF
               If strKeyNo <> .Fields("a0y01") Then Exit Do
               strNoList = strNoList & ", " & .Fields("a0z02")
               intI = intI + 1
               If .Fields("a1k29") <> "Y" Then bolEmail = False '部分收款不發Mail
               .MoveNext
            Loop
            If intI > 1 Then
               '最後一個 ", " 換成 " and "
               strNoList = Left(strNoList, InStrRev(strNoList, ", ") - 1) & " and " & Mid(strNoList, InStrRev(strNoList, ", ") + 2)
            End If
            
            If bolEmail Then
               strRecAmount = ""
               'Modified by Morgan 2016/2/17 若為暫收款要抓原貸方的付款方式
               'Modified by Morgan 2018/10/17 暫收款若沒有收款資料時不要帶 Ex:M10704930,N10700159 --婉莘
               'Modified by Morgan 2023/6/30 暫收款金額合併且不再帶日期 --斯閔
               strExc(0) = "select a.a1p19,sum(a.a1p21) a1p21,'' a1p23,'' a1p24,'' a1p05,1 srt from acc1p0 a where a.a1p04='" & strKeyNo & "' and a.a1p07>0 and a.a1p05='2401' group by a.a1p19" & _
                  " union select a.a1p19,a.a1p21,a.a1p23,a.a1p24,a.a1p05,3 srt from acc1p0 a where a.a1p04='" & strKeyNo & "' and a.a1p07>0 and a.a1p05<>'2401' and a1p05<>'611301' and a.a1p24 is not null" & _
                  " order by srt"
                intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(1) = ""
                  strExc(2) = ""
                  strExc(3) = ""
                  If RsTemp("srt") = 1 Then
                     strRecDate = ""
                     If RsTemp.RecordCount = 1 Then
                        strExc(1) = "credit in the amount of " & RsTemp("a1p19") & Format(RsTemp("a1p21"), FDollar)
                     Else
                        strExc(3) = " and included credit in the amount of " & RsTemp("a1p19") & Format(RsTemp("a1p21"), FDollar)
                        RsTemp.MoveNext
                     End If
                  End If

                  dblAmount = RsTemp("a1p21")

                  '台幣(科目 110204)收款時, 手續費 (科目 611301) 的金額與收款金額相加後, 其金額再列示在收據內(收台幣的時候不會有多筆收款的情況)
                  If RsTemp("a1p05") = "110204" Then
                     strExc(0) = "select a1p21 from acc1p0 where a1p04='" & strKeyNo & "' and a1p05='611301' and a1p07>0"
                     intI = 1
                     Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        dblAmount = dblAmount + adoRecordset(0)
                     End If
                  End If
                  
                  If "" & RsTemp("srt") <> 1 Then
                     'CB
                     If RsTemp("a1p24") = "2" Or RsTemp("a1p24") = "3" Then
                        strExc(1) = "check #" & RsTemp("a1p23")
                        strExc(2) = " in the amount of " & RsTemp("a1p19") & Format(dblAmount, FDollar)
                     'IR
                     Else
                        strExc(1) = "your remittance"
                        strExc(2) = " of " & RsTemp("a1p19") & Format(dblAmount, FDollar)
                     End If
                  End If

                  RsTemp.MoveNext
                  Do While Not RsTemp.EOF
                     intI = intI + 1
                     If RsTemp("a1p24") = "2" Then
                        strExc(1) = strExc(1) & ", #" & RsTemp("a1p23")
                     End If
                     strExc(2) = strExc(2) & ", " & RsTemp("a1p19") & Format(RsTemp("a1p21"), FDollar)
                     RsTemp.MoveNext
                  Loop
                  If intI > 2 Then
                     If InStrRev(strExc(1), ", ") > 0 Then
                        strExc(1) = Left(strExc(1), InStrRev(strExc(1), ", ") - 1) & Replace(strExc(1), ", ", " and ", InStrRev(strExc(1), ", "))
                     End If
                     strExc(2) = Left(strExc(2), InStrRev(strExc(2), ", ") - 1) & Replace(strExc(2), ", ", " and ", InStrRev(strExc(2), ", "))
                  ElseIf intI = 2 Then
                     strExc(1) = Replace(strExc(1), ", ", " and ")
                     strExc(2) = Replace(strExc(2), ", ", " and ")
                  End If

                  strRecAmount = strExc(1) & strExc(2) & strExc(3)
                  
                  If strFNo = "Y52431000" Or strFNo = "Y45848000" Then
                     PUB_SendMail strUserNum, strEMailBox, "", "Ticket # [N-297175] Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , , True, True, True, strEmailCC, strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
                  Else
                     PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , , True, True, True, strEmailCC, strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
                  End If
                  
                  '有寄mail的產生特殊收據清單
                  If Check1.Value = vbChecked Then
                        StrSQLa = "Insert Into Accrpt2460 (ID,R001) Values ('" & strUserNum & "','" & strFNo & "')"
                        cnnConnection.Execute StrSQLa
                  End If
                  
                  bolMailFailNoAlert = False
                  If bolMailSendOk = False Then
                     If strMailFailList(0) <> "" Then
                        ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                     End If
                     strMailFailList(UBound(strMailFailList)) = strFNo & " : " & strEMailBox
                  End If
               Else
                  If strMailFailList(0) <> "" Then
                     ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                  End If
                  strMailFailList(UBound(strMailFailList)) = strFNo & " : " & strEMailBox & " => " & strKeyNo & "無法確認收款內容"
               End If
            End If
         Loop
         End With
         
         If strMailFailList(0) <> "" Then
            strExc(0) = "E-Mail失敗清單：" & vbCrLf & vbCrLf
            For intI = 0 To UBound(strMailFailList)
               strExc(0) = strExc(0) & strMailFailList(intI) & vbCrLf
            Next
            If MsgBox(strExc(0) & vbCrLf & "是否要列印？" & vbCrLf, vbYesNo + vbDefaultButton1) = vbYes Then
               Printer.Print strExc(0)
               Printer.EndDoc
            End If
         End If
   
      Else
         MsgBox MsgText(28), , MsgText(5)
      End If
      adoacc1k0.Close
      Exit Sub
   End If '------------If Text5 = "2" And Text7 <> "020" And Text6 <> "Y" Then 發Email改只以內文說明不再夾帶收據圖檔(大陸地區除外)

'----------以下列印收據Code
On Error GoTo ErrorHandle

   adoacc1k0.CursorLocation = adUseClient

   strExc(0) = "select * from (select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, fa05, fa63, fa64, fa65, fa32, fa33, fa34, fa35, fa36, fa06, fa23, fa18, fa19, fa20, fa21, fa22, fa04, fa17, fa01, fa02, a1k02, FA31, substr(a1k01, 1, 8) as DocNo, a1k08 as Amount, a0y02,fa10,nvl(fa79,fa16) fa16,fa83,fa70,a1k28,a0y19,decode(fa105||fa79,null,'',fa134) as emailcc from acc1k0, acc0z0, acc0y0, fagent where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 " & strSql & " and length(a1k01) = 9" & _
            " union select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu66 as fa33, cu67 as fa34, cu68 as fa35, cu69 as fa36, cu06 as fa06, cu29 as fa23, cu24 as fa18, cu25 as fa19, cu26 as fa20, cu27 as fa21, cu28 as fa22, cu04 as fa04, NVL(CU31,cu23) as fa17, cu01 as fa01, cu02 as fa02, a1k02, cu64 as FA31, substr(a1k01, 1, 8) as DocNo, a1k08 as Amount, a0y02,cu10 as fa10,nvl(cu115,cu20) as fa16, cu119 as fa83,cu102 as fa70,a1k28,a0y19,decode(cu115,null,'',cu200) as emailcc from acc1k0, acc0z0, acc0y0, customer where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 " & strSql & " and length(a1k01) = 9" & _
            ") new "
   strExc(0) = strExc(0) & " union select  a0y01,a0y18,a0y07,a0y08,a0y09,a0y10,a1k18,a0y06,a1k13,a1k14,a1k15,a1k16,a1k01,fa05,fa63,fa64,fa65,fa32,fa33,fa34,fa35,fa36 " & _
            ",fa06,fa23,fa18,fa19,fa20,fa21,fa22,fa04,fa17,fa01,fa02,a1k02,fa31,docno,amount,a0y02,fa10,fa16,fa83,fa70,a1k28,a0y19,emailcc " & _
            " from (select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, fa05, fa63, fa64, fa65, fa32, fa33, fa34, fa35, fa36, fa06, fa23, fa18, fa19, fa20, fa21, fa22, fa04, fa17, fa01, fa02, a1k02, fa31,decode(fa105||fa79,null,'',fa134) as emailcc from acc1k0, acc0z0, acc0y0, fagent where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 " & strSql & " and substr(a1k01, 9, 2) = '00'" & _
            " union select a0y01, a0y18, a0y07, a0y08, a0y09, a0y10, a0y03 as a1k18, a0y06, a1k13, a1k14, a1k15, a1k16, a1k01, cu05 as fa05, cu88 as fa63, cu89 as fa64, cu90 as fa65, cu65 as fa32, cu66 as fa33, cu67 as fa34, cu68 as fa35, cu69 as fa36, cu06 as fa06, cu29 as fa23, cu24 as fa18, cu25 as fa19, cu26 as fa20, cu27 as fa21, cu28 as fa22, cu04 as fa04, NVL(CU31,cu23) as fa17, cu01 as fa01, cu02 as fa02, a1k02, cu64 as fa31,decode(cu115,null,'',cu200) as emailcc from acc1k0, acc0z0, acc0y0, customer where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 " & strSql & " and substr(a1k01, 9, 2) = '00'" & _
            ") new,(select substr(a1k01, 1, 8) as DocNo, sum(a1k08) as Amount,max(a0y02) as a0y02,max(fa10) as fa10,max(nvl(fa79,fa16)) as fa16,max(fa83) fa83,max(fa70) fa70,max(a1k28) a1k28,max(a0y19) a0y19 from acc1k0, acc0z0, acc0y0, fagent where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 " & strSql & " and length(a1k01) = 10 group by substr(a1k01, 1, 8)" & _
            " union select substr(a1k01, 1, 8) as DocNo, sum(a1k08) as Amount,max(a0y02) as a0y02,max(cu10) as fa10,max(nvl(cu115,cu20)) as fa16,max(cu119) fa83,max(cu102) fa70,max(a1k28) a1k28,max(a0y19) a0y19 from acc1k0, acc0z0, acc0y0, customer where a1k01 = a0z02 and a0z01 = a0y01 and decode(a0y18, '1', substr(a0y07, 1, 8), '2', substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = cu01 and decode(a0y18, '1', substr(a0y07, 9, 1), '2', substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = cu02 " & strSql & " and length(a1k01) = 10 group by substr(a1k01, 1, 8)" & _
            ") old where substr(a1k01, 1, 8) = DocNo "
   strSql = "select * from (" & strExc(0) & ") X " & strNation & "order by a0y01, decode(a0y18, '1', a0y07, '2', a0y08, a0y09) asc, a1k13 asc, a1k14 asc, a1k15 asc, a1k16 asc, a1k01 asc"

   adoacc1k0.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoacc1k0.RecordCount = 0 Then
      adoacc1k0.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   '存電子檔時不印
   If Text6 = "Y" Then
      strSavePath = PUB_Getdesktop
      bol2Printer = False
   Else
      strSavePath = App.path
      If Text5 <> "2" Then
         bol2Printer = True
      Else
         bol2Printer = False
      End If
   End If
   
   Call Pub_ChkExcelPath(strPrtPath)
   Call PUB_KillTempFile(strUserNum & "\$*.*")

'Modified by Morgan 2014/5/14 電子化不再留卷
iRound = 1

For iCopy = 1 To iRound
   Erase strMailFailList
   ReDim strMailFailList(0)
   strPicLetter = ""
   strPicFileNames = ""
   strNo = ""
   strDocNo = ""
   douAmount = 0
   intLength = 0
   iPageNo = 0
   strBCC = ""
   adoacc1k0.MoveFirst
   
   'Add By Sindy 2014/3/10 大陸案:加控制,收據要增加案件名稱,申請人
   If Text7 = "020" Then
      bolChina = True
   Else
      bolChina = False
   End If

   Do While adoacc1k0.EOF = False
      
      '定稿語文
      strLangTmp = "" & adoacc1k0("fa31")
      
      '若有下國籍條件
      If Text7 <> MsgText(601) Then
         If adoacc1k0.Fields("fa10").Value < Text7 Then
            GoTo NextSkip
         End If
      End If
      If Text8 <> MsgText(601) Then
         If adoacc1k0.Fields("fa10").Value > Text8 & "z" Then
            GoTo NextSkip
         End If
      End If
      '舊系統請款單
      If Len(adoacc1k0.Fields("a1k01").Value) = 10 Then
         If strDocNo = Mid(adoacc1k0.Fields("a1k01").Value, 1, 8) Then
            GoTo NextSkip
         Else
            strDocNo = Mid(adoacc1k0.Fields("a1k01").Value, 1, 8)
         End If
      End If
      
      '收款單號
      If IsNull(adoacc1k0.Fields("a0y01").Value) Then
         strKeyNo = ""
      Else
         strKeyNo = adoacc1k0.Fields("a0y01").Value
      End If
      
      If adoacc1k0.AbsolutePosition = 1 Then
         bolOpenXls = True
      Else
         bolOpenXls = False
      End If
      
      If strNo <> strKeyNo Then
         If douAmount <> 0 Then
         
            Call PrintExcel_BSum
            
            If bol2File = True Then
               '先存PDF檔(另存新檔)放在桌面，不關EXCEL後面再處理信頭、信尾>>PrintExcel_BFile
               If PUB_PrintExcel2File(xlsRpt, strSavePath, "$" & strNo & ".PDF", strExc(1), False) = True Then
                  strPicFileNames = strSavePath & "\" & strExc(1)
               End If
               
               '當附件
               'Modify by Morgan 2010/6/10 選發EMail且未設要存電子檔才寄送
               If bolEmail = True And Text5.Text = "2" And Text6.Text <> "Y" Then
                  bolMailFailNoAlert = True
                  bolMailSendOk = False
                  
                  If strFNo = "Y52431000" Or strFNo = "Y45848000" Then
                     PUB_SendMail strUserNum, strEMailBox, "", "Ticket # [N-297175] Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
                  Else
                     PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
                  End If
                  '有寄mail的產生特殊收據清單
                  If Check1.Value = vbChecked Then
                        StrSQLa = "Insert Into Accrpt2460 (ID,R001) Values ('" & strUserNum & "','" & strFNo & "')"
                        cnnConnection.Execute StrSQLa
                  End If

                  bolMailFailNoAlert = False
                  If bolMailSendOk = False Then
                     If strMailFailList(0) <> "" Then
                        ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
                     End If
                     strMailFailList(UBound(strMailFailList)) = strFNo & " : " & strEMailBox
                  End If
                  strPicFileNames = ""
               End If
               strPicFileNames = "" '已存電子檔,清空路徑
            End If
            If bol2Printer = True Then
               WksRpt1.PrintOut Copies:=1, Collate:=True '列印
            End If
            iPageNo = 0
            douAmount = 0
         End If
         
         strBCC = ""
         strNoList = ""

         '是否寄收據
         strInform = "" & adoacc1k0("fa83")
         '電子信箱
         strEMailBox = "" & adoacc1k0("fa16")
         strEmailCC = "" & adoacc1k0.Fields("emailcc") '財務副本信箱
         If strEMailBox <> "" And UCase(strEMailBox) <> "NO" Then
            If txtReceiver <> "" Then
               strEMailBox = txtReceiver
               strEmailCC = ""
            End If
            
            bolEmail = True
            bol2File = True
            If "" & adoacc1k0("fa10") = "020" Then
               bolChinese = True
            Else
               bolChinese = False
            End If
         Else
            bolEmail = False
            bol2File = False
         End If
         
         'Add by Morgan 2008/4/16 放在迴圈內是因為要印是否"EMail通知","不寄收據"的註記
         '產生電子檔
         If Text6.Text = "Y" Then
            bol2File = True
         '不發Mail 或 不寄收據 時只列印不產生電子檔
         ElseIf Text5.Text <> "2" Or strInform = "N" Then
            bol2File = False
         End If
         
         If Text5.Text = "1" Then
            If bol2File = True Then
               bol2Printer = False
            'E+印的也要印
            ElseIf strInform = "B" Then
               bol2Printer = True
            ElseIf (strInform = "N" Or bolEmail = True) Then
               bol2Printer = False
            Else
               bol2Printer = True
            End If
         '強制列印
         ElseIf Text5.Text = "3" Then
            If (strInform = "N" Or bol2File = True) Then
               bol2Printer = False
            Else
               bol2Printer = True
            End If
         End If
          
         '是否存電子檔
         If bol2File = True Then
            m_iNo = 0: m_iNo2 = 0
            If strSrvDate(1) >= 智慧所更名日 Then
               '改用EXCEL：信頭、信尾分開來
               PUB_GetLetterPicID "2", , m_iNo, m_iNo2, , , "HALF"
            Else
               m_iNo = 6
            End If
            
         End If
         If PrintExcel_BFile(bolOpenXls, m_iNo, m_iNo2) = False Then
         End If
         
         Call PrintExcel_BHead
         
         strA0Y10 = ""
         If IsNull(adoacc1k0.Fields("a0y10").Value) = False Then '暫收款單號
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "CREDIT"
            WksRpt1.Range(Chr(xCols + 7) & nRow).Value = "" & adoacc1k0.Fields("a1k18").Value '幣別
            WksRpt1.Range(Chr(xCols + 8) & nRow).Value = Format(Val(adoacc1k0.Fields("a0y06").Value), FDollar) '金額
            WksRpt1.Range(Chr(xCols + 8) & nRow).NumberFormatLocal = FDollar
            douAmount = douAmount + Val(adoacc1k0.Fields("a0y06").Value)
            Call PrintExcel_BPage
         End If
         strNo = strKeyNo
         strRecDate = "" & adoacc1k0("a0y02")
         strA1K28 = "" & adoacc1k0("a1k28")
      End If  '------If strNo <> strKeyNo Then
      
      intCounter = 0 '列印欄位Column
      strNoList = strNoList & adoacc1k0.Fields("a1k01").Value & " (" & adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & IIf(adoacc1k0.Fields("a1k15") & adoacc1k0.Fields("a1k16") = "000", "", "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value & ")") & ")" & vbCrLf & String(5, "　")
      '收據抬頭(指定代理人)
      Select Case adoacc1k0.Fields("a0y18").Value
         Case "1"
            If IsNull(adoacc1k0.Fields("a0y07").Value) Then
               strFNo = ""
            Else
               strFNo = adoacc1k0.Fields("a0y07").Value
            End If
         Case "2"
            If IsNull(adoacc1k0.Fields("a0y08").Value) Then
               strFNo = ""
            Else
               strFNo = adoacc1k0.Fields("a0y08").Value
            End If
         Case "3"
            If IsNull(adoacc1k0.Fields("a0y09").Value) Then
               strFNo = ""
            Else
               strFNo = adoacc1k0.Fields("a0y09").Value
            End If
         Case Else
            If IsNull(adoacc1k0.Fields("a0y09").Value) Then
               strFNo = ""
            Else
               strFNo = adoacc1k0.Fields("a0y09").Value
            End If
      End Select
            
      adoquery.CursorLocation = adUseClient
      '收款金額
      If Len(adoacc1k0.Fields("a1k01").Value) = 10 Then
         adoquery.Open "select sum(a0z04) from acc0z0, acc0y0 where a0z01 = a0y01 and substr(a0z02, 1, 8) = '" & Mid(adoacc1k0.Fields("a1k01").Value, 1, 8) & "' and A0y01='" & strKeyNo & "' " & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
      Else
         adoquery.Open "select sum(a0z04) from acc0z0, acc0y0 where a0z01 = a0y01 and a0z02 = '" & adoacc1k0.Fields("a1k01").Value & "' and A0y01='" & strKeyNo & "' " & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
      End If

      If adoquery.RecordCount <> 0 Then
         If adoquery.Fields(0).Value <> 0 Then
            If Not IsNull(adoacc1k0.Fields("A0Y10").Value) Then '暫收款單號
               strA0Y10 = adoacc1k0.Fields("A0Y10").Value
            End If
            
            '我方文號
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & IIf(adoacc1k0.Fields("a1k15").Value & adoacc1k0.Fields("a1k16").Value = "000", "", "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value)
            intCounter = intCounter + 2
            
            '貴方文號
            strPA77 = GetYourRefNo1(adoacc1k0.Fields("a1k13").Value, adoacc1k0.Fields("a1k14").Value, adoacc1k0.Fields("a1k15").Value, adoacc1k0.Fields("a1k16").Value, _
                            IIf(Left(strFNo, 6) = "Y52269", True, False))
            If strPA77 = MsgText(601) Then
               strPA77 = GetYourRefNo2(adoacc1k0.Fields("a1k13").Value, adoacc1k0.Fields("a1k14").Value, adoacc1k0.Fields("a1k15").Value, adoacc1k0.Fields("a1k16").Value)
            End If
            If strPA77 <> MsgText(601) Then
               'Modify by Amy 2014/07/11 只取18個字-婉莘; Excel 欄寬不足先用14
               WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = Left(strPA77, 14)
            End If
            intCounter = intCounter + 2
            '帳單編號
            If Not IsNull(adoacc1k0.Fields("a1k01").Value) Then
               WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = IIf(Len(adoacc1k0.Fields("a1k01").Value) > 10, Mid(adoacc1k0.Fields("a1k01").Value, 3, 6), adoacc1k0.Fields("a1k01").Value)
            End If
            intCounter = intCounter + 2
            
            '幣別
            adocheck.CursorLocation = adUseClient
            adocheck.Open "select distinct a0y03 from acc0z0, acc0y0 where a0z01 = a0y01 and a0z02 = '" & adoacc1k0.Fields("a1k01").Value & "' and A0y01='" & strKeyNo & "' ", adoTaie, adOpenStatic, adLockReadOnly
            If adocheck.RecordCount <> 0 Then
               If IsNull(adocheck.Fields("a0y03").Value) Then
                  strCurrency = "USD"
               Else
                  strCurrency = adocheck.Fields("a0y03").Value
               End If
            Else
               strCurrency = "USD"
            End If
            adocheck.Close
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strCurrency
            intCounter = intCounter + 1
            
            '金額
            If IsNull(adoquery.Fields(0).Value) = False Then
               strAmount = Format(Val(adoquery.Fields(0).Value), FDollar)
               WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strAmount
               WksRpt1.Range(Chr(xCols + intCounter) & nRow).NumberFormatLocal = FDollar
               douAmount = douAmount + Val(adoquery.Fields(0).Value)
            End If
            intCounter = intCounter + 2
            
            If bolChina = True Then
               '案件名稱
               strData = GetPrjName("" & adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value)
               'Excel 欄寬不足先用12
               WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = convForm(CheckStr(strData), 12)
               intCounter = intCounter + 2
               
               '申請人
               strData = GetPrjPeopleNum1("" & adoacc1k0.Fields("a1k13").Value & "-" & adoacc1k0.Fields("a1k14").Value & "-" & adoacc1k0.Fields("a1k15").Value & "-" & adoacc1k0.Fields("a1k16").Value)
               If strData <> "" Then
                  strData = GetPrjPeople1(strData)
                  'Excel 欄寬不足先用12
                  WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = convForm(CheckStr(strData), 12)
               End If
               intCounter = intCounter + 2
            End If
            Call PrintExcel_BPage
         End If
      End If
      adoquery.Close
NextSkip:
      adoacc1k0.MoveNext
   Loop
   
   Call PrintExcel_BSum

   If bol2File = True Then
      'PDF檔放在桌面
      If PUB_PrintExcel2File(xlsRpt, strSavePath, "$" & strNo & ".PDF", strExc(1), False) = True Then
         strPicFileNames = strPicFileNames & strSavePath & "\" & strExc(1) & "*"
      End If
         
      '當附件
      'Modify by Morgan 2010/6/10 選發EMail且未設要存電子檔才寄送
      If bolEmail = True And Text5.Text = "2" And Text6.Text <> "Y" Then
         bolMailFailNoAlert = True
         bolMailSendOk = False
         If strFNo = "Y52431000" Or strFNo = "Y45848000" Then
            PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
         Else
            PUB_SendMail strUserNum, strEMailBox, "", "Receipt Acknowledgement (" & strFNo & ")", GetMailContent, , strPicFileNames, True, True, True, , strAccMailBox, "TAI E INTERNATIONAL PATENT & LAW OFFICE", strAccMailBox, , , strBCC
         End If
         '有寄mail的產生特殊收據清單
         If Check1.Value = vbChecked Then
                StrSQLa = "Insert Into Accrpt2460 (ID,R001) Values ('" & strUserNum & "','" & strFNo & "')"
                cnnConnection.Execute StrSQLa
         End If
         
         bolMailFailNoAlert = False
         If bolMailSendOk = False Then
            If strMailFailList(0) <> "" Then
               ReDim Preserve strMailFailList(UBound(strMailFailList) + 1)
            End If
            strMailFailList(UBound(strMailFailList)) = strFNo & " : " & strEMailBox
         End If
      Else
         If strPicFileNames <> "" Then
            MsgBox "電子檔已存桌面！"
         End If
      End If
      '刪除舊的暫存檔
      Call PUB_KillTempFile(strUserNum & "\$*.*")
   End If

   xlsRpt.Workbooks(1).Save
   If bol2Printer = True Then
      WksRpt1.PrintOut Copies:=1, Collate:=True '列印
   End If
   xlsRpt.Workbooks.Close
   xlsRpt.Quit
   Set xlsRpt = Nothing
   Set WksRpt1 = Nothing
      
   If strMailFailList(0) <> "" Then
      strExc(0) = "E-Mail失敗清單：" & vbCrLf & vbCrLf
      For intI = 0 To UBound(strMailFailList)
         strExc(0) = strExc(0) & strMailFailList(intI) & vbCrLf
      Next
      If MsgBox(strExc(0) & vbCrLf & "是否要列印？" & vbCrLf, vbYesNo + vbDefaultButton1) = vbYes Then
         Printer.Print strExc(0)
         Printer.EndDoc
      End If
   End If
Next

   adoacc1k0.Close
   
   Exit Sub
    
ErrorHandle:
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Resume Next
    End If
End Sub

'Adeed by Lydia 2024/12/19 改用EXCEL：收據票抬頭列印
Private Sub PrintExcel_BHead()
Dim intRow As Integer
Dim strFA17 As String, strFA18 As String, strFA19 As String, strFA20 As String, strFA21 As String, strFA22 As String, strFA70 As String, strFA23 As String
Dim strFA32 As String, strFA33 As String, strFA34 As String, strFA35 As String, strFA36 As String
   
   
   'Modify by Morgan 2006/10/25 沒設定時預設英文
   strLanguage = strLangTmp
   If strLanguage <> "1" Then strLanguage = "2"

   If iPageNo > 0 Then '跨頁
      '最後會跳行
   Else
      
      '發票/INVOICE抬頭
      strExc(1) = strPrtPath & "\$Tmp03.jpg"
      If PUB_ReadDB2File(strExc(1), "17", "M31", "0", IIf(strLanguage = "2", "00", "01")) = True Then
         Set oShape3 = WksRpt1.Shapes.AddPicture(strExc(1), True, True, xlsRpt.CentimetersToPoints(7.5), xlsRpt.CentimetersToPoints(3.9), xlsRpt.CentimetersToPoints(4.36), xlsRpt.CentimetersToPoints(1.23))
      End If
        
      'Add by Morgan 2005/3/18 代理人Y20412要求不寄紙本故加註"收據不印"字樣--婧瑄
      If Left("" & strA1K28, 5) = "Y20412" Then
         WksRpt1.Range("H" & nRow + 2).Value = "收據不寄"
      End If
      If bolChina = True Then
         strExc(1) = "K"
      Else
         strExc(1) = "H"
      End If
      If strLanguage = "1" Then
         '有輸日期起時才帶，沒有將用當天
         If Me.MaskEdBox1.Text = "___/__/__" Then
            WksRpt1.Range(strExc(1) & nRow).Value = Format(strSrvDate(1), "#### 年 ## 月 ## 日")
         Else
            WksRpt1.Range(strExc(1) & nRow).Value = Format(DBDATE(Me.MaskEdBox1.Text), "#### 年 ## 月 ## 日")
         End If
      Else
         '有輸日期起時才帶，沒有將用當天
         If Me.MaskEdBox1.Text = "___/__/__" Then
            WksRpt1.Range(strExc(1) & nRow).Value = Format(AFDate(strSrvDate(1)), "mmm. d, yyyy")
         Else
            WksRpt1.Range(strExc(1) & nRow).Value = Format(AFDate(ChangeTStringToWString(Replace(Me.MaskEdBox1.Text, "/", ""))), "mmm. d, yyyy")
         End If
      End If
      WksRpt1.Range(strExc(1) & nRow).VerticalAlignment = xlBottom
      
      nRow = nRow + 1
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
            
      '地址改為變數判斷
      Select Case strLanguage
         Case "1" '中文(中-->英-->日)
            '代理人名稱
            If IsNull(adoacc1k0.Fields("a0y19").Value) = False Then  '大陸收據抬頭
               WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("a0y19").Value
            Else
               If IsNull(adoacc1k0.Fields("fa04").Value) = False Then
                  WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa04").Value
               ElseIf IsNull(adoacc1k0.Fields("fa05").Value) = False Then
                  WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa05").Value
                  If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa63").Value
                  End If
                  If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa64").Value
                  End If
                  If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
                     nRow = nRow + 1
                     WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa65").Value
                   End If
               ElseIf IsNull(adoacc1k0.Fields("fa06").Value) = False Then
                   '日文名稱
                   WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa06").Value
               End If
            End If
            
            '代理人地址
            '地址,順序:中文->POB->英文->日文 --- Memo by Lydia 2025/09/25
            If strFA17 <> MsgText(601) Then
               nRow = nRow + 1 'Added by Morgan 2025/9/18
               '中文地址
JumpToCAddr1:
               If LenB(strFA17) > 44 Then
                  strExc(0) = PUB_StrToStr(strFA17, 42)
                  strFA17 = Mid(strFA17, Len(strExc(0)) + 1)
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strExc(0)
               Else
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA17
                  strFA17 = ""
               End If
               If strFA17 <> "" Then
                  nRow = nRow + 1
                  GoTo JumpToCAddr1
               End If
               
            'POB1~POB5
            ElseIf Trim(strFA32 & strFA33 & strFA34 & strFA35 & strFA36) <> MsgText(601) Then
               If strFA32 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA32
               End If
               If strFA33 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA33
               End If
               If strFA34 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA34
               End If
               If strFA35 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA35
               End If
               If strFA36 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA36
               End If
            '英文地址1~6
            ElseIf Trim(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70) <> MsgText(601) Then
               If strFA18 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA18
               End If
               If strFA19 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA19
               End If
               If strFA20 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA20
               End If
               If strFA21 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA21
               End If
               If strFA22 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA22
               End If
               If strFA70 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA70
               End If
            ElseIf strFA23 <> MsgText(601) Then
               '日文地址
               nRow = nRow + 1
               WksRpt1.Range(Chr(xCols) & nRow).Value = strFA23
            End If
            
         Case "2" '英文(英-->中-->日)
            '代理人名稱
            If IsNull(adoacc1k0.Fields("fa05").Value) = False Then
               WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa05").Value
               If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa63").Value
               End If
               If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa64").Value
               End If
               If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa65").Value
               End If
            ElseIf IsNull(adoacc1k0.Fields("fa04").Value) = False Then
               WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa04").Value
            ElseIf IsNull(adoacc1k0.Fields("fa06").Value) = False Then
               WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa06").Value
            End If
            '代理人地址
            '地址,順序:POB->英文->中文->日文 --- Memo by Lydia 2025/09/25
            'POB1~POB5
            If Trim(strFA32 & strFA33 & strFA34 & strFA35 & strFA36) <> MsgText(601) Then
               If strFA32 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA32
               End If
               If strFA33 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA33
               End If
               If strFA34 <> "" Then
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA34
                  nRow = nRow + 1
               End If
               If strFA35 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA35
               End If
               If strFA36 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA36
               End If
            '英文地址1~6
            ElseIf Trim(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70) <> MsgText(601) Then
               If strFA18 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA18
               End If
               If strFA19 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA19
               End If
               If strFA20 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA20
               End If
               If strFA21 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA21
               End If
               If strFA22 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA22
               End If
               If strFA70 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA70
               End If
            ElseIf strFA17 <> MsgText(601) Then
               nRow = nRow + 1 'Added by Lydia 2025/09/25
               '中文地址
JumpToCAddr2:
               If LenB(strFA17) > 44 Then
                  strExc(0) = PUB_StrToStr(strFA17, 42)
                  strFA17 = Mid(strFA17, Len(strExc(0)) + 1)
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strExc(0)
               Else
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA17
                  strFA17 = ""
               End If
               If strFA17 <> "" Then
                  nRow = nRow + 1
                  GoTo JumpToCAddr2
               End If
            ElseIf strFA23 <> MsgText(601) Then
               '日文地址
               nRow = nRow + 1
               WksRpt1.Range(Chr(xCols) & nRow).Value = strFA23
            End If
            
         Case "3" '日文(日-->英-->中)
            '代理人名稱
            If IsNull(adoacc1k0.Fields("fa06").Value) = False Then
                WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa06").Value
            ElseIf IsNull(adoacc1k0.Fields("fa05").Value) = False Then
               WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa05").Value
               If IsNull(adoacc1k0.Fields("fa63").Value) = False Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa63").Value
               End If
               If IsNull(adoacc1k0.Fields("fa64").Value) = False Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa64").Value
               End If
               If IsNull(adoacc1k0.Fields("fa65").Value) = False Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa65").Value
               End If
            ElseIf IsNull(adoacc1k0.Fields("fa04").Value) = False Then
                WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoacc1k0.Fields("fa04").Value
            End If

            '代理人地址
            '順序：日文->POB->英文->中文
            '日文
            If strFA23 <> MsgText(601) Then
                nRow = nRow + 1
                WksRpt1.Range(Chr(xCols) & nRow).Value = strFA23
            'POB1~POB5
            ElseIf Trim(strFA32 & strFA33 & strFA34 & strFA35 & strFA36) <> MsgText(601) Then
               If strFA32 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA32
               End If
               If strFA33 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA33
               End If
               If strFA34 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA34
               End If
               If strFA35 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA35
               End If
               If strFA36 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA36
               End If
            '英文地址1~6
            ElseIf Trim(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70) <> MsgText(601) Then
               If strFA18 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA18
               End If
               If strFA19 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA19
               End If
               If strFA20 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA20
               End If
               If strFA21 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA21
               End If
               If strFA22 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA22
               End If
               If strFA70 <> "" Then
                  nRow = nRow + 1
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA70
               End If
            '中文地址
            ElseIf strFA17 <> MsgText(601) Then
               nRow = nRow + 1 'Added by Lydia 2025/09/25
JumpToCAddr3:
               If LenB(strFA17) > 44 Then
                  strExc(0) = PUB_StrToStr(strFA17, 42)
                  strFA17 = Mid(strFA17, Len(strExc(0)) + 1)
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strExc(0)
               Else
                  WksRpt1.Range(Chr(xCols) & nRow).Value = strFA17
                  strFA17 = ""
               End If
               If strFA17 <> "" Then
                  nRow = nRow + 1
                  GoTo JumpToCAddr3
               End If
            End If
      End Select
      If nRow < 6 Then nRow = 6

      nRow = nRow + 2
      If strLanguage = "1" Then
         strExc(1) = ""
         strExc(0) = "select SUM(A0Z04) FROM ACC0Z0 WHERE A0Z01='" & adoacc1k0("A0Y01") & "'"
         intI = 1
         Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "" & adoacc1k0("a1k18") & " " & Format(AdoRecordSet3.Fields(0) + Val("" & adoacc1k0("A0Y06")), FDollar)
         End If

         strExc(0) = "茲通知收到　貴方的付款，金額為 " & strExc(1) & "，已支付下列帳單。"
         WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
         WksRpt1.Range(Chr(xCols) & nRow).Value = strExc(0)
      Else
         WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
         WksRpt1.Range(Chr(xCols) & nRow).Value = String(6, " ") & ReportSum(100)  'We acknowledge with thanks receipt
      End If
   End If  '----------If iPageNo > 0 Then '跨頁
   iPageNo = iPageNo + 1
   nRow = nRow + 2 '跨頁不印公司名稱+地址，只跳2行
   '欄位抬頭
   If bolColTitle = True Then
      intCounter = 0
      If strLanguage = "1" Then
         WksRpt1.Range(Chr(xCols) & nRow).Value = "我方文號"
         intCounter = intCounter + 2
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "貴方文號"
         intCounter = intCounter + 2
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "帳單編號"
         intCounter = intCounter + 2
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "    金額"
         If bolChina = True Then
            intCounter = intCounter + 3
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "案件名稱"
            intCounter = intCounter + 2
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "申請人"
         End If
      Else
         WksRpt1.Range(Chr(xCols) & nRow).Value = "OUR REF"
         intCounter = intCounter + 2
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "YOUR REF"
         intCounter = intCounter + 2
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "DEBIT NOTE."
         intCounter = intCounter + 2
         WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "    AMOUNT"
         If bolChina = True Then
            intCounter = intCounter + 3
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "TITILE"
            intCounter = intCounter + 2
            WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "APPLICANT"
         End If
      End If
      WksRpt1.Range(Chr(xCols) & nRow & ":" & Chr(xColE) & nRow).Borders(xlEdgeBottom).LineStyle = xlContinuous  '儲存格底線
      nRow = nRow + 1
   End If
   
End Sub

'Adeed by Lydia 2024/12/19 改用EXCEL：收據換行
Private Sub PrintExcel_BPage(Optional ByVal pAddLine As Integer = 1)
   nRow = nRow + pAddLine
   If nRow >= xRowE Then
      Call PrintExcel_BFile(False, m_iNo, m_iNo2)
      Call PrintExcel_BHead
   End If
End Sub

'Added by Lydia 2024/12/19 改用EXCEL：收據TOTAL列印
Private Sub PrintExcel_BSum()
Dim strTmp As String

   WksRpt1.Range(Chr(xCols) & nRow & ":" & Chr(xColE) & nRow).Borders(xlEdgeTop).LineStyle = xlContinuous  '儲存格上邊界框線
   Call PrintExcel_BPage
   
   intCounter = 4
   If strLanguage = "1" Then
      WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = "金額合計"
   End If
   intCounter = intCounter + 2
   WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strCurrency
   intCounter = intCounter + 1
   strAmount = Format(douAmount, FDollar)
   WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = strAmount
   WksRpt1.Range(Chr(xCols + intCounter) & nRow).NumberFormatLocal = FDollar
   strRecAmount = strCurrency & " " & strAmount
   
   bolColTitle = False
   Call PrintExcel_BPage
   intCounter = 6
   WksRpt1.Range(Chr(xCols + intCounter) & nRow).Value = String(20, "v")
   
   Call PrintExcel_BPage(6)
   
   '預估:中文加印收據章
   If strLanguage = "1" Then
      If nRow - ((iPageNo - 1) * maxRows) + 4 > maxRows Then
         Call PrintExcel_BPage(4)
      End If
   End If
   
   'Add by Morgan 2007/3/3 不必寄發收據註記
   If strInform = "N" Then
      '不必寄發收據
      strTmp = "●" & strFNo
   ElseIf bol2Printer = True Or bol2File = True Then
      '要發Mail
      strTmp = "＊" & strFNo
   Else
      strTmp = strFNo
   End If
   WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
   WksRpt1.Range(Chr(xCols) & nRow).Value = strTmp  '代理人編號前面加註記
   
   '中文加印收據章
   If strLanguage = "1" Then
      strExc(1) = strPrtPath & "\$Tmp04.png"
      If PUB_ReadDB2File(strExc(1), "16", "M31") = True Then
         If nRow - ((iPageNo - 1) * maxRows) < 8 Then '跨頁不印公司名稱+地址
            Set oShape3 = WksRpt1.Shapes.AddPicture(strExc(1), True, True, xlsRpt.CentimetersToPoints(13), xlsRpt.CentimetersToPoints(3.87 + 1.29 + (29.55 * (iPageNo - 1))), xlsRpt.CentimetersToPoints(5.6), xlsRpt.CentimetersToPoints(4.21))
         Else
            Set oShape3 = WksRpt1.Shapes.AddPicture(strExc(1), True, True, xlsRpt.CentimetersToPoints(13), xlsRpt.CentimetersToPoints(3.87 + 1.29 + Round(0.59 * (nRow - ((iPageNo - 1) * maxRows) - 6), 2) + (29.55 * (iPageNo - 1))), xlsRpt.CentimetersToPoints(5.6), xlsRpt.CentimetersToPoints(4.21))
         End If
      End If
   End If
   
   If strA0Y10 <> "" Then  '暫收款單號
      Call PrintExcel_BPage(2)
      If bol2Printer = True Or bol2File = True Then
         WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
         WksRpt1.Range(Chr(xCols) & nRow).Value = "P.S. : The credit amount shown in this receipt derives from your payment, which "
         Call PrintExcel_BPage
         WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
         WksRpt1.Range(Chr(xCols) & nRow).Value = "          can be deducted from your next payment if requested. However, if you prefer"
         Call PrintExcel_BPage
         WksRpt1.Range(Chr(xCols) & nRow).Font.Size = 14
         WksRpt1.Range(Chr(xCols) & nRow).Value = "          to return this credit, please advise us immediately."
      End If
   End If
   
'-------PrintCoverPage
   Dim stSQL As String, intR As Integer, intRow As Integer
   Dim adoTmp As ADODB.Recordset
   'Add by Amy 2018/10/31
   Dim strFA17 As String, strFA18 As String, strFA19 As String, strFA20 As String, strFA21 As String, strFA22 As String, strFA70 As String, strFA23 As String
   Dim strFA32 As String, strFA33 As String, strFA34 As String, strFA35 As String, strFA36 As String

   If bol2Printer = False Then Exit Sub
   If strA1K28 <> "Y51371000" Then Exit Sub
   stSQL = "select * from fagent where fa01='" & Left(strA1K28, 8) & "' and fa02='" & Mid(strA1K28, 9) & "'"
   intR = 1
   Set adoTmp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      Call PrintExcel_BFile(False, 0, 0)
      Call PrintExcel_BHead
      '去掉頁碼
      If bolChina = True Then
         WksRpt1.Range("F" & xRowE + 1).Value = ""
   
      Else
         WksRpt1.Range("D" & xRowE + 1).Value = ""
      End If

      '收款日期
      If Me.MaskEdBox1.Text = "___/__/__" Then
         WksRpt1.Range(IIf(bolChina = True, "K", "H") & nRow).Value = Format(AFDate(strSrvDate(1)), "mmm. d, yyyy")
      Else
         WksRpt1.Range(IIf(bolChina = True, "K", "H") & nRow).Value = Format(AFDate(ChangeTStringToWString(Replace(Me.MaskEdBox1.Text, "/", ""))), "mmm. d, yyyy")
      End If
      
      '代理人名稱
      nRow = nRow + 1
      If Trim("" & adoTmp.Fields("fa05")) = "" Then
         WksRpt1.Range(Chr(xCols) & nRow).Value = "" & adoTmp.Fields("fa04")
      Else
         WksRpt1.Range(Chr(xCols) & nRow).Value = Trim(adoTmp.Fields("fa05") & " " & adoTmp.Fields("fa63"))
         If Trim("" & adoTmp.Fields("fa64")) <> "" Then
            nRow = nRow + 1
             WksRpt1.Range(Chr(xCols) & nRow).Value = Trim(adoTmp.Fields("fa64") & " " & adoTmp.Fields("fa65"))
         End If
      End If
      
      '代理人地址
      'Add by Amy 2018/10/31 +地址有「竹曆退件」字樣不顯示地址
      strFA17 = "" & adoTmp.Fields("fa17").Value
      strFA18 = "" & adoTmp.Fields("fa18").Value: strFA19 = "" & adoTmp.Fields("fa19").Value: strFA20 = "" & adoTmp.Fields("fa20").Value
      strFA21 = "" & adoTmp.Fields("fa21").Value: strFA22 = "" & adoTmp.Fields("fa22").Value: strFA70 = "" & adoTmp.Fields("fa70").Value
      strFA23 = "" & adoTmp.Fields("fa23").Value
      strFA32 = "" & adoTmp.Fields("fa32").Value: strFA33 = "" & adoTmp.Fields("fa33").Value: strFA34 = "" & adoTmp.Fields("fa34").Value
      strFA35 = "" & adoTmp.Fields("fa35").Value: strFA36 = "" & adoTmp.Fields("fa36").Value
      
      If InStr(strFA17, "竹曆退件") > 0 Then strFA17 = ""
      If InStr(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70, "竹曆退件") > 0 Then
        strFA18 = "": strFA19 = "": strFA20 = "": strFA21 = "": strFA22 = "": strFA70 = ""
      End If
      If InStr(strFA23, "竹曆退件") > 0 Then strFA23 = ""
      If InStr(strFA32 & strFA33 & strFA34 & strFA35 & strFA36, "竹曆退件") > 0 Then
        strFA32 = "": strFA33 = "": strFA34 = "": strFA35 = "": strFA36 = ""
      End If
      'end 2018/10/31
      'POB1~POB5
      If Trim(strFA32 & strFA33 & strFA34 & strFA35 & strFA36) <> MsgText(601) Then
         If strFA32 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strFA32
            nRow = nRow + 1
         End If
         If strFA33 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strFA33
            nRow = nRow + 1
         End If
         If strFA34 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strFA34
            nRow = nRow + 1
         End If
         If strFA35 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strFA35
            nRow = nRow + 1
         End If
         If strFA36 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strFA36
            nRow = nRow + 1
         End If
      '英文地址1~6
      ElseIf Trim(strFA18 & strFA19 & strFA20 & strFA21 & strFA22 & strFA70) <> MsgText(601) Then
         If strFA18 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strFA18
            nRow = nRow + 1
         End If
         If strFA19 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strFA19
            nRow = nRow + 1
         End If
         If strFA20 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strFA20
            nRow = nRow + 1
         End If
         If strFA21 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strFA21
            nRow = nRow + 1
         End If
         If strFA22 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strFA22
            nRow = nRow + 1
         End If
         If strFA70 <> "" Then
            WksRpt1.Range(Chr(xCols) & nRow).Value = strFA70
            nRow = nRow + 1
         End If
      '中文地址
      ElseIf strFA17 <> "" Then
         WksRpt1.Range(Chr(xCols) & nRow).Value = strFA17
         nRow = nRow + 1
      '日文地址
      ElseIf strFA23 <> "" Then
         WksRpt1.Range(Chr(xCols) & nRow).Value = strFA23
         nRow = nRow + 1
      End If
      
      If nRow < 6 Then nRow = 6
      
      nRow = nRow + 3
      WksRpt1.Range(Chr(xCols) & nRow).Value = "Dear Sirs,"
      nRow = nRow + 2
      WksRpt1.Range(Chr(xCols) & nRow).Value = "Thank you for the payment in the amount of " & strCurrency & " " & strAmount & " on " & ChgEngDate(DBDATE(strRecDate)) & "."
      nRow = nRow + 2
      WksRpt1.Range(Chr(xCols) & nRow).Value = "Enclosed herewith please find the receipt, which is in the name of your client."
      nRow = nRow + 2
      WksRpt1.Range(Chr(xCols) & nRow).Value = "Best regards,"
      nRow = nRow + 5
      If bolChina = True Then
         strTmp = "I"
      Else
         strTmp = "F"
      End If
      '財務主管
      WksRpt1.Range(Chr(xCols) & nRow).Value = "Chloe Wu" '---婉莘
      WksRpt1.Range(Chr(xCols) & nRow + 1).Value = "Accounting  Department"
      '國外部主管
      WksRpt1.Range(strTmp & nRow).Value = "Fred C.T. Yen"
      WksRpt1.Range(strTmp & nRow + 1).Value = "Patent Attorney"
      WksRpt1.Range(strTmp & nRow + 2).Value = "Managing Partner"
   End If
   
   Set adoTmp = Nothing
   
End Sub


