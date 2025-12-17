VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Frmacc2490 
   AutoRedraw      =   -1  'True
   Caption         =   "國外帳款對帳單"
   ClientHeight    =   3360
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5388
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   5388
   Begin VB.Frame Frame2 
      Height          =   450
      Left            =   1320
      TabIndex        =   26
      Top             =   2400
      Width           =   3500
      Begin VB.OptionButton Option2 
         Caption         =   "最新"
         Height          =   300
         Index           =   0
         Left            =   72
         TabIndex        =   28
         Top             =   120
         Width           =   900
      End
      Begin VB.OptionButton Option2 
         Caption         =   "當下"
         Height          =   300
         Index           =   1
         Left            =   1050
         TabIndex        =   27
         Top             =   120
         Value           =   -1  'True
         Width           =   2000
      End
   End
   Begin VB.TextBox Text9 
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1080
      Width           =   612
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   1320
      TabIndex        =   20
      Top             =   3360
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Text4 
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
      Left            =   1320
      MaxLength       =   3
      TabIndex        =   2
      Top             =   360
      Width           =   492
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3405
      TabIndex        =   19
      Top             =   360
      Width           =   612
   End
   Begin VB.TextBox Text5 
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
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   3
      Top             =   360
      Width           =   852
   End
   Begin VB.TextBox Text6 
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
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text7 
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
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   5
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   1200
   End
   Begin VB.OptionButton Option1 
      Caption         =   "請款對象"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   60
      Value           =   -1  'True
      Width           =   1200
   End
   Begin VB.CommandButton CmdExcel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel"
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
      Left            =   330
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   2895
      Width           =   4692
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
      Left            =   330
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   4692
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
      Left            =   1320
      TabIndex        =   7
      Top             =   1410
      Width           =   612
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
      Left            =   1320
      TabIndex        =   0
      Top             =   30
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
      Left            =   3240
      TabIndex        =   1
      Top             =   30
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1320
      TabIndex        =   8
      Top             =   2100
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
      Left            =   3240
      TabIndex        =   9
      Top             =   2100
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
   Begin MSForms.TextBox Text8 
      Height          =   330
      Left            =   1320
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   690
      Width           =   4005
      VariousPropertyBits=   671105055
      BackColor       =   14737632
      MaxLength       =   50
      Size            =   "7056;582"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Lbl_Inf2 
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
      Height          =   270
      Left            =   4260
      TabIndex        =   29
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "資料狀態"
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
      Left            =   360
      TabIndex        =   25
      Top             =   2520
      Width           =   975
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
      Height          =   270
      Left            =   4830
      TabIndex        =   24
      Top             =   30
      Width           =   255
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "公司別             (1非智權 2.智權 空白.全部)"
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
      Left            =   360
      TabIndex        =   23
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "(可用A4)"
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
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   450
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   132
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
      Height          =   255
      Left            =   3000
      TabIndex        =   15
      Top             =   2100
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "往來日期"
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
      Left            =   360
      TabIndex        =   14
      Top             =   2100
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "(1.FC往來 2.FC未收         3.CF往來 4.CF未付         5.往來    6.未收未付)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2040
      TabIndex        =   13
      Top             =   1410
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "資料性質"
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
      Left            =   360
      TabIndex        =   12
      Top             =   1410
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
      Left            =   3000
      TabIndex        =   11
      Top             =   30
      Width           =   255
   End
End
Attribute VB_Name = "Frmacc2490"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/09 改成Form2.0 ; Text8
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/2 日期欄已修改
Option Explicit

Public adoquery As New ADODB.Recordset
Dim strSql As String
Dim strWhere(8) As String 'Modify by Amy 2015/06/30 原:6
Dim intCounter As Integer
Dim intRecord As Integer
Dim intPage As Integer
'Add by Morgan 2004/10/26
'Modify By Sindy 2012/12/21 變成陣列值
Const intMax As Integer = 19
Dim i As Integer, intUseI As Integer
Dim m_FCCur(20) As String, m_CFCur(20) As String, m_CFVCur(20) As String
Dim m_FCSum(20) As Double, m_CFSum(20) As Double, m_CFVSum(20) As Double
'2012/12/21 End
'Add By Sindy 2013/1/28
Dim m_j As Integer
Dim PLeft(0 To 9) As Integer
Dim m_i As Integer, iLine As Integer
Dim strTemp(1 To 7) As String
Dim strCurr(20) As String, dblTotFAmt(20) As Double
Dim dblTotNAmt As Double, dblTotFee As Double
'2013/1/28 End
Dim bolExcel As Boolean  'Add by Amy 2014/04/30 是否產生Excel
'Add by Amy 2014/11/18
Dim prnPrint As Printer
Dim strPrint As String
Dim bolSumTxt As Boolean '是否已印合計
Dim MaxLine As Integer '大表行數/A4行數
'Add by Amy 2015/06/30
Dim strWhere2(2) As String, dblSum(1 To 5) As Double
'Add by Amy 2020/09/17
Dim RsQ As New ADODB.Recordset

'Add by Amy 2014/04/30 +產生Excel
Private Sub cmdExcel_Click()
    bolExcel = True
    If FormCheck = False Then
        'MsgBox MsgText(181), , MsgText(5) Mark by Amy 2014/11/18
        Exit Sub
    End If
    'Add by Amy 2020/09/17
    If Option2(1).Value = True Then
        If MaskEdBox2 = MsgText(601) Then
            MaskEdBox2.Text = CFDate(strSrvDate(2))
            MaskEdBox2.Mask = DFormat
            Option2(1).Caption = "(" & MaskEdBox2 & ")"
        End If
    End If
    'end 2020/09/17
    Screen.MousePointer = vbHourglass
    PrintData
    FormClear
    Screen.MousePointer = vbDefault
    bolExcel = False
End Sub

Private Sub Combo1_Click()
    Option1(1).Value = True
    CaseQuery
End Sub

Private Sub Command2_Click()
   Call Text3_LostFocus 'Add By Sindy 2013/1/23
   If FormCheck = False Then
      'MsgBox MsgText(181), , MsgText(5) 'Mark by Amy 2014/11/18
      Exit Sub
   End If
   
   'Add by Amy 2014/11/18
   'Modify by Amy 2015/07/30 +可選PDF Creator & PDF reDirect v2
   For Each prnPrint In Printers
      If prnPrint.DeviceName = Combo2 Then
        If Option1(0).Value = True And Pub_StrUserSt03 <> "M51" Then
            If Combo2 <> "Ledomars 7800II" And Combo2 <> "IBM 5577-KC2" And Combo2 <> "PDFCreator" And Combo2 <> "PDF reDirect v2" Then
                MsgBox "印表機只能選 " & vbCrLf & "Ledomars 7800II 或 IBM 5577-KC2 或" & vbCrLf & _
                                "PDF Creator 或 PDF reDirect v2"
                Exit Sub
            End If
        End If
         Set Printer = prnPrint
      End If
   Next
   'end 2014/11/18
   Screen.MousePointer = vbHourglass
   PrintData
   FormClear   '2013/7/26 modify by sonia 原取消又瑞婷說要加回
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   'Mark by Amy 2020/09/17 不使用列印
'   If KeyCode <> vbKeyEscape Then
'      Frmacc0000.StatusBar1.Panels(1).Text = IIf(Option1(0) = True, MsgText(102), MsgText(101)) 'Modify by Amy 2015/06/30
'   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
'Add by Amy 2014/11/18
Dim rsA As New ADODB.Recordset
Dim ii As Integer
   
   'Modified by Lydia 2021/12/09 改成模組
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 5460
'   Me.Height = 3690 'Modify by Amy 2020/09/17 拿掉列印及印表機 原:4530 'Modify by Amy 2014/11/25
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
   PUB_InitForm Me, 5500, 3800, strBackPicPath4
   'end 2021/12/09
      
   'Add by Amy 2014/11/18 +印表機及本所案號
   PUB_SetPrinter Me.Name, Combo2, strPrint 'Modified by Morgan 2017/11/8 設定印表機改呼叫公用函數,原程式移除
       
   Text4.Enabled = False
   Text5.Enabled = False
   Text6.Enabled = False
   Text7.Enabled = False
   Combo1.AddItem ComboItem(121)
   Combo1.AddItem ComboItem(122)
   Combo1.AddItem ComboItem(123)
   Combo1 = ComboItem(121)
   'end 2014/11/18
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   'end 2020/09/17
   'Modify by Amy 2014/04/30 改產生Excel鈕 原列印抵帳明細拿掉Check1
   'Modify by Amy 2014/08/07 進入按鈕顯示但設不能按
   'CmdExcel.Visible = False
   'CmdExcel.Enabled = False 'Mark by Amy 2020/09/17 只產生Excel
   'Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)'Mark by Amy 2020/09/17不使用 列印
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2014/11/18
   Printer.DrawWidth = 1
   '若印表機變動, 則更新列印設定
   If Me.Combo2.Text <> Me.Combo2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo2.Name, "0", "0", Me.Combo2.Text
   End If
   'end 2014/11/18
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc2490 = Nothing
End Sub

'Add by Amy 2020/09/17
Private Sub Lbl_Inf_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     'Memo 因Excel Sheet 有限制，故只能下前6碼相同
     Lbl_Inf.ToolTipText = "不提供輸空白(不可查全部代理人)"
End Sub

Private Sub Lbl_Inf2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Lbl_Inf2.ToolTipText = "6.未收未付 抵帳專用"
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
    If Option2(1).Value = True And MaskEdBox2 <> MsgText(601) Then
        Option2(1).Caption = "當下 (" & MaskEdBox2.Text & ")"
    Else
        Option2(1).Caption = "當下"
    End If
End Sub
'end 2020/09/17

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Text4 = "": Text5 = "": Text6 = "": Text7 = "": Text8 = ""
        Combo1 = ComboItem(121)
        Text1.Enabled = True
        Text2.Enabled = True
        Text4.Enabled = False
        Text5.Enabled = False
        Text6.Enabled = False
        Text7.Enabled = False
        Text1.SetFocus
    Else
        Text1 = "": Text2 = ""
        Text4.Enabled = True
        Text5.Enabled = True
        Text6.Enabled = True
        Text7.Enabled = True
        Text1.Enabled = False
        Text2.Enabled = False
        Text4.SetFocus
    End If
    'Mark by Amy 2020/09/17 不使用列印
    'Frmacc0000.StatusBar1.Panels(1).Text = IIf(Index = 0, MsgText(102), MsgText(101)) 'Modify by Amy 2015/06/30 從Command2_Click搬過來
End Sub

'Add by Amy 2020/09/17
Private Sub Option2_Click(Index As Integer)
    If Option2(1).Value = True And MaskEdBox2 <> MsgText(601) Then
        Option2(1).Caption = "當下 (" & MaskEdBox2.Text & ")"
    Else
        Option2(1).Caption = "當下"
    End If
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus()
    'Add by Amy 2014/04/30 請款對象前6碼相同產生Excel鈕
    'Modify by Amy 2014/08/07
    'Mark by Amy 2020/09/17 只能產生Excel
'    If Text1 <> "" And Text2 <> "" And Left(Text1, 6) = Left(Text2, 6) and Text3.Text = "6"Then
'      'CmdExcel.Visible = True
'      CmdExcel.Enabled = True
'    Else
'       'CmdExcel.Visible = False
'       CmdExcel.Enabled = False
'    End If
    'end 2014/08/07
    'end 2014/04/30
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

Private Sub Text2_LostFocus()
   'Add by Amy 2014/04/30 請款對象前6碼相同產生Excel鈕
   'Modify by AMy 2014/08/07
   'Mark by Amy 2020/09/17 只能產生Excel
'   If Text1 <> "" And Text2 <> "" And Left(Text1, 6) = Left(Text2, 6) And Text3.Text = "6" Then
'      'CmdExcel.Visible = True
'      CmdExcel.Enabled = True
'   Else
'       'CmdExcel.Visible = False
'       CmdExcel.Enabled = False
'   End If
   'end 2014/04/30
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
   Text1 = ""
   Text2 = ""
   Text3 = ""
   MaskEdBox1.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = ""
   MaskEdBox2.Text = ""
   MaskEdBox2.Mask = DFormat
   'Modify by Amy 2014/04/30 改產生Excel鈕 原列印抵帳明細拿掉
   'Modify by Amy 2014/08/07
   'CmdExcel.Visible = False
   'CmdExcel.Enabled = False 'Mark by Amy 2020/09/17
   'end 2014/04/30
   'Add by Amy 2015/06/30
   Text4 = ""
   Text5 = ""
   Text6 = ""
   Text7 = ""
   Combo1 = ""
   'Add by Amy 2020/09/17
   Text9 = ""
   Option2(1).Caption = "當下"
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

'*************************************************
'  產生對帳資料
'  Amy 2015/06/30 重新整理
'*************************************************
Public Sub PrintData()
Dim strQ(3) As String
Dim strM As String, strP As String, strUpd As String, strDel As String '收款M/付款U's W/更新/刪除語法
Dim Rs As ADODB.Recordset
Dim bolIsHaveData As Boolean
Dim strAGField As String, strCaseField As String, strTField As String
Dim strCmp As String 'Add by Amy 2015/11/25 公司別(特殊出名公司)
   
On Error GoTo Checking
   
    strSql = ""
    '清空陣列值
    For i = 0 To intMax
        m_FCCur(i) = "": m_CFCur(i) = "": m_CFVCur(i) = ""
        m_FCSum(i) = 0: m_CFSum(i) = 0: m_CFVSum(i) = 0
    Next i
   
    For intCounter = 0 To UBound(strWhere)
        strWhere(intCounter) = ""
        If intCounter < 2 Then
            strQ(intCounter) = ""
        End If
    Next intCounter
    If adoquery.State = adStateOpen Then
        adoquery.Close
    End If
    adoquery.CursorLocation = adUseClient
   
    strAGField = ""
    '請款對象
    If Option1(0).Value = True Then
        strAGField = " a1k28 " '請款對象
        strCaseField = " '' " 'Acc0y0抓案件編號
        strTField = " 0 " 'Acc0y0抓規費
        If Text1 <> "" Then
            strWhere(0) = strWhere(0) & " And a1k28 >= '" & Text1 & "'"
            'strWhere(1) = strWhere(1) & " And a1203 >= '" & Text1 & "'" '秀玲說Mark
            strWhere(2) = strWhere(2) & " And a1503 >= '" & Text1 & "'"
            strWhere(3) = strWhere(3) & " And a1603 >= '" & Text1 & "'"
            strWhere(4) = strWhere(4) & " And Decode(a0y18, '1', a0y07, '2', a0y08, a0y09) >= '" & Text1 & "'"
            strWhere(5) = strWhere(5) & " And a1k28 >= '" & Text1 & "'"
            strWhere(6) = strWhere(6) & " And a1503 >= '" & Text1 & "'"
            strWhere(7) = strWhere(7) & " And a1503 >= '" & Text1 & "'"
            strWhere(8) = strWhere(8) & " And a1603 >= '" & Text1 & "'"
            strWhere2(0) = strWhere2(0) & " And a1k28 >= '" & Text1 & "'"
            strWhere2(1) = strWhere2(1) & " And a1503 >= '" & Text1 & "'"
            strWhere2(2) = strWhere2(2) & " And a1603 >= '" & Text1 & "'"
        End If
        If Text2 <> "" Then
            strWhere(0) = strWhere(0) & " And a1k28 <= '" & Text2 & "'"
            'strWhere(1) = strWhere(1) & " And a1203 <= '" & Text2 & "'" '秀玲說Mark
            strWhere(2) = strWhere(2) & " And a1503 <= '" & Text2 & "'"
            strWhere(3) = strWhere(3) & " And a1603 <= '" & Text2 & "'"
            strWhere(4) = strWhere(4) & " And Decode(a0y18, '1', a0y07, '2', a0y08, a0y09) <= '" & Text2 & "'"
            strWhere(5) = strWhere(5) & " And a1k28 <= '" & Text2 & "'"
            strWhere(6) = strWhere(6) & " And a1503 <= '" & Text2 & "'"
            strWhere(7) = strWhere(7) & " And a1503 <= '" & Text2 & "'"
            strWhere(8) = strWhere(8) & " And a1603 <= '" & Text2 & "'"
            strWhere2(0) = strWhere2(0) & " And a1k28 <= '" & Text2 & "'"
            strWhere2(1) = strWhere2(1) & " And a1503 <= '" & Text2 & "'"
            strWhere2(2) = strWhere2(2) & " And a1603 <= '" & Text2 & "'"
        End If
    End If
    '本所案號
    If Option1(1).Value = True Then
        strAGField = " a1k03 " '本所案號抓代理人
        strCaseField = " a1k03||a1k14||a1k15||a1k16 " 'Acc0y0抓案件編號
        strTField = " Decode(nvl(a1k30,0),0,0,Decode(sign(a1k30-a1k09),-1,a1k30,a1k09)) " 'Acc0y0抓規費編號
        If Trim(Text6) = MsgText(601) Then Text6 = "0"
        If Trim(Text7) = MsgText(601) Then Text7 = "00"
        
        strWhere(0) = strWhere(0) & " And a1k13 = '" & Text4 & "' And a1k14='" & Text5 & "' And a1k15='" & Text6 & "' And a1k16='" & Text7 & "' "
        'strWhere(1) = strWhere(1) & " And a1208 = '" & Text4 & Text5 & Text6 & Text7 & "' " '秀玲說Mark
        strWhere(2) = strWhere(2) & " And axf03 = '" & Text4 & Text5 & Text6 & Text7 & "' "
        strWhere(3) = strWhere(3) & " And axg03 = '" & Text4 & Text5 & Text6 & Text7 & "' "
        strWhere(4) = strWhere(4) & " And a1k13 = '" & Text4 & "' And a1k14='" & Text5 & "' And a1k15='" & Text6 & "' And a1k16='" & Text7 & "' "
        strWhere(5) = strWhere(5) & " And a1k13 = '" & Text4 & "' And a1k14='" & Text5 & "' And a1k15='" & Text6 & "' And a1k16='" & Text7 & "' "
        strWhere(6) = strWhere(6) & " And axf03 = '" & Text4 & Text5 & Text6 & Text7 & "' "
        strWhere(7) = strWhere(7) & " And axf03 = '" & Text4 & Text5 & Text6 & Text7 & "' "
        strWhere(8) = strWhere(8) & " And axg03 = '" & Text4 & Text5 & Text6 & Text7 & "' "
        strWhere2(0) = strWhere2(0) & " And a1k13 = '" & Text4 & "' And a1k14='" & Text5 & "' And a1k15='" & Text6 & "' And a1k16='" & Text7 & "' "
        strWhere2(1) = strWhere2(1) & " And axf03 = '" & Text4 & Text5 & Text6 & Text7 & "' "
        strWhere2(2) = strWhere2(2) & " And axg03 = '" & Text4 & Text5 & Text6 & Text7 & "' "
    End If
   
    If MaskEdBox1.Text <> MsgText(29) And MaskEdBox1.Text <> "" Then
        strWhere(0) = strWhere(0) & " And a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
        'strWhere(1) = strWhere(1) & " And a1202 >= " & Val(FCDate(MaskEdBox1.Text)) & "" '秀玲說Mark
        strWhere(2) = strWhere(2) & " And a1502 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
        strWhere(3) = strWhere(3) & " And a1602 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
        strWhere(4) = strWhere(4) & " And a0y02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
        strWhere(5) = strWhere(5) & " And Decode(a1h02, null, a1i03, a1h02) >= " & Val(FCDate(MaskEdBox1.Text))
        strWhere(6) = strWhere(6) & " And a1b03 >= " & Val(FCDate(MaskEdBox1.Text))
        strWhere(7) = strWhere(7) & " And Decode(a1h02, null, a1i03, a1h02) >= " & Val(FCDate(MaskEdBox1.Text))
        strWhere(8) = strWhere(8) & " And a1b03 >= " & Val(FCDate(MaskEdBox1.Text))
    End If
    If MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> "" Then
        strWhere(0) = strWhere(0) & " And a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
        'strWhere(1) = strWhere(1) & " And a1202 <= " & Val(FCDate(MaskEdBox2.Text)) & "" '秀玲說Mark
        strWhere(2) = strWhere(2) & " And a1502 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
        strWhere(3) = strWhere(3) & " And a1602 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
        strWhere(4) = strWhere(4) & " And a0y02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
        strWhere(5) = strWhere(5) & " And Decode(a1h02, null, a1i03, a1h02) <= " & Val(FCDate(MaskEdBox2.Text))
        strWhere(6) = strWhere(6) & " And a1b03 <= " & Val(FCDate(MaskEdBox2.Text))
        strWhere(7) = strWhere(7) & " And Decode(a1h02, null, a1i03, a1h02) <= " & Val(FCDate(MaskEdBox2.Text))
        strWhere(8) = strWhere(8) & " And a1b03 <= " & Val(FCDate(MaskEdBox2.Text))
    End If
    
    'Modify by Amy 2020/09/17 都產生Excel,加收據公司別
    'If bolExcel = True Then
    '最新
    If Option2(0).Value = True Then
        '帳款抵帳明細
        'Modify by Amy 2015/09/14 將fa70欄位拿來存CP45
        'Add by Amy 2015/11/25 + if
        'Modify by Amy 2020/09/17 +2.FC最新
        If Text3 = "2" Or Text3 = "6" Then
            '未結清請款X
            strSql = "Select " & strAGField & " as FagentNo, a1k02 as DocDate, Decode(a1k12, null, Decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (Decode(a1k30,0,NULL,A1K30) - (a1k11 - nvl(a1k06, 0))) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 as ODocNo,'' as CP45,a1k29,a1k27,a1k28 From acc1k0, Fagent Where a1k12 is null And a1k25 is null And SubStr(" & strAGField & ",1, 8) = fa01 (+) And SubStr(" & strAGField & ", 9, 1) = fa02 (+)" & strWhere(0) & " And (a1k29 is null or a1k29 = '')"
            '部分付款M(列於X減項)
            strSql = strSql & " Union all Select Distinct Decode(a0y18, '1', a0y07, '2', a0y08, a0y09) as FagentNo, a0y02 as DocDate, a0y01 as DocNo, a0y03 as Currency, a0z04*(-1) as Famount, 0 as Namount, " & strTField & " as Tamount, ((-1) * a0z04 * Decode(a0y03, 'NT$', 1, a0y04))  as Ramount, 0 as Pamount, 0 as Vamount, a1k13||a1k14||a1k15||a1k16 as caseno, ((a0z04 * Decode(a0y03, 'NT$', 1, a0y04)) / 1000) as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 as ODocNo,'' as CP45,'' as CloseS,'' as a1k27,'' as a1k28 From acc0y0, Fagent, nation, acc0z0, acc1k0, patent Where SubStr(a0y07, 1, 8) = fa01 (+) And SubStr(a0y07, 9, 1) = fa02 (+) And fa10 = na01 (+) And a0y01 = a0z01 And a0z02 = a1k01 And a1k13 = pa01 (+) And a1k14 = pa02 (+) And a1k15 = pa03 (+) And a1k16 = pa04 (+) " & strWhere(0) & " And (a1k29 is null or a1k29 = '')"
            If Text3 = "6" Then strSql = strSql & " Union all "
        End If
        'end 2015/11/25
        
        '資料性質 4-CF未付, 只抓 未結清帳單U及未付款抵帳V-婉莘
        'Add by Amy 2020/09/17 +if
        If Text3 = "4" Or Text3 = "6" Then
            '未結清帳單U
            strSql = strSql & "Select a1503 as FagentNo, a1502 as DocDate, Decode(a1507, null, a1501, a1501||'*') as DocNo, a1505 as Currency, axf04 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, Decode(a1c01, null, 0, axf04 * a1906) as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, '' as ODocNo,CP45,'' as a1k29,'' as a1k27,'' as a1k28 From acc151, acc150, acc1c0, acc190, Fagent,CaseProgress Where a1507 is null And SubStr(a1503, 1, 8) = fa01 (+) And SubStr(a1503, 9, 1) = fa02 (+) And axf01 = a1501 And axf01 = a1c03 (+) And axf01 = a1902 (+)" & strWhere(2) & " And (a1520 is null or a1520 = 0) And axf02=CP09(+)"
            '未付款抵帳V
            strSql = strSql & " Union all Select Distinct a1603 as FagentNo, a1602 as DocDate, a1601 as DocNo, a1605 as Currency, axg04 as Famount, 0 as Namount, 0 as Tamount, Decode(a1c01, null, 0, axg04 * a1906) as Ramount, 0 as Pamount, 0 as Vamount, axg03 as Caseno, 0 as Point, a1604 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1601 as ODocNo,CP45,'' as a1k29,'' as a1k27,'' as a1k28 From acc161, acc160, acc1c0, acc190, Fagent, acc170,CaseProgress Where A1607 is null And SubStr(a1603, 1, 8) = fa01 (+) And SubStr(a1603, 9, 1) = fa02 (+) And axg01 = a1601 And axg01 = a1c03 (+) And axg01 = a1902 (+) And a1901 is null " & strWhere(3) & " And a1702(+)=a1601 And a1701(+)='2' And a1701 is null And axg02=CP09(+)"
        End If
        
        'end 2015/09/14
    Else
        '--選「往來」
        '請款X
        strM = "Select " & strAGField & " as FagentNo, a1k02 as DocDate, Decode(a1k12, null, Decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (Decode(a1k30,0,NULL,A1K30) - (a1k11 - nvl(a1k06, 0))) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 as ODocNo,fa70,a1k29 as CloseS,a1k27,a1k28 From acc1k0, Fagent Where a1k12 is null And a1k25 is null And SubStr(" & strAGField & ",1, 8) = fa01 (+) And SubStr(" & strAGField & ", 9, 1) = fa02 (+) " & strWhere(0)
        '請款收款X's M 在往來期間
        strM = strM & " Union Select Distinct Decode(a0y18, '1', a0y07, '2', a0y08, a0y09) as FagentNo, a0y02 as DocDate, a0y01 as DocNo, a0y03 as Currency, a0z04 as Famount, 0 as Namount, " & strTField & " as Tamount, (a0z04 * Decode(a0y03, 'NT$', 1, a0y04))  as Ramount, 0 as Pamount, 0 as Vamount, a1k13||a1k14||a1k15||a1k16 as caseno, ((a0z04 * Decode(a0y03, 'NT$', 1, a0y04)) / 1000) as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 as ODocNo,fa70,'' as CloseS,'' as a1k27,'' as a1k28 From acc0y0, Fagent, nation, acc0z0, acc1k0, patent Where SubStr(a0y07, 1, 8) = fa01 (+) And SubStr(a0y07, 9, 1) = fa02 (+) And fa10 = na01 (+) And a0y01 = a0z01 And a0z02 = a1k01 And a1k13 = pa01 (+) And a1k14 = pa02 (+) And a1k15 = pa03 (+) And a1k16 = pa04 (+)" & strWhere(4)
        '請款收款X's M 在往來期間,但X不在期間內也要顯示(M's X )
        strM = strM & " Union Select " & strAGField & " as FagentNo, a1k02 as DocDate, Decode(a1k12, null, Decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (Decode(a1k30,0,NULL,A1K30) - (a1k11 - nvl(a1k06, 0))) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 as ODocNO,fa70,a1k29 as CloseS,a1k27,a1k28 From acc0y0, Fagent, nation, acc0z0, acc1k0, patent Where SubStr(a0y07, 1, 8) = fa01 (+) And SubStr(a0y07, 9, 1) = fa02 (+) And fa10 = na01 (+) And a0y01 = a0z01 And a0z02 = a1k01 And a1k13 = pa01 (+) And a1k14 = pa02 (+) And a1k15 = pa03 (+) And a1k16 = pa04 (+)" & strWhere(4)
        '請款抵帳X's Z 在往來期間
        strQ(0) = strQ(0) & " Union  Select " & strAGField & " as FagentNo, Decode(a1h02, null, a1i03, a1h02) as DocDate, a1g01 as DocNo, NVL(A1H03,A1I05) as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k08 - nvl(a1k31, 0)) * nvl(a1g02, 0) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (Decode(a1k30,0,NULL,A1K30) - (a1k11 - nvl(a1k06, 0))) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 as ODocNo,fa70,'' as CloseS,a1k27,a1k28 From acc1k0, Fagent, acc140, acc1g0, acc1h0, acc1i0 Where a1k12 is null And a1k25 is null And SubStr(" & strAGField & ",1, 8) = fa01 (+) And SubStr(" & strAGField & ", 9, 1) = fa02 (+) And a1k01 = a1403 (+) And a1k17 = a1g01 (+) And a1k17 = a1h01 (+) And a1k17 = a1i01 (+) And a1k17 is not null" & strWhere(5)
        '請款抵帳X's Z 在往來期間,但X不在期間內也要顯示(Z's X )
        strQ(0) = strQ(0) & " Union Select " & strAGField & " as FagentNo, a1k02 as DocDate, Decode(a1k12, null, Decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (Decode(a1k30,0,NULL,A1K30) - (a1k11 - nvl(a1k06, 0))) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 as ODocNO,fa70,a1k29 as CloseS,a1k27,a1k28 From acc1k0, Fagent, acc140, acc1g0, acc1h0, acc1i0 Where a1k12 is null And a1k25 is null And SubStr(" & strAGField & ",1, 8) = fa01 (+) And SubStr(" & strAGField & ", 9, 1) = fa02 (+) And a1k01 = a1403 (+) And a1k17 = a1g01 (+) And a1k17 = a1h01 (+) And a1k17 = a1i01 (+) And a1k17 is not null" & strWhere(5)
        
        '帳單U
        strP = "Select a1503 as FagentNo, a1502 as DocDate, Decode(a1507, null, a1501, a1501||'*') as DocNo, a1505 as Currency, axf04 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, 0 as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1501 as ODocNo,fa70,Decode(a1506, a1520, 'Y', '') as CloseS,'' as a1k27,'' as a1k28 From acc151, acc150, acc1c0, acc190, Fagent Where A1507 IS NULL And SubStr(a1503, 1, 8) = fa01 (+) And SubStr(a1503, 9, 1) = fa02 (+) And axf01 = a1501 And axf01 = a1c03 (+) And axf01 = a1902 (+)" & strWhere(2)
        '帳單付款U's W 在往來期間
        strP = strP & " Union Select a1503 as FagentNo, a1b03 as DocDate, a1901 as DocNo, a1903 as Currency, axf04 as Famount, axf04*a1906 as Namount, 0 as Tamount, 0 as Ramount, Decode(a1c01, null, 0, axf04 * a1906) as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1501 as ODocNo,fa70,Decode(a1506, a1520, 'Y', '') as CloseS,'' as a1k27,'' as a1k28 From acc151, acc150, acc1c0, acc1b0, acc190, Fagent Where SubStr(a1503, 1, 8) = fa01 (+) And SubStr(a1503, 9, 1) = fa02 (+) And a1902=a1501 And a1501=axf01(+) And a1908=a1b01 And a1c01 = a1b01 (+) And a1c02 = a1b02 (+) And a1501=a1c03(+) And axf01 = a1902 (+) And a1908 is not null" & strWhere(6)
        '帳單付款U's W 在往來期間,但U不在期間內也要顯示(W's U )
        strP = strP & " Union Select a1503 as FagentNo, a1502 as DocDate, Decode(a1507, null, a1501, a1501||'*') as DocNo, a1505 as Currency, axf04 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, 0 as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1501 as ODocNo,fa70,Decode(a1506, a1520, 'Y', '') as CloseS,'' as a1k27,'' as a1k28 From acc151, acc150, acc1c0, acc1b0, acc190, Fagent Where SubStr(a1503, 1, 8) = fa01 (+) And SubStr(a1503, 9, 1) = fa02 (+) And a1902=a1501 And a1501=axf01(+) And a1908=a1b01 And a1c01 = a1b01 (+) And a1c02 = a1b02 (+) And a1501=a1c03(+) And axf01 = a1902 (+) And a1908 is not null" & strWhere(6)
        '帳單抵帳U's Z 在往來期間
        strQ(1) = strQ(1) & " Union Select a1503 as FagentNo, Decode(a1h02, null, a1i03, a1h02) as DocDate, a1512 as DocNo, a1505 as Currency, axf04 as Famount, a1506 * nvl(a1g03, 0) as Namount, 0 as Tamount, 0 as Ramount, 0 as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1501 as ODocNo,fa70,'' as CloseS,'' as a1k27,'' as a1k28 From acc151, acc150, acc1g0, acc1h0, acc1i0, Fagent Where SubStr(a1503, 1, 8) = fa01 (+) And SubStr(a1503, 9, 1) = fa02 (+) And axf01 = a1501 And a1512=a1g01(+) And a1512 = a1h01 (+) And a1512 = a1i01(+) And a1512 is not null" & strWhere(7)
        '帳單付款U's Z 在往來期間,但U不在期間內也要顯示(Z's U )
        strQ(1) = strQ(1) & " Union Select a1503 as FagentNo, a1502 as DocDate, Decode(a1507, null, a1501, a1501||'*') as DocNo, a1505 as Currency, axf04 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, 0 as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1501 as ODocNo,fa70,Decode(a1506, a1520, 'Y', '') as CloseS,'' as a1k27,'' as a1k28 From acc151, acc150, acc1g0, acc1h0, acc1i0, Fagent Where SubStr(a1503, 1, 8) = fa01 (+) And SubStr(a1503, 9, 1) = fa02 (+) And axf01 = a1501 And a1512=a1g01(+) And a1512 = a1h01 (+) And a1512 = a1i01(+) And a1512 is not null" & strWhere(7)
        '抵帳V
        strQ(1) = strQ(1) & " Union Select Distinct a1603 as FagentNo, a1602 as DocDate, a1601 as DocNo, a1605 as Currency, axg04 as Famount, 0 as Namount, 0 as Tamount, Decode(a1c01, null, 0, axg04 * a1906) as Ramount, 0 as Pamount, 0 as Vamount, axg03 as Caseno, 0 as Point, a1604 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1601 as ODocNo,fa70,'' as CloseS,'' as a1k27,'' as a1k28 From acc161, acc160, acc1c0, acc190, Fagent Where SubStr(a1603, 1, 8) = fa01 (+) And SubStr(a1603, 9, 1) = fa02 (+) And axg01 = a1601 And axg01 = a1c03 (+) And axg01 = a1902 (+)" & strWhere(3)
        '減少支出V's W 在往來期間
        strQ(1) = strQ(1) & " Union Select a1603 as FagentNo, a1b03 as DocDate, a1901 as DocNo, a1605 as Currency, axg04 * (-1) as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, Decode(a1c01, null, 0, axg04 * a1906 *(-1)) as Pamount, 0 as Vamount, axg03 as Caseno, 0 as Point, a1604 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1601 as ODocNo,fa70,'' as CloseS,'' as a1k27,'' as a1k28 From acc161, acc160, acc1c0, acc1b0, acc190, Fagent Where SubStr(a1603, 1, 8) = fa01 (+) And SubStr(a1603, 9, 1) = fa02 (+) And axg01 = a1601 And a1c01 = a1b01 (+) And a1c02 = a1b02 (+) And a1902=a1c03(+) And axg01 = a1902 (+) And a1901 is not null" & strWhere(8)
        '減少支出V's W 在往來期間,但U不在期間內也要顯示(W's V )
        strQ(1) = strQ(1) & " Union Select Distinct a1603 as FagentNo, a1602 as DocDate, a1601 as DocNo, a1605 as Currency, axg04 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, 0 as Pamount, 0 as Vamount, axg03 as Caseno, 0 as Point, a1604 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1601 as ODocNo,fa70,'' as CloseS,'' as a1k27,'' as a1k28 From acc161, acc160, acc1c0, acc1b0, acc190, Fagent Where SubStr(a1603, 1, 8) = fa01 (+) And SubStr(a1603, 9, 1) = fa02 (+) And axg01 = a1601 And a1c01 = a1b01 (+) And a1c02 = a1b02 (+) And a1902=a1c03(+) And axg01 = a1902 (+) And a1901 is not null" & strWhere(8)
        '暫收款N -秀玲說Mark
        'strSql = strSql & " Union all Select Distinct a1203 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1204 as Currency, a1207 as Famount, a1207 * a1205 as Namount, 0 as Tamount, a1207 * a1205 as Ramount, 0 as Pamount, 0 as Vamount, a1208 as Caseno, (a1207 * a1205 / 1000) as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, '' as ODocNo,fa70,'' as CloseS,'' as a1k27,'' as a1k28 From acc120, Fagent Where SubStr(a1203, 1, 8) = fa01 (+) And SubStr(a1203, 9, 1) = fa02 (+)" & strWhere(1)
        
        Select Case Text3
            Case "1" 'FC往來
                strSql = strM & strQ(0)
            Case "2" 'FC 未收
                strSql = strM
            Case "3" 'CF往來
                strSql = strP & strQ(1)
            Case "4" 'CF未付
                strSql = strP
            Case "6" '未收/未付
                strSql = strM & " Union " & strP
            Case Else '往來
                strSql = strM & strQ(0) & " Union " & strP & strQ(1)
        End Select
    End If
    'end 2020/09/17
    
    strSql = "Select '" & strUserNum & "',X.*,0,Decode(substr(ODocNo,1,1), 'X',1,'U',2)||ODocNo|| Decode(substr(DOCNO,1,1),'X',1,'U',2,3) sort From (" & strSql & ") X "

    cnnConnection.Execute "Delete From AccRpt432 Where ID='" & strUserNum & "'"
    'Modify by Amy 2020/09/17 +收據公司別 R43235,於後面才更新,故顯示暫存檔欄位,否則會錯
    cnnConnection.Execute "Insert Into AccRpt432 (ID,R43201,R43202,R43203,R43204,R43205,R43206,R43207,R43208,R43209,R43210," & _
                                                                                "R43211,R43212,R43213,R43214,R43215,R43216,R43217,R43218,R43219,R43220," & _
                                                                                "R43221,R43222,R43223,R43224,R43225,R43226,R43227,R43228,R43229,R43230," & _
                                                                                "R43231,R43232,R43233,R43234) " & strSql
    'Add by Amy 2020/09/17 公司別 原:1.台一 2.智權,原抓出名公司抵帳用(資料性質 4/6 Excel 用),改抓收據公司別
    '***寫入W/U/V單號收據公司別欄位
    'U單號公司別(U10908160 有兩筆,故使用Distinct)
    'Modify by Amy 2020/12/21 條件下:Y52269 3.CF往來 1090915會錯
    '因U10906893 有兩筆收文資料AA8054363(收據: E10829641-1公司)及AA9030090(收據: E10916579-2公司)-莘:抓E單號較大,秀玲:抓a0k02日期最大之E單號
    'strUpd = "Update Accrpt432 Set r43235=(Select Distinct GetA0k11(axf02) From acc151 Where r43203=axf01(+) ) " & _
                    "Where ID='" & strUserNum & "' And SubStr(R43203,1,1)='U' "
    'Modify by Amy 2022/11/22 條件下:Y20284 111/10/31 6.未收未付 選最新 總收文號BB1024723(未結清請款的 U11107177 因為多國案串不到Acc0j0及 Acc0k0,莘說多國案之請款單會掛於母案,導致GetA0k11(空白)回傳為2),公司別抓錯而顯示出來(違反台一與智權不可互抵)
    '調整成AXF02串ACC0J0,串得到的走新語法,串不到的走原語法-秀玲
'    strUpd = "Update Accrpt432 Set r43235=(Select Distinct GetA0k11(SubStr(Max(a0k02||'@'||axf02),InStr(Max(a0k02||'@'||axf02),'@')+1)) From acc151,acc0j0,acc0k0 Where r43203=axf01 And axf02=a0j01 And a0j13=a0k01) " & _
'                    "Where ID='" & strUserNum & "' And SubStr(R43203,1,1)='U' "
    strUpd = "Update Accrpt432 Set r43235=(Select Distinct GetA0k11(axf02) From acc151 Where r43203=axf01(+) And axf01 is null) " & _
                    "Where ID='" & strUserNum & "' And SubStr(R43203,1,1)='U' "
     cnnConnection.Execute strUpd
     strUpd = "Update Accrpt432 Set r43235=(Select Distinct GetA0k11(SubStr(Max(a0k02||'@'||axf02),InStr(Max(a0k02||'@'||axf02),'@')+1)) From acc151,acc0j0,acc0k0 Where r43203=axf01(+) And axf01 is not null And axf02=a0j01(+) And a0j13=a0k01(+) ) " & _
                    "Where ID='" & strUserNum & "' And SubStr(R43203,1,1)='U' "
    cnnConnection.Execute strUpd
    'end 2022/11/22
    
    'V單號公司別
    strUpd = "Update Accrpt432 Set r43235=(Select Distinct GetA0k11(SubStr(Max(a0k02||'@'||axg02),InStr(Max(a0k02||'@'||axg02),'@')+1)) From acc161,acc0j0,acc0k0 Where r43203=axg01 And axg02=a0j01 And a0j13=a0k01) " & _
                    "Where ID='" & strUserNum & "' And SubStr(R43203,1,1)='V' "
    'end 2020/12/21
    cnnConnection.Execute strUpd
    
    'W單號公司別
    strUpd = "Update Accrpt432 Set r43235=(Select a1917 From acc190 Where r43203=a1901(+) And r43228=a1902(+)) " & _
                    "Where ID='" & strUserNum & "' And SubStr(R43203,1,1)='W' "
    cnnConnection.Execute strUpd
    '*** end 收據公司別
    
    If Trim(Text9) <> MsgText(601) Or Text3 = "6" Then
        strCmp = " 'J' "
        
        '產生6.未收未付 且公司別為空白或1.台一
        If Text3 = "6" And (Text9 = "" Or Text9 = "1") Then
            '排除J
            strCmp = "=" & strCmp
        '非J
        ElseIf Text9 = "1" Then
            '排除J
            strCmp = "=" & strCmp
        'J
        Else
            '排除不是 J
            strCmp = "<>" & strCmp
        End If
        strDel = "Delete Accrpt432 Where ID='" & strUserNum & "' And R43235" & strCmp
        cnnConnection.Execute strDel
        'J公司排除 X/M/Z單號 (X/M/Z單號 一律非J-秀玲)
        If Text9 = "2" Then
            strDel = "Delete Accrpt432 Where ID='" & strUserNum & "' And (SubStr(R43203,1,1)='X' or SubStr(R43203,1,1)='M' or SubStr(R43203,1,1)='Z' )"
            cnnConnection.Execute strDel
        End If
    End If
    'end 2020/09/17
    'Modify by Amy 2020/09/17 只能產生Excel,2.FC未收也要可以產生最新資料
'    If bolExcel = True Then
'        'Modify by Amy 2015/11/26 +資料性質 4-FC未收(代理人付款前提供台一或者智權付款明細供其核對)
'        If Text3 = "6" Then
'            '代理人抵帳明細排除 J公司
'            strCmp = "='J' "
'        Else
'            Select Case Text9
'                Case "1"
'                    strCmp = "='J' "
'                Case "2"
'                    strCmp = "<>'J' "
'                Case Else
'                    strCmp = ""
'            End Select
'        End If
'        If strCmp <> MsgText(601) Then
'            '刪除 特殊出名公司
'            strDel = "Delete From accrpt432 Where ID='" & strUserNum & "' And R43203 in (" & _
'                            "Select R43203 From Patent,accrpt432 Where ID='" & strUserNum & "' And Decode(length(R43211),10,substr(R43211,1,1),11,substr(R43211,1,2),12,substr(R43211,1,3),R43211)=PA01 And Decode(length(R43211),10, substr(R43211,2,6),11, substr(R43211,3,6),12, substr(R43211,4,6),R43211)=PA02 And Decode(length(R43211),10, substr(R43211,8,1),11, substr(R43211,9,1),12, substr(R43211,10,1),R43211)=PA03 And Decode(length(R43211),10, substr(R43211,9,2),11, substr(R43211,10,2),12, substr(R43211,11,2),R43211)=PA04 And Nvl(PA161,'1')" & strCmp & _
'                 "Union Select R43203 From TradeMark,accrpt432 Where ID='" & strUserNum & "' And Decode(length(R43211),10,substr(R43211,1,1),11,substr(R43211,1,2),12,substr(R43211,1,3),R43211)=TM01 And Decode(length(R43211),10, substr(R43211,2,6),11, substr(R43211,3,6),12, substr(R43211,4,6),R43211)=TM02 And Decode(length(R43211),10, substr(R43211,8,1),11, substr(R43211,9,1),12, substr(R43211,10,1),R43211)=TM03 And Decode(length(R43211),10, substr(R43211,9,2),11, substr(R43211,10,2),12, substr(R43211,11,2),R43211)=TM04 And Nvl(TM130,'1')" & strCmp & _
'                 "Union Select R43203 From LawCase,accrpt432 Where ID='" & strUserNum & "' And Decode(length(R43211),10,substr(R43211,1,1),11,substr(R43211,1,2),12,substr(R43211,1,3),R43211)=LC01 And Decode(length(R43211),10, substr(R43211,2,6),11, substr(R43211,3,6),12, substr(R43211,4,6),R43211)=LC02 And Decode(length(R43211),10, substr(R43211,8,1),11, substr(R43211,9,1),12, substr(R43211,10,1),R43211)=LC03 And Decode(length(R43211),10, substr(R43211,9,2),11, substr(R43211,10,2),12, substr(R43211,11,2),R43211)=LC04 And Nvl(LC48,'1')" & strCmp & _
'                 "Union Select R43203 From ServicePractice,accrpt432 Where ID='" & strUserNum & "' And Decode(length(R43211),10,substr(R43211,1,1),11,substr(R43211,1,2),12,substr(R43211,1,3),R43211)= SP01 And Decode(length(R43211),10, substr(R43211,2,6),11, substr(R43211,3,6),12, substr(R43211,4,6),R43211)=SP02 And Decode(length(R43211),10, substr(R43211,8,1),11, substr(R43211,9,1),12, substr(R43211,10,1),R43211)=SP03 And Decode(length(R43211),10, substr(R43211,9,2),11, substr(R43211,10,2),12, substr(R43211,11,2),R43211)=SP04 And Nvl(SP85,'1')" & strCmp & " )"
'            cnnConnection.Execute strDel
'        End If
'    ElseIf (Val(Text3) = 2 Or Val(Text3) = 6) Then
    If Val(Text3) = 2 Or Val(Text3) = 6 Then
    'end 2020/09/17
       '刪除已結清(已全收)
       strDel = "Delete From accrpt432 Where ID='" & strUserNum & "' And SubStr(R43234,2,9) in " & _
                    "(Select R43203 From " & _
                    "(Select R43203,R43208 From accrpt432 Where ID='" & strUserNum & "' And Substr(R43203,1,1)='X') x," & _
                    "(Select R43228,Sum(R43208) as R43208 From accrpt432 Where ID='" & strUserNum & "' And Substr(R43203,1,1)='M' Group by R43228) m " & _
                    "Where R43203=R43228 And x.R43208-m.R43208=0)"
        cnnConnection.Execute strDel
    End If
    If Val(Text3) = 4 Or Val(Text3) = 6 Then
        '刪除已結清(已全付)
        strDel = "Delete From accrpt432 Where ID='" & strUserNum & "' And SubStr(R43234,2,9) in " & _
                    "(Select R43203 From " & _
                    "(Select R43203,R43208 From accrpt432 Where ID='" & strUserNum & "' And Substr(R43203,1,1)='U') u," & _
                    "(Select R43228,R43208 From accrpt432 Where ID='" & strUserNum & "' And Substr(R43203,1,1)='W') w " & _
                    "Where R43203=R43228 And u.R43208-w.R43208=0)"
        cnnConnection.Execute strDel
    End If
    'Modify by Amy 2020/09/17 只能產生Excel(原:資料性質 4/6 才可產生最新Excel), 改為 只有6不可產生當下資料
    '列印(列印資料抓畫面期間「當下」情況)
    'If bolExcel = False Then
    '當下 or 最新(2/4)
    If Option2(1).Value = True Or (Option2(0).Value = True And (Text3 = "2" Or Text3 = "4")) Then
        '當下才更新
        If Option2(1).Value = True Then
            '更新資料-X在區間,但M/Z不在區間/U在區間,其他資料不在區間,應不需顯示台幣收款金額及差額
            strUpd = "Update accrpt432 Set R43208=null,R43210=null,R43230=null " & _
                        "Where ID='" & strUserNum & "' And R43203 in " & _
                        "(Select R43203 From accrpt432 x Where ID='" & strUserNum & "' And Substr(R43203,1,1)='X' and Not Exists (Select * From Accrpt432 Where R43228=x.R43203 And Substr(R43203,1,1)<>'X') " & _
              "Union Select R43203 From accrpt432 x Where ID='" & strUserNum & "' And Substr(R43203,1,1)='U' and Not Exists (Select * From Accrpt432 Where R43228=x.R43203 And Substr(R43203,1,1)<>'U') )"
            cnnConnection.Execute strUpd
        End If
    
        'Modify by Amy 2020/09/17 原抓資料程式改至產生Excel 時抓,因代理人編號不同需換頁籤顯示
        strSql = "Select Distinct R43201 as FagentNo From Accrpt432 Where ID='" & strUserNum & "' Order by R43201 asc"
        
    '最新(6.未收未付)
    Else
        'Modify by Amy 2017/05/24 抵帳且代理人不為Y51566或Y52269 則依幣別不同分開顯示 原:Order by R43234 asc,R43201 asc,R43202 asc,R43211 asc,R43229 asc
        'Modify by Amy 2020/09/17 原IIf(bolExcel = True, ",R43234 as Sort ", "")拿掉iif
        strSql = "Select R43201 as FagentNo,R43202 as DocDate,R43203 as DocNo,R43204 as Currency,R43205 as Famount,R43206 as Namount,R43207 as Tamount,R43208 as Ramount,R43209 as Pamount,R43210 as Vamount," & _
                    "R43211 as Caseno,R43212 as Point,R43213 as DNno,R43214 as Checkno,R43215 as fa05,R43216 as fa32,R43217 as fa33,R43218 as fa34,R43219 as fa35,R43220 as fa36,R43221 as fa18,R43222 as fa19,R43223 as fa20," & _
                    "R43224 as fa21,R43225 as fa22,R43226 as fa06,R43227 as fa23,R43228 as a1k01,R43229 as CP45,R43230 as CloseS,R43231 as a1k27,R43232 as a1k28,R43234 as Sort " & _
                    "From AccRpt432 Where ID='" & strUserNum & "' "
        
        'Modify by Amy 2024/07/15 拿掉if 因Y51566及Y52269 不需都帶NTD+台幣金額,故全部使用同一種格式-斯閔
'        If Left(Text1, 6) <> "Y51566" And Left(Text1, 6) <> "Y52269" And Val(Text3) = 6 Then
            strSql = strSql & "Order by SubStr(R43234,1,1) asc,R43204 asc,R43234 asc,R43201 asc,R43202 asc,R43211 asc,R43229 asc"
'        Else
'            strSql = strSql & "Order by R43234 asc,R43201 asc,R43202 asc,R43211 asc,R43229 asc"
'        End If
        'end 2024/07/15
        'end 2017/05/24
    End If
     
   adoquery.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount = 0 Then
      adoquery.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   
   'Add by Amy 2014/04/30 +產生Excel
   If bolExcel = True Then
      'Modify by Amy 2020/09/17 原資料性質4/6最新資料Excel由ExcelSaveNew產生,改只有資料性質6由ExcelSaveNew產生，其他由ExcelSaveNew2產生
      If Text3 = "6" Then
        If ExcelSaveNew = True Then MsgBox ("Excel檔案已產生！")
      Else
        If ExcelSaveNew2 = True Then MsgBox ("Excel檔案已產生！")
      End If
      Exit Sub
   End If
   'end 2014/04/30

   MaxLine = 28 '大報表可印列數
   If Option1(1).Value = True Then
        If Printer.PaperSize = 9 Then SetPrintA4
        PrintData_Case
        Exit Sub
   End If
   
   Call PrintPaper

   adoquery.Close
Checking:
   Set Rs = Nothing
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'Add by Amy 2015/06/30
Private Sub PrintPaper()
    Dim strAmount As String
    Dim intLength As Integer
    Dim strNo As String
    
    Printer.PaperSize = PUB_GetPaperSize(15) '美國標準
    intCounter = 3: intRecord = 1: intPage = 1
    Do While adoquery.EOF = False
        If strNo = "" Then
            PrintHead
        ElseIf strNo <> adoquery.Fields("FagentNo").Value Or intRecord > MaxLine Then
            'Modify by Amy 2020/09/17 +adoquery.Fields("FagentNo").Value
            If strNo <> adoquery.Fields("FagentNo").Value Then PrintSum (adoquery.Fields("FagentNo").Value)
            Printer.NewPage
            intCounter = 3: intRecord = 1: intPage = intPage + 1
            PrintHead
        End If
        '單據日期
        If "" & adoquery.Fields("DocDate") <> MsgText(601) Then
            Printer.CurrentX = 0
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print CFDate(adoquery.Fields("DocDate").Value)
        End If
        '單據編號
        If "" & adoquery.Fields("DocNo") <> MsgText(601) Then
            Printer.CurrentX = 1300
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print adoquery.Fields("DocNo").Value
        End If
        '幣別
        If "" & adoquery.Fields("Currency") <> MsgText(601) Then
            Printer.CurrentX = 2600
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print adoquery.Fields("Currency").Value
        End If
        '外幣金額
        If Val("" & adoquery.Fields("Famount")) <> 0 Then
            strAmount = Format(Val(adoquery.Fields("Famount").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 4800 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '台幣金額
        If Val("" & adoquery.Fields("Namount")) <> 0 Then
            strAmount = Format(Val(adoquery.Fields("Namount").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 6100 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '規費
        If Val("" & adoquery.Fields("Tamount")) <> 0 Then
            strAmount = Format(Val(adoquery.Fields("Tamount").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 7400 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '台幣收款金額
        If Val("" & adoquery.Fields("Ramount")) <> 0 Then
             strAmount = Format(Val(adoquery.Fields("Ramount").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 8700 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '台幣付款金額
        If Val("" & adoquery.Fields("Pamount")) <> 0 Then
            strAmount = Format(Val(adoquery.Fields("Pamount").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 10000 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '差額
        If Val("" & adoquery.Fields("Vamount")) <> 0 Then
            strAmount = Format(Val(adoquery.Fields("Vamount").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 11000 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '本所案號
        If "" & adoquery.Fields("CaseNo") <> MsgText(601) Then
            Printer.CurrentX = 11100
            Printer.CurrentY = 300 + intCounter * 300
            If Mid(adoquery.Fields("CaseNo").Value, Len(adoquery.Fields("CaseNo").Value) - 2, 3) = "000" Then
               Printer.Print Mid(adoquery.Fields("CaseNo").Value, 1, Len(adoquery.Fields("CaseNo").Value) - 9) & "-" & Mid(adoquery.Fields("CaseNo").Value, Len(adoquery.Fields("CaseNo").Value) - 8, 6)
            Else
               Printer.Print Mid(adoquery.Fields("CaseNo").Value, 1, Len(adoquery.Fields("CaseNo").Value) - 9) & "-" & Mid(adoquery.Fields("CaseNo").Value, Len(adoquery.Fields("CaseNo").Value) - 8, 6) & _
                         "-" & Mid(adoquery.Fields("CaseNo").Value, Len(adoquery.Fields("CaseNo").Value) - 2, 1) & "-" & Mid(adoquery.Fields("CaseNo").Value, Len(adoquery.Fields("CaseNo").Value) - 1, 2)
            End If
        End If
        '點數
        If "" & adoquery.Fields("Point") <> MsgText(601) Or Val("" & adoquery.Fields("Point")) <> 0 Then
            strAmount = Format(Val(adoquery.Fields("Point").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 13900 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        '代理人D/N No
        If "" & adoquery.Fields("DNNo") <> MsgText(601) Then
            Printer.CurrentX = 14000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print adoquery.Fields("DNNo").Value
        End If
        '匯票號碼
        If "" & adoquery.Fields("CheckNo") <> MsgText(601) Then
            Printer.CurrentX = 15000
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print adoquery.Fields("CheckNo").Value
        End If
        strNo = adoquery.Fields("FagentNo").Value
        intCounter = intCounter + 1
        intRecord = intRecord + 1
        adoquery.MoveNext
    Loop
    Printer.Line (0, 300 + intCounter * 300 + 350)-(16400, 300 + intCounter * 300 + 350)
    intCounter = intCounter + 1
    Call PrintSum(strNo) 'Modify by Amy 2020/09/17
    Printer.EndDoc
End Sub

'Modify by Amy 2020/09/17 +stFAgent/IsExcel/幣別欄/外幣欄
Private Sub PrintSum(stFAgent As String, Optional IsExcel As Boolean = False, Optional Wks As Worksheet, Optional stFieldCurr As String, Optional strFieldAmt As String)
    Dim adoSum As New ADODB.Recordset
    Dim j As Integer
    Dim StrS(1 To 4) As String, oldSumState As String
    Dim intFreq As Integer  '次數
    Dim IsFirst As Boolean '是否第一次印
    
    For i = LBound(dblSum) To UBound(dblSum)
        dblSum(i) = 0
    Next i
    
    'Modify by Amy 2020/09/17 +stFAgent
    'FC 已收:M+X's Z
    StrS(1) = "Select 'FC已收' as SumState,R43204 as Currency,Sum(R43205) as Famount,Sum(R43206) as Namount,Sum(R43207) as Tamount,Sum(R43208) as Ramount,0 as Pamount From AccRpt432 " & _
                    "Where ID='" & strUserNum & "' And R43201='" & stFAgent & "' And (SubStr(R43203,1,1)='M' OR (SubStr(R43228,1,1)='X' And SubStr(R43203,1,1)='Z')) Group by R43204"
    'FC 未收:X-M(有考慮部分未收)
    'Modify by Amy 2019/11/18 +And R43230 is null(未結清)
    StrS(2) = "Select 'FC未收' as SumState,x.R43204 as Currency,Sum(R43205-Nvl(MFamount,0)) as Famount,0 as Namount,0 as Tamount,0 as Ramount,0 as Pamount From " & _
                    "(Select R43204,Nvl(R43205,0) as R43205,Nvl(R43206,0) as R43206,Nvl(R43207,0) as R43207,Nvl(R43208,0) as R43208,R43203 From AccRpt432 Where ID='" & strUserNum & "' And R43201='" & stFAgent & "' And SubStr(R43203,1,1)='X' And R43230 is null) x, " & _
                    "(Select R43204,Sum(Nvl(R43205,0)) as MFamount,Sum(Nvl(R43208,0)) as MRamount,R43228 From AccRpt432 Where ID='" & strUserNum & "' And R43201='" & stFAgent & "' And SubStr(R43203,1,1)='M' Group by R43228,R43204) m " & _
                    "Where R43203=R43228(+) Group by x.R43204 Having Sum(R43205-Nvl(MFamount,0))<>0"
    'CF 已付:U's W+U's Z+V'sW
    StrS(3) = "Select 'CF已付' as SumState,R43204 as Currency,Sum(R43205) as Famount,0 as Namount,0 as Tamount,Sum(R43208) as Ramount,Sum(R43209) as Pamount From AccRpt432 " & _
                    "Where ID='" & strUserNum & "' And R43201='" & stFAgent & "' And ((SubStr(R43228,1,1)='U' And (SubStr(R43203,1,1)='W' OR SubStr(R43203,1,1)='Z' )) " & _
                    "OR (SubStr(R43228,1,1)='V' And SubStr(R43203,1,1)='W')) Group by R43204"
    'CF 未付:U(沒W資料的)-V
    StrS(4) = "Select 'CF未付' as SumState,R43204 as Currency,Sum(R43205) as Famount,Sum(R43206) as Namount,0 as Tamount,0 as Ramount,0 as Pamount From " & _
                    "(Select R43204, Nvl(R43205,0) as R43205, Nvl(R43206,0) as R43206 From Accrpt432 Where ID='" & strUserNum & "' And R43201='" & stFAgent & "' And SubStr(R43203,1,1)='U' And R43203 not in (Select R43228 From AccRpt432 Where ID='" & strUserNum & "' And R43201='" & stFAgent & "' And SubStr(R43228,1,1)='U' And SubStr(R43203,1,1)<>'U')" & _
                    "Union All Select R43204, Nvl(R43205,0)*(-1) as R43205, Nvl(R43206,0)*(-1) as R43206 From AccRpt432 Where ID='" & strUserNum & "' And R43201='" & stFAgent & "' And SubStr(R43203,1,1)='V' And R43203 not in (Select R43228 From AccRpt432 Where ID='" & strUserNum & "' And R43201='" & stFAgent & "' And SubStr(R43228,1,1)='V' And SubStr(R43203,1,1)<>'V')) " & _
                    "Group by R43204 Having Sum(R43205)<>0"
    'end 2020/09/17
    
    '選「往來」
    Select Case Text3
        Case "1"
            intFreq = 2
        Case "2"
            intFreq = 2
        Case "3"
            intFreq = 4
        Case "4", "6"
            intFreq = 4
            If Val(Text3) = 6 Then i = 2
        Case Else
           intFreq = 4
    End Select
    
    If Val(Text3) = 0 Or Val(Text3) > 4 Then
        If Val(Text3) = 6 Then
            '6.未收/未付
            i = 2
        Else
            '5.往來
            i = 1
        End If
    Else
        i = Val(Text3)
    End If
    
    IsFirst = True
    Do While i <= intFreq
        If adoSum.State <> adStateClosed Then adoSum.Close
        adoSum.CursorLocation = adUseClient
        adoSum.Open StrS(i), cnnConnection, adOpenStatic, adLockReadOnly
        If adoSum.RecordCount > 0 Then
            For j = 0 To adoSum.RecordCount - 1
                'Modify by Amy 2020/09/17 +IsExcel
                If IsExcel = True Then
                    If IsFirst = True Then
                        Wks.Range("A" & intCounter).Value = "合計"
                        Wks.Range("A" & intCounter).HorizontalAlignment = xlLeft
                        IsFirst = False
                    End If
                    Wks.Range("B" & intCounter).Value = adoSum.Fields("SumState")
                    Wks.Range("B" & intCounter).HorizontalAlignment = xlLeft
                    '幣別
                    Wks.Range(stFieldCurr & intCounter).Value = adoSum.Fields("Currency")
                    Wks.Range(stFieldCurr & intCounter).HorizontalAlignment = xlCenter
                    '外幣
                    Wks.Range(strFieldAmt & intCounter).Value = adoSum.Fields("Famount")
                    Wks.Range(strFieldAmt & intCounter).NumberFormatLocal = "#,##0.00"
                    Wks.Range(strFieldAmt & intCounter).HorizontalAlignment = xlRight
                Else
                    If IsFirst = True Then
                        intCounter = intCounter + 1
                        Printer.CurrentX = 0
                        Printer.CurrentY = 300 + intCounter * 300
                        Printer.Print "合      計"
                        IsFirst = False
                    End If
                    If oldSumState <> "" & adoSum.Fields("SumState") Then
                        Printer.CurrentX = 1300
                        Printer.CurrentY = 300 + intCounter * 300
                        Printer.Print "" & adoSum.Fields("SumState")
                    End If
                    dblSum(1) = Val("" & adoSum.Fields("Famount"))
                    dblSum(2) = Val("" & adoSum.Fields("Namount"))
                    dblSum(3) = Val("" & adoSum.Fields("Tamount"))
                    dblSum(4) = Val("" & adoSum.Fields("Ramount"))
                    dblSum(5) = Val("" & adoSum.Fields("Pamount"))
                    If Option1(1) = True Then
                        Call PrintSumVal_Case("" & adoSum.Fields("Currency"))
                    Else
                        Call PrintSumVal("" & adoSum.Fields("Currency"))
                    End If
                End If
                'end 2020/09/17
            Next j
            intCounter = intCounter + 1
            intRecord = intRecord + 1
        End If
        If Val(Text3) = 4 Or Val(Text3) = 6 Then
            i = i + 2
        Else
            i = i + 1
        End If
    Loop

End Sub

Private Sub PrintSumVal(ByVal strCurrency As String)
    Dim strAmount As String
    Dim intLength As Integer
    
    Printer.CurrentX = 2600
    Printer.CurrentY = 300 + intCounter * 300
    Printer.Print strCurrency
    
    '外幣金額
    If dblSum(1) <> 0 Then
        strAmount = Format(dblSum(1), FDollar)
        intLength = Printer.TextWidth(strAmount)
        Printer.CurrentX = 4800 - intLength
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print strAmount
    End If
     '台幣額
    If dblSum(2) <> 0 Then
        strAmount = Format(dblSum(2), FDollar)
        intLength = Printer.TextWidth(strAmount)
        Printer.CurrentX = 6100 - intLength
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print strAmount
    End If
    '規費
    If dblSum(3) <> 0 Then
        strAmount = Format(dblSum(3), FDollar)
        intLength = Printer.TextWidth(strAmount)
        Printer.CurrentX = 7400 - intLength
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print strAmount
    End If
    '台幣收款金額
    If dblSum(4) <> 0 Then
        strAmount = Format(dblSum(4), FDollar)
        intLength = Printer.TextWidth(strAmount)
        Printer.CurrentX = 8700 - intLength
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print strAmount
    End If
    '台幣付款金額
    If dblSum(5) <> 0 Then
        strAmount = Format(dblSum(5), FDollar)
        intLength = Printer.TextWidth(strAmount)
        Printer.CurrentX = 10000 - intLength
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print strAmount
    End If
    
End Sub

Private Sub PrintSumVal_Case(ByVal strCurrency As String)
     Dim strAmount As String
    Dim intLength As Integer
    
    Printer.CurrentX = 2600
    Printer.CurrentY = 300 + intCounter * 300
    Printer.Print strCurrency
    
    '外幣金額合計
    If dblSum(1) <> 0 Then
        strAmount = Format(dblSum(1), FDollar)
        intLength = Printer.TextWidth(strAmount)
        Printer.CurrentX = 4800 - intLength
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print strAmount
    End If
    
    '台幣金額合計
    If dblSum(2) <> 0 Then
        strAmount = Format(dblSum(2), FDollar)
        intLength = Printer.TextWidth(strAmount)
        Printer.CurrentX = 6300 - intLength
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print strAmount
    End If
    
    '規費合計
    If dblSum(3) <> 0 Then
        strAmount = Format(dblSum(3), FDollar)
        intLength = Printer.TextWidth(strAmount)
        Printer.CurrentX = 7600 - intLength
        Printer.CurrentY = 300 + intCounter * 300
        Printer.Print strAmount
    End If
End Sub

'Add by Morgan 2004/10/26
'*************************************************
'  抬頭合計
'
'*************************************************
Private Sub PrintSum_Old()
'   Dim intRow As Integer
'
'   Printer.Line (0, 300 + intCounter * 300 + 150)-(16400, 300 + intCounter * 300 + 150)
'
'   'Modify By Sindy 2012/12/21 改成陣列值
'   intRow = intCounter
'   If m_FCSum(0) > 0 Or m_CFSum(0) > 0 Or m_CFVSum(0) > 0 Then
'      intCounter = intCounter + 1
'      Printer.CurrentX = 2800
'      Printer.CurrentY = 300 + intCounter * 300
'      Printer.Print "合計"
'
'      intCounter = intRow
'      If m_FCSum(0) > 0 Then
'         For i = 0 To intMax
'            If m_FCSum(i) > 0 Then
'               intCounter = intCounter + 1
'               If i = 0 Then
'                  Printer.CurrentX = 3500 '3600
'                  Printer.CurrentY = 300 + intCounter * 300
'                  Printer.Print "FC未收"
'               End If
'               Printer.CurrentX = 4500
'               Printer.CurrentY = 300 + intCounter * 300
'               Printer.Print m_FCCur(i) & " " & Format(m_FCSum(i), FDollar)
'            ElseIf m_FCSum(i) = 0 Then
'               Exit For
'            End If
'         Next i
'      End If
'      intCounter = intRow
'      If m_CFSum(0) > 0 Then
'         For i = 0 To intMax
'            If m_CFSum(i) > 0 Then
'               intCounter = intCounter + 1
'               If i = 0 Then
'                  Printer.CurrentX = 6500
'                  Printer.CurrentY = 300 + intCounter * 300
'                  Printer.Print "CF未付"
'               End If
'               Printer.CurrentX = 7500
'               Printer.CurrentY = 300 + intCounter * 300
'               Printer.Print m_CFCur(i) & " " & Format(m_CFSum(i), FDollar)
'            ElseIf m_CFSum(i) = 0 Then
'               Exit For
'            End If
'         Next i
'      End If
'      intCounter = intRow
'      If m_CFVSum(0) > 0 Then
'         For i = 0 To intMax
'            If m_CFVSum(i) > 0 Then
'               intCounter = intCounter + 1
'               If i = 0 Then
'                  Printer.CurrentX = 9500
'                  Printer.CurrentY = 300 + intCounter * 300
'                  Printer.Print "抵帳單合計"
'               End If
'               Printer.CurrentX = 11000
'               Printer.CurrentY = 300 + intCounter * 300
'               Printer.Print m_CFVCur(i) & " " & Format(m_CFVSum(i), FDollar)
'            ElseIf m_CFVSum(i) = 0 Then
'               Exit For
'            End If
'         Next i
'      End If
'   End If
'   '2012/12/21 End
'   'Modify By Sindy 2012/12/21 清空陣列值
'   For i = 0 To intMax
'      m_FCCur(i) = "": m_CFCur(i) = "": m_CFVCur(i) = ""
'      m_FCSum(i) = 0: m_CFSum(i) = 0: m_CFVSum(i) = 0
'   Next i
'   '2012/12/21 End
End Sub

'*************************************************
'  抬頭列印
'
'*************************************************
Private Sub PrintHead()
Dim strSelect As String

   Printer.FontSize = 14
   Printer.CurrentX = 5000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print ReportTitle(209)
   Printer.FontSize = 12
   intCounter = intCounter + 2
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印人員: " & StaffQuery(strUserNum)
   Printer.CurrentX = 12000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印日期: " & CFDate(ACDate(ServerDate))
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "代理人編號: " & adoquery.Fields("FagentNo").Value
   Printer.CurrentX = 4000
   Printer.CurrentY = 300 + intCounter * 300
   '2012/2/22 MODIFY BY SONIA Y47804無英文名稱
   'If IsNull(adoquery.Fields("fa05").Value) Then
   '   Printer.Print "代理人名稱(英): "
   'Else
   '   Printer.Print "代理人名稱(英): " & adoquery.Fields("fa05").Value
   'End If
   Printer.Print "代理人名稱: " & adoquery.Fields("fa05").Value
   '2012/2/22 END
   Printer.CurrentX = 12000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "頁次: " & intPage
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "代理人地址: "
   If IsNull(adoquery.Fields("fa32").Value) Then
      If IsNull(adoquery.Fields("fa18").Value) = False Then
         Printer.CurrentX = 1300
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa18").Value
      End If
   Else
      Printer.CurrentX = 1300
      Printer.CurrentY = 300 + intCounter * 300
      Printer.Print adoquery.Fields("fa32").Value
   End If
   intCounter = intCounter + 1
   If IsNull(adoquery.Fields("fa32").Value) Then
      If IsNull(adoquery.Fields("fa19").Value) = False Then
         Printer.CurrentX = 1300
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa19").Value
      End If
   Else
      Printer.CurrentX = 1300
      Printer.CurrentY = 300 + intCounter * 300
      Printer.Print "" & adoquery.Fields("fa33").Value
   End If
   intCounter = intCounter + 1
   If IsNull(adoquery.Fields("fa32").Value) Then
      If IsNull(adoquery.Fields("fa20").Value) = False Then
         Printer.CurrentX = 1300
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa20").Value
      End If
   Else
      Printer.CurrentX = 1300
      Printer.CurrentY = 300 + intCounter * 300
      Printer.Print "" & adoquery.Fields("fa34").Value
   End If
   intCounter = intCounter + 1
   If IsNull(adoquery.Fields("fa32").Value) Then
      If IsNull(adoquery.Fields("fa21").Value) = False Then
         Printer.CurrentX = 1300
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa21").Value
      End If
   Else
      Printer.CurrentX = 1300
      Printer.CurrentY = 300 + intCounter * 300
      Printer.Print "" & adoquery.Fields("fa35").Value
   End If
   intCounter = intCounter + 1
   If IsNull(adoquery.Fields("fa32").Value) Then
      If IsNull(adoquery.Fields("fa22").Value) = False Then
         Printer.CurrentX = 1300
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa22").Value
      End If
      
      'Add by Morgan 2011/5/25
      '英文地址6
      If IsNull(adoquery.Fields("fa70").Value) = False Then
         intCounter = intCounter + 1
         Printer.CurrentX = 1300
         Printer.CurrentY = 300 + intCounter * 300
         Printer.Print adoquery.Fields("fa70").Value
      End If
      
   Else
      Printer.CurrentX = 1300
      Printer.CurrentY = 300 + intCounter * 300
      Printer.Print "" & adoquery.Fields("fa36").Value
   End If
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "往來日期: " & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
   Select Case Text3
      Case "1"
         strSelect = "1. FC往來"
      Case "2"
         strSelect = "2. FC未收"
      Case "3"
         strSelect = "3. CF往來"
      Case "4"
         strSelect = "4. CF未付"
      Case "5"
         strSelect = "5. 往來"
      Case "6"
         strSelect = "6. 未收未付"
      Case Else
         strSelect = ""
   End Select
   Printer.CurrentX = 4000
   Printer.CurrentY = 300 + intCounter * 300
   '2008/7/15 modify by sonia
   'Printer.Print "資料性質: " & strSelect
   Printer.Print "資料性質: " & strSelect '& " (請款單之外幣金額為美金金額)"
   intCounter = intCounter + 1
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "單據日期"
   Printer.CurrentX = 1300
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "單據編號"
   Printer.CurrentX = 2600
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "幣別"
   Printer.CurrentX = 3600
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "外幣金額"
   Printer.CurrentX = 4900
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "台幣金額"
   Printer.CurrentX = 6200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "規費"
   Printer.CurrentX = 7500
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "台幣收款"
   Printer.CurrentX = 8800
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "台幣付款"
   Printer.CurrentX = 10100
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "差額"
   Printer.CurrentX = 11100
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "本所案號"
   Printer.CurrentX = 13000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "點數"
   Printer.CurrentX = 14000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "代理人"
   Printer.CurrentX = 15000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "匯票號碼"
   intCounter = intCounter + 1
   Printer.CurrentX = 7500
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "    金額"
   Printer.CurrentX = 8800
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "    金額"
   Printer.CurrentX = 14000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "D/N No."
   Printer.Line (0, 300 + intCounter * 300 + 350)-(16400, 300 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   Dim bolCancel As Boolean 'Add by Amy 2020/09/17
   
   'Modify by Amy 2014/11/18
   FormCheck = False
   '資料性質必填
   If Trim(Text3) = MsgText(601) Then
        MsgBox Label2.Caption & MsgText(52), , MsgText(5)
        Text3.SetFocus
        Exit Function
   End If
   'Add by Amy 2020/12/21 若資料性質為1、3、5則起、迄日期必填,以免資料量過大無法顯示
   If Val(Text3) Mod 2 = 1 Then
        If MaskEdBox1.Text = MsgText(601) Or MaskEdBox1.Text = MsgText(29) Then
            MsgBox "往來日期起日不可為空", , MsgText(5)
            MaskEdBox1.SetFocus
            Exit Function
        End If
        If MaskEdBox2.Text = MsgText(601) And MaskEdBox2.Text = MsgText(29) Then
            MsgBox "往來日期迄日不可為空", , MsgText(5)
            MaskEdBox2.SetFocus
            Exit Function
        End If
   End If
   'Add by Amy 2020/09/17
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) And MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
        If Val(FCDate(MaskEdBox1.Text)) > Val(FCDate(MaskEdBox2.Text)) Then
            MsgBox "往來日期起日不可大於止日", , MsgText(5)
            Exit Function
        End If
   End If
   'end 2020/09/17
   
   If Option1(0).Value = True Then
        'Modify by Amy 2020/09/17 因只產生Excel,故請款對象必輸,且為前6碼相同
'        '若選請款對象不輸編號,則查全部
'        FormCheck = True
        If Trim(Text1) = MsgText(601) Or Trim(Text2) = MsgText(601) Then
            MsgBox "請款對象" & MsgText(52), , MsgText(5)
            Exit Function
        End If
        If Left(Text1, 6) <> Left(Text2, 6) Then
            MsgBox "請款對象前6碼必需相同", , MsgText(5)
            Exit Function
        Else
            FormCheck = True
            Exit Function
        End If
        'end 2020/09/17
   ElseIf Option1(1).Value = True Then
        '本所案號查詢必需輸本所案號
        If Trim(Text4) = MsgText(601) Or Trim(Text5) = MsgText(601) Then
            MsgBox "本所案號" & MsgText(52), , MsgText(5)
            Exit Function
        Else
            FormCheck = True
            Exit Function
        End If
   End If

'   If Text1 <> MsgText(601) Then
'        FormCheck = True
'        Exit Function
'   End If
'   If Text2 <> MsgText(601) Then
'        FormCheck = True
'        Exit Function
'   End If
'   If Text3 <> MsgText(601) Then
'      FormCheck = True
'      Exit Function
'   End If
'   If MaskEdBox1.Text <> MsgText(29) Then
'      FormCheck = True
'      Exit Function
'   End If
'   If MaskEdBox2.Text <> MsgText(29) Then
'      FormCheck = True
'      Exit Function
'   End If
   'FormCheck = False
   'end 2014/11/18
End Function

'Add By Sindy 2013/1/23
Private Sub Text3_LostFocus()
   'Modify by Amy 2014/04/30 改請款對象前6碼相同產生Excel鈕 原列印抵帳明細拿掉
   'Modify by Amy 2014/08/07
   'Modify by Amy 2015/11/25 +資料性質 4也顯示Excel鈕
   'Mark by Amy 2020/09/17 只能產生Excel
'   If (Text1 <> "" And Text2 <> "" And Left(Text1, 6) = Left(Text2, 6)) And (Text3.Text = "6" Or Text3.Text = "4") Then
'      'CmdExcel.Visible = True
'      CmdExcel.Enabled = True
'   Else
'      'CmdExcel.Visible = False
'      CmdExcel.Enabled = False
'   End If
End Sub

'Add By Sindy 2013/1/23
Sub GetPleft(intItem As Integer)
   Erase PLeft
   If intItem = 0 Then
      PLeft(0) = 500
      PLeft(1) = 1800
      PLeft(2) = 3300
      PLeft(3) = 5500
      PLeft(4) = 7000
      PLeft(5) = 9000
      PLeft(6) = 9500
   Else
      PLeft(0) = 500
      PLeft(1) = 1800
      PLeft(2) = 4000
      PLeft(3) = 6500
      PLeft(4) = 7000
   End If
End Sub

'Add By Sindy 2013/1/23 未收抵帳明細
Private Sub PrintList1()
   Call GetPleft(0)
   Printer.PaperSize = 9
   Printer.Orientation = 1
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 14
   Printer.Font.Bold = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("Statement of Account") / 2)
   Printer.CurrentY = 600
   Printer.Print "Statement of Account"
   iLine = 3
   Printer.Font.Size = 12
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("In favor of Tai E International Patent & Law Office") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "In favor of Tai E International Patent & Law Office"
   Printer.Font.Size = 9
   Printer.Font.Bold = False
   iLine = iLine + 1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iLine * 300
   Printer.Print "列印人　：" & strUserName
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iLine * 300
   Printer.Print "製表日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iLine = iLine + 1
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   iLine = iLine + 1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iLine * 300
   Printer.Print String(115, "-")
   iLine = iLine + 1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iLine * 300
   Printer.Print "Date(Y/M/D)"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "Debit Note No"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "Currency"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("Amount")
   Printer.CurrentY = iLine * 300
   Printer.Print "Amount"
   Printer.CurrentX = PLeft(4) - Printer.TextWidth("Amount/NT")
   Printer.CurrentY = iLine * 300
   Printer.Print "Amount/NT"
   Printer.CurrentX = PLeft(5) - Printer.TextWidth("Official Fee/NT")
   Printer.CurrentY = iLine * 300
   Printer.Print "Official Fee/NT"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iLine * 300
   Printer.Print "Our Ref"
   iLine = iLine + 1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iLine * 300
   Printer.Print String(115, "-")
   iLine = iLine + 1
End Sub

'Add By Sindy 2013/1/28
Sub PrintDetail1()
For m_j = 0 To 6
   If m_j >= 3 And m_j <= 5 Then
      Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j + 1))
   Else
      Printer.CurrentX = PLeft(m_j)
   End If
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(m_j + 1)
Next m_j
iLine = iLine + 1
End Sub

'Add By Sindy 2013/1/23 未付抵帳明細
Private Sub PrintList2(strFagentName As String)
   Call GetPleft(1)
   Printer.PaperSize = 9
   Printer.Orientation = 1
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 14
   Printer.Font.Bold = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("Statement of Account") / 2)
   Printer.CurrentY = 600
   Printer.Print "Statement of Account"
   iLine = 3
   Printer.Font.Size = 12
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("In favor of " & strFagentName) / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "In favor of " & strFagentName
   Printer.Font.Size = 9
   Printer.Font.Bold = False
   iLine = iLine + 1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iLine * 300
   Printer.Print "列印人　：" & strUserName
   Printer.CurrentX = 9000 'PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "製表日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iLine = iLine + 1
   Printer.CurrentX = 9000 'PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "頁　　次：" & Printer.Page
   iLine = iLine + 1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iLine * 300
   Printer.Print String(115, "-")
   iLine = iLine + 1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iLine * 300
   Printer.Print "Date(Y/M/D)"
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "Invoice No"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "Currency"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("Amount")
   Printer.CurrentY = iLine * 300
   Printer.Print "Amount"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "Our Ref"
   iLine = iLine + 1
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iLine * 300
   Printer.Print String(115, "-")
   iLine = iLine + 1
End Sub

'Add By Sindy 2013/1/28
Sub PrintDetail2()
For m_j = 0 To 4
   If m_j = 3 Then
      Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j + 1))
   Else
      Printer.CurrentX = PLeft(m_j)
   End If
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(m_j + 1)
Next m_j
iLine = iLine + 1
End Sub

'*************************************************
'  轉成Excel檔案:  6-未收未付(抵帳單) /4-未付(109/09/17 不使用此格式,改用ExcelSaveNew2格式)
'  Add by Amy 2014/04/30
'*************************************************
Private Function ExcelSaveNew() As Boolean
Dim xlsAgentPoint As New Excel.Application
Dim wksrpt As New Worksheet
Dim xlsFileName As String
Dim i As Integer, j As Integer, intR As Integer, TaieRow As Integer, AgentRow As Integer
Dim strTitle1, intWidth1, strFormat1, strTitle2, intWidth2, strFormat2, intSumField1, intSumField2 '欄位 名稱/寬/格式/加總欄名
Dim strMaxField As String, intHeadRow As Integer '最大欄位編號/表頭總列
Dim bolIsChinese As Boolean, strAgentName As String '是否為中文/代理人名稱
Dim ReportTitle As String '報表名稱
'Add by Amy 2017/05/24
Dim strOldCurr As String, strSum(1) As String '幣別/加總欄位
Dim intStart As Integer  '起始Row
Dim strCurr_T, strCurr_A   'for 最後顯示幣別 台一/代理
Dim strMoney_T() As String, strMoney_A() As String 'for 最後顯示金額 台一/代理
Dim strTmp(4) As String

On Error GoTo ErrHnd

'設定Excel欄位名稱/寬度/格式
'Modify by Amy 2014/09/04 增加Your Ref. No(彼所案號)
strTitle1 = Array("Our Debit Notes", "Date", "Our Ref", "Your Ref. No.", "Amount")
strFormat1 = Array("", "", "@", "yyyy/m/d", "#,##0.00")
intSumField1 = Array(4) '台一加總(從0算第n欄),主要加總寫第一個(for Information 有底色區)

'Modify by Amy 2017/05/24 改A4直印,改版加各幣別分別列示
strTitle2 = Array("Your Debit Note", "Date", "Our Ref", "Your Ref. No.", "Amount", "SoA")
'Modify by Amy 2015/11/25 +資料性質4 欄位大小
'Modify by Amy 2024/06/25 選6.未收未付 增加 SoA:U單號,4-未付已不使用此格式
'If Text3 = 6 Then
    intWidth1 = Array(9, 8.5, 7.5, 1.5, 13)
    intWidth2 = Array(11.5, 7.5, 6.5, 1.5, 10.5, 10)
'Else
'    intWidth1 = Array(12, 8.5, 10, 21, 13)
'    intWidth2 = Array(15.5, 14.5, 16.5, 24, 18.5)
'End If
strFormat2 = Array("@", "yyyy/m/d", "", "@", "#,##0.00", "")
intSumField2 = Array(4) '代理人加總(從0算第n欄),主要加總寫第一個(for Information 有底色區)
'end 2024/06/25
'end 2014/09/04

'定稿語文(fa31)若為中文則使用中文版，其他使用英文版
bolIsChinese = fa31IsChinese(Left(Text1, 8), strAgentName)
'Memo by Amy 2024/07/15 斯閔取消:Y51566 及Y52269 用中文格式且不管請款幣別為何,請款(左邊)都帶NTD+台幣金額
'Modify by Amy 2015/11/25 +Y52404 用中文格式
If Left(Text1, 6) = "Y51566" Or Left(Text1, 6) = "Y52269" Or Left(Text1, 6) = "Y52404" Then
    bolIsChinese = True
End If
'Modify by Amy 2015/11/25 +資料性質4-對帳單檔名
'Modify by Amy 2024/06/25 Mark if 因4-未付,已不使用此格式
If bolIsChinese = True Then
'    If Text3 = 6 Then
        ReportTitle = "&""新細明體,粗體""&14抵帳明細表" & Chr(10) & _
                            Year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日"
        xlsFileName = Left(Text1, 6) & "帳款抵帳明細" & MsgText(43)
'    Else
'        ReportTitle = "&""新細明體,粗體""&14對帳單" & Chr(10) & _
'                            Year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日"
'        xlsFileName = Left(Text1, 6) & "對帳單" & Val(FCDate(MaskEdBox2.Text)) & MsgText(43)
'    End If
Else
'    If Text3 = 6 Then
        ReportTitle = "&""新細明體,粗體""&14Details of mutaul account offset" & Chr(10) & _
                            Year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日"
                            
        xlsFileName = Left(Text1, 6) & " Mutual account offset" & MsgText(43)
'    Else
'        ReportTitle = "&""新細明體,粗體""&14Statement of account" & Chr(10) & _
'                            Year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日"
'
'        xlsFileName = Left(Text1, 6) & " Statement of account" & Val(FCDate(MaskEdBox2.Text)) & MsgText(43)
'    End If
End If
'end 2015/11/25

If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
    If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
    End If
Else
    Kill strExcelPath & xlsFileName
End If
 
xlsAgentPoint.SheetsInNewWorkbook = 3  'Moidfy by Amy 2019/11/19 改回預設工作表數量
xlsAgentPoint.Workbooks.add
Set wksrpt = xlsAgentPoint.Worksheets(1)

strMaxField = Chr(UBound(strTitle1) + UBound(strTitle2) + 67)
intHeadRow = 1

'設定台一/代理人抬頭
wksrpt.Range("a" & intHeadRow).Value = IIf(bolIsChinese = True, "台一", "Taie E") & " Debit Notes Details"
wksrpt.Range("a" & intHeadRow & ":" & Chr(UBound(strTitle1) + 65) & intHeadRow).MergeCells = True
wksrpt.Range("a" & intHeadRow & ":" & Chr(UBound(strTitle1) + 65) & intHeadRow).HorizontalAlignment = xlCenter

wksrpt.Range(Chr(UBound(strTitle1) + 67) & intHeadRow).Value = strAgentName & " Invoices Details"
wksrpt.Range(Chr(UBound(strTitle1) + 67) & intHeadRow & ":" & strMaxField & intHeadRow).MergeCells = True
wksrpt.Range(Chr(UBound(strTitle1) + 67) & intHeadRow & ":" & strMaxField & intHeadRow).HorizontalAlignment = xlCenter
intHeadRow = intHeadRow + 1

'設定台一/代理人欄位
For i = 0 To UBound(strTitle1)
    wksrpt.Columns(Chr(i + 65) & ":" & Chr(i + 65)).ColumnWidth = intWidth1(i)
    wksrpt.Range(Chr(i + 65) & intHeadRow).Value = strTitle1(i)
Next i
'間隔欄
wksrpt.Columns(Chr(UBound(strTitle1) + 66) & ":" & Chr(UBound(strTitle1) + 66)).ColumnWidth = 1
For i = 0 To UBound(strTitle2)
    wksrpt.Columns(Chr(UBound(strTitle1) + i + 67) & ":" & Chr(UBound(strTitle1) + i + 67)).ColumnWidth = intWidth2(i)
    wksrpt.Range(Chr(UBound(strTitle1) + i + 67) & intHeadRow).Value = strTitle2(i)
Next i

'欄位名稱置中
wksrpt.Range("a" & intHeadRow & ":" & strMaxField & intHeadRow).HorizontalAlignment = xlCenter

TaieRow = intHeadRow + 1: AgentRow = intHeadRow + 1
intStart = TaieRow
'抓資料
adoquery.MoveFirst
For i = 1 To adoquery.RecordCount
    '台一
    If Left(adoquery.Fields("sort"), 1) = "1" Then
        'Modify by Amy 2024/07/15 拿掉if 因Y51566及Y52269 不需都帶NTD+台幣金額,故全部使用同一種格式-斯閔
        '非Y51566及Y52269 幣別不同需, 分開列示並加總 - 莘
        'If Left(Text1, 6) <> "Y51566" And Left(Text1, 6) <> "Y52269" Then
            If strOldCurr <> MsgText(601) And strOldCurr <> "" & adoquery.Fields("Currency") Then
                wksrpt.Range("a" & TaieRow).Value = "小計"
                wksrpt.Range(Chr(UBound(strTitle1) + 65) & TaieRow).NumberFormatLocal = """" & strOldCurr & """#,##0.00_);[紅色](" & """" & strOldCurr & """#,##0.00)"
                'Modify by Amy 2019/11/19 +if 沒資料時,公式會加到欄位名稱
                If intStart = TaieRow Then
                    wksrpt.Range(Chr(UBound(strTitle1) + 65) & TaieRow).Value = 0
                Else
                    wksrpt.Range(Chr(UBound(strTitle1) + 65) & TaieRow).Formula = "=sum(" & Chr(UBound(strTitle1) + 65) & intStart & ":" & Chr(UBound(strTitle1) + 65) & TaieRow - 1 & ")"
                End If
                wksrpt.Range(Chr(UBound(strTitle1) + 65) & TaieRow - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous '下框線
                strSum(0) = strSum(0) & ";" & strOldCurr & "-" & Chr(UBound(strTitle1) + 65) & TaieRow
                TaieRow = TaieRow + 2: intStart = TaieRow
            End If
        'End If
        'end 2024/07/15
        '資料
        For j = 0 To UBound(strTitle1)
            If strFormat1(j) <> "" Then
                If strTitle1(j) = "Amount" Then
                    'Modify by Amy 2024/07/15 拿掉if 因Y51566及Y52269 不需都帶NTD+台幣金額,故全部使用同一種格式-斯閔
'                    If Left(Text1, 6) = "Y51566" Or Left(Text1, 6) = "Y52269" Then
'                        'Y51566 及Y52269 用中文格式且不管請款幣別為何，請款(左邊)都帶NTD+台幣金額
'                        strFormat1(j) = """NTD""#,##0.00_);[紅色](" & """NTD""#,##0.00)"
'                    Else
                        strFormat1(j) = """" & adoquery.Fields("Currency") & """#,##0.00_);[紅色](" & """" & adoquery.Fields("Currency") & """#,##0.00)"
'                    End If
                    'end 2024/07/15
                End If
                wksrpt.Range(Chr(j + 65) & TaieRow).NumberFormatLocal = strFormat1(j)
            End If
            Select Case j
                Case 0
                    strExc(0) = "" & adoquery.Fields("a1k01")
                'Modify by Amy 2015/07/28 調位置 原:Amount 前
                Case 1
                    strExc(0) = ChangeTStringToWDateString(adoquery.Fields("DocDate"))
                Case 2
                    If Not IsNull(adoquery.Fields("CaseNo")) Then
                        strExc(0) = Left(adoquery.Fields("CaseNo"), Len(adoquery.Fields("CaseNo")) - 3)
                    Else
                        strExc(0) = ""
                    End If
                'Add by Amy 2014/09/04 +Your Reference No(彼所案號)
                Case 3
                    'Modfiy by Amy 2016/03/31 +巨京沒彼所案號抓分所案號
                    strExc(0) = GetBaseYourRef("" & adoquery.Fields("CaseNo"), IIf(Left(Text1, 6) = "Y52269", True, False))
                Case 4
                    'Modify by Amy 2024/07/15 拿掉if 因Y51566及Y52269 不需都帶NTD+台幣金額,故全部使用同一種格式-斯閔
'                    If Left(Text1, 6) = "Y51566" Or Left(Text1, 6) = "Y52269" Then
'                        'Y51566 及Y52269 用中文格式且不管請款幣別為何,都帶NTD+台幣金額
'                        strExc(0) = adoquery.Fields("Namount")
'                    Else
                        strExc(0) = adoquery.Fields("Famount")
'                    End If
                    'end 2024/0715
               Case Else
            End Select
            wksrpt.Range(Chr(j + 65) & TaieRow).Value = strExc(0)
        Next j
        TaieRow = TaieRow + 1
        
    '代理人資料
    Else
        'Modify by Amy 2024/07/15 拿掉if 因Y51566及Y52269 同其他代理人也要顯示小計,故全部使用同一種格式-斯閔
        '非Y51566及Y52269 幣別不同需, 分開列示並加總 - 莘
        'If Left(Text1, 6) <> "Y51566" And Left(Text1, 6) <> "Y52269" Then
            If Text3 = "6" And AgentRow = intHeadRow + 1 Then
                '顯示最後一筆台一加總
                wksrpt.Range("a" & TaieRow).Value = "小計"
                wksrpt.Range(Chr(UBound(strTitle1) + 65) & TaieRow).NumberFormatLocal = """" & strOldCurr & """#,##0.00_);[紅色](" & """" & strOldCurr & """#,##0.00)"
                'Modify by Amy 2019/11/19 +if 沒資料時,公式會加到欄位名稱
                If intStart = TaieRow Then
                    wksrpt.Range(Chr(UBound(strTitle1) + 65) & TaieRow).Value = 0
                Else
                    wksrpt.Range(Chr(UBound(strTitle1) + 65) & TaieRow).Formula = "=sum(" & Chr(UBound(strTitle1) + 65) & intStart & ":" & Chr(UBound(strTitle1) + 65) & TaieRow - 1 & ")"
                End If
                wksrpt.Range(Chr(UBound(strTitle1) + 65) & TaieRow - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous '下框線
                strSum(0) = strSum(0) & ";" & strOldCurr & "-" & Chr(UBound(strTitle1) + 65) & TaieRow
                TaieRow = TaieRow + 2
                intStart = AgentRow: strOldCurr = ""
            End If
            If strOldCurr <> MsgText(601) And strOldCurr <> "" & adoquery.Fields("Currency") Then
                wksrpt.Range(Chr(UBound(strTitle1) + 67) & AgentRow).Value = "小計"
                'Modify by Amy 2024/06/25 原67 ex:Chr(UBound(strTitle1) + UBound(strTitle2) + 67,增加SoA欄位
                wksrpt.Range(Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow).NumberFormatLocal = """" & strOldCurr & """#,##0.00_);[紅色](" & """" & strOldCurr & """#,##0.00)"
                'Modify by Amy 2019/11/19 +if 沒資料時,公式會加到欄位名稱
                If intStart = AgentRow Then
                    wksrpt.Range(Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow).Value = 0
                Else
                    wksrpt.Range(Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow).Formula = "=sum(" & Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & intStart & ":" & Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow - 1 & ")"
                End If
                wksrpt.Range(Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous '下框線
                strSum(1) = strSum(1) & ";" & strOldCurr & "-" & Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow
                'end 2024/06/25
                AgentRow = AgentRow + 2: intStart = AgentRow
            End If
        'End If
        'end 2024/07/15
        '資料
        For j = 0 To UBound(strTitle2)
            If strFormat2(j) <> "" Then
                If strTitle2(j) = "Amount" Then
                    strFormat2(j) = """" & adoquery.Fields("Currency") & """#,##0.00_);[紅色](" & """" & adoquery.Fields("Currency") & """#,##0.00)"
                End If
                wksrpt.Range(Chr(UBound(strTitle1) + j + 67) & AgentRow).NumberFormatLocal = strFormat2(j)
            End If
            Select Case j
                Case 0
                    strExc(0) = "" & adoquery.Fields("DNNo")
                Case 1
                    If Not IsNull(adoquery.Fields("DocDate")) Then
                        strExc(0) = ChangeTStringToWDateString(adoquery.Fields("DocDate"))
                    Else
                        strExc(0) = ""
                    End If
                Case 2
                    If Not IsNull(adoquery.Fields("CaseNo")) Then
                        strExc(0) = Left(adoquery.Fields("CaseNo"), Len(adoquery.Fields("CaseNo")) - 3)
                    Else
                        strExc(0) = ""
                    End If
                'Add by Amy 2014/09/04 +Your Reference No(彼所案號)
                Case 3
                    'Modify by Amy 2015/09/14 直接抓cp45
                    'Modfify by Amy 2015/09/11+參數CaseNo 避免同U單號,抓到多筆資料只帶一筆
                    'strExc(0) = GetYourRefNo2("" & adoquery.Fields("DocNo"), "" & adoquery.Fields("CaseNo"))
                    strExc(0) = "" & adoquery.Fields("CP45")
                Case 4
                    'Modify by Amy 2016/07/14 +判斷抵帳單要顯示負的 ex:Y47740 期間:~105.6.30 CFP026917000 V10500015/6 顯示負
                    If Left(adoquery.Fields("DocNo"), 1) = "V" Then
                        strExc(0) = Val(adoquery.Fields("Famount")) * -1
                    Else
                        strExc(0) = adoquery.Fields("Famount")
                    End If
               'Add by Amy 2024/06/25 增加 SoA:U單號
               Case 5
                  strExc(0) = ""
                  If Left(adoquery.Fields("DocNo"), 1) = "U" Then
                     strExc(0) = adoquery.Fields("DocNo")
                  End If
               Case Else
            End Select
            wksrpt.Range(Chr(UBound(strTitle1) + j + 67) & AgentRow).Value = strExc(0)
        Next j
        AgentRow = AgentRow + 1
    End If
    strOldCurr = "" & adoquery.Fields("Currency")
    adoquery.MoveNext
Next i

wksrpt.Range("a1:" & strMaxField & intHeadRow - 1).Font.Bold = True
'合計
'資料性質 6(未收未付)
If Text3 = "6" Then
    'Modify by Amy 2024/07/15 拿掉if 因Y51566及Y52269 不需都帶NTD+台幣金額,故全部使用同一種格式-斯閔
    '非特定代理不需小計
    'If Left(Text1, 6) <> "Y51566" And Left(Text1, 6) <> "Y52269" Then
        'Add by Amy 2019/11/19 當只有台一資料時,會沒小計,ReDim strMoney_T(UBound(strCurr_T))為-1 導致出現「找不到檔案」
        If strSum(0) = MsgText(601) Then
            '顯示最後一筆台一加總
                wksrpt.Range("a" & TaieRow).Value = "小計"
                If TaieRow = intHeadRow + 1 Then
                    wksrpt.Range(Chr(UBound(strTitle1) + 65) & TaieRow).Value = 0
                Else
                    wksrpt.Range(Chr(UBound(strTitle1) + 65) & TaieRow).Formula = "=sum(" & Chr(UBound(strTitle1) + 65) & intStart & ":" & Chr(UBound(strTitle1) + 65) & TaieRow - 1 & ")"
                End If
                wksrpt.Range(Chr(UBound(strTitle1) + 65) & TaieRow - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous '下框線
                strSum(0) = strSum(0) & ";" & strOldCurr & "-" & Chr(UBound(strTitle1) + 65) & TaieRow
                TaieRow = TaieRow + 2
        End If
        'end 2019/11/19
        '顯示最後一筆代理人加總
        wksrpt.Range(Chr(UBound(strTitle1) + 67) & AgentRow).Value = "小計"
        'Modify by Amy 2024/06/25 原67 ex:Chr(UBound(strTitle1) + 67,增加SoA欄位
        wksrpt.Range(Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow).NumberFormatLocal = """" & strOldCurr & """#,##0.00_);[紅色](" & """" & strOldCurr & """#,##0.00)"
        'Modify by Amy 2019/11/19 +if 沒資料時,公式會加到欄位名稱
        If intStart = AgentRow Then
            wksrpt.Range(Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow).Value = 0
        Else
            wksrpt.Range(Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow).Formula = "=sum(" & Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & intStart & ":" & Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow - 1 & ")"
        End If
        wksrpt.Range(Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous '下框線
        strSum(1) = strSum(1) & ";" & strOldCurr & "-" & Chr(UBound(strTitle1) + UBound(strTitle2) + 66) & AgentRow
        'end 2024/06/25
        AgentRow = AgentRow + 2
        '顯示 Total
        strCurr_T = Split(Mid(strSum(0), 2), ";")
        strCurr_A = Split(Mid(strSum(1), 2), ";")
        ReDim strMoney_T(UBound(strCurr_T))
        ReDim strMoney_A(UBound(strCurr_A))
    'End If
    'end 2024/07/15
    
    '多空一行
    If AgentRow >= TaieRow Then
        AgentRow = AgentRow + 1
        intStart = AgentRow
    Else
        TaieRow = TaieRow + 1
        intStart = TaieRow
    End If
    'Modify by Amy 2024/07/15 拿掉if 因Y51566及Y52269 不需都帶NTD+台幣金額,故全部使用同一種格式-斯閔
    '非特定代理依不同幣別Total顯示
'    If Left(Text1, 6) <> "Y51566" And Left(Text1, 6) <> "Y52269" Then
        '台一 Total
        For i = LBound(strCurr_T) To UBound(strCurr_T)
            strMoney_T(i) = Mid(strCurr_T(i), InStr(strCurr_T(i), "-") + 1)
            strCurr_T(i) = Replace(strCurr_T(i), "-" & strMoney_T(i), "")
            
            wksrpt.Range("a" & intStart + i).Value = "Total"
            wksrpt.Range(Chr(UBound(strTitle1) + 65) & intStart + i).NumberFormatLocal = """" & strCurr_T(i) & """#,##0.00_);[紅色](" & """" & strCurr_T(i) & """#,##0.00)"
            wksrpt.Range(Chr(UBound(strTitle1) + 65) & intStart + i).Value = "=" & strMoney_T(i)
        Next i
        '代理人 Total
        For i = LBound(strCurr_A) To UBound(strCurr_A)
            strMoney_A(i) = Mid(strCurr_A(i), InStr(strCurr_A(i), "-") + 1)
            strCurr_A(i) = Replace(strCurr_A(i), "-" & strMoney_A(i), "")
            
            wksrpt.Range(Chr(UBound(strTitle1) + 67) & intStart + i).Value = "Total"
            'Modify by Amy 2017/11/24
            For j = LBound(intSumField2) To UBound(intSumField2)
                wksrpt.Range(Chr(intSumField2(j) + UBound(strTitle1) + 67) & intStart + i).NumberFormatLocal = """" & strCurr_A(i) & """#,##0.00_);[紅色](" & """" & strCurr_A(i) & """#,##0.00)"
                wksrpt.Range(Chr(intSumField2(j) + UBound(strTitle1) + 67) & intStart + i).Value = "=" & strMoney_A(i)
            Next j
            'end 2024/06/25
        Next i
        If UBound(strCurr_T) >= UBound(strCurr_A) Then
            TaieRow = intStart + UBound(strCurr_T)
        Else
            TaieRow = intStart + UBound(strCurr_A)
        End If
'    '特定代理人Total顯示
'    Else
'        TaieRow = intStart
'        wksrpt.Range("a" & TaieRow).Value = "Total"
'        For i = 0 To UBound(intSumField1)
'            wksrpt.Range(Chr(intSumField1(i) + 65) & TaieRow).Formula = "=sum(" & Chr(intSumField1(i) + 65) & intHeadRow + 1 & ":" & Chr(intSumField1(i) + 65) & TaieRow - 1 & ")"
'            strSum(0) = "," & Chr(intSumField1(i) + 65) & TaieRow
'        Next
'
'        wksrpt.Range(Chr(UBound(strTitle1) + 67) & TaieRow).Value = "Total"
'        For i = 0 To UBound(intSumField2)
'            wksrpt.Range(Chr(intSumField2(i) + UBound(strTitle1) + 67) & TaieRow).Formula = "=sum(" & Chr(intSumField2(i) + UBound(strTitle1) + 67) & intHeadRow + 1 & ":" & Chr(intSumField2(i) + UBound(strTitle1) + 67) & TaieRow - 1 & ")"
'            strSum(1) = "," & Chr(intSumField2(i) + UBound(strTitle1) + 67) & TaieRow
'        Next
'        'end 2024/06/25
'    End If
    'end 2024/07/15
    TaieRow = TaieRow + 2
  
'*** Information 有底色區塊 ***
    If bolIsChinese = True Then
        strTmp(0) = "抵帳明細如下："
        strTmp(1) = "    台一應付" & strAgentName & "款項為"
        strTmp(2) = "(約當USD)"
        strTmp(3) = "    " & strAgentName & "應付台一款項為"
        strTmp(4) = "    抵帳後應付金額為"
    Else
        strTmp(0) = "The detail of offset are as follows:"
        strTmp(1) = "    In favor of  " & strAgentName & " is"
        strTmp(2) = "(about USD)"
        strTmp(3) = "    In favor of  Tai E is"
        strTmp(4) = "    Total balance due to"
    End If
    wksrpt.Range("a" & TaieRow).Value = strTmp(0)
    TaieRow = TaieRow + 1: intStart = TaieRow
    'Modify by Amy 2024/07/15 拿掉if 因Y51566及Y52269 不需都帶NTD+台幣金額,故全部使用同一種格式-斯閔
    '非特定代理依不同幣別Total顯示
'    If Left(Text1, 6) <> "Y51566" And Left(Text1, 6) <> "Y52269" Then
        strSum(0) = "": strSum(1) = ""
        For i = LBound(strCurr_A) To UBound(strCurr_A)
            wksrpt.Range("a" & TaieRow).Value = strTmp(1)
            wksrpt.Range("e" & TaieRow).NumberFormatLocal = """" & strCurr_A(i) & """#,##0.00_);[紅色](" & """" & strCurr_A(i) & """#,##0.00)"
            wksrpt.Range("e" & TaieRow).Value = "=" & strMoney_A(i)
            If strCurr_A(i) <> "USD" Then wksrpt.Range("f" & TaieRow).Value = strTmp(2)
            strSum(1) = strSum(1) & ",e" & TaieRow
            TaieRow = TaieRow + 1
        Next i
        For i = LBound(strCurr_T) To UBound(strCurr_T)
            wksrpt.Range("a" & TaieRow).Value = strTmp(3)
            wksrpt.Range("e" & TaieRow).NumberFormatLocal = """" & strCurr_T(i) & """#,##0.00_);[紅色](" & """" & strCurr_T(i) & """#,##0.00)"
            wksrpt.Range("e" & TaieRow).Value = "=" & strMoney_T(i)
            strSum(0) = strSum(0) & ",e" & TaieRow
            If strCurr_T(i) <> "USD" Then wksrpt.Range("f" & TaieRow).Value = strTmp(2)
            If i = UBound(strCurr_T) Then
                wksrpt.Range("e" & TaieRow).Borders(xlEdgeBottom).LineStyle = xlContinuous  '下框線
            End If
            TaieRow = TaieRow + 1
        Next i
'    '特定代理人Total顯示
'    Else
'        wksrpt.Range("a" & TaieRow).Value = strTmp(1)
'        wksrpt.Range("e" & TaieRow).Value = "=" & Mid(strSum(1), 2)
'        wksrpt.Range("f" & TaieRow).Value = strTmp(2)
'        TaieRow = TaieRow + 1
'        wksrpt.Range("a" & TaieRow).Value = strTmp(3)
'        wksrpt.Range("e" & TaieRow).Value = "=" & Mid(strSum(0), 2)
'        wksrpt.Range("e" & TaieRow).Borders(xlEdgeBottom).LineStyle = xlContinuous  '下框線
'        TaieRow = TaieRow + 1
'    End If
    wksrpt.Range("a" & TaieRow).Value = strTmp(4)
    
'    If Left(Text1, 6) = "Y51566" Or Left(Text1, 6) = "Y52269" Then
'        'Y51566 及Y52269 不管請款幣別為何,都帶NTD+台幣金額
'        strExc(0) = "NTD"
'    Else
        '因幣別可能不同故帶第一筆的幣別-莘
        strExc(0) = IIf(strCurr_T(0) <> MsgText(601), strCurr_T(0), strCurr_A(0))
'    End If
    'end 2024/07/15
    wksrpt.Range("e" & TaieRow).NumberFormatLocal = """" & strExc(0) & """#,##0.00_);[紅色](" & """" & strExc(0) & """#,##0.00)"
    wksrpt.Range("e" & TaieRow).Value = "=" & IIf(strSum(1) <> "", "Sum(" & Mid(strSum(1), 2) & ")", "") & _
                                                                    "-" & IIf(strSum(0) = "", 0, "Sum(" & Mid(strSum(0), 2) & ")")
   
    '設定粗體/底色
    wksrpt.Range("a" & intStart - 1 & ":" & strMaxField & TaieRow).Font.Bold = True
    wksrpt.Range("a" & intStart - 1 & ":" & strMaxField & TaieRow).Interior.ColorIndex = 15
'*** end Information 有底色區塊 ***
    
    TaieRow = TaieRow + 2
    'Information 文字
    If bolIsChinese = True Then
        wksrpt.Range("a" & TaieRow).Value = "請在確認明細後,儘速付款."
        wksrpt.Range("a" & TaieRow + 1).Value = "待 貴公司確認後,將會盡速安排付款."
        wksrpt.Range("a" & TaieRow + 3).Value = "台一銀行明細："
    Else
        wksrpt.Range("a" & TaieRow).Value = "Please confirm the above. Upon receiving your confirmation, the payment will be settled."
        wksrpt.Range("a" & TaieRow + 1).Value = "Please settle this payment at your early convenience."
        wksrpt.Range("a" & TaieRow + 3).Value = "Our bank details:"
    End If
    wksrpt.Range("a" & TaieRow & ":a" & TaieRow + 1).Font.Bold = True
    
    TaieRow = TaieRow + 4
    'Modify by Amy 2014/08/29 修改大小寫顯示及拿掉Account No. EUR-婉莘
    'wksrpt.Range("a" & TaieRow).Value = "Name of Account ：TAI E INTERNATIONAL PATENT & LAW OFFICE": TaieRow = TaieRow + 1
    'wksrpt.Range("a" & TaieRow).Value = "Name of Bank ：Bank of  Taiwan": TaieRow = TaieRow + 1
    'wksrpt.Range("a" & TaieRow).Value = "SWIFT CODE：BKTW TWTP": TaieRow = TaieRow + 1
    'wksrpt.Range("a" & TaieRow).Value = "Account No. USD：003007052646": TaieRow = TaieRow + 1
    'wksrpt.Range("a" & TaieRow).Value = "Account No. EUR：003007085127"
    wksrpt.Range("a" & TaieRow).Value = "Name of Account ：Tai E International Patent & Law Office": TaieRow = TaieRow + 1
    wksrpt.Range("a" & TaieRow).Value = "Name of Bank ：Bank of  Taiwan": TaieRow = TaieRow + 1
    wksrpt.Range("a" & TaieRow).Value = "SWIFT CODE：BKTW TWTP": TaieRow = TaieRow + 1
    wksrpt.Range("a" & TaieRow).Value = "Account No.：003007052646"
    'end Information 文字
'Mark by Amy 2024/06/25 4 (CF未付)不使用此格式
''資料性質 4 (CF未付)不會有台一資料
'Else
'    TaieRow = AgentRow
'    wksrpt.Range("a" & TaieRow).Value = "Total"
'    wksrpt.Range(Chr(UBound(strTitle1) + 67) & TaieRow).Value = "Total"
'    For i = 0 To UBound(intSumField1)
'        wksrpt.Range(Chr(intSumField1(i) + 65) & TaieRow).Formula = "=sum(" & Chr(intSumField1(i) + 65) & intHeadRow + 1 & ":" & Chr(intSumField1(i) + 65) & TaieRow - 1 & ")"
'    Next
'    For i = 0 To UBound(intSumField2)
'        wksrpt.Range(Chr(intSumField2(i) + UBound(strTitle1) + 67) & TaieRow).Formula = "=sum(" & Chr(intSumField2(i) + UBound(strTitle1) + 67) & intHeadRow + 1 & ":" & Chr(intSumField2(i) + UBound(strTitle1) + 67) & TaieRow - 1 & ")"
'    Next
End If

'設定字型
wksrpt.Range("a1:" & strMaxField & TaieRow).Font.Name = "新細明體"
wksrpt.Range("a1:" & strMaxField & TaieRow).Font.Name = "Times New Roman"
wksrpt.Range("a1:" & strMaxField & TaieRow).Font.Size = 10
'Mark by Amy 2024/06/25 4-CF未付,不使用此格式
''Add by Amy 2015/11/25 資料性質 4-CF未付,不顯示台一資料
'If Text3 = "4" Then wksrpt.Range("A:" & Chr(UBound(strTitle1) + 66)).Delete Shift:=xlToLeft

'原:抵帳單為橫印(2017/05/24改前)
wksrpt.PageSetup.PaperSize = xlPaperA4 'A4
wksrpt.PageSetup.Orientation = xlPortrait '直印
wksrpt.PageSetup.PrintTitleRows = "$1:$" & intHeadRow '標題列
wksrpt.PageSetup.CenterHeader = ReportTitle '頁首
wksrpt.PageSetup.LeftMargin = 27 '邊界
wksrpt.PageSetup.RightMargin = 27
wksrpt.PageSetup.TopMargin = 90
wksrpt.PageSetup.BottomMargin = 20
wksrpt.PageSetup.PrintGridlines = True '列印格線
wksrpt.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
'end 2017/05/24
adoquery.Close
'Modify by Amy 2015/11/25 判斷若版本2007以上改變存格式,97才能開
If Val(xlsAgentPoint.Version) < 12 Then
    xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
Else
    xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
End If
xlsAgentPoint.Workbooks.Close
xlsAgentPoint.Quit
ExcelSaveNew = True
Exit Function
    
ErrHnd:
    ExcelSaveNew = False
    adoquery.Close
    'Modify by Amy 2015/11/25 判斷若版本2007以上改變存格式,97才能開
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Kill strExcelPath & Left(Text1, 6) & xlsFileName
    If Err.Number <> 0 Then
        MsgBox "未產生Excel (錯誤:" & Err.Description & ")", vbCritical
    End If
End Function

'*************************************************
'  轉成Excel檔案
'  Add by Amy 2020/09/17
'*************************************************
Private Function ExcelSaveNew2() As Boolean
    Dim xlsAgentPoint As New Excel.Application
    Dim wksrpt As New Worksheet
    Dim strQ As String, xlsFileName As String, strAllField As String, strWkName As String
    Dim strCurrF As String, strFAmtF As String, strTmp As String
    Dim strTitle, intWidth, strFormat
    Dim i As Integer, intQ As Integer, intHeadRow As Integer, intXlsSheet As Integer
    Dim IsFirst As Boolean
    
On Error GoTo ErrHnd
    
    strAllField = "本所案號;"
    If Option1(1).Value = True Then strAllField = "代理人編號;"
    strAllField = "日期<br>Date;單據編號<br>DN No.;幣別<br>Currency;外幣金額<br>Amount;公司別;" & _
                      "台幣金額;規費;台幣收款<br>金額;台幣付款<br>金額;" & _
                      "差額;點數;" & strAllField & "代理人帳單號<br>Invoice No.;匯票號碼"
    '設定Excel欄位名稱/寬度/格式
    strTitle = Split(strAllField, ";")
    intWidth = Array(7.5, 9, 7, 10, 6, 9, 9, 9, 9, 9, _
                               6, 11, 11.5, 11.5)
    strFormat = Array("", "", "", "#,##0.00", "", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", "#,##0.00", _
                                "#,##0.00", "", "@", "")
    '檔名
    If Option1(1).Value = True Then
        '本所案號
        xlsFileName = Text4 & Text5 & IIf(Text6 <> "", Text6 & Text7, "")
    Else
        '代理人
        xlsFileName = Left(Text1, 6)
    End If
    If Trim(Text9) = "1" Then
        xlsFileName = "-非J"
    ElseIf Trim(Text9) = "2" Then
        xlsFileName = "-J"
    End If
    xlsFileName = "代理人對帳單 " & xlsFileName & "-" & GetType & strSrvDate(1) & ServerTime & MsgText(43)
    
    If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & xlsFileName
    End If
    
    '工作表
    intXlsSheet = 1
    xlsAgentPoint.SheetsInNewWorkbook = 3  '預設工作表數量
    xlsAgentPoint.Workbooks.add
    'xlsAgentPoint.Visible = True
    
    adoquery.MoveFirst
    Do While adoquery.EOF = False
        If strWkName = MsgText(601) Then strWkName = Left(xlsAgentPoint.Worksheets(1).Name, Len(xlsAgentPoint.Worksheets(1).Name) - 1)
        'Modified by Morgan 2025/2/6
        'If intXlsSheet > 3 And (intXlsSheet <> 1 And Val(xlsAgentPoint.Version) = 15) Then
        '    xlsAgentPoint.Worksheets.add   '插入sheet
        If intXlsSheet > 3 Then
            xlsAgentPoint.Worksheets.add After:=xlsAgentPoint.Worksheets(xlsAgentPoint.Worksheets.Count)
        'end 2025/2/6
        End If
        Set wksrpt = xlsAgentPoint.Worksheets(strWkName & intXlsSheet)
        wksrpt.Activate
        
        intCounter = 1: IsFirst = True
        
        '從PrintData搬過來修改,代理人編號不同換頁
        strQ = "Select R43201 as FagentNo,R43202 as DocDate,R43203 as DocNo,R43204 as Currency,R43205 as Famount,R43206 as Namount,R43207 as Tamount,R43208 as Ramount,R43209 as Pamount,R43210 as Vamount," & _
                    "R43211 as Caseno,R43212 as Point,R43213 as DNno,R43214 as Checkno,R43215 as fa05,R43216 as fa32,R43217 as fa33,R43218 as fa34,R43219 as fa35,R43220 as fa36,R43221 as fa18,R43222 as fa19,R43223 as fa20," & _
                    "R43224 as fa21,R43225 as fa22,R43226 as fa06,R43227 as fa23,R43228 as a1k01,R43229 as fa70,R43230 as CloseS,R43231 as a1k27,R43232 as a1k28,R43235 as Cmp " & _
                    "From AccRpt432 Where ID='" & strUserNum & "' And R43201='" & adoquery.Fields("FagentNo") & "' Order by R43201 asc,R43234 asc,R43202 asc"
         intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
        If intQ = 1 Then
            RsQ.MoveFirst
            Do While RsQ.EOF = False
                If IsFirst = True Then
                    '設定抬頭
                    Call SetTitle(IIf(Option1(0).Value = True, 1, 2), wksrpt, intCounter, strTitle, intWidth)
                    intHeadRow = intCounter - 1
                    IsFirst = False
                End If
                '資料內容
                For i = LBound(strTitle) To UBound(strTitle)
                    strTmp = ""
                    Select Case strTitle(i)
                        Case "日期<br>Date"
                            If "" & RsQ.Fields("DocDate") <> MsgText(601) Then
                                strTmp = CFDate("" & RsQ.Fields("DocDate"))
                            End If
                        Case "單據編號<br>DN No."
                            strTmp = "" & RsQ.Fields("DocNo")
                        Case "幣別<br>Currency"
                            strTmp = "" & RsQ.Fields("Currency")
                            If intCounter = intHeadRow + 1 Then
                                '記錄 幣別 欄位
                                strCurrF = Chr(i + 65)
                            End If
                        Case "外幣金額<br>Amount"
                            strTmp = "" & RsQ.Fields("Famount")
                            If intCounter = intHeadRow + 1 Then
                                '記錄 外幣金額 欄位
                                strFAmtF = Chr(i + 65)
                            End If
                        Case "公司別"
                            strTmp = "" & RsQ.Fields("Cmp")
                        Case "台幣金額"
                            strTmp = "" & RsQ.Fields("Namount")
                        Case "規費"
                            strTmp = "" & RsQ.Fields("Tamount")
                        Case "台幣收款<br>金額"
                            strTmp = "" & RsQ.Fields("Ramount")
                        Case "台幣付款<br>金額"
                            strTmp = "" & RsQ.Fields("Pamount")
                        Case "差額"
                            strTmp = "" & RsQ.Fields("Vamount")
                        Case "點數"
                            strTmp = "" & RsQ.Fields("Point")
                        Case "代理人編號"
                            strTmp = "" & RsQ.Fields("FagentNo")
                        Case "本所案號"
                            strTmp = "" & RsQ.Fields("CaseNo")
                        Case "代理人帳單號<br>Invoice No."
                            strTmp = "" & RsQ.Fields("DNNo")
                        Case "匯票號碼"
                            strTmp = "" & RsQ.Fields("CheckNo")
                    End Select
                    wksrpt.Range(Chr(i + 65) & intCounter).Value = strTmp
                    wksrpt.Range(Chr(i + 65) & intCounter).NumberFormatLocal = strFormat(i)
                    If strFormat(i) = "#,##0.00" Then
                        wksrpt.Range(Chr(i + 65) & intCounter).HorizontalAlignment = xlRight
                    Else
                        wksrpt.Range(Chr(i + 65) & intCounter).HorizontalAlignment = xlCenter
                    End If
                Next i
                intCounter = intCounter + 1
                RsQ.MoveNext
            Loop
        End If
        '框線
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlEdgeLeft).Weight = xlMedium '線
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlEdgeBottom).Weight = xlMedium '線
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlInsideVertical).Weight = xlThin '細線
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlInsideHorizontal).Weight = xlThin '細線
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlEdgeTop).Weight = xlMedium '線
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        wksrpt.Range(Chr(LBound(strTitle) + 65) & intHeadRow & ":" & Chr(UBound(strTitle) + 65) & intCounter - 1).Borders(xlEdgeRight).Weight = xlMedium '線
        '合計
        Call PrintSum(adoquery.Fields("FagentNo"), True, wksrpt, strCurrF, strFAmtF)
       
        '設定字型
        wksrpt.Range("a1:" & Chr(UBound(strTitle) + 65) & intCounter).Font.Name = "新細明體"
        wksrpt.Range("a1:" & Chr(UBound(strTitle) + 65) & intCounter).Font.Name = "Times New Roman"
        wksrpt.Range("a1:" & Chr(UBound(strTitle) + 65) & intCounter).Font.Size = 10
       
        wksrpt.PageSetup.PaperSize = xlPaperA4 'A4
        wksrpt.PageSetup.Orientation = xlLandscape '橫印
        wksrpt.PageSetup.PrintTitleRows = "$1:$" & intHeadRow '標題列
        wksrpt.PageSetup.LeftMargin = 27 '邊界
        wksrpt.PageSetup.RightMargin = 27
        wksrpt.PageSetup.TopMargin = 90
        wksrpt.PageSetup.BottomMargin = 20
        wksrpt.PageSetup.PrintGridlines = True '列印格線
        wksrpt.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
        
        wksrpt.Name = adoquery.Fields("FagentNo")
        intXlsSheet = intXlsSheet + 1
        adoquery.MoveNext
    Loop
    adoquery.Close

    '判斷版本
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    ExcelSaveNew2 = True
    Exit Function

ErrHnd:
    ExcelSaveNew2 = False
    adoquery.Close
    '判斷版本2007
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Kill strExcelPath & Left(Text1, 6) & xlsFileName
    If Err.Number <> 0 Then
        MsgBox "未產生Excel (錯誤:" & Err.Description & ")", vbCritical
    End If
End Function

Private Sub SetTitle(intChoose As Integer, ByRef Wks As Worksheet, ByRef intHR As Integer, ByRef strTitle, ByRef intWidth)
    Dim i As Integer
    Dim strAddr As String
    
    '代理人格式
    If intChoose = 1 Then
        If IsNull(RsQ.Fields("fa32").Value) Then
            '英文地址
            strAddr = "" & RsQ.Fields("fa18")
            If IsNull(RsQ.Fields("fa19")) = False Then strAddr = strAddr & " " & RsQ.Fields("fa19")
            If IsNull(RsQ.Fields("fa20")) = False Then strAddr = strAddr & " " & RsQ.Fields("fa20")
            If IsNull(RsQ.Fields("fa21")) = False Then strAddr = strAddr & " " & RsQ.Fields("fa21")
            If IsNull(RsQ.Fields("fa22")) = False Then strAddr = strAddr & " " & RsQ.Fields("fa22")
            If IsNull(RsQ.Fields("fa70")) = False Then strAddr = strAddr & " " & RsQ.Fields("fa70")
        Else
            'POB
            strAddr = "" & RsQ.Fields("fa32")
            If IsNull(RsQ.Fields("fa33")) = False Then strAddr = strAddr & " " & RsQ.Fields("fa33")
            If IsNull(RsQ.Fields("fa34")) = False Then strAddr = strAddr & " " & RsQ.Fields("fa34")
            If IsNull(RsQ.Fields("fa35")) = False Then strAddr = strAddr & " " & RsQ.Fields("fa35")
            If IsNull(RsQ.Fields("fa36")) = False Then strAddr = strAddr & " " & RsQ.Fields("fa36")
        End If
    End If
               
    Wks.Range(Chr(LBound(strTitle) + 65) & intHR).Value = ReportTitle(209)
    Wks.Range(Chr(LBound(strTitle) + 65) & intHR & ":" & Chr(UBound(strTitle) + 65) & intHR).MergeCells = True
    Wks.Range(Chr(LBound(strTitle) + 65) & intHR & ":" & Chr(UBound(strTitle) + 65) & intHR).HorizontalAlignment = xlCenter
    intHR = intHR + 1
    '代理人格式
    If intChoose = 1 Then
        Wks.Range(Chr(LBound(strTitle) + 65) & intHR).Value = "代理人編號：" & RsQ.Fields("FagentNo")
        Wks.Range(Chr(UBound(strTitle) + 62) & intHR).Value = "列印日期：" & CFDate(ACDate(ServerDate))
        intHR = intHR + 1
        Wks.Range(Chr(LBound(strTitle) + 65) & intHR).Value = "代理人名稱：" & RsQ.Fields("fa05")
        Wks.Range(Chr(UBound(strTitle) + 62) & intHR).Value = "列印人員：" & StaffQuery(strUserNum)
        intHR = intHR + 1
        Wks.Range(Chr(LBound(strTitle) + 65) & intHR).Value = "代理人地址：" & strAddr
        intHR = intHR + 2
        Wks.Range(Chr(LBound(strTitle) + 65) & intHR).Value = "往來日期：" & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text & IIf(Option2(1).Value = True, " (當下)", " (最新)")
        Wks.Range(Chr(UBound(strTitle) + 62) & intHR).Value = "資料性質：" & Text3 & "." & GetType
        intHR = intHR + 1
    '本所案號格式
    Else
        Wks.Range(Chr(LBound(strTitle) + 65) & intHR).Value = "本所案號：" & Text4 & "-" & Text5 & "-" & Text6 & "-" & Text7
        Wks.Range(Chr(UBound(strTitle) + 62) & intHR).Value = "列印日期：" & CFDate(ACDate(ServerDate))
        intHR = intHR + 1
        Wks.Range(Chr(LBound(strTitle) + 65) & intHR).Value = "往來日期：" & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text & IIf(Option2(1).Value = True, " (當下)", " (最新)")
        Wks.Range(Chr(UBound(strTitle) + 62) & intHR).Value = "列印人員：" & StaffQuery(strUserNum)
        intHR = intHR + 1
        Wks.Range(Chr(LBound(strTitle) + 65) & intHR).Value = "資料性質：" & Text3 & "." & GetType
        intHR = intHR + 1
    End If
    intHR = intHR + 1
     
    For i = LBound(strTitle) To UBound(strTitle)
        Wks.Range(Chr(i + 65) & intHR).Value = Replace(strTitle(i), "<br>", vbCrLf)
        Wks.Columns(Chr(i + 65) & ":" & Chr(i + 65)).ColumnWidth = intWidth(i)
    Next i
    Wks.Range(Chr(LBound(strTitle) + 65) & intHR & ":" & Chr(UBound(strTitle) + 65) & intHR).HorizontalAlignment = xlCenter
    intHR = intHR + 1
End Sub

Private Function GetType() As String
    Select Case Text3
        Case "1"
            GetType = "FC往來"
        Case "2"
            GetType = "FC未收"
        Case "3"
            GetType = "CF往來"
        Case "4"
            GetType = "CF未付"
        Case "5"
            GetType = "往來"
        Case "6"
            GetType = "未收未付"
    End Select
End Function
'end 2020/09/17

'Add by Amy 2014/04/30 判斷fa31是否為中文，並回傳名稱
Private Function fa31IsChinese(ByRef strNo As String, ByRef strName As String) As Boolean
    Dim Rs As ADODB.Recordset
        
    strName = ""
    strExc(0) = "Select fa31,Decode(fa31,1,Nvl(fa04,Decode(fa05||fa63||fa64||fa65,null,fa06,fa05||' '||fa63||' '||fa64||' '||fa65)),Decode(fa05||fa63||fa64||fa65,null,Nvl(fa04,fa06),fa05||' '||fa63||' '||fa64||' '||fa65)) as FName From Fagent " & _
                     "Where fa01='" & strNo & "' "
    intI = 1
    Set Rs = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        If Rs.Fields("fa31") = "1" Then
            fa31IsChinese = True
        Else
            fa31IsChinese = False
        End If
        strName = Rs.Fields("FName")
        Exit Function
    End If
    fa31IsChinese = False
    
    Set Rs = Nothing
End Function

'Modify by Amy 2016/03/31 依各系統別取得基本檔 彼所案號(Taie)
Private Function GetBaseYourRef(strCaseNo As String, m_bolONo As Boolean) As String
     Dim stCaseNo(1 To 4) As String
     
     If strCaseNo = "" Then Exit Function
     
     If Mid(strCaseNo, Len(strCaseNo) - 2, 3) = "000" Then
        stCaseNo(1) = Mid(strCaseNo, 1, Len(strCaseNo) - 9)
        stCaseNo(2) = Mid(strCaseNo, Len(strCaseNo) - 8, 6)
        stCaseNo(3) = "0"
        stCaseNo(4) = "00"
    Else
        stCaseNo(1) = Mid(strCaseNo, 1, Len(strCaseNo) - 9)
        stCaseNo(2) = Mid(strCaseNo, Len(strCaseNo) - 8, 6)
        stCaseNo(3) = Mid(strCaseNo, Len(strCaseNo) - 2, 1)
        stCaseNo(4) = Mid(strCaseNo, Len(strCaseNo) - 1, 2)
    End If
    GetBaseYourRef = GetYourRefNo1(stCaseNo(1), stCaseNo(2), stCaseNo(3), stCaseNo(4), m_bolONo)
End Function

'Mark by Amy 2015/09/14
''依U單號抓取Acc151的總收文號再抓CaseProgress的 彼所案號
''Modify by Amy 2015/09/11 +參數 strCaseNo
Private Function GetYourRefNo2(strAXF01 As String, strCaseNo As String) As String
'    Dim rs As ADODB.Recordset
'    Dim strQuery As String
'    Dim intR As Integer
'
'    GetYourRefNo2 = ""
'    If strAXF01 = "" Then Exit Function
'    'Add by Amy 2015/09/11 +本所案號,避免同U單號多筆抓錯
'    If strCaseNo <> "" Then
'        strQuery = "And CP01='" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' And CP02='" & Mid(strCaseNo, (Len(strCaseNo) - 9) + 1, 6) & "' " & _
'                        "And CP03='" & Mid(strCaseNo, (Len(strCaseNo) - 3) + 1, 1) & "' And CP04='" & Mid(strCaseNo, (Len(strCaseNo) - 2) + 1, 2) & "'"
'    End If
'    strQuery = "Select CP45 From Acc151,CaseProgress Where AXF01='" & strAXF01 & "' And AXF02=CP09(+) " & strQuery
'    'end 2015/09/11
'    intR = 1
'    Set rs = ClsLawReadRstMsg(intR, strQuery)
'    If intR = 1 Then
'        GetYourRefNo2 = "" & rs.Fields("CP45")
'    End If
'
'    Set rs = Nothing
End Function
''end 2014/09/04
'end 2015/09/14

'Add by Amy 2020/09/17
Private Sub Text3_Validate(Cancel As Boolean)
    Option2(0).Enabled = True
    Option2(1).Enabled = True
    If Val(Text3) = 0 Then Exit Sub
    If Val(Text3) Mod 2 = 1 Then
        Option2(0).Enabled = False
        If Option2(0).Value = True Then
            MsgBox "往來只能選「 當下」"
            Option2(1).Value = True
        End If
    ElseIf Text3 = "6" And Option2(1).Value = True Then
        MsgBox "未收未付只能選「 最新」"
        Option2(1).Enabled = False
        Option2(0).Value = True
    End If
End Sub

'Add by Amy 2014/11/18
Private Sub Text4_GotFocus()
    TextInverse Text4
    CloseIme
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
    CaseQuery
End Sub

Private Sub Text5_GotFocus()
    'Option1(1).Value = True
    InverseTextBox Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
    If Trim(Text4) <> MsgText(601) And Trim(Text5) <> MsgText(601) Then
        Text6 = "0"
        Text7 = "00"
    End If
    
End Sub

Private Sub Text6_GotFocus()
    'Option1(1).Value = True
    InverseTextBox Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
    CaseQuery
End Sub

Private Sub Text7_GotFocus()
    'Option1(1).Value = True
    InverseTextBox Text7
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
    CaseQuery
End Sub

Private Sub CaseQuery()
   Select Case Combo1
      Case ComboItem(121)
         Text8 = CaseNameShow(Text4, Text5, Text6, Text7, 1)
      Case ComboItem(122)
         Text8 = CaseNameShow(Text4, Text5, Text6, Text7, 2)
      Case ComboItem(123)
         Text8 = CaseNameShow(Text4, Text5, Text6, Text7, 3)
   End Select
End Sub

'Add by Amy 2014/11/18 選擇本所案號產生相關明細表
Private Sub PrintData_Case()
    Dim strAmount As String
    Dim intLength As Integer
    Dim Rs As ADODB.Recordset
    
    intCounter = 3
    intRecord = 1
    intPage = 0
    Do While adoquery.EOF = False
        If intPage = 0 Or intRecord > MaxLine Then
            If intPage <> 0 Then
                Printer.NewPage
                intCounter = 3
                intRecord = 1
            End If
            intPage = intPage + 1
            PrintHead_Case
        End If
        
        '單據日期
        Printer.CurrentX = 0
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoquery.Fields("DocDate").Value) = False Then
            Printer.Print CFDate(adoquery.Fields("DocDate").Value)
        End If
        
        '單據編號
        Printer.CurrentX = 1300
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoquery.Fields("DocNo").Value) = False Then
            Printer.Print adoquery.Fields("DocNo").Value
            'Mark by Amy 2015/06/30
'            If InStr("2,4,6", Text3.Text) > 0 Then
'                If Left(adoquery.Fields("DocNo").Value, 1) = "X" Then
'                    intUseI = 0
'                    For i = 0 To intMax
'                        If m_FCSum(i) > 0 Then
'                            If m_FCCur(i) = adoquery.Fields("Currency").Value Then
'                                intUseI = i '累計
'                                Exit For
'                            End If
'                        ElseIf m_FCSum(i) = 0 Then
'                            intUseI = i '使用新陣列
'                            Exit For
'                        End If
'                    Next i
'                    m_FCSum(intUseI) = m_FCSum(intUseI) + Val("" & adoquery.Fields("Famount").Value)
'                    '若有已收金額要扣除
'                    If Val("" & adoquery.Fields("Ramount").Value) > 0 Then
'                        strExc(0) = "SELECT sum(a0z04) FROM acc0z0 WHERE a0z02='" & adoquery.Fields("DocNo").Value & "'"
'                        intI = 1
'                        Set rs = ClsLawReadRstMsg(intI, strExc(0))
'                        If intI = 1 Then
'                            m_FCSum(intUseI) = m_FCSum(intUseI) - Val("" & rs.Fields(0).Value)
'                        End If
'                        rs.Close
'                    End If
'                    m_FCCur(intUseI) = adoquery.Fields("Currency").Value
'                ElseIf Left(adoquery.Fields("DocNo").Value, 1) = "U" Then
'                    intUseI = 0
'                    For i = 0 To intMax
'                        If m_CFSum(i) > 0 Then
'                            If m_CFCur(i) = adoquery.Fields("Currency").Value Then
'                                intUseI = i '累計
'                                Exit For
'                            End If
'                        ElseIf m_CFSum(i) = 0 Then
'                            intUseI = i '使用新陣列
'                            Exit For
'                        End If
'                    Next i
'                    m_CFSum(intUseI) = m_CFSum(intUseI) + Val("" & adoquery.Fields("Famount").Value)
'                    m_CFCur(intUseI) = adoquery.Fields("Currency").Value
'                ElseIf Left(adoquery.Fields("DocNo").Value, 1) = "V" Then
'                     intUseI = 0
'                     For i = 0 To intMax
'                         If m_CFVSum(i) > 0 Then
'                             If m_CFVCur(i) = adoquery.Fields("Currency").Value Then
'                                 intUseI = i '累計
'                                 Exit For
'                             End If
'                         ElseIf m_CFVSum(i) = 0 Then
'                             intUseI = i '使用新陣列
'                             Exit For
'                         End If
'                     Next i
'                     m_CFVSum(intUseI) = m_CFVSum(intUseI) + Val("" & adoquery.Fields("Famount").Value)
'                     m_CFVCur(intUseI) = adoquery.Fields("Currency").Value
'                End If
'            End If
        End If
        
        '幣別
        Printer.CurrentX = 2600
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoquery.Fields("Currency").Value) = False Then
            Printer.Print adoquery.Fields("Currency").Value
        End If
        
        '外幣金額
        If IsNull(adoquery.Fields("Famount").Value) = False Then
            strAmount = Format(Val(adoquery.Fields("Famount").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 4800 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        
        '台幣金額
        If IsNull(adoquery.Fields("Namount").Value) = False And Val("" & adoquery.Fields("Namount").Value) > 0 Then
            strAmount = Format(Val(adoquery.Fields("Namount").Value), FDollar)
        ElseIf IsNull(adoquery.Fields("Ramount").Value) = False And Val("" & adoquery.Fields("Ramount").Value) > 0 Then
            strAmount = Format(Val(adoquery.Fields("Ramount").Value), FDollar)
        ElseIf IsNull(adoquery.Fields("Pamount").Value) = False And Val("" & adoquery.Fields("Pamount").Value) > 0 Then
            strAmount = Format(Val(adoquery.Fields("Pamount").Value), FDollar)
        Else
            strAmount = ""
        End If
        If strAmount <> "" Then
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 6300 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        
        '規費
        If IsNull(adoquery.Fields("Tamount").Value) = False And Val("" & adoquery.Fields("Tamount").Value) > 0 Then
            strAmount = Format(Val(adoquery.Fields("Tamount").Value), FDollar)
            intLength = Printer.TextWidth(strAmount)
            Printer.CurrentX = 7600 - intLength
            Printer.CurrentY = 300 + intCounter * 300
            Printer.Print strAmount
        End If
        
        '是否結清
        Printer.CurrentX = 8500
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoquery.Fields("CloseS").Value) = False Then
            Printer.Print adoquery.Fields("CloseS").Value
        End If
        
        '代理人
        Printer.CurrentX = 9300
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoquery.Fields("FagentNo").Value) = False Then
            Printer.Print "" & adoquery.Fields("FagentNo").Value
        End If
        
        '列印對象
        Printer.CurrentX = 11100
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoquery.Fields("A1K27").Value) = False Then
            Printer.Print adoquery.Fields("A1K27").Value
        End If
        
        '請款對象
        Printer.CurrentX = 12800
        Printer.CurrentY = 300 + intCounter * 300
        If IsNull(adoquery.Fields("A1K28").Value) = False Then
            Printer.Print adoquery.Fields("A1K28").Value
      End If
      intCounter = intCounter + 1
      intRecord = intRecord + 1
      adoquery.MoveNext
    Loop
    '合計
    'Modify by Amy 2015/06/30
    'bolSumTxt = False
    'PrintSum_Case
    Printer.Line (0, 300 + intCounter * 300 + 350)-(16400, 300 + intCounter * 300 + 350)
    intCounter = intCounter + 1
    Call PrintSum(adoquery.Fields("FAgentNo").Value) 'Modify by Amy 2020/09/17 +adoquery.Fields("FAgentNo").Value
    Printer.EndDoc
End Sub

Private Sub SetPrintA4()
    MaxLine = 30 'A4橫印可印列數
    Printer.Orientation = 2 '橫印
End Sub

Private Sub PrintHead_Case()
Dim strSelect As String

   Printer.FontSize = 14
   Printer.CurrentX = 7500
   Printer.CurrentY = 300 + intCounter * 300
   Select Case Text3
      Case "1"
         strSelect = "FC案件往來"
      Case "2"
         strSelect = "FC案件未收"
      Case "3"
         strSelect = "CF案件往來"
      Case "4"
         strSelect = "CF案件未付"
      Case "5"
         strSelect = "案件往來"
      Case "6"
         strSelect = "案件未收未付"
      Case Else
         strSelect = ""
   End Select
   Printer.Print strSelect & "明細表"
   Printer.FontSize = 12
   intCounter = intCounter + 2
   
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印人員: " & StaffQuery(strUserNum)
   
   Printer.CurrentX = 12000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印日期: " & CFDate(ACDate(ServerDate))
   intCounter = intCounter + 1
   
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   If Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 2, 3) = "000" Then
        Printer.Print "本所案號: " & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "-" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 8, 6)
   Else
        Printer.Print "本所案號: " & Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "-" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 8, 6) & "-" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 2, 1) & "-" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 1, 2)
   End If
   Printer.CurrentX = 4000
   Printer.CurrentY = 300 + intCounter * 300
   'Modify by Amy 2016/03/31 +巨京沒彼所案號抓分所案號
   Printer.Print "彼所案號: " & GetBaseYourRef("" & adoquery.Fields("CaseNo"), IIf(Left("" & adoquery.Fields("FagentNo"), 6) = "Y52269", True, False))
  
   Printer.CurrentX = 12000
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "頁次: " & intPage
   intCounter = intCounter + 1

   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "往來日期: " & MaskEdBox1.Text & " ~ " & MaskEdBox2.Text
  
   Printer.CurrentX = 4000
   Printer.CurrentY = 300 + intCounter * 300
   Select Case Text3
      Case "1"
         strSelect = "1. FC往來"
      Case "2"
         strSelect = "2. FC未收"
      Case "3"
         strSelect = "3. CF往來"
      Case "4"
         strSelect = "4. CF未付"
      Case "5"
         strSelect = "5. 往來"
      Case "6"
         strSelect = "6. 未收未付"
      Case Else
         strSelect = ""
   End Select
     Printer.Print "資料性質: " & strSelect
   intCounter = intCounter + 2
   
   Printer.CurrentX = 0
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "單據日期"
   Printer.CurrentX = 1300
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "單據編號"
   Printer.CurrentX = 2600
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "幣別"
   Printer.CurrentX = 3600
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "外幣金額"
   Printer.CurrentX = 5200
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "台幣金額"
   Printer.CurrentX = 6900
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "規費"
   Printer.CurrentX = 7800
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "是否結清"
   Printer.CurrentX = 9500
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "代理人"
   Printer.CurrentX = 11150
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "列印對象"
   Printer.CurrentX = 12850
   Printer.CurrentY = 300 + intCounter * 300
   Printer.Print "請款對象"
   
   Printer.Line (0, 300 + intCounter * 300 + 350)-(16400, 300 + intCounter * 300 + 350)
   intCounter = intCounter + 2
End Sub

Private Sub PrintSum_Case()
'Mark by Amy 2015/06/30 重改後此不使用
'Dim AdoSum1 As New ADODB.Recordset, AdoSum2 As New ADODB.Recordset
'Dim stSQL As String
'Dim strTemp(0 To 3) As String
'
'bolSumTxt = False
'Printer.Line (0, 300 + intCounter * 300 + 200)-(16400, 300 + intCounter * 300 + 200)
'If intRecord > MaxLine Then
'    Printer.NewPage
'    intCounter = 3
'    intRecord = 1
'    intPage = intPage + 1
'    PrintHead_Case
'End If
'intCounter = intCounter + 1
'
''**************** FC ****************
'If Text3 <> "3" And Text3 <> "4" Then
'    If Text3 <> "2" And Text3 <> "6" Then
'        '*** FC 已收 ***
'        stSQL = "Select  a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount from acc0y0, fagent, nation, acc0z0, acc1k0, patent where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04" & strWhere(0) & strWhere(4) & " Group by  a0y03, a0y06"
'        stSQL = stSQL & " union Select  a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount from acc0y0, fagent, nation, acc0z0, acc1k0, trademark where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04" & strWhere(0) & strWhere(4) & " Group by  a0y03, a0y06"
'        stSQL = stSQL & " union Select  a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount from acc0y0, fagent, nation, acc0z0, acc1k0, servicepractice where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04" & strWhere(0) & strWhere(4) & " Group by  a0y03, a0y06"
'        stSQL = stSQL & " union Select  a0y03 as Currency, sum(a0z04) as Famount, sum(a0z04 * a0y04) as Namount, sum(decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09))) as Tamount, a0y06 as Oamount from acc0y0, fagent, nation, acc0z0, acc1k0, LAWCASE where decode(a0y18, 1, substr(a0y07, 1, 8), 2, substr(a0y08, 1, 8), substr(a0y09, 1, 8)) = fa01 (+) and decode(a0y18, 1, substr(a0y07, 9, 1), 2, substr(a0y08, 9, 1), substr(a0y09, 9, 1)) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = LC01 and a1k14 = LC02 and a1k15 = LC03 and a1k16 = LC04" & strWhere(0) & strWhere(4) & " Group by  a0y03, a0y06"
'        stSQL = "Select Currency,sum(Famount) as Famount,sum(Namount) as Namount,sum(Tamount) as Tamount,sum(Oamount) as Oamount From (" & stSQL & ") Group by Currency Order by Currency"
'        If AdoSum1.State = adStateOpen Then AdoSum1.Close
'        AdoSum1.CursorLocation = adUseClient
'        AdoSum1.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'
'        '印合計(依幣別)
'        If AdoSum1.RecordCount > 0 Then
'            Erase strTemp
'            If bolSumTxt = False Then
'                PrintSumXY_Case 99, "合    計       FC 已收", intCounter
'                bolSumTxt = True
'            End If
'            AdoSum1.MoveFirst
'            Do While Not AdoSum1.EOF
'                If intRecord > MaxLine Then
'                    Printer.NewPage
'                    intCounter = 3
'                    intRecord = 1
'                    intPage = intPage + 1
'                    PrintHead_Case
'                End If
'                '幣別
'                If strTemp(0) <> "" & AdoSum1.Fields("Currency").Value Then
'                    PrintSumXY_Case 0, "" & AdoSum1.Fields("Currency").Value, intCounter
'                End If
'
'                strTemp(1) = Format("0", FDollar): strTemp(2) = Format("0", FDollar): strTemp(3) = Format("0", FDollar)
'                If Val("" & AdoSum1.Fields("Famount").Value) <> "0" Then strTemp(1) = AdoSum1.Fields("Famount").Value
'                If Val("" & AdoSum1.Fields("Namount").Value) <> "0" Then strTemp(2) = AdoSum1.Fields("Namount").Value
'                If Val("" & AdoSum1.Fields("Tamount").Value) <> "0" Then strTemp(3) = AdoSum1.Fields("Tamount").Value
'
'                If strTemp(1) <> "0" Then strTemp(1) = Format(strTemp(1), FDollar)
'                If strTemp(2) <> "0" Then strTemp(2) = Format(strTemp(2), FDollar)
'                If strTemp(3) <> "0" Then strTemp(3) = Format(strTemp(3), FDollar)
'
'                PrintSumXY_Case 1, strTemp(1), intCounter '外幣金額合計
'                PrintSumXY_Case 2, strTemp(2), intCounter '台幣金額合計
'                PrintSumXY_Case 3, strTemp(3), intCounter '規費合計
'
'                intCounter = intCounter + 1
'                strTemp(0) = "" & AdoSum1.Fields("Currency").Value
'                AdoSum1.MoveNext
'            Loop
'        End If
'        '*** End FC 已收 ***
'    End If
'
'    '*** FC未收 ***
'    stSQL = "select a1k18 as Currency,(a1k08 - nvl(a1k31, 0)) as Famount,Decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,Decode(nvl(a1k30,0),0,a1k09,Decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Tamount from acc1k0, patent where a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'    stSQL = stSQL & " union select a1k18 as Currency,(a1k08 - nvl(a1k31, 0)) as Famount,Decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,Decode(nvl(a1k30,0),0,a1k09,Decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Tamount from acc1k0, trademark where a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'    stSQL = stSQL & " union select a1k18 as Currency,(a1k08 - nvl(a1k31, 0)) as Namount,Decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,Decode(nvl(a1k30,0),0,a1k09,Decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Tamount from acc1k0, servicepractice where a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'    stSQL = stSQL & " union select a1k18 as Currency,(a1k08 - nvl(a1k31, 0)) as Famount,Decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,Decode(nvl(a1k30,0),0,a1k09,Decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Tamount from acc1k0 where (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'    stSQL = stSQL & " union select a1k18 as Currency,a0z04 * (-1) as Famount,Decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,Decode(nvl(a1k30,0),0,a1k09,Decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Tamount from acc0y0, acc0z0, acc1k0, patent where a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 and a1k14 = pa02 and a1k15 = pa03 and a1k16 = pa04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k30>0"
'    stSQL = stSQL & " union select a1k18 as Currency,a0z04 * (-1) as Famount,Decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,Decode(nvl(a1k30,0),0,a1k09,Decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Tamount from acc0y0, acc0z0, acc1k0, trademark where a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = tm01 and a1k14 = tm02 and a1k15 = tm03 and a1k16 = tm04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k30>0"
'    stSQL = stSQL & " union select a1k18 as Currency,a0z04 * (-1) as Famount,Decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,Decode(nvl(a1k30,0),0,a1k09,Decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Tamount from acc0y0, acc0z0, acc1k0, servicepractice where a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = sp01 and a1k14 = sp02 and a1k15 = sp03 and a1k16 = sp04 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k30>0"
'    stSQL = stSQL & " union select a1k18 as Currency,a0z04 * (-1) as Famount,Decode(nvl(a1k30,0),0,(a1k11 - nvl(a1k06, 0)),round((a1k11 - nvl(a1k06, 0) - nvl(a1k30,0)),2)) as Namount,Decode(nvl(a1k30,0),0,a1k09,Decode(sign(a1k30-a1k09),-1,a1k09-a1k30,0)) as Tamount from acc0y0, acc0z0, acc1k0 where a0y01 = a0z01 and a0z02 = a1k01 and (a1k12 is null or a1k12 = 0) and a1k25 is null " & strWhere(0) & " and (a1k29 is null or a1k29 = '') and a1k30>0"
'    stSQL = "select Currency,sum(Famount) as Famount,sum(Namount) as Namount,sum(Tamount) as Tamount from (" & stSQL & ") group by Currency order by Currency"
'
'    If AdoSum1.State = adStateOpen Then AdoSum1.Close
'    AdoSum1.CursorLocation = adUseClient
'    AdoSum1.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'
'    '印合計(依幣別)
'    If AdoSum1.RecordCount > 0 Then
'        Erase strTemp
'        If bolSumTxt = False Then
'            PrintSumXY_Case 99, "合    計       FC 未收", intCounter
'            bolSumTxt = True
'        Else
'            PrintSumXY_Case 99, "　    　       FC 未收", intCounter
'        End If
'        AdoSum1.MoveFirst
'        Do While Not AdoSum1.EOF
'            If intRecord > MaxLine Then
'                Printer.NewPage
'                intCounter = 3
'                intRecord = 1
'                intPage = intPage + 1
'                PrintHead_Case
'            End If
'            '幣別
'            If strTemp(0) <> "" & AdoSum1.Fields("Currency").Value Then
'                PrintSumXY_Case 0, "" & AdoSum1.Fields("Currency").Value, intCounter
'            End If
'
'            strTemp(1) = Format("0", FDollar): strTemp(2) = Format("0", FDollar): strTemp(3) = Format("0", FDollar)
'            If Val("" & AdoSum1.Fields("Famount").Value) <> "0" Then strTemp(1) = AdoSum1.Fields("Famount").Value
'            If Val("" & AdoSum1.Fields("Namount").Value) <> "0" Then strTemp(2) = AdoSum1.Fields("Namount").Value
'            If Val("" & AdoSum1.Fields("Tamount").Value) <> "0" Then strTemp(3) = AdoSum1.Fields("Tamount").Value
'
'            If strTemp(1) <> "0" Then strTemp(1) = Format(strTemp(1), FDollar)
'            If strTemp(2) <> "0" Then strTemp(2) = Format(strTemp(2), FDollar)
'            If strTemp(3) <> "0" Then strTemp(3) = Format(strTemp(3), FDollar)
'
'            PrintSumXY_Case 1, strTemp(1), intCounter '外幣金額合計
'            PrintSumXY_Case 2, strTemp(2), intCounter '台幣金額合計
'            PrintSumXY_Case 3, strTemp(3), intCounter '規費合計
'
'            intCounter = intCounter + 1
'            strTemp(0) = "" & AdoSum1.Fields("Currency").Value
'            AdoSum1.MoveNext
'        Loop
'    End If
'    '*** end FC未收 ***
'End If
''**************** End FC ****************
'
''**************** CF ****************
'If Text3 <> "1" And Text3 <> "2" Then
'    If Text3 <> "4" And Text3 <> "6" Then
'        '*** CF已付 ***
'        stSQL = "Select   a1903 as Currency,sum(axf04) as Famount, sum(axf04*a1906) as Namount, 0 as Tamount, 0 as Oamount From acc190, acc180, fagent, nation, acc151, acc150, patent, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(2) & " group by a1903 "
'        stSQL = stSQL & " Union Select  a1903 as Currency,sum(axf04) as Famount, sum(axf04*a1906) as Namount, 0 as Tamount, 0 as Oamount From acc190, acc180, fagent, nation, acc151, acc150, trademark, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(2) & " group by a1903 "
'        stSQL = stSQL & " Union Select a1903 as Currency, sum(axf04) as Famount, sum(axf04*a1906) as Namount, 0 as Tamount, 0 as Oamount From acc190, acc180, fagent, nation, acc151, acc150, servicepractice, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(2) & " group by a1903 "
'        stSQL = stSQL & " Union Select  a1903 as Currency,sum(axf04) as Famount, sum(axf04*a1906) as Namount, 0 as Tamount, 0 as Oamount From acc190, acc180, fagent, nation, acc151, acc150, LAWCASE, acc1b0 where a1901 = a1801 and substr(a1803, 1, 8) = fa01 (+) and substr(a1803, 9, 1) = fa02 (+) and fa10 = na01 (+) and a1902 = a1501 and a1501 = axf01 and substr(axf03, 1, length(axf03) - 9) = LC01 and substr(axf03, length(axf03) - 8, 6) = LC02 and substr(axf03, length(axf03) - 2, 1) = LC03 and substr(axf03, length(axf03) - 1, 2) = LC04 and a1908=a1b01(+) and A1908 IS not NULL" & strWhere(2) & " group by a1903 "
'        stSQL = "Select Currency,sum(Famount) as Famount,sum(Namount) as Namount,sum(Tamount) as Tamount,sum(Oamount) as Oamount From (" & stSQL & ") Group by Currency Order by Currency"
'        If AdoSum1.State = adStateOpen Then AdoSum1.Close
'        AdoSum1.CursorLocation = adUseClient
'        AdoSum1.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'
'        '印合計(依幣別)
'        If AdoSum1.RecordCount > 0 Then
'            Erase strTemp
'            If bolSumTxt = False Then
'                PrintSumXY_Case 99, "合    計       CF 已付", intCounter
'                bolSumTxt = True
'            Else
'                PrintSumXY_Case 99, "　    　       CF 已付", intCounter
'            End If
'            AdoSum1.MoveFirst
'            Do While Not AdoSum1.EOF
'                If intRecord > MaxLine Then
'                    Printer.NewPage
'                    intCounter = 3
'                    intRecord = 1
'                    intPage = intPage + 1
'                    PrintHead_Case
'                End If
'                '幣別
'                If strTemp(0) <> "" & AdoSum1.Fields("Currency").Value Then
'                    PrintSumXY_Case 0, "" & AdoSum1.Fields("Currency").Value, intCounter
'                End If
'
'                strTemp(1) = Format("0", FDollar): strTemp(2) = Format("0", FDollar): strTemp(3) = Format("0", FDollar)
'                If Val("" & AdoSum1.Fields("Famount").Value) <> "0" Then strTemp(1) = AdoSum1.Fields("Famount").Value
'                If Val("" & AdoSum1.Fields("Namount").Value) <> "0" Then strTemp(2) = AdoSum1.Fields("Namount").Value
'                If Val("" & AdoSum1.Fields("Tamount").Value) <> "0" Then strTemp(3) = AdoSum1.Fields("Tamount").Value
'
'                If strTemp(1) <> "0" Then strTemp(1) = Format(strTemp(1), FDollar)
'                If strTemp(2) <> "0" Then strTemp(2) = Format(strTemp(2), FDollar)
'                If strTemp(3) <> "0" Then strTemp(3) = Format(strTemp(3), FDollar)
'
'                PrintSumXY_Case 1, strTemp(1), intCounter '外幣金額合計
'                PrintSumXY_Case 2, strTemp(2), intCounter '台幣金額合計
'                PrintSumXY_Case 3, strTemp(3), intCounter '規費合計
'
'                intCounter = intCounter + 1
'                strTemp(0) = "" & AdoSum1.Fields("Currency").Value
'                AdoSum1.MoveNext
'            Loop
'        End If
'        '*** End CF已付 ***
'    End If
'
'    '*** CF未付 ***
'    stSQL = "select a1505,sum(axf04) as Famount from acc151, acc150, patent, acc190 where a1902(+) = axf01 and axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = pa01 and substr(axf03, length(axf03) - 8, 6) = pa02 and substr(axf03, length(axf03) - 2, 1) = pa03 and substr(axf03, length(axf03) - 1, 2) = pa04 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null group by a1505"
'    stSQL = stSQL & " union select a1605 as a1505,sum(axg04 * (-1)) as Famount from acc161, acc160, patent where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = pa01 and substr(axg03, length(axg03) - 8, 6) = pa02 and substr(axg03, length(axg03) - 2, 1) = pa03 and substr(axg03, length(axg03) - 1, 2) = pa04 and a1607 is null" & strWhere(3) & " group by a1605"
'    stSQL = stSQL & " union select a1505,sum(axf04) as Famount from acc151, acc150, trademark, acc190 where a1902(+) = axf01 and axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = tm01 and substr(axf03, length(axf03) - 8, 6) = tm02 and substr(axf03, length(axf03) - 2, 1) = tm03 and substr(axf03, length(axf03) - 1, 2) = tm04 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null group by a1505"
'    stSQL = stSQL & " union select a1605 as a1505,sum(axg04 * (-1)) as Famount from acc161, acc160, trademark where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = tm01 and substr(axg03, length(axg03) - 8, 6) = tm02 and substr(axg03, length(axg03) - 2, 1) = tm03 and substr(axg03, length(axg03) - 1, 2) = tm04 and a1607 is null" & strWhere(3) & " group by a1605"
'    stSQL = stSQL & " union select a1505,sum(axf04) as Famount from acc151, acc150, servicepractice, acc190 where a1902(+) = axf01 and axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = sp01 and substr(axf03, length(axf03) - 8, 6) = sp02 and substr(axf03, length(axf03) - 2, 1) = sp03 and substr(axf03, length(axf03) - 1, 2) = sp04 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null group by a1505"
'    stSQL = stSQL & " union select a1605 as a1505,sum(axg04 * (-1)) as Famount from acc161, acc160, servicepractice where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = sp01 and substr(axg03, length(axg03) - 8, 6) = sp02 and substr(axg03, length(axg03) - 2, 1) = sp03 and substr(axg03, length(axg03) - 1, 2) = sp04 and a1607 is null" & strWhere(3) & " group by a1605"
'    stSQL = stSQL & " union select a1505,sum(axf04) as Famount from acc151, acc150, lawcase, acc190 where a1902(+) = axf01 and axf01 = a1501 and substr(axf03, 1, length(axf03) - 9) = lc01 and substr(axf03, length(axf03) - 8, 6) = lc02 and substr(axf03, length(axf03) - 2, 1) = lc03 and substr(axf03, length(axf03) - 1, 2) = lc04 and (a1507 is null or a1507 = 0)" & strWhere(2) & " and (a1520 = 0 or a1520 is null) and a1512 is null and a1908 is null group by a1505"
'    stSQL = stSQL & " union select a1605 as a1505,sum(axg04 * (-1)) as Famount from acc161, acc160, LAWCASE where axg01 = a1601 and substr(axg03, 1, length(axg03) - 9) = lc01 and substr(axg03, length(axg03) - 8, 6) = lc02 and substr(axg03, length(axg03) - 2, 1) = lc03 and substr(axg03, length(axg03) - 1, 2) = lc04 and a1607 is null" & strWhere(3) & " group by a1605"
'    stSQL = "select a1505 as Currency,sum(Famount) as Famount from (" & stSQL & ") group by a1505 order by a1505"
'    If AdoSum1.State = adStateOpen Then AdoSum1.Close
'    AdoSum1.CursorLocation = adUseClient
'    AdoSum1.Open stSQL, adoTaie, adOpenStatic, adLockReadOnly
'
'    '印合計(依幣別)
'    If AdoSum1.RecordCount > 0 Then
'        Erase strTemp
'        If bolSumTxt = False Then
'            PrintSumXY_Case 99, "合    計       CF 未付", intCounter
'            bolSumTxt = True
'        Else
'            PrintSumXY_Case 99, "　    　       CF 未付", intCounter
'        End If
'        AdoSum1.MoveFirst
'        Do While Not AdoSum1.EOF
'            If intRecord > MaxLine Then
'                Printer.NewPage
'                intCounter = 3
'                intRecord = 1
'                intPage = intPage + 1
'                PrintHead_Case
'            End If
'            '幣別
'            If strTemp(0) <> "" & AdoSum1.Fields("Currency").Value Then
'                PrintSumXY_Case 0, "" & AdoSum1.Fields("Currency").Value, intCounter
'            End If
'
'            strTemp(1) = Format("0", FDollar): strTemp(2) = Format("0", FDollar): strTemp(3) = Format("0", FDollar)
'            If Val("" & AdoSum1.Fields("Famount").Value) <> "0" Then strTemp(1) = AdoSum1.Fields("Famount").Value
'
'            If strTemp(1) <> "0" Then strTemp(1) = Format(strTemp(1), FDollar)
'
'            PrintSumXY_Case 1, strTemp(1), intCounter '外幣金額合計
'
'            intCounter = intCounter + 1
'            strTemp(0) = "" & AdoSum1.Fields("Currency").Value
'            AdoSum1.MoveNext
'        Loop
'    End If
'    '*** End CF未付 ***
'End If
''**************** End CF ****************
End Sub

Private Sub PrintSumXY_Case(ByVal intChoose As Integer, ByVal stAmount As String, ByVal intCount As Integer)
'Mark by Amy 2015/06/30 重改後改不用
'    Dim intLength As Integer
'
'    Select Case intChoose
'        Case 0
'            '幣別
'            Printer.CurrentX = 2600
'            Printer.CurrentY = 300 + intCounter * 300
'            Printer.Print stAmount
'        Case 1
'             '外幣金額合計
'            intLength = Printer.TextWidth(stAmount)
'            Printer.CurrentX = 4800 - intLength
'            Printer.CurrentY = 300 + intCounter * 300
'            Printer.Print stAmount
'        Case 2
'            '台幣金額合計
'            intLength = Printer.TextWidth(stAmount)
'            Printer.CurrentX = 6300 - intLength
'            Printer.CurrentY = 300 + intCounter * 300
'            Printer.Print stAmount
'        Case 3
'            '規費合計
'            intLength = Printer.TextWidth(stAmount)
'            Printer.CurrentX = 7600 - intLength
'            Printer.CurrentY = 300 + intCounter * 300
'            Printer.Print stAmount
'        Case Else
'            Printer.CurrentX = 300
'            Printer.CurrentY = 300 + intCounter * 300
'            Printer.Print stAmount '單據日期欄
'    End Select
    
End Sub

Public Sub PrintData_Old()
'Mark by Amy 2015/06/30
''add by nickc 2007/02/08
'Dim strNo As String
'Dim strAmount As String
'Dim intLength As Integer
'Dim rs As ADODB.Recordset 'Add By Sindy 2012/12/21
'Dim bolIsHaveData As Boolean 'Add By Sindy 2013/1/30
''Add by Amy 2014/11/18
'Dim strAGField As String, strCaseField As String, strTField As String
'
'On Error GoTo Checking
'
'   strSql = ""
'   'Modify By Sindy 2012/12/21 清空陣列值
'   For i = 0 To intMax
'      m_FCCur(i) = "": m_CFCur(i) = "": m_CFVCur(i) = ""
'      m_FCSum(i) = 0: m_CFSum(i) = 0: m_CFVSum(i) = 0
'   Next i
'   '2012/12/21 End
'   For intCounter = 0 To 5
'      strWhere(intCounter) = ""
'   Next intCounter
'   If adoquery.State = adStateOpen Then
'      adoquery.Close
'   End If
'   adoquery.CursorLocation = adUseClient
'
'   'Modify by Amy 2014/11/18 +Option
'   strAGField = ""
'   If Option1(0).Value = True Then
'    strAGField = " a1k28 " '請款對象
'    strCaseField = " '' " 'Acc0y0抓案件編號
'    strTField = " 0 " 'Acc0y0抓規費
'    If Text1 <> "" Then
'      strWhere(0) = strWhere(0) & " and A1K28 >= '" & Text1 & "'"   '2006/3/27 MODIFY BY SONIA a1k03-->A1K28
'      strWhere(1) = strWhere(1) & " and a1203 >= '" & Text1 & "'"
'      strWhere(2) = strWhere(2) & " and a1503 >= '" & Text1 & "'"
'      strWhere(3) = strWhere(3) & " and a1603 >= '" & Text1 & "'"
'      strWhere(4) = strWhere(4) & " and decode(a0y18, '1', a0y07, '2', a0y08, a0y09) >= '" & Text1 & "'"
'    End If
'    If Text2 <> "" Then
'      strWhere(0) = strWhere(0) & " and A1K28 <= '" & Text2 & "'"   '2006/3/27 MODIFY BY SONIA a1k03-->A1K28
'      strWhere(1) = strWhere(1) & " and a1203 <= '" & Text2 & "'"
'      strWhere(2) = strWhere(2) & " and a1503 <= '" & Text2 & "'"
'      strWhere(3) = strWhere(3) & " and a1603 <= '" & Text2 & "'"
'      strWhere(4) = strWhere(4) & " and decode(a0y18, '1', a0y07, '2', a0y08, a0y09) <= '" & Text2 & "'"
'    End If
'   End If
'   If Option1(1).Value = True Then
'        strAGField = " a1k03 " '本所案號抓代理人
'        strCaseField = " a1k03||a1k14||a1k15||a1k16 " 'Acc0y0抓案件編號
'        strTField = " decode(nvl(a1k30,0),0,0,decode(sign(a1k30-a1k09),-1,a1k30,a1k09)) " 'Acc0y0抓規費編號
'        If Trim(Text6) = MsgText(601) Then Text6 = "0"
'        If Trim(Text7) = MsgText(601) Then Text7 = "00"
'
'        strWhere(0) = strWhere(0) & " and A1K13 = '" & Text4 & "' And A1K14='" & Text5 & "' And A1K15='" & Text6 & "' And A1K16='" & Text7 & "' "
'        strWhere(1) = strWhere(1) & " and a1208 = '" & Text4 & Text5 & Text6 & Text7 & "' "
'        strWhere(2) = strWhere(2) & " and axf03 = '" & Text4 & Text5 & Text6 & Text7 & "' "
'        strWhere(3) = strWhere(3) & " and axg03 = '" & Text4 & Text5 & Text6 & Text7 & "' "
'        strWhere(4) = strWhere(4) & strWhere(0)
'   End If
'   'end 2014/11/18
'
'   If MaskEdBox1.Text <> MsgText(29) And MaskEdBox1.Text <> "" Then
'      strWhere(0) = strWhere(0) & " and a1k02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'      strWhere(1) = strWhere(1) & " and a1202 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'      strWhere(2) = strWhere(2) & " and a1502 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'      strWhere(3) = strWhere(3) & " and a1602 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'      strWhere(4) = strWhere(4) & " and a0y02 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> "" Then
'      strWhere(0) = strWhere(0) & " and a1k02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'      strWhere(1) = strWhere(1) & " and a1202 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'      strWhere(2) = strWhere(2) & " and a1502 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'      strWhere(3) = strWhere(3) & " and a1602 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'      strWhere(4) = strWhere(4) & " and a0y02 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   '2006/3/27 MODIFY BY SONIA 已作廢資料不印
'   '2007/11/7 MODIFY BY SONIA 已銷帳資料不印
'   'Modify by Morgan 2007/5/14 union --> union all (否則相同帳單相同案號多筆收文號且金額相同時會只顯示一筆)
'   '2010/10/28 modify by sonia 因請款金額都顯示美金,婧瑄說那幣別A1K18全部顯示USD不抓A1K18
'   'Modify By Sindy 2012/12/7 (a1k08 - nvl(a1k06, 0)) as Famount ==> (a1k08 - nvl(a1k31, 0)) as Famount
'   '                          (a1k11 - nvl(a1k06, 0) * a1k10) as Namount ==> (a1k11 - nvl(a1k06, 0)) as Namount
'   '                          (DECODE(a1k30,0,NULL,A1K30) - (a1k11 - nvl(a1k06, 0) * a1k10)) as Vamount ==> (DECODE(a1k30,0,NULL,A1K30) - (a1k11 - nvl(a1k06, 0))) as Vamount
'   Select Case Text3
'      Case "1"
'         'Modify by Morgan 2008/4/28 請款單的表頭改抓請款對象
'         'strSQL = "select a1k03 as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (a1k30 - (a1k11 - nvl(a1k06, 0) * a1k10)) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+)" & strWhere(0)
'         '2011/3/25 MODIFY BY SONIA X09500811之A1K30=0時會出現差額-16100
'         'strSql = "select a1k28 as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, 'USD' as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (a1k30 - (a1k11 - nvl(a1k06, 0) * a1k10)) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+)" & strWhere(0)
'         'Modify by Morgan 2011/5/25 +fa70
'         '2012/2/22 MODIFY BY SONIA 無英文名稱抓中文否則抓日文,地址欄也是 Y47804
'         'Modify By Sindy 2012/12/7 原程式抓a0z03改抓a0y03
'         'Modify by Amy 2014/11/18 改FagentNo (選請款對象代理人抓a1k28, 本所案號抓a1k03 原:a1k28) +A1K29/A1K27/A1K28/strTField
'         strSql = "select " & strAGField & " as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (DECODE(a1k30,0,NULL,A1K30) - (a1k11 - nvl(a1k06, 0))) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01,fa70,a1k29,a1k27,a1k28 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(" & strAGField & ", 1, 8) = fa01 (+) and substr(" & strAGField & ", 9, 1) = fa02 (+)" & strWhere(0)
'         strSql = strSql & " union all select distinct a1203 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1204 as Currency, a1207 as Famount, a1207 * a1205 as Namount, 0 as Tamount, a1207 * a1205 as Ramount, 0 as Pamount, 0 as Vamount, a1208 as Caseno, (a1207 * a1205 / 1000) as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, '' as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc120, fagent where substr(a1203, 1, 8) = fa01 (+) and substr(a1203, 9, 1) = fa02 (+)" & strWhere(1)
'         strSql = strSql & " union all select distinct decode(a0y18, '1', a0y07, '2', a0y08, a0y09) as FagentNo, a0y02 as DocDate, a0y01 as DocNo, a0y03 as Currency, a0z04 as Famount, 0 as Namount, " & strTField & " as Tamount, (a0z04 * decode(a0y03, 'NT$', 1, a0y04))  as Ramount, 0 as Pamount, 0 as Vamount, a1k13||a1k14||a1k15||a1k16 as caseno, ((a0z04 * decode(a0y03, 'NT$', 1, a0y04)) / 1000) as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc0y0, fagent, nation, acc0z0, acc1k0, patent where substr(a0y07, 1, 8) = fa01 (+) and substr(a0y07, 9, 1) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 (+) and a1k14 = pa02 (+) and a1k15 = pa03 (+) and a1k16 = pa04 (+)" & strWhere(4)
'      Case "2"
'         'Modify by Morgan 2008/4/28 請款單的表頭改抓請款對象
'         'strSQL = "select a1k03 as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (a1k30 - (a1k11 - nvl(a1k06, 0) * a1k10)) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         '2011/3/25 MODIFY BY SONIA X09500811之A1K30=0時會出現差額-16100
'         'strSql = "select a1k28 as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, 'USD' as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (a1k30 - (a1k11 - nvl(a1k06, 0) * a1k10)) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         'Modify by Morgan 2011/5/25 +fa70
'         'Modify by Amy 2014/11/18 選請款對象代理人抓a1k28, 本所案號抓a1k03 原:a1k28
'         strSql = "select " & strAGField & " as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (DECODE(a1k30,0,NULL,A1K30) - (a1k11 - nvl(a1k06, 0))) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01,fa70,a1k29,a1k27,a1k28 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(" & strAGField & ", 1, 8) = fa01 (+) and substr(" & strAGField & ", 9, 1) = fa02 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'      Case "3"
'         'Modify by Morgan 2011/5/25 +fa70
'         strSql = "select a1503 as FagentNo, a1502 as DocDate, decode(a1507, null, a1501, a1501||'*') as DocNo, a1505 as Currency, axf04 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, 0 as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1501 as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc151, acc150, acc1c0, acc190, fagent where A1507 IS NULL AND substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and axf01 = a1501 and axf01 = a1c03 (+) and axf01 = a1902 (+)" & strWhere(2)
'         strSql = strSql & " union all select distinct a1603 as FagentNo, a1602 as DocDate, a1601 as DocNo, a1605 as Currency, axg04 as Famount, 0 as Namount, 0 as Tamount, decode(a1c01, null, 0, axg04 * a1906) as Ramount, 0 as Pamount, 0 as Vamount, axg03 as Caseno, 0 as Point, a1604 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1601 as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc161, acc160, acc1c0, acc190, fagent where substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and axg01 = a1601 and axg01 = a1c03 (+) and axg01 = a1902 (+)" & strWhere(3)
'         '2007/11/27 MODIFY BY SONIA 因W09601477之U09604115有二案號,故將外幣金額由A1904改為AXF04
'         'strSQL = strSQL & " union all select a1503 as FagentNo, a1b03 as DocDate, a1901 as DocNo, a1505 as Currency, a1904 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, decode(a1c01, null, 0, axf04 * a1906) as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, a1c01 as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1501 as a1k01 from acc151, acc150, acc1c0, acc1b0, acc190, fagent where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and axf01 = a1501 and axf01 = a1c03 (+) and a1c01 = a1b01 (+) and a1c02 = a1b02 (+) and axf01 = a1902 (+) and a1901 is not null" & strWhere(2)
'         'Modify by Morgan 2011/5/25 +fa70
'         strSql = strSql & " union all select a1503 as FagentNo, a1b03 as DocDate, a1901 as DocNo, a1505 as Currency, axf04 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, decode(a1c01, null, 0, axf04 * a1906) as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1501 as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc151, acc150, acc1c0, acc1b0, acc190, fagent where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and axf01 = a1501 and axf01 = a1c03 (+) and a1c01 = a1b01 (+) and a1c02 = a1b02 (+) and axf01 = a1902 (+) and a1901 is not null" & strWhere(2)
'      Case "4"
'         'Modify by Morgan 2011/5/25 +fa70
'         strSql = "select a1503 as FagentNo, a1502 as DocDate, decode(a1507, null, a1501, a1501||'*') as DocNo, a1505 as Currency, axf04 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, decode(a1c01, null, 0, axf04 * a1906) as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, '' as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc151, acc150, acc1c0, acc190, fagent where A1507 IS NULL AND substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and axf01 = a1501 and axf01 = a1c03 (+) and axf01 = a1902 (+)" & strWhere(2) & " and (a1520 is null or a1520 = 0)"
'         '2006/3/27 MODIFY BY SONIA
'         'strSQL = strSQL & " union select a1603 as FagentNo, a1602 as DocDate, a1601 as DocNo, a1605 as Currency, axg04 as Famount, 0 as Namount, 0 as Tamount, decode(a1c01, null, 0, axg04 * a1906) as Ramount, 0 as Pamount, 0 as Vamount, axg03 as Caseno, 0 as Point, a1604 as DNno, a1c01 as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1601 as a1k01 from acc161, acc160, acc1c0, acc190, fagent, acc170 where substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and axg01 = a1601 and axg01 = a1c03 (+) and axg01 = a1902 (+)" & strWhere(3) & " and a1702(+)=a1601 and a1701(+)='2' and a1701 is null"
'         'Modify by Morgan 2011/5/25 +fa70
'         strSql = strSql & "  union all select distinct a1603 as FagentNo, a1602 as DocDate, a1601 as DocNo, a1605 as Currency, axg04 as Famount, 0 as Namount, 0 as Tamount, decode(a1c01, null, 0, axg04 * a1906) as Ramount, 0 as Pamount, 0 as Vamount, axg03 as Caseno, 0 as Point, a1604 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1601 as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc161, acc160, acc1c0, acc190, fagent, acc170 where A1607 IS NULL AND substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and axg01 = a1601 and axg01 = a1c03 (+) and axg01 = a1902 (+) AND A1901 IS NULL " & strWhere(3) & " and a1702(+)=a1601 and a1701(+)='2' and a1701 is null"
'         '2006/3/27 END
'      Case "6"
'         'Modify by Morgan 2008/4/28 請款單的表頭改抓請款對象
'         'strSQL = "select a1k03 as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (a1k30 - (a1k11 - nvl(a1k06, 0) * a1k10)) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         '2011/3/25 MODIFY BY SONIA X09500811之A1K30=0時會出現差額-16100
'         'strSql = "select a1k28 as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, 'USD' as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (a1k30 - (a1k11 - nvl(a1k06, 0) * a1k10)) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         'Modify by Morgan 2011/5/25 +fa70
'         'Modify by Sindy 2013/1/23 +FRamount
'         'Modify by Amy 2014/11/18 選請款對象代理人抓a1k28, 本所案號抓a1k03 原:a1k28
'         strSql = "select " & strAGField & " as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (DECODE(a1k30,0,NULL,A1K30) - (a1k11 - nvl(a1k06, 0))) as Vamount, DECODE(nvl(a1k30,0),0,0,A1K30/A1K10) as FRamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01,fa70,a1k29,a1k27,a1k28 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(" & strAGField & ",1, 8) = fa01 (+) and substr(" & strAGField & ", 9, 1) = fa02 (+)" & strWhere(0) & " and (a1k29 is null or a1k29 = '')"
'         strSql = strSql & " union all select a1503 as FagentNo, a1502 as DocDate, decode(a1507, null, a1501, a1501||'*') as DocNo, a1505 as Currency, axf04 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, decode(a1c01, null, 0, axf04 * a1906) as Pamount, 0 as Vamount, 0 as FRamount, axf03 as Caseno, 0 as Point, a1504 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, '' as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc151, acc150, acc1c0, acc190, fagent where A1507 IS NULL AND substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and axf01 = a1501 and axf01 = a1c03 (+) and axf01 = a1902 (+)" & strWhere(2) & " and (a1520 is null or a1520 = 0)"
'         '2006/3/27 MODIFY BY SONIA
'         'strSQL = strSQL & " union select a1603 as FagentNo, a1602 as DocDate, a1601 as DocNo, a1605 as Currency, axg04 as Famount, 0 as Namount, 0 as Tamount, decode(a1c01, null, 0, axg04 * a1906) as Ramount, 0 as Pamount, 0 as Vamount, axg03 as Caseno, 0 as Point, a1604 as DNno, a1c01 as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1601 as a1k01 from acc161, acc160, acc1c0, acc190, fagent, acc170 where substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and axg01 = a1601 and axg01 = a1c03 (+) and axg01 = a1902 (+)" & strWhere(3) & " and a1702(+)=a1601 and a1701(+)='2' and a1701 is null"
'         'Modify by Morgan 2011/5/25 +fa70
'         strSql = strSql & "  union all select distinct a1603 as FagentNo, a1602 as DocDate, a1601 as DocNo, a1605 as Currency, axg04 as Famount, 0 as Namount, 0 as Tamount, decode(a1c01, null, 0, axg04 * a1906) as Ramount, 0 as Pamount, 0 as Vamount, 0 as FRamount, axg03 as Caseno, 0 as Point, a1604 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1601 as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc161, acc160, acc1c0, acc190, fagent, acc170 where A1607 IS NULL AND substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and axg01 = a1601 and axg01 = a1c03 (+) and axg01 = a1902 (+) AND A1901 IS NULL " & strWhere(3) & " and a1702(+)=a1601 and a1701(+)='2' and a1701 is null"
'         '2006/3/27 END
'      Case Else
'         'Modify by Morgan 2008/4/28 請款單的表頭改抓請款對象
'         'strSQL = "select a1k03 as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (a1k30 - (a1k11 - nvl(a1k06, 0) * a1k10)) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(a1k03, 1, 8) = fa01 (+) and substr(a1k03, 9, 1) = fa02 (+)" & strWhere(0)
'         '2011/3/25 MODIFY BY SONIA X09500811之A1K30=0時會出現差額-16100
'         'strSql = "select a1k28 as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, 'USD' as Currency, (a1k08 - nvl(a1k06, 0)) as Famount, (a1k11 - nvl(a1k06, 0) * a1k10) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (a1k30 - (a1k11 - nvl(a1k06, 0) * a1k10)) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(a1k28, 1, 8) = fa01 (+) and substr(a1k28, 9, 1) = fa02 (+)" & strWhere(0)
'         'Modify by Morgan 2011/5/25 +fa70
'         'Modify By Sindy 2012/12/7 原程式抓a0z03改抓a0y03
'         'Modify by Amy 2014/11/18 選請款對象代理人抓a1k28, 本所案號抓a1k03 原:a1k28 /strTField
'         strSql = "select " & strAGField & " as FagentNo, a1k02 as DocDate, decode(a1k12, null, decode(a1k07, null, a1k01, a1k01||'@'), a1k01||'*') as DocNo, a1k18 as Currency, (a1k08 - nvl(a1k31, 0)) as Famount, (a1k11 - nvl(a1k06, 0)) as Namount, a1k09 as Tamount,  a1k30 as Ramount, 0 as Pamount, (DECODE(a1k30,0,NULL,A1K30) - (a1k11 - nvl(a1k06, 0))) as Vamount, a1k13||a1k14||a1k15||a1k16 as Caseno, a1k30 / 1000 as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01,fa70,a1k29,a1k27,a1k28 from acc1k0, fagent where A1K12 IS NULL AND A1K25 IS NULL AND substr(" & strAGField & ",1, 8) = fa01 (+) and substr(" & strAGField & ", 9, 1) = fa02 (+)" & strWhere(0)
'         strSql = strSql & " union all select distinct a1203 as FagentNo, a1202 as DocDate, a1201 as DocNo, a1204 as Currency, a1207 as Famount, a1207 * a1205 as Namount, 0 as Tamount, a1207 * a1205 as Ramount, 0 as Pamount, 0 as Vamount, a1208 as Caseno, (a1207 * a1205 / 1000) as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, '' as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc120, fagent where substr(a1203, 1, 8) = fa01 (+) and substr(a1203, 9, 1) = fa02 (+)" & strWhere(1)
'         strSql = strSql & " union all select distinct decode(a0y18, '1', a0y07, '2', a0y08, a0y09) as FagentNo, a0y02 as DocDate, a0y01 as DocNo, a0y03 as Currency, a0z04 as Famount, 0 as Namount, " & strTField & " as Tamount, (a0z04 * decode(a0y03, 'NT$', 1, a0y04))  as Ramount, 0 as Pamount, 0 as Vamount, a1k13||a1k14||a1k15||a1k16 as caseno, ((a0z04 * decode(a0y03, 'NT$', 1, a0y04)) / 1000) as Point, '' as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc0y0, fagent, nation, acc0z0, acc1k0, patent where substr(a0y07, 1, 8) = fa01 (+) and substr(a0y07, 9, 1) = fa02 (+) and fa10 = na01 (+) and a0y01 = a0z01 and a0z02 = a1k01 and a1k13 = pa01 (+) and a1k14 = pa02 (+) and a1k15 = pa03 (+) and a1k16 = pa04 (+)" & strWhere(4)
'         strSql = strSql & " union all select a1503 as FagentNo, a1502 as DocDate, decode(a1507, null, a1501, a1501||'*') as DocNo, a1505 as Currency, axf04 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, 0 as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, '' as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1501 as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc151, acc150, acc1c0, acc190, fagent where A1507 IS NULL AND substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and axf01 = a1501 and axf01 = a1c03 (+) and axf01 = a1902 (+)" & strWhere(2)
'         strSql = strSql & " union all select distinct a1603 as FagentNo, a1602 as DocDate, a1601 as DocNo, a1605 as Currency, axg04 as Famount, 0 as Namount, 0 as Tamount, decode(a1c01, null, 0, axg04 * a1906) as Ramount, 0 as Pamount, 0 as Vamount, axg03 as Caseno, 0 as Point, a1604 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1601 as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc161, acc160, acc1c0, acc190, fagent where substr(a1603, 1, 8) = fa01 (+) and substr(a1603, 9, 1) = fa02 (+) and axg01 = a1601 and axg01 = a1c03 (+) and axg01 = a1902 (+)" & strWhere(3)
'         '2007/11/27 MODIFY BY SONIA 因W09601477之U09604115有二案號,故將外幣金額由A1904改為AXF04
'         'strSQL = strSQL & " union all select a1503 as FagentNo, a1b03 as DocDate, a1901 as DocNo, a1505 as Currency, a1904 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, decode(a1c01, null, 0, axf04 * a1906) as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, a1c01 as Checkno, fa05, fa32, fa33, fa34, fa35, fa36, fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1501 as a1k01 from acc151, acc150, acc1c0, acc1b0, acc190, fagent where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and axf01 = a1501 and axf01 = a1c03 (+) and a1c01 = a1b01 (+) and a1c02 = a1b02 (+) and axf01 = a1902 (+) and a1901 is not null" & strWhere(2)
'         strSql = strSql & " union all select a1503 as FagentNo, a1b03 as DocDate, a1901 as DocNo, a1505 as Currency, axf04 as Famount, 0 as Namount, 0 as Tamount, 0 as Ramount, decode(a1c01, null, 0, axf04 * a1906) as Pamount, 0 as Vamount, axf03 as Caseno, 0 as Point, a1504 as DNno, a1c01 as Checkno, NVL(FA05,NVL(FA04,FA06)) fa05, fa32, fa33, fa34, fa35, fa36, NVL(FA18,NVL(FA17,FA23)) fa18, fa19, fa20, fa21, fa22, fa06, fa23, a1501 as a1k01,fa70,'' as a1k29,'' as a1k27,'' as a1k28 from acc151, acc150, acc1c0, acc1b0, acc190, fagent where substr(a1503, 1, 8) = fa01 (+) and substr(a1503, 9, 1) = fa02 (+) and axf01 = a1501 and axf01 = a1c03 (+) and a1c01 = a1b01 (+) and a1c02 = a1b02 (+) and axf01 = a1902 (+) and a1901 is not null" & strWhere(2)
'   End Select
'   'Modify by Amy 2014/06/13 修改產生Excel的排序
'   If bolExcel = True Then
'        adoquery.Open "select X.*,decode(substr(DOCNO,1,1),'X',1,'U',2,3) sort from (" & strSql & ") X order by sort asc,FagentNo asc, DocDate asc, DocNo asc", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'        'Modify by Morgan 2006/3/16
'        '改依單據編號 'X','U',V' 的順序排
'        'adoquery.Open strSQL & " order by FagentNo asc, a1k01 asc, DocDate asc, DocNo asc", adoTaie, adOpenStatic, adLockReadOnly
'        adoquery.Open "select X.*,decode(substr(DOCNO,1,1),'X',1,'U',2,3) sort from (" & strSql & ") X order by FagentNo asc, sort asc, DocDate asc, DocNo asc", adoTaie, adOpenStatic, adLockReadOnly
'   End If
'   'end 2014/06/13
'   If adoquery.RecordCount = 0 Then
'      adoquery.Close
'      MsgBox MsgText(28), , MsgText(5)
'      Exit Sub
'   End If
'
'   'Add by Amy 2014/04/30 +產生Excel
'   If bolExcel = True Then
'      If ExcelSaveNew = True Then
'          MsgBox ("Excel檔案已產生！")
'      End If
'      Exit Sub
'   End If
'   'end 2014/04/30
'
'   'Add by Amy 2014/11/18 選擇本所案號報表格式
'   MaxLine = 28 '大報表可印列數
'   If Option1(1).Value = True Then
'        If Printer.PaperSize = 9 Then SetPrintA4
'        PrintData_Case
'        Exit Sub
'   End If
'   'end 2014/11/18
'
'   intCounter = 3
'   intRecord = 1
'   intPage = 0
'   Do While adoquery.EOF = False
'      'Modify by Morgan 2004/10/26 加印合計
'      'If strNo <> adoquery.Fields("FagentNo").Value Or intRecord > 30 Then
'      If strNo <> adoquery.Fields("FagentNo").Value Or intRecord > 28 Then
'         If strNo <> "" Then
'            'Add by Morgan 2004/10/26 印合計
'            If strNo <> adoquery.Fields("FagentNo").Value Then
'               PrintSum
'            End If
'            '2004/10/26
'            Printer.NewPage
'         End If
'         intCounter = 3
'         intRecord = 1
'         intPage = intPage + 1
'         PrintHead
'         strNo = adoquery.Fields("FagentNo").Value
'      End If
'      Printer.CurrentX = 0
'      Printer.CurrentY = 300 + intCounter * 300
'      If IsNull(adoquery.Fields("DocDate").Value) Then
'         Printer.Print ""
'      Else
'         Printer.Print CFDate(adoquery.Fields("DocDate").Value)
'      End If
'      Printer.CurrentX = 1300
'      Printer.CurrentY = 300 + intCounter * 300
'      If IsNull(adoquery.Fields("DocNo").Value) Then
'         Printer.Print ""
'      Else
'         Printer.Print adoquery.Fields("DocNo").Value
'         'Add by Morgan 2004/10/26
'         If InStr("2,4,6", Text3.Text) > 0 Then
'            If Left(adoquery.Fields("DocNo").Value, 1) = "X" Then
'               'Modify By Sindy 2012/12/21 變成陣列值
'               intUseI = 0
'               For i = 0 To intMax
'                  If m_FCSum(i) > 0 Then
'                     If m_FCCur(i) = adoquery.Fields("Currency").Value Then
'                        intUseI = i '累計
'                        Exit For
'                     End If
'                  ElseIf m_FCSum(i) = 0 Then
'                     intUseI = i '使用新陣列
'                     Exit For
'                  End If
'               Next i
'               m_FCSum(intUseI) = m_FCSum(intUseI) + Val("" & adoquery.Fields("Famount").Value)
'               'Add By Sindy 2012/12/21 若有已收金額要扣除
'               If Val("" & adoquery.Fields("Ramount").Value) > 0 Then
'                   strExc(0) = "SELECT sum(a0z04) FROM acc0z0 WHERE a0z02='" & adoquery.Fields("DocNo").Value & "'"
'                   intI = 1
'                   Set rs = ClsLawReadRstMsg(intI, strExc(0))
'                   If intI = 1 Then
'                      m_FCSum(intUseI) = m_FCSum(intUseI) - Val("" & rs.Fields(0).Value)
'                   End If
'                   rs.Close
'               End If
'               'End
'               m_FCCur(intUseI) = adoquery.Fields("Currency").Value
'               '2012/12/21 End
'            ElseIf Left(adoquery.Fields("DocNo").Value, 1) = "U" Then
'               'Modify By Sindy 2012/12/21 變成陣列值
'               intUseI = 0
'               For i = 0 To intMax
'                  If m_CFSum(i) > 0 Then
'                     If m_CFCur(i) = adoquery.Fields("Currency").Value Then
'                        intUseI = i '累計
'                        Exit For
'                     End If
'                  ElseIf m_CFSum(i) = 0 Then
'                     intUseI = i '使用新陣列
'                     Exit For
'                  End If
'               Next i
'               m_CFSum(intUseI) = m_CFSum(intUseI) + Val("" & adoquery.Fields("Famount").Value)
'               m_CFCur(intUseI) = adoquery.Fields("Currency").Value
'               '2012/12/21 End
'            ElseIf Left(adoquery.Fields("DocNo").Value, 1) = "V" Then
'               'Modify By Sindy 2012/12/21 變成陣列值
'               intUseI = 0
'               For i = 0 To intMax
'                  If m_CFVSum(i) > 0 Then
'                     If m_CFVCur(i) = adoquery.Fields("Currency").Value Then
'                        intUseI = i '累計
'                        Exit For
'                     End If
'                  ElseIf m_CFVSum(i) = 0 Then
'                     intUseI = i '使用新陣列
'                     Exit For
'                  End If
'               Next i
'               m_CFVSum(intUseI) = m_CFVSum(intUseI) + Val("" & adoquery.Fields("Famount").Value)
'               m_CFVCur(intUseI) = adoquery.Fields("Currency").Value
'               '2012/12/21 End
'            End If
'         End If
'      End If
'      Printer.CurrentX = 2600
'      Printer.CurrentY = 300 + intCounter * 300
'      If IsNull(adoquery.Fields("Currency").Value) Then
'         Printer.Print ""
'      Else
'         Printer.Print adoquery.Fields("Currency").Value
'      End If
'      If IsNull(adoquery.Fields("Famount").Value) = False Then
'         strAmount = Format(Val(adoquery.Fields("Famount").Value), FDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 4800 - intLength
'         Printer.CurrentY = 300 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoquery.Fields("Namount").Value) = False Then
'         strAmount = Format(Val(adoquery.Fields("Namount").Value), DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 6100 - intLength
'         Printer.CurrentY = 300 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoquery.Fields("Tamount").Value) = False Then
'         strAmount = Format(Val(adoquery.Fields("Tamount").Value), DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 7400 - intLength
'         Printer.CurrentY = 300 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoquery.Fields("Ramount").Value) = False Then
'         strAmount = Format(Val(adoquery.Fields("Ramount").Value), DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 8700 - intLength
'         Printer.CurrentY = 300 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoquery.Fields("Pamount").Value) = False Then
'         strAmount = Format(Val(adoquery.Fields("Pamount").Value), DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 10000 - intLength
'         Printer.CurrentY = 300 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      If IsNull(adoquery.Fields("Vamount").Value) = False Then
'         strAmount = Format(Val(adoquery.Fields("Vamount").Value), DDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 11000 - intLength
'         Printer.CurrentY = 300 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      Printer.CurrentX = 11100
'      Printer.CurrentY = 300 + intCounter * 300
'      If IsNull(adoquery.Fields("Caseno").Value) Then
'         Printer.Print ""
'      Else
'         If Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 2, 3) = "000" Then
'            Printer.Print Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "-" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 8, 6)
'         Else
'            Printer.Print Mid(adoquery.Fields("Caseno").Value, 1, Len(adoquery.Fields("Caseno").Value) - 9) & "-" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 8, 6) & "-" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 2, 1) & "-" & Mid(adoquery.Fields("Caseno").Value, Len(adoquery.Fields("Caseno").Value) - 1, 2)
'         End If
'      End If
'      If IsNull(adoquery.Fields("Point").Value) = False Or adoquery.Fields("Point").Value <> 0 Then
'         strAmount = Format(Val(adoquery.Fields("Point").Value), FDollar)
'         intLength = Printer.TextWidth(strAmount)
'         Printer.CurrentX = 13900 - intLength
'         Printer.CurrentY = 300 + intCounter * 300
'         Printer.Print strAmount
'      End If
'      Printer.CurrentX = 14000
'      Printer.CurrentY = 300 + intCounter * 300
'      If IsNull(adoquery.Fields("DNno").Value) Then
'         Printer.Print ""
'      Else
'         Printer.Print adoquery.Fields("DNno").Value
'      End If
'      Printer.CurrentX = 15000
'      Printer.CurrentY = 300 + intCounter * 300
'      If IsNull(adoquery.Fields("Checkno").Value) Then
'         Printer.Print ""
'      Else
'         Printer.Print adoquery.Fields("Checkno").Value
'      End If
'      intCounter = intCounter + 1
'      intRecord = intRecord + 1
'      adoquery.MoveNext
'   Loop
'   PrintSum 'Add by Morgan 2004/10/26
'   Printer.EndDoc
'
'   'Mark by Amy 2014/04/30 改產生Excel
''   'Add By Sindy 2013/1/23
''   If Check1.Visible = True And Check1.Value = 1 Then '列印抵帳明細
''      '清空變數值
''      For i = 0 To intMax
''         strCurr(i) = ""
''         dblTotFAmt(i) = 0
''      Next i
''      dblTotNAmt = 0
''      dblTotFee = 0
''      bolIsHaveData = False: iLine = 0
''      '未收
''      With adoquery
''         .MoveFirst
''         PrintList1
''         Do While Not .EOF
''            If Left(Trim(.Fields("DocNo")), 1) = "X" Then
''               bolIsHaveData = True
''               For m_i = 1 To 7
''                  strTemp(m_i) = ""
''               Next m_i
''               strTemp(1) = CheckStr(DBDATE(.Fields("DocDate")))
''               strTemp(2) = CheckStr(.Fields("DocNo"))
''               strTemp(3) = CheckStr(.Fields("Currency"))
''               strTemp(4) = Format(Val(.Fields("Famount")) - Val(.Fields("FRamount")), FDollar)
''               For i = 0 To intMax
''                  If strCurr(i) <> "" Then
''                     If strCurr(i) = strTemp(3) Then
''                        intUseI = i '累計
''                        Exit For
''                     End If
''                  Else
''                     intUseI = i '使用新陣列
''                     Exit For
''                  End If
''               Next i
''               strCurr(intUseI) = strTemp(3)
''               dblTotFAmt(intUseI) = dblTotFAmt(intUseI) + Val(.Fields("Famount")) - Val(.Fields("FRamount"))
''               strTemp(5) = Format(Val(.Fields("Namount")) - Val("" & .Fields("Ramount")), DDollar)
''               dblTotNAmt = dblTotNAmt + Val(.Fields("Namount")) - Val("" & .Fields("Ramount"))
''               strTemp(6) = Format(Val(.Fields("Tamount")), DDollar)
''               dblTotFee = dblTotFee + Val(.Fields("Tamount"))
''               strTemp(7) = CheckStr(.Fields("Caseno"))
''               If Right(strTemp(7), 3) = "000" Then
''                  strTemp(7) = Left(strTemp(7), Len(strTemp(7)) - 3)
''               End If
''               PrintDetail1
''               If iLine >= 52 Then
''                  'If .AbsolutePosition <> .RecordCount Then
''                     Printer.NewPage
''                     PrintList1
''                  'End If
''               End If
''            End If
''            .MoveNext
''         Loop
''         If bolIsHaveData = True Then
''            Printer.CurrentX = PLeft(0)
''            Printer.CurrentY = iLine * 300
''            Printer.Print String(115, "-")
''            iLine = iLine + 1
''            For i = 0 To intMax
''               If i = 0 Then
''                  strTemp(1) = "TOTAL"
''                  strTemp(2) = strCurr(i)
''                  strTemp(3) = Format(dblTotFAmt(i), FDollar)
''                  strTemp(4) = Format(dblTotNAmt, DDollar)
''                  strTemp(5) = Format(dblTotFee, DDollar)
''               Else
''                  If strCurr(i) <> "" Then
''                     strTemp(1) = ""
''                     strTemp(2) = strCurr(i)
''                     strTemp(3) = Format(dblTotFAmt(i), FDollar)
''                     strTemp(4) = ""
''                     strTemp(5) = ""
''                  Else
''                     Exit For
''                  End If
''               End If
''               If iLine >= 52 Then
''                  'If .AbsolutePosition <> .RecordCount Then
''                     Printer.NewPage
''                     PrintList1
''                  'End If
''               End If
''               For m_j = 1 To 5
''                  If m_j >= 3 And m_j <= 5 Then
''                     Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
''                  Else
''                     Printer.CurrentX = PLeft(m_j)
''                  End If
''                  Printer.CurrentY = iLine * 300
''                  Printer.Print strTemp(m_j)
''               Next m_j
''               iLine = iLine + 1
''            Next i
''            Printer.EndDoc
''         End If
''      End With
''
''      '清空變數值
''      For i = 0 To intMax
''         strCurr(i) = ""
''         dblTotFAmt(i) = 0
''      Next i
''      dblTotNAmt = 0
''      dblTotFee = 0
''      strTitle = ""
''      bolIsHaveData = False: iLine = 0
''      '未付
''      With adoquery
''         .MoveFirst
''         Do While Not .EOF
''            If Left(Trim(.Fields("DocNo")), 1) = "U" Then
''               bolIsHaveData = True
''               If iLine >= 52 Or strTitle <> .Fields("fa05") Then
''                  If strTitle <> .Fields("fa05") And strTitle <> "" Then
''                     '合計
''                     Printer.CurrentX = PLeft(0)
''                     Printer.CurrentY = iLine * 300
''                     Printer.Print String(115, "-")
''                     iLine = iLine + 1
''                     For i = 0 To intMax
''                        If i = 0 Then
''                           strTemp(1) = "TOTAL"
''                           strTemp(2) = strCurr(i)
''                           strTemp(3) = Format(dblTotFAmt(i), FDollar)
''                        Else
''                           If strCurr(i) <> "" Then
''                              strTemp(1) = ""
''                              strTemp(2) = strCurr(i)
''                              strTemp(3) = Format(dblTotFAmt(i), FDollar)
''                           Else
''                              Exit For
''                           End If
''                        End If
''                        If iLine >= 52 Then
''                           'If .AbsolutePosition <> .RecordCount Then
''                              Printer.NewPage
''                              Call PrintList2(strTitle)
''                           'End If
''                        End If
''                        For m_j = 1 To 3
''                           If m_j = 3 Then
''                              Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
''                           Else
''                              Printer.CurrentX = PLeft(m_j)
''                           End If
''                           Printer.CurrentY = iLine * 300
''                           Printer.Print strTemp(m_j)
''                        Next m_j
''                        iLine = iLine + 1
''                     Next i
''                     '清空變數值
''                     For i = 0 To intMax
''                        strCurr(i) = ""
''                        dblTotFAmt(i) = 0
''                     Next i
''                     dblTotNAmt = 0
''                     dblTotFee = 0
''                  End If
''                  'If .AbsolutePosition <> .RecordCount Then
''                     If strTitle <> "" Then Printer.NewPage
''                     Call PrintList2(.Fields("fa05"))
''                  'End If
''               End If
''               For m_i = 1 To 5
''                  strTemp(m_i) = ""
''               Next m_i
''               strTemp(1) = CheckStr(DBDATE(.Fields("DocDate")))
''               strTemp(2) = CheckStr(.Fields("DNno"))
''               strTemp(3) = CheckStr(.Fields("Currency"))
''               strTemp(4) = Format(Val(.Fields("Famount")) - Val(.Fields("Pamount")), FDollar)
''               For i = 0 To intMax
''                  If strCurr(i) <> "" Then
''                     If strCurr(i) = strTemp(3) Then
''                        intUseI = i '累計
''                        Exit For
''                     End If
''                  Else
''                     intUseI = i '使用新陣列
''                     Exit For
''                  End If
''               Next i
''               strCurr(intUseI) = strTemp(3)
''               dblTotFAmt(intUseI) = dblTotFAmt(intUseI) + Val(.Fields("Famount")) - Val(.Fields("Pamount"))
''               strTemp(5) = CheckStr(.Fields("Caseno"))
''               If Right(strTemp(5), 3) = "000" Then
''                  strTemp(5) = Left(strTemp(5), Len(strTemp(5)) - 3)
''               End If
''               PrintDetail2
''               strTitle = .Fields("fa05")
''            End If
''            .MoveNext
''         Loop
''         If bolIsHaveData = True Then
''            Printer.CurrentX = PLeft(0)
''            Printer.CurrentY = iLine * 300
''            Printer.Print String(115, "-")
''            iLine = iLine + 1
''            For i = 0 To intMax
''               If i = 0 Then
''                  strTemp(1) = "TOTAL"
''                  strTemp(2) = strCurr(i)
''                  strTemp(3) = Format(dblTotFAmt(i), FDollar)
''               Else
''                  If strCurr(i) <> "" Then
''                     strTemp(1) = ""
''                     strTemp(2) = strCurr(i)
''                     strTemp(3) = Format(dblTotFAmt(i), FDollar)
''                  Else
''                     Exit For
''                  End If
''               End If
''               If iLine >= 52 Then
''                  'If .AbsolutePosition <> .RecordCount Then
''                     Printer.NewPage
''                     Call PrintList2(strTitle)
''                  'End If
''               End If
''               For m_j = 1 To 3
''                  If m_j = 3 Then
''                     Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
''                  Else
''                     Printer.CurrentX = PLeft(m_j)
''                  End If
''                  Printer.CurrentY = iLine * 300
''                  Printer.Print strTemp(m_j)
''               Next m_j
''               iLine = iLine + 1
''            Next i
''            Printer.EndDoc
''         End If
''      End With
''   End If
''   '2013/1/23 End
'    'end 2014/04/30
'   adoquery.Close
'
'Checking:
'   Set rs = Nothing 'Add by Sindy 2012/12/21
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'end 2015/06/30
End Sub

'Add by Amy 2015/11/25 +公司別
Private Sub Text9_GotFocus()
    TextInverse Text2
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
     KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      Beep
      KeyAscii = 0
    End If
End Sub


