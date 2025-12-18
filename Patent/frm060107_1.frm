VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060107_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "翻譯完稿輸入"
   ClientHeight    =   5640
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8952
   Begin VB.CheckBox Check1 
      Caption         =   "產生電子檔"
      Height          =   225
      Left            =   1710
      TabIndex        =   64
      Top             =   150
      Width           =   1335
   End
   Begin VB.CommandButton CmdPrint2 
      Caption         =   "會稿承辦單"
      Height          =   400
      Left            =   3450
      TabIndex        =   61
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "翻譯承辦單"
      Height          =   400
      Left            =   4920
      TabIndex        =   59
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox txtEP04 
      Height          =   270
      Left            =   1800
      MaxLength       =   6
      TabIndex        =   2
      Top             =   2790
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "案件資料"
      Height          =   3705
      Left            =   270
      TabIndex        =   23
      Top             =   510
      Width           =   8385
      Begin VB.TextBox txtEP31 
         Height          =   270
         Left            =   6735
         MaxLength       =   7
         TabIndex        =   6
         Top             =   3330
         Width           =   855
      End
      Begin VB.TextBox txtTF32 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1530
         MaxLength       =   7
         TabIndex        =   5
         Top             =   3330
         Width           =   855
      End
      Begin VB.TextBox txtCP113 
         Height          =   270
         Left            =   4800
         MaxLength       =   4
         TabIndex        =   0
         Top             =   1965
         Width           =   615
      End
      Begin VB.TextBox txtCaseNo 
         Enabled         =   0   'False
         Height          =   270
         Index           =   4
         Left            =   3105
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtCaseNo 
         Enabled         =   0   'False
         Height          =   270
         Index           =   3
         Left            =   2865
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox txtCaseNo 
         Enabled         =   0   'False
         Height          =   270
         Index           =   2
         Left            =   2025
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtCaseNo 
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   1545
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtEP09T 
         Height          =   270
         Left            =   6735
         MaxLength       =   7
         TabIndex        =   1
         Top             =   1965
         Width           =   1100
      End
      Begin VB.TextBox txtEP08T 
         Height          =   270
         Left            =   6735
         MaxLength       =   7
         TabIndex        =   3
         Top             =   2265
         Width           =   1100
      End
      Begin MSForms.TextBox txtCP64 
         Height          =   675
         Left            =   1530
         TabIndex        =   4
         Top             =   2610
         Width           =   6690
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "11800;1191"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Lbl2 
         AutoSize        =   -1  'True
         Caption         =   "Claims完稿日:"
         Height          =   180
         Index           =   1
         Left            =   5535
         TabIndex        =   63
         Top             =   3330
         Width           =   1065
      End
      Begin VB.Label Lbl2 
         AutoSize        =   -1  'True
         Caption         =   "只交Claims期限:"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   3330
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "翻譯時數:"
         Height          =   180
         Index           =   20
         Left            =   3960
         TabIndex        =   60
         Top             =   2010
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "發文日:"
         Height          =   180
         Index           =   16
         Left            =   2925
         TabIndex        =   55
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label lblCP27T 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   240
         Left            =   3645
         TabIndex        =   54
         Top             =   1410
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   53
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(外):"
         Height          =   180
         Index           =   0
         Left            =   1020
         TabIndex        =   52
         Top             =   1110
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Left            =   1020
         TabIndex        =   51
         Top             =   825
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Left            =   1020
         TabIndex        =   50
         Top             =   540
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱"
         Height          =   180
         Left            =   225
         TabIndex        =   49
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請日:"
         Height          =   180
         Index           =   1
         Left            =   6015
         TabIndex        =   48
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lblAppDate 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   240
         Left            =   6735
         TabIndex        =   47
         Top             =   240
         Width           =   1500
      End
      Begin MSForms.Label lblCaseName 
         Height          =   285
         Index           =   1
         Left            =   1545
         TabIndex        =   46
         Top             =   540
         Width           =   6675
         BackColor       =   -2147483643
         VariousPropertyBits=   746604571
         Size            =   "11774;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCaseName 
         Height          =   285
         Index           =   2
         Left            =   1545
         TabIndex        =   45
         Top             =   825
         Width           =   6675
         BackColor       =   -2147483643
         VariousPropertyBits=   746604571
         Size            =   "11774;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCaseName 
         Height          =   285
         Index           =   3
         Left            =   1545
         TabIndex        =   44
         Top             =   1110
         Width           =   6675
         BackColor       =   -2147483643
         VariousPropertyBits=   746604571
         Size            =   "11774;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCP05T 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   240
         Left            =   1545
         TabIndex        =   43
         Top             =   1410
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收文日:"
         Height          =   180
         Index           =   2
         Left            =   780
         TabIndex        =   42
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質:"
         Height          =   180
         Index           =   3
         Left            =   600
         TabIndex        =   41
         Top             =   1710
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人:"
         Height          =   180
         Index           =   4
         Left            =   780
         TabIndex        =   40
         Top             =   2010
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "核稿期限:"
         Height          =   180
         Index           =   6
         Left            =   5835
         TabIndex        =   39
         Top             =   2310
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "完稿日:"
         Height          =   180
         Index           =   7
         Left            =   6015
         TabIndex        =   38
         Top             =   2010
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "總收文號:"
         Height          =   180
         Index           =   8
         Left            =   5835
         TabIndex        =   37
         Top             =   1410
         Width           =   765
      End
      Begin VB.Label lblCP10T 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   240
         Left            =   1530
         TabIndex        =   36
         Top             =   1710
         Width           =   1665
      End
      Begin MSForms.Label lblCP14T 
         Height          =   285
         Left            =   2430
         TabIndex        =   35
         Top             =   2010
         Width           =   1440
         BackColor       =   -2147483643
         VariousPropertyBits=   746604571
         Size            =   "2540;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCP09 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   240
         Left            =   6735
         TabIndex        =   34
         Top             =   1380
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "核稿人:"
         Height          =   180
         Index           =   5
         Left            =   780
         TabIndex        =   33
         Top             =   2310
         Width           =   585
      End
      Begin MSForms.Label lblEP04T 
         Height          =   285
         Left            =   2430
         TabIndex        =   32
         Top             =   2310
         Width           =   1440
         BackColor       =   -2147483643
         VariousPropertyBits=   746604571
         Size            =   "2540;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCP14 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   240
         Left            =   1545
         TabIndex        =   31
         Top             =   2010
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利種類:"
         Height          =   180
         Index           =   9
         Left            =   5835
         TabIndex        =   30
         Top             =   1710
         Width           =   765
      End
      Begin VB.Label lblPA08T 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   240
         Left            =   6735
         TabIndex        =   29
         Top             =   1710
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "進度備註:"
         Height          =   180
         Index           =   10
         Left            =   600
         TabIndex        =   28
         Top             =   2610
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "翻譯費用資料"
      Height          =   1275
      Left            =   270
      TabIndex        =   17
      Top             =   4290
      Visible         =   0   'False
      Width           =   8385
      Begin VB.TextBox txtTF20 
         Height          =   270
         Left            =   5355
         MaxLength       =   12
         TabIndex        =   14
         Top             =   915
         Width           =   1455
      End
      Begin VB.TextBox txtTF19 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   3285
         MaxLength       =   6
         TabIndex        =   13
         Top             =   915
         Width           =   720
      End
      Begin VB.TextBox txtTF23 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   12
         Top             =   915
         Width           =   720
      End
      Begin VB.TextBox txtTF06 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   3285
         MaxLength       =   3
         TabIndex        =   10
         Top             =   615
         Width           =   720
      End
      Begin VB.TextBox txtTF18 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   5355
         MaxLength       =   3
         TabIndex        =   11
         Top             =   615
         Width           =   720
      End
      Begin VB.TextBox txtTF05 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   9
         Top             =   615
         Width           =   720
      End
      Begin VB.TextBox txtTF04 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   3285
         MaxLength       =   6
         TabIndex        =   8
         Top             =   300
         Width           =   720
      End
      Begin VB.TextBox txtTF03 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   7
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "相似案號:"
         Height          =   180
         Index           =   19
         Left            =   4500
         TabIndex        =   58
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "相似度:                   %"
         Height          =   180
         Index           =   18
         Left            =   2610
         TabIndex        =   57
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "原文字數:"
         Height          =   180
         Index           =   17
         Left            =   240
         TabIndex        =   56
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "瑕疵折扣:                   %"
         Height          =   180
         Index           =   14
         Left            =   2430
         TabIndex        =   21
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "加成比率:                   %"
         Height          =   180
         Index           =   15
         Left            =   4500
         TabIndex        =   22
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "相似折扣:                   %"
         Height          =   180
         Index           =   13
         Left            =   225
         TabIndex        =   20
         Top             =   660
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "數學式字數:"
         Height          =   180
         Index           =   12
         Left            =   2250
         TabIndex        =   19
         Top             =   345
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日文字數:"
         Height          =   180
         Index           =   11
         Left            =   225
         TabIndex        =   18
         Top             =   345
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7575
      TabIndex        =   16
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6750
      TabIndex        =   15
      Top             =   60
      Width           =   800
   End
End
Attribute VB_Name = "frm060107_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/11 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Modify by Morgan 2011/5/31 核搞人發文後也要能改--靜芳
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Public bolTfOnly As Boolean
Dim bolIsValidate As Boolean, stST15 As String, stCP10 As String, stCP06 As String
Dim bolActived As Boolean 'Add by Morgan 2010/3/22
Dim stCP60 As String 'Add by Morgan2010/6/18
Dim pa() As String, intWhere As Integer   '2011/11/30 add by sonia
'Added by Morgan 2013/9/11
Dim m_UpdateCP09 As String '待更新之新型案檢視中說收文號
Dim m_oldEP09T As String
Dim m_UpdateCP06 As String '待更新之新型案檢視中說的所限
'Added by Lydia 2018/09/13
Dim m_TF30 As String '英文本收文號
Dim m_oldEP31T As String 'Added by Lydia 2019/04/17
Dim strMurgitroyd As String 'Added by Lydia 2021/01/06 Murgitroyd案的代理人
'Add By Sindy 2022/5/12
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2022/5/12 END
Dim bolEP04trigger As Boolean 'Added by Lydia 2024/05/07 輸入核稿人是否系統之工程師組別

'Add By Sindy 2022/5/12
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

'Modify By Sindy 2023/10/18
'Private Sub cmdBack_Click()
Public Sub cmdBack_Click()
'2023/10/18 END
    Call frm060107.SetGrid(False)
    frm060107.Show
    Unload Me
End Sub

Public Sub SetData(ByRef rstGrid As ADODB.Recordset, ByVal iRow As Integer)
    
    Dim ii As Integer, stPA08 As String
    
    With frm060107
        For ii = 1 To 4
            txtCaseNo(ii) = .txtCaseNo(ii)
        Next ii
        lblAppDate = .lblAppDate
        For ii = 1 To 3
            lblCaseName(ii) = .lblCaseName(ii)
        Next ii
    End With
    
    'Added by Lydia 2021/01/06
    pa(1) = txtCaseNo(1): pa(2) = txtCaseNo(2): pa(3) = txtCaseNo(3): pa(4) = txtCaseNo(4)
    If txtCaseNo(1) = "FCP" Or txtCaseNo(1) = "P" Then
       Call ClsPDReadPatentDatabase(pa(), intWhere)
    ElseIf txtCaseNo(1) = "FG" Or txtCaseNo(1) = "PS" Then
       Call ClsPDReadServicePracticeDatabase(pa(), intWhere)
    End If
    strMurgitroyd = Pub_GetSpecMan("外專MURGITROYD設定")
    'end 2021/01/06
    
    With rstGrid
        .Move iRow - 1, adBookmarkFirst
        lblCP27T = "" & .Fields("CP27T") 'Add by Morgan 2011/5/31
        lblCP05T = "" & .Fields("CP05T")
        lblCP09 = "" & .Fields("CP09")
        lblCP10T = "" & .Fields("CP10T")
        lblCP14 = "" & .Fields("CP14")
        lblCP14T = "" & .Fields("CP14T")
        txtEP09T = "" & .Fields("EP09T")
        m_oldEP09T = txtEP09T 'Added by Morgan 2013/9/11
        txtEP04 = "" & .Fields("EP04")
        txtEP04.Tag = txtEP04
        lblEP04T = "" & .Fields("EP04T")
        txtEP08T = "" & .Fields("CP48T")
        'Add by Morgan 2005/9/12
        txtEP08T.Tag = txtEP08T
        'Added by Lydia 2018/05/07 承辦人為所內員工上班譯，新增翻譯時數欄位。
        Label1(20).Visible = False: txtCP113.Visible = False
        'Modified by Lydia 2018/06/05 排除外翻F編號(=所內員工下班譯,或外翻)
        'If Left("" & lblCP14, 1) = "A" Then
        If Left("" & lblCP14, 1) <> "F" And Trim(lblCP14) <> "" Then
            Label1(20).Visible = True
            txtCP113.Visible = True
            txtCP113 = "" & .Fields("CP113")
        End If
        'end 2018/05/07
   
        lblPA08T = "" & .Fields("PA08T")
        txtCP64 = "" & .Fields("CP64")
        stPA08 = "" & .Fields("PA08")
        stST15 = "" & .Fields("ST15")
        stCP10 = "" & .Fields("CP10")
        stCP06 = "" & .Fields("CP06")
        stCP60 = "" & .Fields("CP60")
    End With
    'Modify by Morgan 2005/5/25 承辦人為外專工程師(F21)時不可改核稿人,2008/4/8加F81
    'If stPA08 = "3" Then txtEP04.Enabled = False
    '2009/5/4 modify by sonia 取消承辦人為外專工程師(F21,F81)時不可改核稿人
    'If stPA08 = "3" Or stST15 = "F21" Or stST15 = "F81" Then txtEP04.Enabled = False
    If stPA08 = "3" Then txtEP04.Enabled = False
    
   'Add by Morgan 2007/6/5
   Frame1.Visible = False

   If Left("" & lblCP14, 1) = "F" Then
      Frame1.Visible = True
   'Mark by Lydia 2021/01/01/20 欄位已自行控制是否顯示；FCP-64205無法輸入”Claims完稿日”
   'Else
   '   Me.Height = 3750
   'end 2021/01/20
   End If
   'end 2007/6/5
   'Added by Lydia 2019/04/17 增加Claims期限和完稿日
   Lbl2(0).Visible = False: Lbl2(1).Visible = False
   txtTF32.Visible = False: txtEP31.Visible = False

   'Move by Lydia 2018/09/13 從Frame1.Visible = True移下來
   'Modified by Lydia 2019/04/17 +EP31
    'strExc(0) = "select * from TransFee where TF01='" & lblCP09 & "'"
    strExc(0) = "select b.*,ep31  from TransFee b, engineerprogress c where TF01='" & lblCP09 & "' and tf01=ep02(+) "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        txtTF03.Text = "" & RsTemp.Fields("TF03")
        txtTF04.Text = "" & RsTemp.Fields("TF04")
        txtTF05.Text = "" & RsTemp.Fields("TF05")
        txtTF06.Text = "" & RsTemp.Fields("TF06")
        txtTF18.Text = "" & RsTemp.Fields("TF18")
        'Added by Lydia 2017/05/17 原文字數TF23、相似度TF19、相似案號TF20
        txtTF23.Text = "" & RsTemp.Fields("TF23")
        txtTF19.Text = "" & RsTemp.Fields("TF19")
        txtTF20.Text = "" & RsTemp.Fields("TF20")
        'end 2017/05/17
        m_TF30 = "" & RsTemp.Fields("TF30") 'Added by Lydia 2018/09/13 英文本收文號
        'Modified by Lydia 2018/09/13 判斷下班翻或外翻才可修改
        'If "" & RsTemp.Fields("TF07") <> "" Then
        If "" & RsTemp.Fields("TF07") <> "" Or Left("" & lblCP14, 1) <> "F" Then
            txtTF03.Enabled = False
            txtTF04.Enabled = False
            txtTF05.Enabled = False
            txtTF06.Enabled = False
            txtTF18.Enabled = False
        End If
        'Added by Lydia 2019/04/17 增加Claims期限和完稿日
        If "" & RsTemp.Fields("TF32") <> "" Then
            Lbl2(0).Visible = True: Lbl2(1).Visible = True
            txtTF32.Visible = True: txtEP31.Visible = True
            txtTF32.Text = TransDate("" & RsTemp.Fields("TF32"), 1)
            txtTF32.Tag = txtTF32.Text
            txtEP31.Text = TransDate("" & RsTemp.Fields("EP31"), 1)
            m_oldEP31T = txtEP31.Text
        End If
    End If
End Sub

Private Function FormSave() As Boolean

   Dim stEP04 As String, stEP09 As String, stCP09 As String, stEP08  As String
   Dim stSQL As String, stUpdateEP As String, stUpdateCP As String
   'Modifeid by Lydia 2017/05/17 記錄原文字數、相似度、相似案號
   'Dim stTF(3 To 18) As String
   Dim stTF(3 To 23) As String
   Dim msgTxt As String 'Added by Lydia 2018/05/31
   Dim strTemp As String 'Added by Lydia 2020/02/24
   
On Error GoTo flgError

cnnConnection.BeginTrans
   
   'Modify by Morgan 2009/6/30
   '發文後未付翻譯費前可修改字數
   
   stUpdateEP = ""
   stUpdateCP = ""
   If txtEP04.Enabled = True Then
      stEP04 = txtEP04
      stUpdateEP = stUpdateEP & ",EP04='" & stEP04 & "'"
   End If
   
   If txtEP08T.Enabled = True Then
      stEP08 = ChangeTStringToWString(txtEP08T)
      stUpdateEP = stUpdateEP & ",EP08=" & CNULL(stEP08)
   End If
   
   If txtEP09T.Enabled = True Then
      stEP09 = ChangeTStringToWString(txtEP09T)
      stUpdateEP = stUpdateEP & ",EP09=" & CNULL(stEP09)
   End If
   
   'Added by Lydia 2018/05/07 工作時數
   If txtCP113.Visible = True And Val(txtCP113) > 0 Then
       'Modified by Lydia 2019/06/17 + stUpdatcp
       'stUpdateCP = " CP113='" & Val(txtCP113) & "'"
       stUpdateCP = stUpdateCP & ", CP113='" & Val(txtCP113) & "'"
   End If
   
   'Added by Lydia 2019/04/17 Claims完稿日
   If txtEP31.Visible = True And txtEP31.Text <> m_oldEP31T Then
       stUpdateEP = stUpdateEP & ",EP31=" & CNULL(DBDATE(txtEP31))
       '記錄在中說進度備註(201),於翻譯承辦單判斷EP31列出
       If txtEP31.Text <> "" Then 'Added by Lydia 2019/05/14 取消日期不用加備註
          'strExc(1) = "交稿Claims:" & IIf(txtEP31.Text <> "", ChangeTStringToTDateString(txtEP31.Text) & "已交稿Claims", ChangeTStringToTDateString(strSrvDate(2)) & "取消交稿日") & ";'"
          strExc(1) = IIf(txtEP31.Text <> "", ChangeTStringToTDateString(txtEP31.Text) & "已交稿Claims", ChangeTStringToTDateString(strSrvDate(2)) & "取消交稿日") & ";"
          txtCP64 = strExc(1) & txtCP64
       End If
   End If
   
   'Move by Lydia 2019/04/17 從工作時數上面移下來
   If txtCP64.Enabled = True Then
      'Modified by Lydia 2019/06/17 + stUpdatcp (Sharon反應FCP-059967有輸入翻譯時數，到Daphone輸入核稿日卻不見)
      'stUpdateCP = " CP64='" & ChgSQL(txtCP64) & "'"
      stUpdateCP = stUpdateCP & ", CP64='" & ChgSQL(txtCP64) & "'"
   End If
   
   If stUpdateEP <> "" Or stUpdateCP <> "" Then
      stCP09 = lblCP09
      'Modify by Morgan 2005/9/8 核稿期限改放ep08(原cp48)
      'stSQL = " Begin" 'Mark by Lydia 2024/05/07
      If stUpdateEP <> "" Then
         If bolEP04trigger = False Then cnnConnection.Execute "begin user_data.user_notrigger:=1; end;"  'Added by Lydia 2024/05/07 控制 Trigger 不被觸發
           'Modified by Lydia 2024/05/07
           'stSQL = stSQL & " Update ENGINEERPROGRESS Set " & Mid(stUpdateEP, 2) & " Where EP02='" & stCP09 & "'; "
           stSQL = " Update ENGINEERPROGRESS Set " & Mid(stUpdateEP, 2) & " Where EP02='" & stCP09 & "' "
           cnnConnection.Execute stSQL 'Added by Lydia 2024/05/07
         If bolEP04trigger = False Then cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"  'Added by Lydia 2024/05/07 控制 Trigger 不被觸發
      End If
      If stUpdateCP <> "" Then
         'Modified by Lydia 2019/06/17
         'stSQL = stSQL & " Update CASEPROGRESS Set " & stUpdateCP & " Where CP09='" & stCP09 & "';"
         'Modified by Lydia 2024/05/07 拿掉;
         'stSQL = stSQL & " Update CASEPROGRESS Set " & Mid(stUpdateCP, 2) & " Where CP09='" & stCP09 & "'; "
         stSQL = " Update CASEPROGRESS Set " & Mid(stUpdateCP, 2) & " Where CP09='" & stCP09 & "' "
         cnnConnection.Execute stSQL 'Added by Lydia 2024/05/07
      End If
      'Mark by Lydia 2024/05/07
      'stSQL = stSQL & " End;"
      'cnnConnection.Execute stSQL
      'end 2024/05/07
   End If
    
   'Added by Morgan 2013/9/11
   If m_UpdateCP09 <> "" Then
      stSQL = "update engineerprogress set ep06=" & CNULL(ChangeTStringToWString(txtEP09T), True) & " where ep02='" & m_UpdateCP09 & "'"
      cnnConnection.Execute stSQL, intI
      'Modified by Lydia 2018/05/31 判斷所限與承辦期限 (ex.FCP-58496發明案輸入完稿日，更新FCP-58557新型案的承辦期限)
      'stSQL = "update caseprogress set cp48=" & CNULL(ChangeTStringToWString(txtEP08T), True) & " where cp09='" & m_UpdateCP09 & "' "
      strExc(1) = ChangeTStringToWString(txtEP08T)
      If Val(m_UpdateCP06) = 0 Or Val(m_UpdateCP06) > Val(strExc(1)) Then '無所限或所限>承辦期限
           stSQL = strExc(1)
      Else '承辦期限>所限, 設承辦期限=>所限
           stSQL = m_UpdateCP06
      End If
      stSQL = "update caseprogress set cp48=" & stSQL & " where cp09='" & m_UpdateCP09 & "' "
      'end 2018/05/31
      cnnConnection.Execute stSQL, intI
   End If
   'end 2013/9/11
   
   'Added by Lydia 2018/03/15 FCP新案翻譯Key核稿人後，檢查有尚未發文的A類會稿，則會稿924的承辦人更新為核稿人並且自動上已分案。
   If txtEP04.Enabled = True And Trim(txtEP04.Text) <> "" And pa(1) = "FCP" Then
        strExc(1) = "": strExc(2) = ""
        If PUB_ChkCPExist(pa, "924", 1, strExc(1), strExc(2), "A") = True Then
             stSQL = "update caseprogress set cp14=" & CNULL(Trim(txtEP04.Text)) & ", cp122='Y' where cp09=" & CNULL(strExc(1))
             cnnConnection.Execute stSQL, intI
        End If
   End If
   'end 2018/03/15
   
   'Add by Morgan 2007/5/21
   If Frame1.Visible = True And txtTF03.Enabled = True Then
      If RTrim(txtTF03) = "" Then
         If txtEP09T = "" Then
            stTF(3) = "Null"
         Else
            stTF(3) = "0"
         End If
      Else
         stTF(3) = Val(txtTF03)
      End If
       
      If RTrim(txtTF04) = "" Then
         If txtEP09T = "" Then
            stTF(4) = "Null"
         Else
            stTF(4) = "0"
         End If
      Else
         stTF(4) = Val(txtTF04)
      End If
   
      If RTrim(txtTF05) = "" Or Val(txtTF05) = 100 Then
         stTF(5) = "Null"
      Else
         stTF(5) = Val(txtTF05)
      End If
   
      If RTrim(txtTF06) = "" Or Val(txtTF06) = 100 Then
         stTF(6) = "Null"
      Else
         stTF(6) = Val(txtTF06)
      End If
      
      If RTrim(txtTF18) = "" Or Val(txtTF18) = 100 Then
         stTF(18) = "Null"
      Else
         stTF(18) = Val(txtTF18)
      End If
      
      'Added by Lydia 2017/05/17 記錄原文字數、相似度、相似案號
      If Trim(txtTF23) = "" Then
         stTF(23) = "NULL"
      Else
         stTF(23) = Val(txtTF23)
      End If
      If Trim(txtTF19) = "" Or Val(txtTF19) = 100 Then
         stTF(19) = "NULL"
      Else
         stTF(19) = Val(txtTF19)
      End If
      If Trim(txtTF20) = "" Then
         stTF(20) = "NULL"
      Else
         'Modified by Lydia 2017/05/23
         'stTF(20) = Trim(txtTF20)
         stTF(20) = "'" & Trim(txtTF20) & "'"
      End If
      'end 2017/05/15
       
      'Modified by Lydia 2017/05/17 +原文字數、相似度、相似案號
      'stSQL = "Update TransFee set TF03=" & stTF(3) & ",TF04=" & stTF(4) & ",TF05=" & stTF(5) & ",TF06=" & stTF(6) & ",TF18=" & stTF(18) & " where TF01='" & lblCP09 & "' and tf07 is null"
      stSQL = "Update TransFee set TF03=" & stTF(3) & ",TF04=" & stTF(4) & ",TF05=" & stTF(5) & ",TF06=" & stTF(6) & ",TF18=" & stTF(18) & _
              ",TF19=" & stTF(19) & ",TF20=" & stTF(20) & ",TF23=" & stTF(23) & " where TF01='" & lblCP09 & "' and tf07 is null"
      'end 2017/05/17
      cnnConnection.Execute stSQL, intI
      If intI = 0 Then
         'Modified by Lydia 2017/05/17 +原文字數、相似度、相似案號
         'stSQL = "insert into TransFee(TF01,TF03,TF04,TF05,TF06,TF18) values('" & lblCP09 & "'," & stTF(3) & "," & stTF(4) & "," & stTF(5) & "," & stTF(6) & "," & stTF(18) & ")"
         stSQL = "insert into TransFee(TF01,TF03,TF04,TF05,TF06,TF18,TF19,TF20,TF23) " & _
                 "values('" & lblCP09 & "'," & stTF(3) & "," & stTF(4) & "," & stTF(5) & "," & stTF(6) & "," & stTF(18) & "," & stTF(19) & "," & stTF(20) & "," & stTF(23) & ")"
         'end 2017/05/17
         cnnConnection.Execute stSQL, intI
      End If
   'Remove by Lydia 2018/09/13 已與Morgan確認,修改Account翻譯費轉應付的條件,不用刪除記錄
   'Else
   '   stSQL = "delete TransFee where TF01='" & lblCP09 & "' and tf07 is null"
   '   cnnConnection.Execute stSQL, intI
   End If
   
   'Add by Morgan 2010/6/17
   '若已開請款單則換承辦人或核稿人時發Mail通知靜芳
   If stCP60 > "X" Then
      strExc(1) = txtCaseNo(1) & "-" & txtCaseNo(2) & IIf(txtCaseNo(3) & txtCaseNo(4) = "000", "", "-" & txtCaseNo(3) & "-" & txtCaseNo(4))
      PUB_PointReAssignInform strExc(1), stCP60, , , txtEP04.Tag, txtEP04
   End If
   
   'Added by Morgan 2015/10/12
   '更新核稿人至會稿或寄中說的承辦人
   'Modified by Lydia 2016/06/21 模組化 PUB_UpdateFCP924
   'If txtEP04 <> "" Then
   '   stSQL = "update caseprogress set cp14='" & txtEP04 & "',cp122='Y' where cp01='" & txtCaseNo(1) & "'" & _
         " and cp02='" & txtCaseNo(2) & "' and cp03='" & txtCaseNo(3) & "'" & _
         " and cp04='" & txtCaseNo(4) & "' and cp10 in ('924','949') and cp57||cp27 is null and (cp14 is null or cp14<>'" & txtEP04 & "')"
   '   cnnConnection.Execute stSQL, intI
   'End If
   Call PUB_UpdateFCP924(txtCaseNo(1).Text, txtCaseNo(2).Text, txtCaseNo(3).Text, txtCaseNo(4).Text, stCP10, txtEP04)
   'end 2015/10/12
   
   'Added by Lydia 2018/05/31 輸入核稿人檢查有無會稿, 會稿自動掛承辦期限 (承辦人在前面已更新,所以傳空白)
   If txtEP04 <> "" Then
       msgTxt = PUB_Update924CP(txtCaseNo(1).Text, txtCaseNo(2).Text, txtCaseNo(3).Text, txtCaseNo(4).Text, "", stCP06)
      'Added by Lydia 2018/09/13 因為已不列印紙本,在輸入核稿人發email通知到卷宗區下載提申本
      'Modified by Lydia 2018/11/06 排除FMP案不通知核稿人(ex.P-121250)
      'If txtEP04.Tag <> txtEP04.Text Then
      If txtEP04.Tag <> txtEP04.Text And pa(1) <> "P" Then
           strExc(5) = "" 'email:內文
           If m_TF30 <> "" And m_TF30 <> "Y" Then
                 '英文參考本：新案建檔有設「英文本收文號」，抓該收文號在卷宗區掛的sep.pdf。沒有就沒附件
                 '因為有時sep.pdf只有文字，所以加抓圖檔(DWG.pdf)
                 stSQL = "SELECT '1' ord1,CPP01,CPP02,CPP14,CPP06,CPP07,sqldatet(cp05) cp05t,sqldatet(cp27) cp27t," & IIf(pa(9) <> "000", "cpm04", "cpm03") & " as cp10n " & _
                              "FROM CASEPROGRESS A,CASEPAPERPDF B, casepropertymap " & _
                              "WHERE CP09='" & m_TF30 & "' AND CP159=0 AND CP09=CPP01(+) " & _
                              "and cp01=cpm01(+) and cp10=cpm02(+) AND NVL(CPP10,'N') <> 'D' AND (UPPER(CPP02) LIKE '%.SEP.PDF' OR UPPER(CPP02) LIKE '%.DWG.PDF') "
                '含-序列表
                'Modified by Lydia 2018/10/02 +.TBL.
                'Modified by Lydia 2019/10/24 拿掉 OR UPPER(CPP02) LIKE '%.PWD.%'
                'Modified by Lydia 2023/05/08 因應智慧局4/25起對序列表翻譯的變更,無需再抓取檔案, 卷宗區.SEQ.pdf及原始檔區.xml檔案=>拿掉UPPER(CPP02) LIKE '%.SEQ.%' OR
                stSQL = stSQL & " Union  SELECT '2' ord1,CPP01,CPP02,CPP14,CPP06,CPP07,sqldatet(cp05) cp05t,sqldatet(cp27) cp27t," & IIf(pa(9) <> "000", "cpm04", "cpm03") & " as cp10n " & _
                              "FROM CASEPROGRESS A,CASEPAPERPDF B, casepropertymap " & _
                              "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP159=0 AND CP09=CPP01(+) " & _
                              "and cp01=cpm01(+) and cp10=cpm02(+) AND NVL(CPP10,'N') <> 'D' AND (UPPER(CPP02) LIKE '%.TBL.%' ) " & _
                              "ORDER BY ord1 Asc,CPP06 DESC, CPP07 DESC "
           Else
                '其他(一般翻譯):抓卷宗區最後上傳的外文本(*.ORI.PDF、*.ORI.REP.PDF、*.ORI.FIX.PDF)
                'Modified by Lydia 2018/09/18 ORI.FIX=>改成FIX.ORI
                'Modified by Lydia 2018/11/30 判斷最後一道.ORI.%.PDF ,因為.FIX有人加後面
                'stSQL = "SELECT '2' ord1,CPP01,CPP02,CPP14,CPP06,CPP07,sqldatet(cp05) cp05t,sqldatet(cp27) cp27t," & IIf(pa(9) <> "000", "cpm04", "cpm03") & "  as cp10n " & _
                              "FROM CASEPROGRESS A,CASEPAPERPDF B, casepropertymap " & _
                               "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP159=0 AND CP09=CPP01(+) " & _
                               "and cp01=cpm01(+) and cp10=cpm02(+) AND NVL(CPP10,'N') <> 'D' AND (UPPER(CPP02) LIKE '%.ORI.PDF' OR UPPER(CPP02) LIKE '%.ORI.REP%.PDF' " & _
                               "OR UPPER(CPP02) LIKE '%.FIX%.ORI.PDF' OR UPPER(CPP02) LIKE '%.SEQ.%' OR UPPER(CPP02) LIKE '%.PWD.%' OR UPPER(CPP02) LIKE '%.TBL.%' ) " & _
                               "ORDER BY ord1 Asc,CPP06 DESC, CPP07 DESC "
                'Modified by Lydia 2019/10/24 拿掉 OR UPPER(CPP02) LIKE '%.PWD.%'
                'Modified by Lydia 2023/05/08 因應智慧局4/25起對序列表翻譯的變更,無需再抓取檔案, 卷宗區.SEQ.pdf及原始檔區.xml檔案=>拿掉OR UPPER(CPP02) LIKE '%.SEQ.%'
                stSQL = "SELECT '2' ord1,CPP01,CPP02,CPP14,CPP06,CPP07,sqldatet(cp05) cp05t,sqldatet(cp27) cp27t," & IIf(pa(9) <> "000", "cpm04", "cpm03") & "  as cp10n " & _
                              "FROM CASEPROGRESS A,CASEPAPERPDF B, casepropertymap " & _
                               "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP159=0 AND CP09=CPP01(+) " & _
                               "and cp01=cpm01(+) and cp10=cpm02(+) AND NVL(CPP10,'N') <> 'D' AND ((UPPER(CPP02) LIKE '%.ORI.%' AND UPPER(CPP02) LIKE '%.PDF' ) " & _
                               "OR UPPER(CPP02) LIKE '%.TBL.%' ) " & _
                               "ORDER BY ord1 Asc,CPP06 DESC, CPP07 DESC "
           End If
           intI = 1
           strExc(5) = ""
           Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
           If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                    If "" & RsTemp.Fields("ord1") = "1" Then '英文本收文號
                        If "" & RsTemp.Fields("CPP01") <> "" And "" & RsTemp.Fields("CPP02") <> "" And "" & RsTemp.Fields("CPP14") <> "" Then
                             strExc(5) = strExc(5) & RsTemp.Fields("cp05t") & "  " & convForm(RsTemp.Fields("cp10n"), 10) & "  " & RsTemp.Fields("CPP02") & vbCrLf
                        End If
                    Else  '最新的提申本
                        '說明書
                        'Modified by Lydia 2018/09/18 ORI.FIX=>改成FIX ; ORI.REP => 改成REP
                        'Modified by Lydia 2018/11/30 判斷最後一道.ORI.%.PDF ,因為.FIX有人加後面
                        'If InStr(UCase("" & RsTemp.Fields("CPP02")), ".ORI.") > 0 And InStr(UCase(strExc(5)), ".ORI.PDF") = 0 And _
                                     InStr(UCase(strExc(5)), ".REP") = 0 And InStr(UCase(strExc(5)), ".FIX") = 0 Then
                        If InStr(UCase("" & RsTemp.Fields("CPP02")), ".ORI.") > 0 And InStr(UCase(strExc(5)), ".ORI.") = 0 Then
                               strExc(5) = strExc(5) & RsTemp.Fields("cp05t") & "  " & convForm(RsTemp.Fields("cp10n"), 10) & "  " & RsTemp.Fields("CPP02") & vbCrLf
                        End If
                        '序列表
                        'Mark by Lydia 2023/05/08 因應智慧局4/25起對序列表翻譯的變更,無需再抓取檔案, 卷宗區.SEQ.pdf及原始檔區.xml檔案
                        'If InStr(UCase("" & RsTemp.Fields("CPP02")), ".SEQ.") > 0 And InStr(UCase(strExc(5)), ".SEQ.") = 0 Then
                        '       strExc(5) = strExc(5) & RsTemp.Fields("cp05t") & "  " & convForm(RsTemp.Fields("cp10n"), 10) & "  " & RsTemp.Fields("CPP02") & vbCrLf
                        'End If
                        ''密碼檔
                        'if InStr(UCase("" & RsTemp.Fields("CPP02")), ".PWD.") > 0 And InStr(UCase(strExc(5)), ".PWD.") = 0 Then
                        '       strExc(5) = strExc(5) & RsTemp.Fields("cp05t") & "  " & convForm(RsTemp.Fields("cp10n"), 10) & "  " & RsTemp.Fields("CPP02") & vbCrLf
                        'End If
                        'end 2023/05/08
                        
                        'Added by Lydia 2018/10/02 需提供外翻非說明書部分之其他檔案,例如:技術用語對照表
                        If InStr(UCase("" & RsTemp.Fields("CPP02")), ".TBL.") > 0 And InStr(UCase(strExc(5)), ".TBL.") = 0 Then
                               strExc(5) = strExc(5) & RsTemp.Fields("cp05t") & "  " & convForm(RsTemp.Fields("cp10n"), 10) & "  " & RsTemp.Fields("CPP02") & vbCrLf
                        End If
                        'end 2018/10/02
                    End If
                    RsTemp.MoveNext
               Loop
           End If
           
           strExc(0) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & " 核稿"
           'Modified by Lydia 2018/11/05 輸入核稿人時，該案尚未提申(ex.FCP-59622)
           'If strExc(5) = "" Or (strExc(5) <> "" And InStr(UCase(strExc(5)), ".SEP.PDF") = 0 And InStr(UCase(strExc(5)), ".ORI.") = 0) Then
          '      strExc(5) = "請到卷宗區上傳提申本後，通知核稿人(" & lblEP04T & ")！"
          '      strExc(1) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) '收件人：FCP管制人, CC:核稿人
          '      stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                             " values( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd')" & _
                             ",to_char(sysdate,'hh24miss'),'" & strExc(0) & "，卷宗區無提申本" & "','" & strExc(5) & "','" & txtEP04 & "')"
           If Val(pa(10)) = 0 Or strExc(5) = "" Or (strExc(5) <> "" And InStr(UCase(strExc(5)), ".SEP.PDF") = 0 And InStr(UCase(strExc(5)), ".ORI.") = 0) Then
                'Added by Lydia 2020/02/24 English_Vers檔案：放在原始檔區，記錄收文號
                If PUB_ChkCPExist(pa, cntEnglish_Vers, , strTemp, , "D") = True Then
                    strExc(5) = "本案尚未提申，請至〔原始檔區〕\English_Vers(" & strTemp & ") 進行核稿。"
                    stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                 " values( '" & strUserNum & "','" & txtEP04 & "',to_char(sysdate,'yyyymmdd')" & _
                                 ",to_char(sysdate,'hh24miss'),'" & strExc(0) & "，請至〔原始檔區〕\English_Vers(" & strTemp & ") 進行核稿" & "','" & strExc(5) & "',null)"
                Else
                'end 2020/02/24
                    'Modified by Lydia 2024/07/22 改用變數
                    'strExc(5) = "本案尚未提申，請至\\TYPING2\English_Vers最終版本進行核稿。"
                    'stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                 " values( '" & strUserNum & "','" & txtEP04 & "',to_char(sysdate,'yyyymmdd')" & _
                                 ",to_char(sysdate,'hh24miss'),'" & strExc(0) & "，請至\\TYPING2\English_Vers進行核稿" & "','" & strExc(5) & "',null)"
                    strExc(5) = "本案尚未提申，請至\\" & strTyping2Path & "\English_Vers最終版本進行核稿。"
                    stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                 " values( '" & strUserNum & "','" & txtEP04 & "',to_char(sysdate,'yyyymmdd')" & _
                                 ",to_char(sysdate,'hh24miss'),'" & strExc(0) & "，請至\\" & strTyping2Path & "\English_Vers進行核稿" & "','" & strExc(5) & "',null)"
                    'end 2024/07/22
                End If 'Added by Lydia 2020/02/24
           'end 2018/11/05
           Else
                strExc(5) = "核稿期限：" & ChangeTStringToTDateString(txtEP08T.Text) & vbCrLf & vbCrLf & _
                                 convForm("收文日", 11) & convForm("案件性質", 11) & " 檔案名稱" & vbCrLf & strExc(5)
                stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                             " values( '" & strUserNum & "','" & txtEP04 & "',to_char(sysdate,'yyyymmdd')" & _
                             ",to_char(sysdate,'hh24miss'),'" & strExc(0) & "，請到卷宗區進行核稿作業" & "','" & strExc(5) & "',null)"
           End If
           cnnConnection.Execute stSQL, intI
      End If
      'end 2018/09/13
   End If
   'end 2018/05/31
   
   'Modify By Sindy 2024/1/2 mark,改在操作歷程翻譯交稿檢查
'   'Add by Sindy 2022/5/12
'   If m_strIR01 <> "" Then
'      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm060107"
'   End If
'   '2022/5/12 END
   
   cnnConnection.CommitTrans
   
   'Modify By Sindy 2024/1/2 mark,改在操作歷程翻譯交稿檢查
'   'Add By Sindy 2022/5/12
'   If Me.m_strIR01 <> "" Then
'      Unload frm060107
'      If Not m_PrevForm Is Nothing Then
'         Call m_PrevForm.GoNext
'      End If
'      Unload Me
'   End If
'   '2022/5/12 END
   
   FormSave = True
   Exit Function
   
flgError:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
      cnnConnection.RollbackTrans
   End If

End Function

Private Function TxtValidate() As Boolean

   Dim bolCancel As Boolean
    
   bolCancel = False: bolIsValidate = False
   
   'Modify by Morgan 2009/6/30
   '發文後未付翻譯費前可修改字數
   If txtEP09T.Enabled = True Then
      If txtEP09T.Tag = "" Then bolIsValidate = True
      Call txtEP09T_Validate(bolCancel)
      If bolCancel Then GoTo flgFail
   End If
   
   If txtEP04.Enabled = True Then
      bolIsValidate = False
      Call txtEP04_Validate(bolCancel)
      If bolCancel Then GoTo flgFail
   End If
      
   If txtEP08T.Enabled = True Then
      Call txtEP08T_Validate(bolCancel)
      If bolCancel Then GoTo flgFail
      'Add by Morgan 2008/10/9 拿掉完稿日時提示核稿期限也要清空(因為會有例外如巨京..所以不強制)
      If txtEP09T = "" Then
         If txtEP08T <> "" Then
            If MsgBox("核稿期限尚未清空，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
               txtEP08T.SetFocus
               txtEP08T_GotFocus
               GoTo flgFail
            End If
         End If
      End If
   End If
   
   'Added by Lydia 2018/05/07
   If txtCP113.Visible = True Then
       If Val(txtCP113) <= 0 And Val(txtEP09T) >= "1070507" Then
            MsgBox "已有完稿日期，請輸入翻譯時數!", vbCritical
            txtCP113.SetFocus
            txtCP113_GotFocus
            GoTo flgFail
       End If
   End If
   'end 2018/05/07
   
   'Add by Morgan 2007/5/22
   If Frame1.Visible = True Then
      If txtTF05.Enabled = True Then
         Call txtTF05_Validate(bolCancel)
         If bolCancel Then GoTo flgFail
      End If
      If txtTF06.Enabled = True Then
         Call txtTF06_Validate(bolCancel)
         If bolCancel Then GoTo flgFail
      End If
      If txtEP09T = "" Then
         If txtTF03.Enabled = False Then
            MsgBox "本收文已算過翻譯費，請確認是否發生錯誤！", vbExclamation
            GoTo flgFail
         ElseIf (txtTF03 <> "" Or txtTF04 <> "" Or txtTF06 <> "") Then
            MsgBox "【完稿日】空白時【日文字數】、【數學式數】及【瑕疵折扣】也要一併清空！", vbExclamation
            If txtTF03 <> "" Then
               txtTF03.SetFocus
               txtTF03_GotFocus
            ElseIf txtTF04 <> "" Then
               txtTF04.SetFocus
               txtTF04_GotFocus
            Else
               txtTF06.SetFocus
               txtTF06_GotFocus
            End If
            GoTo flgFail
         End If
      'Add by Morgan 2008/1/7
      ElseIf txtTF03.Enabled = True And txtTF03 = "" Then
         '2010/1/8 MODIFY BY SONIA
         'strExc(0) = "select st16 from staff where st01='" & lblCP14 & "'"
         'intI = 1
         'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         'If intI = 1 Then
         '   If "" & RsTemp(0) = "3" Then
         '      MsgBox "當有完稿日時，若承辦人為日文組則日文字數不可空白！", vbExclamation
         '      txtTF03.SetFocus
         '      GoTo flgFail
         '   End If
         'End If
         
         'Removed by Morgan 2019/8/16 108.8.15起翻譯費改以原文字數計算,取消此管控--Sharon
         'If PUB_GetStaffST16(lblCP14) = "3" Then
         '   MsgBox "當有完稿日時，若承辦人為日文組則日文字數不可空白！", vbExclamation
         '   txtTF03.SetFocus
         '   GoTo flgFail
         'End If
         'end 2019/8/16
         
         '2010/1/8 END
      'end 2008/1/7
      End If
   End If
   'end 2007/5/22
   
   'Added by Morgan 2013/9/11
   '當輸入或更新完稿日時檢查(一案兩請發明案)
   m_UpdateCP09 = ""
   m_UpdateCP06 = "" 'Added by Lydia 2018/05/31
   
   'Modified by Morgan 2013/11/6 +235核對中說格式
   'Modified by Lydia 2018/05/31 + CP06所限
   If txtCaseNo(1) = "FCP" And stCP10 = "201" And txtEP09T <> "" And txtEP09T <> m_oldEP09T Then
      strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CaseNo" & _
         ",cp09,cp06 from (select cm01,cm02,cm03,cm04 from casemap where cm10='3'" & _
         " and cm05='" & txtCaseNo(1) & "' and cm06='" & txtCaseNo(2) & "' and cm07='" & txtCaseNo(3) & "' and cm08='" & txtCaseNo(4) & "'" & _
         " union select cm05,cm06,cm07,cm08 from casemap where cm10='3'" & _
         " and cm01='" & txtCaseNo(1) & "' and cm02='" & txtCaseNo(2) & "' and cm03='" & txtCaseNo(3) & "' and cm04='" & txtCaseNo(4) & "'" & _
         "),caseprogress,engineerprogress where cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03 and cp04(+)=cm04" & _
         " and cp10 in ('209','235') and cp27 is null and ep02(+)=cp09"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "一案兩請之新型案號為 " & RsTemp(0) & "，新型案之 [檢視中說] 將自動上 [齊備日] 及 [承辦期限] (同發明案之完稿日及核稿期限)！", vbInformation
         m_UpdateCP09 = RsTemp("cp09")
         m_UpdateCP06 = "" & RsTemp("cp06")  'Added by Lydia 2018/05/31
      End If
   End If
   'end 2013/9/11
   
   'Added by Lydia 2017/05/17 檢查原文字數、相似度、相似案號
   If Frame1.Visible = True And Trim(Replace(txtTF23 & txtTF19 & txtTF20, " ", "")) <> "" Then
      If Trim(txtTF23) = "" Then
         MsgBox "請輸入原文字數！", vbCritical
         txtTF23.SetFocus
         Exit Function
      End If
      'Modified by Lydia 2017/12/29 淑華要求107/1/1先開放單獨輸入原文字數
      'If Trim(txtTF19) = "" Then
     '    MsgBox "請輸入相似度！", vbCritical
     '    txtTF19.SetFocus
     '    Exit Function
      'ElseIf Val(txtTF19) > 100 Then
      If Val(txtTF19) > 100 Then
      'end 2017/12/29
         MsgBox "相似度不可大於100！"
         txtTF19.SetFocus
         Exit Function
      End If
      'Added by Lydia 2021/05/06 外專新案翻譯有相似度並且譯者為F外譯編號(排除F5588舜禹，F5698迅達，F5653捷恩凱)，於輸入翻譯完稿日後，計算相似折扣TF05(%)＝100－相似度TF19(%)。
      'Modified by Lydia 2025/03/13 改用模組取得
      'If txtEP09T.Tag = "" And txtEP09T.Text <> "" And stCP10 = "201" And Val(txtTF19) <> 0 And Left(lblCP14, 1) = "F" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, lblCP14) = 0 Then
      If txtEP09T.Tag = "" And txtEP09T.Text <> "" And stCP10 = "201" And Val(txtTF19) <> 0 And Left(lblCP14, 1) = "F" And InStr(Pub_SetF51Order("F", ""), lblCP14) = 0 Then
          txtTF05 = 100 - Val(txtTF19)
      End If
      'end 2021/05/06
   End If
   'end 2017/05/17
   
   TxtValidate = True
    
flgFail:
    bolIsValidate = True
    
End Function

Private Sub cmdok_Click()
   Call Fun_CmdOk
End Sub

'Modify By Sindy 2023/12/6
'Private Sub Fun_CmdOk(Optional bolExCall As Boolean = False)
Private Function Fun_CmdOk(Optional bolExCall As Boolean = False) As Boolean
'2023/12/6 END
'Dim strFullFileName As String
Dim objOutLook As Object
Dim objMail As Object
Dim myForward As Object
Dim jj As Integer
Dim ArrStr As Variant
Dim strContent As String, strSubject As String
   
   If TxtValidate Then
'      'Add By Sindy 2022/7/19
'      If Me.m_strIR01 <> "" Then
'         If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName) = True Then
'            Exit Sub
'         End If
'      End If
'      '2022/7/19 END
      
      'Added by Lydia 2024/05/07
      bolEP04trigger = False
      If txtEP04.Text <> txtEP04.Tag And Trim(txtEP04.Text) <> "" Then
         If PUB_GetST03(txtEP04.Text) = "F21" And PUB_GetStaffST16(txtEP04.Text) <> PUB_GetStaffST16(txtEP04.Tag) Then
            If MsgBox("是否變更工程師組別？", vbExclamation + vbYesNo + vbDefaultButton2, "工程師分組控管") = vbYes Then
               bolEP04trigger = True
            End If
         End If
      End If
      'end 2024/05/07
      
      m_strContactSheetA4 = "" 'Modify By Sindy 2022/8/12 借用此變數來傳回電子檔名
      If FormSave() = True Then
         Fun_CmdOk = True 'Add By Sindy 2023/12/6
         
         'Added by Lydia 2018/04/09 若完稿日有變更，自動列印承辦單
         If txtEP09T <> m_oldEP09T Then
            'Add By Sindy 2023/9/19
            If CmdPrint.Visible = True And CmdPrint.Caption = "翻譯承辦單" Then
            '2023/9/19 END
               cmdPrint_Click
               Sleep 1000
            End If
         End If
         'end 2018/04/09
         'Added by Lydia 2019/04/22 輸入Claims完稿日，自動列印會稿Claims承辦單
         'Modified by Lydia 2019/05/14 有翻譯完稿日不印
         'If txtEP09T = m_oldEP09T And txtEP31.Text <> "" And txtEP31.Text <> m_oldEP31T Then
         If txtEP09T = "" And txtEP31.Text <> "" And txtEP31.Text <> m_oldEP31T Then
            'Add By Sindy 2023/9/19
            If CmdPrint2.Visible = True Then
            '2023/9/19 END
               cmdPrint2_Click
               Sleep 1000
            End If
         End If
         'end 2019/04/22
         
         'Add By Sindy 2023/9/19
         If strSrvDate(1) < 外專承辦歷程啟用日 Then
         '2023/9/19 END
            'Modify By Sindy 2022/9/8 淑華:若是直接從系統翻譯完稿輸入去輸完稿日，一樣自動帶出outlook草稿，內容規則都相同
            'Modify By Sindy 2022/9/12 淑華說修改核稿人,不需產生Outlook草稿
            If Not (txtEP04.Tag <> txtEP04.Text) Then
               '產生outlook草稿
      '            Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", , strFullFileName, True) '下載信件檔,上傳卷宗區
      '            If strFullFileName <> "" Then
      '               '啟動轉寄功能
               Set objOutLook = CreateObject("Outlook.Application")
               Set objMail = objOutLook.CreateItem(0)
      '               Set objMail = objOutLook.CreateItemFromTemplate(strFullFileName) 'oForm.txtPathIPDept.Text & "\" & oFile.Name
      '               '*** 轉寄 *** 會用inbound名義寄出
      '               Set myForward = objMail.Forward '轉寄
      '               'Set myForward = objMail.ReplyAll ...不能用
      '               '移除原信的收件人及副本;密件副本不會留在msg中
      '               For jj = myForward.Recipients.Count To 1 Step -1
      '                  myForward.Recipients.Remove jj
      '               Next jj
      '               '副本
      '               myForward.cc = ""
      '               myForward.BCC = ""
      '               myForward.To = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
      '               '移除附件
      '               For jj = myForward.Attachments.Count To 1 Step -1
      '                  If UCase(Mid(Trim(myForward.Attachments(jj)), InStrRev(Trim(myForward.Attachments(jj)), "."))) = ".DOC" Or UCase(Mid(Trim(myForward.Attachments(jj)), InStrRev(Trim(myForward.Attachments(jj)), "."))) = ".XLS" Or UCase(Mid(Trim(myForward.Attachments(jj)), InStrRev(Trim(myForward.Attachments(jj)), "."))) = ".DOCX" Or UCase(Mid(Trim(myForward.Attachments(jj)), InStrRev(Trim(myForward.Attachments(jj)), "."))) = ".XLSX" Then
      '                     myForward.Attachments.Remove jj
      '                  End If
      '               Next jj
      '               '加入附件
      '               ArrStr = Split(m_FilePath, ";")
      '               For jj = 0 To UBound(ArrStr)
      '                  If Dir(ArrStr(jj)) <> "" Then
      '                     myForward.Attachments.add ArrStr(jj)
      '                  End If
      '               Next jj
      '               'myForward.senderemailaddress = "ipdept@taie.com.tw"
      '               myForward.sentonbehalfofname = ""
      '               myForward.Sender.address = ""
      '               myForward.HTMLBody = myForward.HTMLBody
      '               myForward.Display
      '               'myForward.Send
               
               'Added by Lydia 2023/10/04 FMP案待客戶最終指示相關控管：加註
               Dim strTmp As String
               If PUB_ChkFMP970mail("2", pa(1), pa(2), pa(3), pa(4), strTmp) = True Then
                  If strTmp <> "" Then strTmp = "1." & strTmp & vbCrLf
               End If
               'end 2023/10/04
            
               If txtEP31.Visible = True And txtEP31.Text <> m_oldEP31T Then
                  'Modified by Lydia 2023/10/04
                  'strSubject = "【Claims翻譯交稿】" & pa(1) & pa(2) '主旨
                  'strContent = "Claims翻譯已交稿，請通知工程師主管進行分案" & vbCrLf
                  strSubject = IIf(strTmp <> "", "【待最終指示】", "") & "【Claims翻譯交稿】" & pa(1) & pa(2) '主旨
                  strContent = strTmp & IIf(strTmp <> "", "2.", "") & "Claims翻譯已交稿，請通知工程師主管進行分案" & vbCrLf
                  If strTmp <> "" And Trim(txtCP64) <> "" Then
                     strContent = strContent & "3.進度備註:" & Trim(txtCP64) & vbCrLf
                  End If
                  'end 2023/10/04
               Else
                  'Modified by Lydia 2023/10/04
                  'strSubject = "【翻譯交稿】" & pa(1) & pa(2) '主旨
                  strSubject = IIf(strTmp <> "", "【待最終指示】", "") & "【翻譯交稿】" & pa(1) & pa(2) '主旨
                  If InStr(m_strContactSheetA4, ";") > 0 Then
                     'Modified by Lydia 2023/10/04
                     'strContent = "1.翻譯已交稿，請進行翻譯核稿流程" & vbCrLf & _
                                  "2.另有會稿說明書承辦單，請通知工程師主管進行分案" & vbCrLf
                     strContent = strTmp & IIf(strTmp <> "", "2.", "1.") & "翻譯已交稿，請進行翻譯核稿流程" & vbCrLf & _
                                  IIf(strTmp <> "", "3.", "2.") & "另有會稿說明書承辦單，請通知工程師主管進行分案" & vbCrLf
                     If strTmp <> "" And Trim(txtCP64) <> "" Then
                        strContent = strContent & "4.進度備註:" & Trim(txtCP64) & vbCrLf
                     End If
                     'end 2023/10/04
                  Else
                     'Modified by Lydia 2023/10/04
                     'strContent = "翻譯已交稿，請進行翻譯核稿流程" & vbCrLf
                     strContent = strTmp & IIf(strTmp <> "", "2.", "") & "翻譯已交稿，請進行翻譯核稿流程" & vbCrLf
                     If strTmp <> "" And Trim(txtCP64) <> "" Then
                        strContent = strContent & "3.進度備註:" & Trim(txtCP64) & vbCrLf
                    End If
                     'end 2023/10/04
                  End If
               End If
               
               '轉HTML格式
               strContent = Replace(strContent, "新細明體", "Times New Roman")
               '&nbsp; 不換行空格
               '&thinsp; 窄空格
               '單純只是想要輸入空白？ &nbsp; 就對了
               '&emsp; 全形空格
               '&ensp; 半形空格
               'strContent = Replace(strContent, "　", "&emsp;") '&emsp; 全形空格
               strContent = Replace(strContent, " ", "&thinsp;") '&ensp; 半形空格
               strContent = Replace(strContent, vbCrLf, "<BR>")
         '      If TypeName(objOutLook.Assistant) <> "Nothing" Then
         '         objOutLook.ActiveWindow.WindowState = 1 '0.最大化 1.視窗小點
         '      End If
               With objMail
                  '.BodyFormat = 2 '2=olFormatHTML 1=olFormatPlain 3=olFormatRichText
                  .To = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
                  .cc = ""
                  '加入附件
                  If m_strContactSheetA4 <> "" Then
                     ArrStr = Split(m_strContactSheetA4, ";")
                     For jj = 0 To UBound(ArrStr)
                        If Dir(ArrStr(jj)) <> "" Then
                           .Attachments.add ArrStr(jj) '加附件
                        End If
                     Next jj
                  End If
                  .Subject = strSubject
                  .HTMLBody = strContent
                  .Display
               End With
      '               strContent = strContent & "<BR><BR>------------------------------------------------------------------------------------------------------------------------------------------<BR>"
      '               myForward.HTMLBody = strContent & myForward.HTMLBody
      '               myForward.Display
      '               DoEvents
      '               Set myForward = Nothing
               Set objMail = Nothing
               Set objOutLook = Nothing
               '*** END
            End If
         End If
         
         If bolExCall = False Then
            'Add By Sindy 2022/5/12
            If Me.m_strIR01 <> "" Then
               Unload frm060107
               If Not m_PrevForm Is Nothing Then
                  Call m_PrevForm.GoNext
               End If
               Unload Me
            Else
            '2022/5/12 END
               cmdBack_Click
            End If
         End If
      Else
         MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
      End If
   End If
End Function

Private Sub Form_Activate()
   'Modify by Morgan 2009/6/30
   '發文後未付翻譯費前可修改字數
   If bolTfOnly = True Then
      txtEP09T.Enabled = False
      txtEP08T.Enabled = False
      'Add by Morgan 2009/7/30 預設在數學式字數
      If txtEP04 <> "" Then
         SendKeys "{Tab}"
         SendKeys "{Tab}"
      End If
   End If

   'Add by Morgan 2010/3/22
   If Not bolActived Then
      bolActived = True
'Removed by Morgan 2015/10/12 改核稿人自動設定至會稿或寄中說的承辦人
'      If txtEP04 = "" Then
'         'MODIFY BY SONIA 2014/6/23 +949寄中說
'         strExc(0) = "select * from caseprogress where cp01='" & txtCaseNo(1) & "'" & _
'            " and cp02='" & txtCaseNo(2) & "' and cp03='" & txtCaseNo(3) & "'" & _
'            " and cp04='" & txtCaseNo(4) & "' and cp10 in ('924','949') and cp14 is null and cp57 is null"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            MsgBox "本案會稿或寄中說尚未分案！"
'         End If
'      End If
'end 2015/10/12
   End If
End Sub

'2011/11/30 add by sonia
Private Sub Form_Initialize()
   ReDim pa(1 To TF_PA) As String
End Sub
'2011/11/30 end

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC    '2011/11/30 add by sonia
   bolIsValidate = True
   
   lblAppDate.BackColor = &H8000000F
   lblCaseName(1).BackColor = &H8000000F
   lblCaseName(2).BackColor = &H8000000F
   lblCaseName(3).BackColor = &H8000000F
   lblCP05T.BackColor = &H8000000F
   lblCP10T.BackColor = &H8000000F
   lblCP09.BackColor = &H8000000F
   lblPA08T.BackColor = &H8000000F
   lblCP14.BackColor = &H8000000F
   lblCP14T.BackColor = &H8000000F
   'Add by Morgan 2011/5/31
   lblEP04T.BackColor = &H8000000F
   lblCP27T.BackColor = &H8000000F
   
   'Add By Sindy 2022/5/12
   m_strIR01 = frm060107.m_strIR01
   m_strIR02 = frm060107.m_strIR02
   m_strIR03 = frm060107.m_strIR03
   m_strIR04 = frm060107.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2022/5/12 END
   
   'Add By Sindy 2023/9/15
   If strSrvDate(1) >= 外專承辦歷程啟用日 Then
      Check1.Visible = False '產生電子檔
      CmdPrint.Caption = "承辦歷程"
      CmdPrint2.Visible = False
   End If
   '2023/9/15 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2010/6/18
   
   'Add By Sindy 2022/5/12
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   '2022/5/12 END
   
   Set frm060107_1 = Nothing
End Sub

Private Sub txtEP08T_GotFocus()
    TextInverse txtEP08T
End Sub

Private Sub txtEP08T_Validate(Cancel As Boolean)
    If txtEP08T = "" Then
        '存檔檢查時若有完稿日則不可空白
        'Modified by Morgan 2013/4/10 P案不必檢查--靜芳 Ex.P-104940 其他翻譯
        If Not bolIsValidate And txtEP09T <> "" And txtCaseNo(1) <> "P" Then
            MsgBox "有完稿日時，核稿期限不可空白！", vbCritical
            Cancel = True
            txtEP08T.SetFocus
        End If
    ElseIf Not ChkDate(txtEP08T) Then
        Cancel = True
        If Not bolIsValidate Then txtEP08T.SetFocus
        Call txtEP08T_GotFocus
    End If
End Sub

Private Sub txtCP64_Validate(Cancel As Boolean)
    Cancel = Not CheckLengthIsOK(txtCP64, 2000)
End Sub

Private Sub txtEP04_GotFocus()
    TextInverse txtEP04
End Sub

Private Sub txtEP04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtEP04_Validate(Cancel As Boolean)
'Dim m_Team As String '2011/11/30 add by sonia 'Remove by Lydia 2021/01/06
    
    lblEP04T = ""
    If txtEP04 = "" Then
'        '存檔檢查時不可空白
'        If Not bolIsValidate Then
'            MsgBox "核稿人不可空白！", vbCritical
'            Cancel = True
'            txtEP04.SetFocus
'        End If
    Else
        Dim strName As String
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetStaff(txtEP04, strName) Then
'Mark by Lydia 2021/01/06 改到SetData
'        'Move by Lydia 2016/06/20
'        pa(1) = txtCaseNo(1): pa(2) = txtCaseNo(2): pa(3) = txtCaseNo(3): pa(4) = txtCaseNo(4)
'        'Modified by Lydia 2018/09/13 + P案
'        If txtCaseNo(1) = "FCP" Or txtCaseNo(1) = "P" Then
'           If ClsPDReadPatentDatabase(pa(), intWhere) Then m_Team = pa(150)
'        'Modified by Lydia 2018/09/13 + PS案
'        ElseIf txtCaseNo(1) = "FG" Or txtCaseNo(1) = "PS" Then
'           If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then m_Team = pa(79)
'        End If
'end 2021/01/06

'Modified by Lydia 2016/06/20 改成模組
'        If ClsPDGetStaff(txtEP04, strName) Then
'             '2011/11/30 ADD BY SONIA 林信昌因分組故自動帶與案件組別的編號
'            If InStr(strName, "林信昌") > 0 Then
'               'Modified by Lydia 2016/06/20 移到上方
'               Select Case m_Team
'                  Case "1"
'                     If Left(txtEP04, 1) = "6" Then txtEP04 = "68091"
'                     If Left(txtEP04, 1) = "F" Then txtEP04 = "F5644"
'                  Case "2"
'                     If Left(txtEP04, 1) = "6" Then txtEP04 = "68092"
'                     If Left(txtEP04, 1) = "F" Then txtEP04 = "F5645"
'                  Case Else
'                     If Left(txtEP04, 1) = "6" Then txtEP04 = "68007"
'                     If Left(txtEP04, 1) = "F" Then txtEP04 = "F5162"
'               End Select
'               If ClsPDGetStaff(txtEP04, strName) Then
'               End If
'            End If
'            '2011/11/30 END
'            lblEP04T = strName
'            'Add by Morgan 2008/2/22 控制只能為外專工程師   2008/4/8 加 F81
'            If InStr("F21,F52,F81", GetStaffDepartment(txtEP04)) = 0 Then
'               MsgBox "核稿人僅能輸外專工程師！"
'               Cancel = True
'               If Not bolIsValidate Then txtEP04.SetFocus
'               Call txtEP04_GotFocus
'            End If
'        Else
'            lblEP04T = ""
'            Cancel = True
'            If Not bolIsValidate Then txtEP04.SetFocus
'            Call txtEP04_GotFocus
'        End If
        If PUB_FCPGetCP14EP04("EP04", pa, txtEP04, lblEP04T, Cancel) Then
        End If
        If Cancel = True Then
            If Not bolIsValidate Then txtEP04.SetFocus
            Call txtEP04_GotFocus
        End If
    End If
'end 2016/06/20

End Sub

Private Sub txtEP09T_GotFocus()
    TextInverse txtEP09T
    txtEP09T.Tag = "" 'Add by Morgan 2010/4/29
End Sub

Private Sub txtEP09T_Validate(Cancel As Boolean)

    If txtEP09T = "" Then
        'Modify by Morgan 2004/3/8
        '不清除核稿人承辦期限
        'txtEP08T = ""
        
'        '存檔檢查時不可空白
'        If Not bolIsValidate Then
'            MsgBox "完稿日不可空白！", vbCritical
'            Cancel = True
'            txtEP09T.SetFocus
'        End If
    ElseIf ChkDate(txtEP09T) Then
      'Modified by Lydia 2016/06/27
      'If bolIsValidate Then Call SetEP08
      If bolIsValidate Then
         'Added by Lydia 2021/01/06 Murgitroyd案：新案發文日＋2個月+15個日曆天再往前推3個工作天
         If stCP10 = "201" And strMurgitroyd <> "" And InStr(strMurgitroyd, ChangeCustomerL(pa(75))) > 0 Then
              If txtEP08T.Tag = "" Then
                  If PUB_FCPsetEP08M(txtCaseNo(1).Text, txtCaseNo(2).Text, txtCaseNo(3).Text, txtCaseNo(4).Text, stCP06, stCP10, lblAppDate.Caption, strExc(1), True) = True Then
                       If strExc(1) <> "" Then txtEP08T = TransDate(strExc(1), 1)
                  End If
              End If
         Else
         'end 2021/01/06
              Call PUB_FCPsetEP08(txtCaseNo(1).Text, txtCaseNo(2).Text, txtCaseNo(3).Text, txtCaseNo(4).Text, stCP06, stCP10, lblAppDate.Caption, lblCP14.Caption, txtEP04.Text, txtEP08T, txtEP09T)
         End If 'Added by Lydia 2021/01/06
         'Added by Lydia 2021/05/06 外專新案翻譯有相似度並且譯者為F外譯編號(排除F5588舜禹，F5698迅達，F5653捷恩凱)，於輸入翻譯完稿日後，計算相似折扣TF05(%)＝100－相似度TF19(%)。
         'Modified by Lydia 2025/03/13 改用模組取得
         'If stCP10 = "201" And Val(txtTF19) <> 0 And Left(lblCP14, 1) = "F" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, lblCP14) = 0 Then
         If stCP10 = "201" And Val(txtTF19) <> 0 And Left(lblCP14, 1) = "F" And InStr(Pub_SetF51Order("F", ""), lblCP14) = 0 Then
             txtTF05 = 100 - Val(txtTF19)
         End If
         'end 2021/05/06
      End If
      'end 2016/06/27
    Else
        Cancel = True
        If Not bolIsValidate Then txtEP09T.SetFocus
        Call txtEP09T_GotFocus
    End If
    If Not Cancel Then txtEP09T.Tag = txtEP09T
End Sub
'Remove by Lydia 2016/06/21 改成模組 PUB_FCPsetEP08
'Private Sub SetEP08()
'   Dim dtEP09 As Date, dtEP08 As Date, dtTmp1 As Date, dtTmp2 As Date
'
'   'Add by Morgan 2005/9/12 若原先已有核稿期限時不再重算
'   If txtEP08T.Tag = "" Then
'
'      txtEP08T = ""
'      dtEP09 = ChangeTStringToWDateString(txtEP09T)
'      If lblAppDate = "" Then
'          MsgBox "尚未輸入申請日！", vbExclamation
'          dtTmp1 = 0
'      Else
'          dtTmp1 = DateAdd("m", 6, ChangeTStringToWDateString(lblAppDate)) - 4
'      End If
'      Select Case stCP10
'          '翻譯
'          Case "201"
'              'Add by Morgan 2008/8/21
'              'P案翻譯核稿期限=完稿日+5個工作天
'              If txtCaseNo(1) = "P" Then
'                 dtTmp2 = CDate(ChangeWStringToWDateString(CompWorkDay(5, DBDATE(txtEP09T))))
'              Else
'              'end 2008/8/21
'                 'Modify by Morgan 2008/10 改成外翻22個工作天,工程師10個工作天
'                 '外翻：核稿承辦期限=完稿日+4週
'                 'Modify by Morgan 2008/11/4 改判斷承辦人與核稿人是否相同--與靜芳討論後暫訂
'                 'If stST15 = "F51" Then
'                 If lblCP14 <> txtEP04 Then
'                     'dtTmp2 = DateAdd("ww", 4, dtEP09)
'                     dtTmp2 = CDate(ChangeWStringToWDateString(CompWorkDay(22, DBDATE(txtEP09T))))
'                 '內翻：核稿承辦期限=完稿日+10天
'                 '2008/6/10 modify by sonia 靜芳說改7天
'                 '2008/8/19 modify by sonia 靜芳說再改回10天
'                 Else
'                     'dtTmp2 = DateAdd("d", 10, dtEP09)
'                     dtTmp2 = CDate(ChangeWStringToWDateString(CompWorkDay(10, DBDATE(txtEP09T))))
'                 End If
'              End If
'
'      'Remove by Morgan 2008/10/21 改由分案輸入齊備日
'      '        '檢視中說
'      '        Case "209"
'      '            '核稿承辦期限=完稿日+3週
'      '            dtTmp2 = DateAdd("ww", 3, dtEP09)
'      '        '製作中說
'      '        Case "210"
'      '            '核稿承辦期限=NVL(本所期限,申請案本所期限)
'      '            If stCP06 = "" Then
'      '                dtTmp2 = GetAppDate
'      '            Else
'      '                dtTmp2 = ChangeWDateStringToWString(stCP06)
'      '            End If
'
'      End Select
'
'      '核稿承辦期限<=申請日+6個月-4天
'      If dtTmp2 <> 0 Then
'         'Add by Morgan 2008/8/21
'         'P案不必控制
'         If txtCaseNo(1) = "P" Then
'            dtEP08 = dtTmp2
'         Else
'         'end 2008/8/21
'            If dtTmp1 = 0 Then
'                dtEP08 = dtTmp2
'            ElseIf DateDiff("d", dtTmp1, dtTmp2) > 0 Then
'                dtEP08 = dtTmp1
'            Else
'                dtEP08 = dtTmp2
'            End If
'         End If
'
'         '2005/7/22 ADD BY SONIA 承辦期限不可大於本所期限
'         'Modify by Morgan 2005/9/4 先判斷有本所期限
'         'If stCP06 <> "" And dtEP08 > ChangeWStringToWDateString(stCP06) Then
'         If stCP06 <> "" Then
'             If dtEP08 > ChangeWStringToWDateString(stCP06) Then
'                dtEP08 = ChangeWStringToWDateString(stCP06)
'             End If
'         End If
'         txtEP08T = Format(Val(Format(dtEP08, "YYYYMMDD")) - 19110000)
'      End If
'
'   End If
'
'   'Add by Morgan 2010/4/29 若核稿期限大於會稿的本所期限時提醒會更新
'   If txtEP08T <> "" Then
'      'MODIFY BY SONIA 2014/6/23 +949寄中說
'      strExc(0) = "select cp06 from caseprogress where cp01='" & txtCaseNo(1) & "'" & _
'         " and cp02='" & txtCaseNo(2) & "' and cp03='" & txtCaseNo(3) & "'" & _
'         " and cp04='" & txtCaseNo(4) & "' and cp10 in ('924','949') and cp57||cp27 is null and cp06<" & DBDATE(txtEP08T)
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         MsgBox "將更新核稿期限為【會稿】或【寄中說】之本所期限！"
'         txtEP08T = TransDate(RsTemp(0), 1)
'      End If
'   End If
'
'End Sub

Private Function GetAppDate() As Date

On Error GoTo flgErr

    Dim stSQL As String
    
    stSQL = "SELECT CP06 FROM CASEPROGRESS WHERE CP57 IS NULL AND CP06 IS NOT NULL AND CP31='Y'" & _
        " AND CP01='" & txtCaseNo(1) & "' AND CP02='" & txtCaseNo(2) & "'" & _
        " AND CP03='" & txtCaseNo(3) & "' AND CP04='" & txtCaseNo(4) & "'"
        
    CheckOC
    adoRecordset.CursorLocation = adUseClient
    adoRecordset.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
    
    If adoRecordset.RecordCount > 0 Then
        GetAppDate = ChangeWDateStringToWString("" & adoRecordset.Fields("CP06"))
    End If
    CheckOC
    
flgErr:
    
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical
    End If
    
End Function

Private Sub txtTF03_GotFocus()
   TextInverse txtTF03
End Sub

Private Sub txtTF03_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtTF04_GotFocus()
   TextInverse txtTF04
End Sub

Private Sub txtTF04_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtTF05_GotFocus()
   TextInverse txtTF05
End Sub

Private Sub txtTF05_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtTF05_Validate(Cancel As Boolean)
   If Val(txtTF05) > 100 Then
      MsgBox "不可大於100！", vbCritical
      Cancel = True
   End If
End Sub

Private Sub txtTF06_GotFocus()
   TextInverse txtTF06
End Sub

Private Sub txtTF06_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtTF06_Validate(Cancel As Boolean)
   If Val(txtTF06) > 100 Then
      MsgBox "不可大於100！", vbCritical
      Cancel = True
   End If
End Sub

Private Sub txtTF18_GotFocus()
   TextInverse txtTF18
End Sub

Private Sub txtTF18_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2017/05/17
Private Sub txtTF23_GotFocus()
   TextInverse txtTF23
End Sub

Private Sub txtTF23_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtTF19_GotFocus()
   TextInverse txtTF19
End Sub

Private Sub txtTF19_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtTF20_GotFocus()
   TextInverse txtTF20
End Sub

Private Sub txtTF20_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtTF23_Validate(Cancel As Boolean)
   If txtTF23 <> "" Then
      txtTF23 = Val(txtTF23)
   End If
End Sub

Private Sub txtTF19_Validate(Cancel As Boolean)
   If txtTF19 <> "" Then
      If Val(txtTF19) > 100 Then
         MsgBox "相似度不可大於100！"
         Cancel = True
         TextInverse txtTF19
      End If
      txtTF19 = Val(txtTF19)
   End If
End Sub

Private Sub txtTF20_Validate(Cancel As Boolean)
   If txtTF20 <> "" Then
      Call ChgCaseNo(txtTF20, strExc)
      If ClsPDCheckCaseCodeIsExist(strExc(1), strExc(2), strExc(3), strExc(4)) = False Then
         Cancel = True
         TextInverse txtTF20
      Else
         txtTF20.Text = strExc(1) & strExc(2) & strExc(3) & strExc(4)
      End If
   End If
End Sub
'end 2017/05/17

'Added by Lydia 2018/04/11 外專翻譯承辦單列印
Private Sub cmdPrint_Click()
Dim strColName() As String, strColText() As String 'Add By Sindy 2023/9/14
   
   'Add By Sindy 2023/9/19
   If strSrvDate(1) >= 外專承辦歷程啟用日 Then
      If Fun_CmdOk(True) = False Then Exit Sub
      If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub
      frm090202_2.Hide
      If txtEP09T = "" And txtEP31.Text <> "" Then
         If PUB_ChkCPExist(pa, "924", 1, strExc(1), strExc(2), "A") = True Then
            frm090202_2.m_EEP01 = strExc(1) '會稿的總收文號
         Else
            frm090202_2.m_EEP01 = lblCP09 '翻譯的總收文號
         End If
      Else
         frm090202_2.m_EEP01 = lblCP09 '翻譯的總收文號
      End If
      frm090202_2.m_FlowUserNum = strUserNum '案件流程所屬人員
      frm090202_2.intReceiveKind = 99
      'Add By Sindy 2024/1/2
      frm090202_2.m_strIR01 = Me.m_strIR01
      frm090202_2.m_strIR02 = Me.m_strIR02
      frm090202_2.m_strIR03 = Me.m_strIR03
      frm090202_2.m_strIR04 = Me.m_strIR04
      frm090202_2.SetParent_IR m_PrevForm
      '2024/1/2 END
      frm090202_2.SetParent Me
      Unload frm060107
      If frm090202_2.QueryData = True Then
         frm090202_2.Show
         Me.Hide
      End If
   Else
   '2023/9/19 END
      '不輸入Claims完稿日,於後面列印翻譯承辦單+會稿說明書承辦單
      'Modified by Lydia 2020/04/10 +產生電子檔
      'Call Pub_PrintFCP201Form(txtCaseNo(1), txtCaseNo(2), txtCaseNo(3), txtCaseNo(4), lblCP09.Caption)
      'Modify By Sindy 2023/9/14 +, strColName, strColText
      Call Pub_PrintFCP201Form(txtCaseNo(1), txtCaseNo(2), txtCaseNo(3), txtCaseNo(4), lblCP09.Caption, strColName, strColText, IIf(Check1.Value = 1, True, False))
   End If
End Sub

'Added by Lydia 2019/04/17 外專會稿承辦單列印
Private Sub cmdPrint2_Click()
Dim strColName() As String, strColText() As String 'Add By Sindy 2023/9/14
   'Modified by Lydia 2019/09/23 改用案號尋找
   'Call Pub_PrintFCP924Form(txtCaseNo(1), txtCaseNo(2), txtCaseNo(3), txtCaseNo(4), lblCP09.Caption)
   'Modified by Lydia 2020/04/10 +產生電子檔
   'Call Pub_PrintFCP924Form(txtCaseNo(1), txtCaseNo(2), txtCaseNo(3), txtCaseNo(4), "")
   'Modify By Sindy 2023/9/14 +, strColName, strColText
   Call Pub_PrintFCP924Form(txtCaseNo(1), txtCaseNo(2), txtCaseNo(3), txtCaseNo(4), "", strColName, strColText, IIf(Check1.Value = 1, True, False))
End Sub

'Added by Lydia 2018/05/07
Private Sub txtCP113_GotFocus()
    TextInverse txtCP113
End Sub

Private Sub txtCP113_KeyPress(KeyAscii As Integer)
    'Remove by Lydia 2018/12/12
    'KeyAscii = Pub_NumAscii(KeyAscii)
End Sub
'end 2018/05/07

'Added by Lydia 2018/12/12
Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

'Added by Lydia 2019/04/17
Private Sub txtTF32_GotFocus()
    TextInverse txtTF32
End Sub

Private Sub txtTF32_Validate(Cancel As Boolean)
     If txtTF32.Text <> "" Then
         If ChkDate(txtTF32.Text) = True Then
         Else
             GoTo JumpExit
         End If
     End If
     
     Exit Sub
     
JumpExit:
    Cancel = True
    txtTF32.SetFocus
    Call txtTF32_GotFocus
End Sub

Private Sub txtEP31_GotFocus()
    TextInverse txtEP31
End Sub

Private Sub txtEP31_Validate(Cancel As Boolean)
     If txtEP31.Text <> "" Then
         If ChkDate(txtEP31.Text) = True Then
             If Val(txtTF32) = 0 Then
                  MsgBox "尚未輸入只交Claims期限!", vbCritical
                  GoTo JumpExit
             'Modified by Lydia 2019/05/14
             'ElseIf Val(txtEP31) < Val(txtTF32) Then
             '     MsgBox "完稿日期不可小於交Claims期限!", vbCritical
             ElseIf Val(txtEP31) > Val(strSrvDate(2)) Then
                  MsgBox "完稿日期不可大於系統日!", vbCritical
             'end 2019/05/14
                  GoTo JumpExit
             End If
         Else
             GoTo JumpExit
         End If
     End If
     
     Exit Sub
     
JumpExit:
    Cancel = True
    txtEP31.SetFocus
    Call txtEP31_GotFocus
End Sub

