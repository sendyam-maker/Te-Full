VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03010301_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "非爭議案核准輸入"
   ClientHeight    =   6048
   ClientLeft      =   1680
   ClientTop       =   1536
   ClientWidth     =   9144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6048
   ScaleWidth      =   9144
   Begin VB.ComboBox Combo1 
      Height          =   276
      Left            =   4176
      TabIndex        =   87
      Text            =   "Combo1"
      Top             =   5064
      Width           =   1620
   End
   Begin VB.TextBox textCP30 
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   5670
      Visible         =   0   'False
      Width           =   2900
   End
   Begin VB.TextBox textCP18 
      Height          =   264
      Left            =   7770
      TabIndex        =   19
      Top             =   5387
      Width           =   972
   End
   Begin VB.CheckBox Check1 
      Caption         =   "公報"
      Height          =   285
      Index           =   2
      Left            =   3348
      TabIndex        =   22
      Top             =   5700
      Width           =   1000
   End
   Begin VB.CheckBox Check1 
      Caption         =   "核准通知書"
      Height          =   285
      Index           =   1
      Left            =   2004
      TabIndex        =   21
      Top             =   5700
      Width           =   1332
   End
   Begin VB.CheckBox Check1 
      Caption         =   "公告通知書"
      Height          =   285
      Index           =   0
      Left            =   660
      TabIndex        =   20
      Top             =   5700
      Width           =   1332
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Left            =   8010
      TabIndex        =   2
      Text            =   "個月"
      Top             =   2923
      Width           =   975
   End
   Begin VB.TextBox textCP53 
      Height          =   285
      Left            =   6024
      MaxLength       =   7
      TabIndex        =   5
      Top             =   3231
      Visible         =   0   'False
      Width           =   1308
   End
   Begin VB.TextBox textCP54 
      Height          =   285
      Left            =   7704
      MaxLength       =   7
      TabIndex        =   6
      Top             =   3231
      Visible         =   0   'False
      Width           =   1308
   End
   Begin VB.TextBox textTM27 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1691
      Width           =   2532
   End
   Begin VB.TextBox textTM16S 
      Height          =   285
      Left            =   1410
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3540
      Width           =   405
   End
   Begin VB.TextBox textTM17 
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   8
      Top             =   3847
      Width           =   372
   End
   Begin VB.TextBox textFee_2 
      Height          =   285
      Left            =   5640
      TabIndex        =   18
      Top             =   5387
      Width           =   972
   End
   Begin VB.TextBox textFee_1 
      Height          =   285
      Left            =   1620
      TabIndex        =   17
      Top             =   5387
      Width           =   972
   End
   Begin VB.TextBox textCP48 
      Height          =   285
      Left            =   5640
      MaxLength       =   8
      TabIndex        =   13
      Top             =   4463
      Width           =   2532
   End
   Begin VB.TextBox textCP07 
      Height          =   285
      Left            =   5640
      MaxLength       =   8
      TabIndex        =   11
      Top             =   4155
      Width           =   2532
   End
   Begin VB.TextBox textCP14 
      Height          =   285
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   12
      Top             =   4463
      Width           =   732
   End
   Begin VB.TextBox textCP06 
      Height          =   285
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   10
      Top             =   4155
      Width           =   2532
   End
   Begin VB.TextBox textCF15_2 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1980
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   3847
      Width           =   1692
   End
   Begin VB.TextBox textCF15 
      Height          =   285
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   9
      Top             =   3847
      Width           =   732
   End
   Begin VB.TextBox textCP26 
      Height          =   285
      Left            =   7560
      MaxLength       =   1
      TabIndex        =   16
      Top             =   5079
      Width           =   372
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   444
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1383
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   444
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1383
      Width           =   2532
   End
   Begin VB.TextBox textTM22S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   6480
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1691
      Width           =   1692
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1999
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1999
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   2307
      Width           =   2532
   End
   Begin VB.TextBox textCP45 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2307
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2615
      Width           =   2412
   End
   Begin VB.TextBox textCP25 
      Height          =   285
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2923
      Width           =   1215
   End
   Begin VB.TextBox textTM14 
      Height          =   285
      Left            =   5640
      MaxLength       =   8
      TabIndex        =   1
      Top             =   2923
      Width           =   1215
   End
   Begin VB.TextBox textTMBM07_1 
      Height          =   285
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   3
      Top             =   3231
      Width           =   732
   End
   Begin VB.TextBox textTMBM07_2 
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   4
      Top             =   3231
      Width           =   732
   End
   Begin VB.TextBox textPrint 
      Height          =   285
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   15
      Top             =   5079
      Width           =   732
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8160
      TabIndex        =   27
      Top             =   24
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6000
      TabIndex        =   25
      Top             =   24
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6960
      TabIndex        =   26
      Top             =   24
      Width           =   1152
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "變更事項(R)"
      Height          =   400
      Left            =   4800
      TabIndex        =   24
      Top             =   24
      Width           =   1152
   End
   Begin VB.Label Label29 
      Caption         =   "定稿案件性質 :"
      Height          =   240
      Left            =   2904
      TabIndex        =   86
      Top             =   5112
      Width           =   1260
   End
   Begin MSForms.TextBox textPS 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Top             =   4770
      Width           =   7815
      VariousPropertyBits=   -1475330021
      MaxLength       =   2000
      Size            =   "13785;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14_2 
      Height          =   285
      Left            =   1980
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   4470
      Width           =   2535
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5640
      TabIndex        =   84
      Top             =   2610
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   1050
      Width           =   7035
      VariousPropertyBits=   671105055
      Size            =   "12409;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1170
      TabIndex        =   82
      Top             =   750
      Width           =   7875
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13891;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label28 
      Caption         =   "美國移轉登記號："
      Height          =   285
      Left            =   4680
      TabIndex        =   81
      Top             =   5700
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label Label14 
      Caption         =   "點數 :"
      Height          =   285
      Left            =   7140
      TabIndex        =   80
      Top             =   5387
      Width           =   525
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3810
      TabIndex        =   79
      Top             =   444
      Width           =   645
   End
   Begin VB.Label Label8 
      Caption         =   "附件 :"
      Height          =   285
      Left            =   120
      TabIndex        =   78
      Top             =   5700
      Width           =   588
   End
   Begin VB.Label Label2 
      Caption         =   "公告期間 : "
      Height          =   285
      Left            =   7080
      TabIndex        =   77
      Top             =   2923
      Width           =   888
   End
   Begin VB.Label Label21 
      Caption         =   "(1:准 , 2:駁)"
      Height          =   285
      Left            =   1980
      TabIndex        =   73
      Top             =   3540
      Width           =   948
   End
   Begin VB.Label Label4 
      Caption         =   "質權設定期間 :"
      Height          =   285
      Index           =   0
      Left            =   4680
      TabIndex        =   76
      Top             =   3231
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "－"
      Height          =   285
      Index           =   1
      Left            =   7440
      TabIndex        =   75
      Top             =   3231
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label Label27 
      Caption         =   "案件目前准駁 :"
      Height          =   285
      Left            =   120
      TabIndex        =   74
      Top             =   3540
      Width           =   2292
   End
   Begin VB.Label Label19 
      Caption         =   "專用權是否存在 :"
      Height          =   285
      Left            =   4680
      TabIndex        =   72
      Top             =   3847
      Width           =   1572
   End
   Begin VB.Label Label20 
      Caption         =   "(Y / N)"
      Height          =   285
      Left            =   6720
      TabIndex        =   71
      Top             =   3847
      Width           =   612
   End
   Begin VB.Label Label18 
      Caption         =   "領證費 :"
      Height          =   285
      Left            =   4680
      TabIndex        =   70
      Top             =   5387
      Width           =   852
   End
   Begin VB.Label Label17 
      Caption         =   "委託代理人費用 :"
      Height          =   285
      Left            =   120
      TabIndex        =   69
      Top             =   5387
      Width           =   1425
   End
   Begin VB.Label Label9 
      Caption         =   "列印備註 :"
      Height          =   285
      Left            =   120
      TabIndex        =   68
      Top             =   4771
      Width           =   972
   End
   Begin VB.Label Label26 
      Caption         =   "承辦期限 :"
      Height          =   285
      Left            =   4680
      TabIndex        =   67
      Top             =   4463
      Width           =   852
   End
   Begin VB.Label Label25 
      Caption         =   "法定期限 :"
      Height          =   285
      Left            =   4680
      TabIndex        =   66
      Top             =   4155
      Width           =   852
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   285
      Left            =   120
      TabIndex        =   65
      Top             =   4463
      Width           =   852
   End
   Begin VB.Label Label5 
      Caption         =   "本所期限 :"
      Height          =   285
      Left            =   120
      TabIndex        =   64
      Top             =   4155
      Width           =   852
   End
   Begin VB.Label Lbl4 
      Caption         =   "下一程序 :"
      Height          =   285
      Left            =   120
      TabIndex        =   63
      Top             =   3847
      Width           =   852
   End
   Begin VB.Label Label16 
      Caption         =   "是否算案件數 :"
      Height          =   288
      Left            =   6264
      TabIndex        =   62
      Top             =   5076
      Width           =   1212
   End
   Begin VB.Label Label15 
      Caption         =   "(N:不算)"
      Height          =   288
      Left            =   8064
      TabIndex        =   61
      Top             =   5076
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   59
      Top             =   444
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   285
      Left            =   120
      TabIndex        =   58
      Top             =   744
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   285
      Left            =   120
      TabIndex        =   57
      Top             =   1044
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   56
      Top             =   1383
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   55
      Top             =   444
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   285
      Index           =   3
      Left            =   4680
      TabIndex        =   54
      Top             =   1383
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "正商標號數 :"
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   53
      Top             =   1695
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "正商標專用期止日 :"
      Height          =   285
      Index           =   5
      Left            =   4680
      TabIndex        =   52
      Top             =   1691
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   285
      Index           =   6
      Left            =   120
      TabIndex        =   51
      Top             =   1999
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   285
      Index           =   7
      Left            =   4680
      TabIndex        =   50
      Top             =   1999
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   285
      Index           =   8
      Left            =   120
      TabIndex        =   49
      Top             =   2307
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   285
      Index           =   9
      Left            =   4680
      TabIndex        =   48
      Top             =   2307
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   285
      Index           =   10
      Left            =   120
      TabIndex        =   47
      Top             =   2615
      Width           =   1212
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   285
      Index           =   11
      Left            =   4680
      TabIndex        =   46
      Top             =   2615
      Width           =   900
   End
   Begin VB.Label Label7 
      Caption         =   "核准通知日 :"
      Height          =   285
      Left            =   120
      TabIndex        =   45
      Top             =   2923
      Width           =   1092
   End
   Begin VB.Label Label10 
      Caption         =   "公告日 :"
      Height          =   285
      Left            =   4680
      TabIndex        =   44
      Top             =   2923
      Width           =   732
   End
   Begin VB.Label Label11 
      Caption         =   "公報卷期 :"
      Height          =   285
      Left            =   120
      TabIndex        =   43
      Top             =   3231
      Width           =   972
   End
   Begin VB.Label Label12 
      Caption         =   "卷"
      Height          =   285
      Left            =   2040
      TabIndex        =   42
      Top             =   3231
      Width           =   252
   End
   Begin VB.Label Label13 
      Caption         =   "期"
      Height          =   285
      Left            =   3480
      TabIndex        =   41
      Top             =   3231
      Width           =   252
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   285
      Left            =   120
      TabIndex        =   40
      Top             =   5079
      Width           =   972
   End
   Begin VB.Label Label23 
      Caption         =   "(N:不印)"
      Height          =   285
      Left            =   2040
      TabIndex        =   39
      Top             =   5079
      Width           =   852
   End
End
Attribute VB_Name = "frm03010301_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/12 改成Form2.0 ; textTM23、cmbTM05、textCP13、textCP14_2、textPS
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 來函收文日
Dim m_CP05 As String
' 原法定期限
Dim m_CP07 As String
' 收文號
Dim m_CP09 As String
' 原案件性質
Dim m_CP10 As String
' 原業務區
Dim m_CP12 As String
' 原智權人員代號
Dim m_CP13 As String
' 原授權期間(迄)
Dim m_CP54 As String
' 原移轉申請人代號
Dim m_CP56 As String
'Add By Sindy 2013/1/11
Dim m_CP89 As String
Dim m_CP90 As String
Dim m_CP91 As String
Dim m_CP92 As String
'2013/1/11 End
' 商標種類代碼
Dim m_TM08 As String
' 國家代碼
Dim m_TM10 As String
' 原專用期限起日
Dim m_TM21 As String
' 原專用期限止日
Dim m_TM22 As String
' 原申請人代號
Dim m_TM23 As String
'Add By Sindy 2013/1/11
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
'2013/1/11 End
' 申請國家的延展年度
Dim m_NA14 As Integer
'add by nickc 2005/05/31
Dim IsAppNpData As Boolean
Dim SeekNewCp09 As String
Dim oClsPrtForm001 As New ClsPrtForm001
'add by nickc 2005/08/04
'Dim m_blnClkChgButton As Boolean '是否有按變更事項鈕
Public m_blnClkChgButton As Boolean '是否有按變更事項鈕 'Modify By Sindy 2012/2/6 Dim->Public
'Add By Sindy 2023/4/27
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2023/4/27 END
Dim m_strCP09 As String, m_strCP10 As String  'Added by Lydia 2023/09/12 定稿收文號和案件性質
Dim NewCP09 As String 'Added by Lydia 2023/09/12 核准產生之收文號
Dim bolJumpChgEvent As Boolean 'Added by Lydia 2024/03/21 不用輸入變更事項
Dim m_CP24 As String 'Added by Lydia 2025/09/12

'Add By Sindy 2023/4/27
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm03010301_02.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm03010301_02
   Unload frm03010301_01
   Unload Me
End Sub

Private Sub cmdMod_Click()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   'add by nickc 2005/08/04
   m_blnClkChgButton = True
    
   'Add By Sindy 2010/4/12
   '變更事項確認鍵, 會直接儲存前畫面, 所以先檢查資料
   If CheckDataValid = False Then Exit Sub
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   '2010/4/12 End
   
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount <= 0 Then
      rsTmp.Close
      strMsg = "無變更事項記錄"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   m_blnClkChgButton = False 'Add By Sindy 2012/2/14
   DisplayNextForm
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub cmdok_Click()
Dim rsA As New ADODB.Recordset
   
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      
'cancel by sonia 2021/11/24移到OnSaveData，否則frm03010301_04直接呼叫OnSaveData就不會詢問了
'      If textCF15 = "" Then  '2015/1/13 ADD BY SONIA 有下一程序則不詢問
'         'add by nickc 2005/04/22
'          Pub_EndModCashMsg m_TM10
'      End If
'end 2021/11/24
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 寫檔
      'edit by nick 2004/11/03
      'OnSaveData
      'add by nickc 2005/05/31
      IsAppNpData = False
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
      'Modified by Morgan 2021/12/8 從 OnSaveData 內移出來
      ' 列印定稿
      If textPrint <> "N" Then
         PrintLetter
      End If
      'end 2021/12/8
      
      'Add By Sindy 2023/3/9 T091286 移轉(501)或變更(301)申請人自請撤回須還原申請人
      '於自請撤回核准輸入時彈提醒修改申請人資料
      If m_CP10 = "306" Then
         strSql = "Select CP09,CP10 From CaseProgress Where CP09=(Select CP43 From CaseProgress Where CP09='" & m_CP09 & "')"
         rsA.CursorLocation = adUseClient
         rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If rsA.Fields("CP10") = "501" Or rsA.Fields("CP10") = "301" Then
               strExc(10) = rsA.Fields("CP09")
               If rsA.Fields("CP10") = "301" Then
                  '有變更申請人
                  strSql = "Select CE01 From ChangeEvent Where CE01='" & strExc(10) & "'" & _
                           " AND CE04||CE05||CE06||CE07||CE08 is not null"
                  rsA.CursorLocation = adUseClient
                  rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount = 0 Then
                     strExc(10) = ""
                  End If
               End If
               If strExc(10) <> "" Then
                  MsgBox "請注意，修改申請人資料！", vbCritical, "自請撤回核准"
               End If
            End If
         End If
      End If
      '2023/3/9 END
      
      'add by nickc 2005/05/31
      If IsAppNpData Then
         'add by nickc 2005/09/27
         If MsgBox("準備列印回覆單!!!", vbExclamation + vbOKCancel) = vbOK Then
            Call oClsPrtForm001.PrintReturnSheet(SeekNewCp09, textCF15, DBDATE(textCP07), False)
         End If
      End If
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      'Add By Sindy 2023/4/27
      If Me.m_strIR01 <> "" Then
         Unload frm03010301_02
         Unload frm03010301_01
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
      Else
      '2023/4/27 END
         Unload Me
         Unload frm03010301_02
         frm03010301_01.Show
      End If
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM27.BackColor = &H8000000F
   textTM22S.BackColor = &H8000000F
   
   textCP05.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP45.BackColor = &H8000000F
   textCF15_2.BackColor = &H8000000F
  
   MoveFormToCenter Me
   'Add By nickc 2005/08/04
   'm_blnClkChgButton = False
   
   'Add By Sindy 2023/4/27
   m_strIR01 = frm03010301_02.m_strIR01
   m_strIR02 = frm03010301_02.m_strIR02
   m_strIR03 = frm03010301_02.m_strIR03
   m_strIR04 = frm03010301_02.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2023/4/27 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
             'add by nickc 2005/08/04
            strSql = "SELECT * FROM ChangeEvent " & _
                     "WHERE CE01 = '" & m_CP09 & "' "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <= 0 Then
               m_blnClkChgButton = True
            Else
               m_blnClkChgButton = False
            End If
            rsTmp.Close
   End Select
End Sub

' 讀取商標基本檔
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim strSub As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSub As ADODB.Recordset
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
         ' 延產年度
         m_NA14 = GetNationExtentYear(rsTmp.Fields("TM10"))
         'Add By Sindy 2013/3/19
         If m_TM10 = "013" Then '香港
            Text1 = "3個月"
         ElseIf m_TM10 = "014" Then '新加坡
            Text1 = "2個月"
         End If
         '2013/3/19 End
      End If
      ' 審定號
      'If IsNull(rsTmp.Fields("TM15")) = False Then
      '   textTM15 = rsTmp.Fields("TM15")
      'End If
      ' 商標名稱(中)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 商標名稱(英)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 商標名稱(日)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         m_TM08 = rsTmp.Fields("TM08")
         If m_TM10 < "010" Then
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
         Else
            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
         End If
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
      End If
      'Add By Sindy 2013/1/11
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = rsTmp.Fields("TM78")
      End If
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = rsTmp.Fields("TM79")
      End If
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = rsTmp.Fields("TM80")
      End If
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = rsTmp.Fields("TM81")
      End If
      '2013/1/11 End
      ' 正商標號數
      If IsNull(rsTmp.Fields("TM27")) = False Then
         textTM27 = rsTmp.Fields("TM27")
      End If
      'Add By Cheng 2002/07/22
      '顯示目前准駁
      Me.textTM16S.Text = "" & rsTmp.Fields("TM16").Value
         
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("TM29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
      
      ' 正商標專用期止日
      Set rsSub = New ADODB.Recordset
      strSub = "SELECT * FROM TradeMark " & _
               "WHERE TM15 = '" & textTM27 & "' AND " & _
                     "TM10 = '" & m_TM10 & "' "
      rsSub.CursorLocation = adUseClient
      rsSub.Open strSub, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsSub.RecordCount > 0 Then
         rsSub.MoveFirst
         If IsNull(rsSub.Fields("TM22")) = False Then
            textTM22S = rsSub.Fields("TM22")
         End If
      End If
      rsSub.Close
      Set rsSub = Nothing
      ' 審定號
      'If IsNull(rsTmp.Fields("TM15")) = False Then
      '   textTM15 = rsTmp.Fields("TM15")
      'End If
      ' 公告日
      If IsNull(rsTmp.Fields("TM14")) = False Then
         textTM14 = DBDATE(rsTmp.Fields("TM14"))
      End If
      ' 專用權是否存在
      If IsNull(rsTmp.Fields("TM17")) = False Then
         textTM17 = rsTmp.Fields("TM17")
      End If
      ' 專用期限 (起)
      m_TM21 = "" & rsTmp.Fields("TM21")
      ' 專用期限 (止)
      m_TM22 = "" & rsTmp.Fields("TM22")
   End If
   rsTmp.Close

   Set rsTmp = Nothing
End Sub

' 讀取案件進度檔
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset

   ' 來函收文日
   textCP05S = m_CP05

   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = DBDATE(rsTmp.Fields("CP05"))
      End If
      ' 機關文號
      'If IsNull(rsTmp.Fields("CP08")) = False Then
      '   textCP08 = rsTmp.Fields("CP08")
      'End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         textCP09 = rsTmp.Fields("CP09")
      End If
      
      'Add By Sindy 2020/1/21
      m_CP07 = ""
      If IsNull(rsTmp.Fields("CP07")) = False Then
         m_CP07 = rsTmp.Fields("CP07")
      End If
      '2020/1/21 END

      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      'Add By Cheng 2002/06/14
      '若案件性質為"授權"
      If m_CP10 = "502" Then
         Me.Label4(0).Visible = True
         Me.Label4(1).Visible = True
         Me.Label4(0).Caption = "授權期間："
         Me.textCP53.Visible = True
         Me.textCP54.Visible = True
         Me.textCP53.MaxLength = 8
         Me.textCP54.MaxLength = 8
         Me.textCP53.Text = "" & DBDATE("" & rsTmp.Fields("CP53"))
         Me.textCP54.Text = "" & DBDATE("" & rsTmp.Fields("CP54"))
      '若案件性質為"設定質權"
      ElseIf m_CP10 = "506" Then
         Me.Label4(0).Visible = True
         Me.Label4(1).Visible = True
         Me.Label4(0).Caption = "質權設定期間："
         Me.textCP53.Visible = True
         Me.textCP54.Visible = True
         Me.textCP53.MaxLength = 8
         Me.textCP54.MaxLength = 8
         Me.textCP53.Text = "" & DBDATE("" & rsTmp.Fields("CP53"))
         Me.textCP54.Text = "" & DBDATE("" & rsTmp.Fields("CP54"))
      'add by sonia 2020/10/27 '若案件性質為"移轉"且為美國案才顯示美國移轉登記號
      ElseIf m_TM10 = "101" And m_CP10 = "501" Then
         Label28.Visible = True
         textCP30.Visible = True
         textCP30.Locked = False
         textCP30.Text = "第號/第格"
      'end 2020/10/27
      End If
      
      ' 業務區
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True) 'Modified by Lydia 2016/03/25 離職人員也顯示
      End If
      ' 承辦人
      'Added by Lydia 2016/03/11 CFT改成模組判斷
      'Modified by Lydia 2016/03/25 全部套用
      'If m_TM01 = "CFT" Then
         Dim strNA69 As String
         'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
         Call GetNA69("", m_TM10, "" & rsTmp.Fields("CP13"), strNA69, m_TM01, m_TM02, m_TM03, m_TM04)
         textCP14 = strNA69
         'add by sonia 2024/2/1 申請台灣案的英文證明之核准由程序操作故承辦人改為操作人員
         If Pub_StrUserSt03 = "F12" Or Pub_StrUserSt93 = "T32" Then
            textCP14 = strUserNum
         End If
         'end 2024/2/1
         textCP14_2 = GetStaffName(textCP14)
'      Else
'      'end 2016/03/11
'        If IsNull(rsTmp.Fields("CP14")) = False Then
'           textCP14 = rsTmp.Fields("CP14")
'           textCP14_2 = GetStaffName(textCP14)
'        End If
'      End If
      'end 2016/03/25
       
      'CANCEL BY SONIA 2015/11/20 阿蓮說不要帶前次審查報告通知日CFT-015723
      '' 核准通知日
      'If IsNull(rsTmp.Fields("CP25")) = False Then
      '   textCP25 = rsTmp.Fields("CP25")
      'End If
      'END 2015/11/20
      ' 彼所案號
      If IsNull(rsTmp.Fields("CP45")) = False Then
         textCP45 = rsTmp.Fields("CP45")
      End If
      ' 授權期間(迄)
      If IsNull(rsTmp.Fields("CP54")) = False Then
         m_CP54 = DBDATE(rsTmp.Fields("CP54"))
      End If
      ' 移轉申請人代號
      m_CP56 = Empty
      If IsNull(rsTmp.Fields("CP56")) = False Then
         m_CP56 = rsTmp.Fields("CP56")
      End If
      'Add By Sindy 2013/1/11
      m_CP89 = Empty
      If IsNull(rsTmp.Fields("CP89")) = False Then
         m_CP89 = rsTmp.Fields("CP89")
      End If
      m_CP90 = Empty
      If IsNull(rsTmp.Fields("CP90")) = False Then
         m_CP90 = rsTmp.Fields("CP90")
      End If
      m_CP91 = Empty
      If IsNull(rsTmp.Fields("CP91")) = False Then
         m_CP91 = rsTmp.Fields("CP91")
      End If
      m_CP92 = Empty
      If IsNull(rsTmp.Fields("CP92")) = False Then
         m_CP92 = rsTmp.Fields("CP92")
      End If
      '2013/1/11 End
      ' 若此收文號之實際結果為1時, 則將准駁日置於核准通知日欄位
      If IsNull(rsTmp.Fields("CP24")) = False Then
         If rsTmp.Fields("CP24") = "1" Then
            If IsNull(rsTmp.Fields("CP25")) = False Then
               If IsEmptyText(rsTmp.Fields("CP25")) = False And rsTmp.Fields("CP25") <> "0" Then
                  textCP25 = DBDATE(rsTmp.Fields("CP25"))
               End If
            End If
         End If
      End If
      m_CP24 = "" & rsTmp.Fields("CP24") 'Added by Lydia 2025/09/12
      
      ' 若案件性質為延展時, 則將授權期間放入專用期限欄位
      'If m_CP10 = "102" Then
      '   If IsNull(rsTmp.Fields("CP53")) = False Then
      '      textTM21 = DBDATE(rsTmp.Fields("CP53"))
      '   End If
      '   If IsNull(rsTmp.Fields("CP54")) = False Then
      '      textTM22 = DBDATE(rsTmp.Fields("CP54"))
      '   End If
      'End If
   End If
   rsTmp.Close

   Set rsTmp = Nothing
End Sub

Public Sub QueryData()
   Dim strDay As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   
   ' 讀取商標基本檔
   QueryTradeMark
   ' 讀取案件進度檔
   QueryCaseProgress
   
   ' 以下一程序代號計算承辦期限
''''ediy by nickc 2007/10/11 改抓有時效性的
''''   strDay = Empty
   Select Case frm03010301_02.GetSelectResult
      Case "1":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1001")
            textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1001", DBDATE(m_CP05), DBDATE(textCP06), textCP09)
      Case "2":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1707")
            textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1707", DBDATE(m_CP05), DBDATE(textCP06), textCP09)
   End Select
''''   If IsEmptyText(strDay) = False Then
''''      ' 90.07.03 modify by louis (承辦期限以實際的工作天數來計算)
''''      'textCP48 = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''      textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''      If IsEmptyText(textCP06) = False Then
''''         If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
''''            textCP48 = DBDATE(textCP06)
''''         End If
''''      End If
''''   End If
   
   ' 案件性質為申請, 申請國家為台灣時, 以審定號數+商標種類代號抓商標公報檔, 帶出卷期
   'If m_CP10 = "101" And m_TM10 < "010" Then
   '   strSQL = "SELECT * FROM TMBULLETIN " & _
   '            "WHERE TMBM01 = '" & textTM15 & "' AND " & _
   '                  "TMBM02 = '" & m_TM08 & "' "
   '   rsTmp.CursorLocation = adUseClient
   '   rsTmp.Open strSQL, cnnConnection, adOpenDynamic
   '   If rsTmp.RecordCount > 0 Then
   '      rsTmp.MoveFirst
   '      If IsNull(rsTmp.Fields("TMBM07")) = False Then
   '         textTMBM07_1 = Mid(rsTmp.Fields("TMBM07"), 1, 2)
   '         textTMBM07_2 = Mid(rsTmp.Fields("TMBM07"), 3, 3)
   '      End If
   '   End If
   '   rsTmp.Close
   'End If

'2011/9/23 CANCEL BY SONIA 因為現行3個國家都是專用10年延展10年,在每10年內只要提一次使用宣誓,在發註冊證掛期限即可
'   ' 申請國家為菲律賓或柬埔寨時, 若案件性質為使用宣誓時
'   '92.10.22 modify by sonia 改為抓 na39有值者
'   If m_CP10 = "105" Then
'      strSql = "SELECT * FROM Nation WHERE NA01 = '" & m_TM10 & "'and na39 is not null "
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
'      If rsTmp.RecordCount > 0 Then
'         rsTmp.MoveFirst
'         If IsNull(rsTmp.Fields("NA39")) = False And IsEmptyText(m_CP07) = False Then
'            ' 計算下次宣誓的期限
'             'Modify By Cheng 2003/09/02
''               strTemp = ChangeWDateStringToWString(DateSerial(Mid(strTemp, 1, 4) + Val(rsTmp.Fields("NA39")), Mid(m_CP07, 5, 2), Mid(m_CP07, 7, 2)))
'            strTemp = ChangeWDateStringToWString(DateAdd("yyyy", Val(rsTmp.Fields("NA39")), ChangeWStringToWDateString(DBDATE(strTemp))))
'            If IsEmptyText(m_TM22) = False And Val(strTemp) < Val(m_TM22) Then
'               ' 下一程序為使用宣誓
'               textCF15 = "105"
'               ' 法定期限
'             'Modify By Cheng 2003/09/02
''                  textCP07 = ChangeWDateStringToWString(DateSerial(Mid(m_CP07, 1, 4) + Val(rsTmp.Fields("NA39")), Mid(m_CP07, 5, 2), Mid(m_CP07, 7, 2)))
'               textCP07 = ChangeWDateStringToWString(DateAdd("yyyy", Val(rsTmp.Fields("NA39")), ChangeWStringToWDateString(DBDATE(m_CP07))))
'               ' 本所期限
'             'Modify By Cheng 2003/09/02
''                  textCP06 = ChangeWDateStringToWString(DateSerial(Mid(strTemp, 1, 4), Mid(strTemp, 5, 2), Mid(strTemp, 7, 2) - 2))
'               '92.10.22 modify by sonia
'               'textCP06 = ChangeWDateStringToWString(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strTemp))))
'               '若申請國家為"菲筆賓"時, 本所期限 = 法定期限 - 半年, 其他國家則 本所期限 = 法定期限 - 1年
'               If m_TM10 = "030" Then
'                  textCP06 = DBDATE(DateAdd("m", -6, ChangeWStringToWDateString(DBDATE(textCP07))))
'               Else
'                  textCP06 = DBDATE(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(textCP07))))
'               End If
'               '92.10.22 end
'               End If
'         End If
'      End If
'      rsTmp.Close
'   End If
   
   ' 案件性質為延展時, 才可輸入專用期限
   'If m_CP10 = "102" Then
   '   EnableTextBox textTM21, True
   '   EnableTextBox textTM22, True
   'Else
   '   EnableTextBox textTM21, False
   '   EnableTextBox textTM22, False
   'End If
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
'   'Add By Cheng 2002/07/12
'   '若案件性質為"申請"(101)時
'   If m_CP10 = "101" Then
'      Me.textTM16S.Text = "Y"
'   '其他案件性質
'   Else
'      Me.textTM16S.Text = "N"
'   End If
   'Add By Cheng 2002/07/22
   '若案件性質為"申請"(101)時, 目前准駁預設為"1"(准)
   If m_CP10 = "101" Or m_CP10 = "107" Then
      Me.textTM16S.Text = "1"
   End If
   
   'Add By Sindy 2013/1/11
   '若該筆移轉或讓與的受讓人(5個),與基本檔不符時,顯示訊息且不可輸入核准函
   cmdOK.Enabled = True
   If m_CP10 = "501" Then
      If m_TM23 <> m_CP56 Or m_TM78 <> m_CP89 Or m_TM79 <> m_CP90 Or m_TM80 <> m_CP91 Or m_TM81 <> m_CP92 Then
         MsgBox "此案基本檔申請人與此程序受讓人不同，請確認資料！"
         cmdOK.Enabled = False
      End If
   End If
   '2013/1/11 End
   
   'Added by Lydia 2023/09/12 更正註冊證1701核准：更正的核准，若為註冊證的更正，產生寄正確註冊證定稿如附。可參考FCT之更正核准的做法。
   'Memo by Lydia 2024/04/17 更正延展註冊證1713核准：參考”更正註冊證核准”
   m_strCP09 = "": m_strCP10 = ""
   Combo1.Clear
   bolJumpChgEvent = False 'Added by Lydia 2024/03/21
   If m_CP10 = "302" Then '更正
     strSql = "select cp09,cp10,cp43 from caseprogress where cp09='" & m_CP09 & "'"
     intI = 1
     Set rsTmp = ClsLawReadRstMsg(intI, strSql)
     If intI = 1 Then
        strTemp = "" & rsTmp.Fields("cp43")
        Do While strTemp <> ""
           strSql = "select cp09,cp10,cp43,cpm03 from caseprogress,casepropertymap where cp09='" & strTemp & "' and cp01=cpm01(+) and cp10=cpm02(+) "
           intI = 1
           Set rsTmp = ClsLawReadRstMsg(intI, strSql)
           '非C類的相關總收文號
           If Left(strTemp, 1) < "C" Then
              Exit Do '無資料,離開迴圈,程式結束
           Else
              If intI = 1 Then
                 strTemp = "" & rsTmp.Fields("cp43")
                 'Modified by Lydia 2024/04/17 更正延展註冊證核准：參考”更正註冊證核准”
                 'If Left("" & rsTmp.Fields("cp09"), 1) = "C" And "" & rsTmp.Fields("cp10") = "1701" Then '目前只針對更正註冊證的核准
                 If Left("" & rsTmp.Fields("cp09"), 1) = "C" And InStr("1701,1713", "" & rsTmp.Fields("cp10")) > 0 Then
                    m_strCP09 = "" & rsTmp.Fields("cp09")
                    m_strCP10 = "" & rsTmp.Fields("cp10")
                    Combo1.Clear
                    Combo1.AddItem rsTmp.Fields("cp10") & " " & rsTmp.Fields("cpm03")
                    Combo1.ListIndex = 0
                    bolJumpChgEvent = True 'Added by Lydia 2024/03/21 更正的核准，若為註冊證的更正，不用輸入變更事項；比照「註冊證的更正」發文
                    Exit Do
                 End If
              Else
                 strTemp = ""
              End If
           End If
        Loop
     End If
   End If
   'end 2023/09/12
   
   Set rsTmp = Nothing
End Sub

Private Sub DisplayNextForm()
   frm03010301_04.SetData 0, m_TM01, True
   frm03010301_04.SetData 1, m_TM02, False
   frm03010301_04.SetData 2, m_TM03, False
   frm03010301_04.SetData 3, m_TM04, False
   frm03010301_04.SetData 4, m_CP09, False
   'Add By Sindy 2023/8/7
   If Not m_PrevForm Is Nothing Then
      Call frm03010301_04.SetParent(m_PrevForm)
   End If
   frm03010301_04.m_strIR01 = m_strIR01
   frm03010301_04.m_strIR02 = m_strIR02
   frm03010301_04.m_strIR03 = m_strIR03
   frm03010301_04.m_strIR04 = m_strIR04
   '2023/8/7 END
   Me.Hide
   frm03010301_04.Show
   frm03010301_04.QueryData
End Sub

'edit b nick 2004/11/03
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strSql As String
   Dim strCP06 As String
   Dim strCP07 As String
   Dim strCP09 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strCP16 As String
   Dim strCP18 As String
   Dim strCP27 As String
   Dim strCP30 As String      'add by sonia 2020/10/27
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP09 As String
   Dim strNP14 As String
   Dim strNP15 As String
   Dim strNP22 As String
   Dim strNP10 As String 'Add By Sindy 2014/9/11
   
   '2021/11/24 ADD BY SONIA 從cmdOK_Click移下，否則frm03010301_04直接呼叫OnSaveData就不會詢問了
   If textCF15 = "" Then   '2015/1/13 ADD BY SONIA有下一程序則不詢問
      Pub_EndModCashMsg m_TM10
   End If
   'end 2021/11/24

'911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   'Modified by Lydia 2016/03/11 +案號
   'Call GetNP69("", m_TM10, m_CP13, strNP10) 'Add By Sindy 2014/9/11
   'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
   Call GetNA69("", m_TM10, m_CP13, strNP10, m_TM01, m_TM02, m_TM03, m_TM04)
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新原案件進度資料的實際結果為准及准駁日
   'Modify By Cheng 2002/07/12
   '不論前一畫面的結果欄為何, 皆要更新
   'modify by sonia 2024/3/21 阿蓮又要求還原，僅核准時才更新
   If frm03010301_02.GetSelectResult = "1" Then
      strSql = "UPDATE CaseProgress SET CP24 = '1', CP25 = " & DBDATE(textCP25) & " " & _
               "WHERE CP09 = '" & m_CP09 & "' AND " & _
                     "(CP24 IS NULL OR CP24 = '' OR CP24 = ' ')"
      cnnConnection.Execute strSql
   End If
   'Add By Cheng 2002/06/14
   If m_CP10 = "502" Or m_CP10 = "506" Then
      strSql = "UPDATE CaseProgress SET CP53 = " & DBDATE(Me.textCP53.Text) & ", CP54 = " & DBDATE(Me.textCP54.Text) & " " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Modify By Cheng 2002/07/22
   '取消更新專用權是否存在
'   ' 更新商標基本檔之專用權是否存在
'   strSQL = "UPDATE TradeMark SET TM17 = '" & textTM17 & "' " & _
'            "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                  "TM02 = '" & m_TM02 & "' AND " & _
'                  "TM03 = '" & m_TM03 & "' AND " & _
'                  "TM04 = '" & m_TM04 & "' "
'   cnnConnection.Execute strSQL
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若案件性質為延展時, 更新商標基本檔之專用期限欄位
   'If m_CP10 = "102" Then
   '   strSQL = "UPDATE TradeMark SET TM21 = " & DBDATE(textTM21) & ", " & _
   '                                 "TM22 = " & DBDATE(textTM22) & " " & _
   '         "WHERE TM01 = '" & m_TM01 & "' AND " & _
   '               "TM02 = '" & m_TM02 & "' AND " & _
   '               "TM03 = '" & m_TM03 & "' AND " & _
   '               "TM04 = '" & m_TM04 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 案件性質為申請時
   '2005/3/23 modify by sonia
   'If m_CP10 = "101" Then
   '2006/1/3 MODIFY BY SONIA 再加跨類
   If (m_CP10 = "101" Or m_CP10 = "308" Or m_CP10 = "107") Then
      ' 更新審定號, 公告日
      'strSQL = "UPDATE TradeMark SET TM15 = '" & textTM15 & "'," & _
      '                              "TM14 = " & DBDATE(textTM14) & " " & _
      '         "WHERE TM01 = '" & m_TM01 & "' AND " & _
      '               "TM02 = '" & m_TM02 & "' AND " & _
      '               "TM03 = '" & m_TM03 & "' AND " & _
      '               "TM04 = '" & m_TM04 & "' "
      '2005/3/23 modify by sonia
      'StrSql = "UPDATE TradeMark SET TM14 = " & DBNullDate(textTM14) & " " & _
      '         "WHERE TM01 = '" & m_TM01 & "' AND " & _
      '               "TM02 = '" & m_TM02 & "' AND " & _
      '               "TM03 = '" & m_TM03 & "' AND " & _
      '               "TM04 = '" & m_TM04 & "' "
      'cnnConnection.Execute StrSql
      ''Modify By Cheng 2002/07/22
      ''當案件性質為商申時(101), 更新目前准/駁及審定來函日兩個欄位
'     ' ' 當使用者輸入要更新基本檔之准/駁時, 更新目前准/駁及審定來函日兩個欄位
'     ' If textTM16S = "Y" Then
      'If m_CP10 = "101" Then
      '   StrSql = "UPDATE TradeMark SET TM16='1'," & _
      '                                 "TM13=" & DBDATE(textCP25) & " " & _
      '            "WHERE TM01 = '" & m_TM01 & "' AND " & _
      '                  "TM02 = '" & m_TM02 & "' AND " & _
      '                  "TM03 = '" & m_TM03 & "' AND " & _
      '                  "TM04 = '" & m_TM04 & "' "
      '   cnnConnection.Execute StrSql
      'End If
         strSql = "UPDATE TradeMark SET TM16='1'," & _
                                       "TM13=" & DBDATE(textCP25) & ", " & _
                                       "TM14 = " & DBNullDate(textTM14) & " " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "' "
      cnnConnection.Execute strSql
      '2005/3/23 end
   End If
         
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'add by nick 2004/09/14
   ' 若案件性質為307時, 更新商標基本檔之是否閉卷=Y
   If m_CP10 = "307" Then
      strSql = "UPDATE TradeMark SET TM29='Y' " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
      cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '  新增資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   NewCP09 = strCP09 'Added by Lydia 2023/09/12
   ' 案件性質為核准
   strCP10 = "1001"
   Select Case frm03010301_02.GetSelectResult
      Case "1": strCP10 = "1001"
      Case "2": strCP10 = "1707"
   End Select
   ' 業務區別 91.8.26 MODIFY BY SONIA
   'strCP12 = GetStaffDepartment(m_CP13)
   ' 費用
   strCP16 = str(Val(textFee_1) + Val(textFee_2))
   ' 點數
   strCP18 = str(Val(textFee_1) + Val(textFee_2) / 1000)
   '92.6.14 MODIFY BY SONIA
   ' 發文日
   'strCP27 = "NULL"
   'If IsEmptyText(textCP14) = True Then
   '   strCP27 = DBDATE(SystemDate())
   'End If
   '2008/12/11 MODIFY BY SONIA 有期限則不上發文日
   'strCP27 = DBDATE(SystemDate())
   If IsEmptyText(textCP06) = False Then
      strCP27 = ""
   Else
      strCP27 = DBDATE(SystemDate())
   End If
   '2008/12/11 END
   '92.6.14 END
   
   'add by sonia 2020/10/27 存美國移轉登記號於大陸申請案號欄cp30
   strCP30 = ""
   If m_TM10 = "101" And m_CP10 = "501" And textCP30 <> "" Then
      strCP30 = textCP30
   End If
   '2020/10/27
   
       'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng 2004/02/04
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP16,CP18,CP20,CP26,CP27,CP32,CP43) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
'                    strCP16 & "," & strCP18 & ",'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "') "
   '2008/12/11 MODIFY BY SONIA 合併CP14,CP48,CP06,CP07
   'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP16,CP18,CP20,CP26,CP27,CP32,CP43) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                    strCP16 & "," & strCP18 & ",'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "') "
   'Modify By Sindy 2012/11/28 CP16,CP17,CP18不存
   'modify by sonia 2020/10/27 +cp30
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP16,CP18,CP20,CP26,CP27,CP32,CP43,CP14,CP48,CP06,CP07,CP30) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                    "null,null,'" & "N" & "','" & textCP26 & "'," & CNULL(DBDATE(strCP27)) & ",'" & "N" & "','" & m_CP09 & "'," & CNULL(textCP14) & "," & CNULL(DBDATE(textCP48)) & "," & CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & "," & CNULL(textCP30) & ") "
                    'strCP16 & "," & strCP18 & ",'" & "N" & "','" & textCP26 & "'," & CNULL(DBDATE(strCP27)) & ",'" & "N" & "','" & m_CP09 & "'," & CNULL(textCP14) & "," & CNULL(DBDATE(textCP48)) & "," & CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & ") "
   '2008/12/11 END
    'End
   cnnConnection.Execute strSql
   
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   '2008/12/11 CANCEL BY SONIA 移至 INSERT INTO CASEPROGRESS
   '' 若有輸入承辦人時
   'If IsEmptyText(textCP14) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP14 = '" & textCP14 & "' " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 若有輸入承辦期限時
   'If IsEmptyText(textCP48) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP48 = " & DBDATE(textCP48) & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 更新新增的案件進度檔其本所期限及法定期限
   'strCP06 = Empty
   'strCP07 = Empty
   'If IsEmptyText(textCF15) = False Then
   '   If IsEmptyText(textCP06) = False Then: strCP06 = DBDATE(textCP06)
   '   If IsEmptyText(textCP07) = False Then: strCP07 = DBDATE(textCP07)
   'End If
   '' 本所期限
   'If IsEmptyText(strCP06) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP06 = " & strCP06 & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '' 法定期限
   'If IsEmptyText(strCP07) = False Then
   '   strSQL = "UPDATE CaseProgress SET CP07 = " & strCP07 & " " & _
   '            "WHERE CP09 = '" & strCP09 & "' "
   '   cnnConnection.Execute strSQL
   'End If
   '2008/12/11 END
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有輸入下一程序時, 新增資料到下一程序檔
   strNP22 = GetNextProgressNo()
   If IsEmptyText(textCF15) = False Then
      strNP14 = Empty
      strNP14 = GetRelatedPerson(m_CP09)
      'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
      '          "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
      '                    strCP06 & "," & strCP07 & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "'," & strNP22 & ")"
      'Modify By Cheng 2002/09/25
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP14,NP22) " & _
'                "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
'                          strCP06 & "," & strCP07 & ",'" & m_CP13 & "','" & strNP14 & "'," & strNP22 & ")"
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP14,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                          CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & ChgSQL(strNP14) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      Select Case textCF15
         Case "102", "105", "702", "708", "305", "998", "997":
         Case Else:
            'add by nickc 2005/05/31
            IsAppNpData = True
            SeekNewCp09 = strCP09
            'Modify By Cheng 2003/07/09
            '改成整批列印
'            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新下一程序檔案件性質為催審的資料
   'Modify By Sindy 2009/06/10 同時更新下一程序檔案件性質為997.收達998.提申的資料
   'modify by sonia 2018/11/23 通知公告日時只更新997,998,不更新催審
   'strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP07 in (305,997,998) "
   If frm03010301_02.GetSelectResult = "2" Then   '通知公告日只更新997,998
      strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
               "WHERE NP01 = '" & m_CP09 & "' AND " & _
                     "NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 in (997,998) "
   Else
      strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
               "WHERE NP01 = '" & m_CP09 & "' AND " & _
                     "NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 in (305,997,998) "
   End If
   'end 2018/11/23
   cnnConnection.Execute strSql
   
   'add by sonia 2019/11/11 若有加速審查311的收達,提申,催審,一併上Y
   strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP07 IN (305,997,998) AND (NP01,NP02,NP03,NP04,NP05) IN" & _
                  "(SELECT CP09,CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE " & _
                  "CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' AND " & _
                  "CP10='311')"
   cnnConnection.Execute strSql
   'end 2019/11/11
   'Added by Lydia 2025/09/12 CFT審查報告、通知公告日更新商申催審期限：非爭議案核准輸入(通知公告日)，除已有國家檔設定NA64=>CFT公告日更新商申催審月數，以公告日加3個月更新至「商申」的催審期限
   If frm03010301_02.GetSelectResult = "2" And m_CP10 = "101" And (IsNull(m_CP24) Or Trim(m_CP24) = "") Then
      intI = Val(Pub_GetField("NATION", " NA01='" & m_TM10 & "' ", "NA64"))
      If intI = 0 Then intI = 3
      strNP09 = DBDATE(DateAdd("m", intI, ChangeWStringToWDateString(DBDATE(IIf(textTM14.Text = "", strSrvDate(1), DBDATE(textTM14.Text))))))
      strSql = "UPDATE NextProgress SET NP08 = " & PUB_GetWorkDay1(strNP09, True) & ", NP09 = " & strNP09 & " " & _
               "WHERE NP01 = '" & m_CP09 & "' AND NP02 = '" & m_TM01 & "' AND NP03 = '" & m_TM02 & "' AND NP04 = '" & m_TM03 & "' AND NP05 = '" & m_TM04 & "' AND NP07 ='305' AND NP06 IS NULL "
      cnnConnection.Execute strSql
   End If
   'end 2025/09/12
   
   'Add By Sindy 2013/6/3 CFT美國案申請程序之通知公告日時,同時新增下一程序1711
   If m_TM01 = "CFT" And m_TM10 = "101" And m_CP10 = "101" And frm03010301_02.GetSelectResult = "2" Then
      '本所期限=法定期限=畫面上輸入之公告日+3個月
      strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(textTM14)), Val(DBMONTH(textTM14)) + 3, Val(DBDAY(textTM14)))))
      '智權人員 = 畫面上輸入之承辦人
      'Modify By Sindy 2014/9/11 textCP14=>strNP10
      'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','1711'," & _
               CNULL(strNP08) & "," & CNULL(strNP08) & ",'" & strNP10 & "'," & GetNextProgressNo() & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','1711'," & _
               CNULL(PUB_GetWorkDay1(strNP08, True)) & "," & CNULL(strNP08) & ",'" & strNP10 & "'," & GetNextProgressNo() & ")"
      cnnConnection.Execute strSql
   End If
   '2013/6/3 End
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 依案件性質來決定是否要新增一筆資料到下一程序檔
   Select Case m_CP10
      'Case "102":
      '   strNP09 = DBDATE(textTM22)
      '   strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP09)), Val(DBMONTH(strNP09)), Val(DBDAY(strNP09)) - 2)))
      '   strNP22 = GetNextProgressNo()
      '   strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
      '            "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & "102" & "," & _
      '                    strNP08 & "," & strNP09 & ",'" & m_CP13 & "'," & strNP22 & ")"
      '   cnnConnection.Execute strSQL
      
      '92.4.23 cancel by sonia
      '' 當案件性質為授權,再授權,設定質權時, 新增一筆資料到下一程序檔
      'Case "502", "504", "506":
      '   Select Case m_CP10
      '      Case "502": strNP07 = "503"
      '      Case "504": strNP07 = "505"
      '      Case "506": strNP07 = "507"
      '   End Select
      '   strNP08 = 0
      '   strNP09 = 0
      '   If IsEmptyText(m_CP54) = False Then
      '      '下一程序法定期限
      '      strNP09 = m_CP54
      '      'Modify By Cheng 2002/06/07
      '      '下一程序本所期限
      '      '本所期限 = 法定期限 - 六個月
'     '       strNP08 = m_CP54
      '      strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP09)), Val(DBMONTH(strNP09)) - 6, Val(DBDAY(strNP09)))))
      '   End If
      '   strNP14 = Empty
      '   strNP14 = GetRelatedPerson(m_CP09)
      '   strNP15 = Empty
      '   strNP15 = GetCaseTypeName(m_TM01, m_CP10, 0) & "終止日"
      '   strNP22 = GetNextProgressNo()
      '92.4.23 end
      
         'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
         '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
         '                 strNP08 & "," & strNP09 & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "','" & strNP15 & "'," & strNP22 & ")"
         'Modify By Cheng 2002/09/25
'         strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP14,NP15,NP22) " & _
'                  "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                          strNP08 & "," & strNP09 & ",'" & m_CP13 & "','" & strNP14 & "','" & strNP15 & "'," & strNP22 & ")"
            'Modify By Cheng 2003/04/03
            '智權人員存最近收文A類接洽記錄單的智權人員
         
      '92.4.23 cancel by sonia
      '   strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP14,NP15,NP22) " & _
      '            "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
      '                    strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & ChgSQL(strNP14) & "','" & strNP15 & "'," & strNP22 & ")"
      '   cnnConnection.Execute strSQL
      '92.4.23 end
      
      Case "501":
        'Modify By Cheng 2003/03/07
        '不在此更新移轉人及移轉申請人資料
'         ' 以商標基本檔的申請人代號更新案件進度檔的移轉人代號
'         strSQL = "UPDATE CaseProgress SET CP55 = '" & m_TM23 & "' " & _
'                  "WHERE CP09 = '" & strCP09 & "' "
'         cnnConnection.Execute strSQL
'         ' 以案件進度檔的移轉申請人代號更新商標基本檔的申請人代號, 更新卷宗性質為1
'         strSQL = "UPDATE TradeMark SET TM23 = '" & m_CP56 & "'," & _
'                                       "TM28 = '" & "1" & "' " & _
'                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                        "TM02 = '" & m_TM02 & "' AND " & _
'                        "TM03 = '" & m_TM03 & "' AND " & _
'                        "TM04 = '" & m_TM04 & "' "
'         cnnConnection.Execute strSQL
         '2006/6/1 ADD BY SONIA 保留更新卷宗性質
         strSql = "UPDATE TradeMark SET TM28 = '1' " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
         cnnConnection.Execute strSql
         '2006/6/1 END
   End Select
          'add by nickc 2005/04/22
          Pub_UpdateEndModCash m_TM01, m_TM02, m_TM03, m_TM04
    'Move By Cheng 2002/11/29
 '911106 nick transation
   
   'Add by Sindy 2023/4/27
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm03010301_01", strCP09
   End If
   '2023/4/27 END
   
   cnnConnection.CommitTrans

   'Modify By Cheng 2002/11/29
' '911106 nick transation
'     cnnConnection.CommitTrans
   Exit Function
   
CheckingErr:
   MsgBox (Err.Description)
   cnnConnection.RollbackTrans
   OnSaveData = False
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Add By Cheng 2002/07/19
   Set frm03010301_03 = Nothing
End Sub

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCF15_2 = Empty
   If IsEmptyText(textCF15) = False Then
      '2014/11/26 ADD BY SONIA
      If Len(Me.textCF15.Text) <> 3 Then
         Cancel = True
         MsgBox "下一程序欄位值必須為三碼!!!", vbExclamation
         textCF15_GotFocus
         Exit Sub
      End If
      '2014/11/26 END
      If m_TM10 < "010" Then
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
      Else
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 1)
      End If
      If IsEmptyText(textCF15_2) = True Then
         strTit = "檢核資料"
         strMsg = "下一程序代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15_GotFocus
         GoTo EXITSUB
      End If
   End If
   
EXITSUB:
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      ' 檢查是否為西元日期
      If CheckIsDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = PUB_GetWorkDay1(textCP06, True)
      'end 2020/07/09
      End If
      'Add By Cheng 2002/03/11
      'Modify By Sindy 2009/09/18
      'If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
      If Val(Me.textCP06.Text) < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Cancel = True
         textCP06_GotFocus
         GoTo EXITSUB
      End If
      
      ' 以下一程序代號計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''      strDay = Empty
      Select Case frm03010301_02.GetSelectResult
         Case "1":
''''            strDay = GetWorkDays(m_TM01, m_TM10, "1001")
            textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1001", DBDATE(m_CP05), DBDATE(textCP06), textCP09)
         Case "2":
''''            strDay = GetWorkDays(m_TM01, m_TM10, "1707")
            textCP48 = Pub_GetHandleDay(m_TM01, m_TM10, "1707", DBDATE(m_CP05), DBDATE(textCP06), textCP09)
      End Select
''''      If IsEmptyText(strDay) = False Then
''''         ' 90.07.03 modify by louis (承辦期限以實際的工作天數來計算)
''''         'textCP48 = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''         textCP48 = DBDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''         If IsEmptyText(textCP06) = False Then
''''            If Val(DBDATE(textCP48)) > Val(DBDATE(textCP06)) Then
''''               textCP48 = DBDATE(textCP06)
''''            End If
''''         End If
''''      End If
   End If
EXITSUB:
End Sub

' 法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      ' 檢查是否為民國年
      If CheckIsDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2010/11/29
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14)
      If IsEmptyText(textCP14_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "承辦人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP14_GotFocus
      End If
   End If
End Sub

' 核准通知日
Private Sub textCP25_Validate(Cancel As Boolean)
   Dim SysDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP25) = False Then
      ' 檢查是否為民國年
      If CheckIsDate(textCP25, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的核准通知日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP25_GotFocus
      End If
      ' 核准通知日不可超過系統日
      'edit by nickc 2006/03/17 要跟伺服器日期比對
      'sysDate = ChangeWDateStringToWString(Date)
      SysDate = strSrvDate(1)
      If Val(textCP25) > Val(SysDate) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "核准通知日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP25_GotFocus
      End If
   End If
End Sub

' 是否算案件數
Private Sub textCP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP26) = False Then
      Select Case textCP26
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
   End If
End Sub

' 承辦人期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為民國日期
      If CheckIsDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "承辦期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
         Exit Sub
      End If
   End If
   'Add By Cheng 2002/05/07
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         Cancel = True
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         textCP48_GotFocus
         Exit Sub
      End If
   End If

End Sub
' 委託代理人費
Private Sub textFee_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textFee_1) = False Then
      If IsNumeric(textFee_1) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "委託代理人費用不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFee_1_GotFocus
      End If
   End If
End Sub

' 領證費
Private Sub textFee_2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textFee_2) = False Then
      If IsNumeric(textFee_2) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "領證費不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFee_2_GotFocus
      End If
   End If
End Sub

' 點數
Private Sub textCP18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP18) = False Then
      If IsNumeric(textCP18) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "點數不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP18_GotFocus
      End If
   End If
End Sub

' 是否列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case "", " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

' 公告日
Private Sub textTM14_Validate(Cancel As Boolean)
   Dim SysDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM14) = False Then
      ' 檢查是否為民國年
      If CheckIsDate(textTM14, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的公告日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14_GotFocus
      End If
      ' 公告日不可超過系統日
      'sysDate = ChangeWDateStringToWString(Date)
      'If Val(textTM14) > Val(sysDate) Then
      '   Cancel = True
      '   strTit = "資料檢核"
      '   strMsg = "公告日不可超過系統日"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textTM14_GotFocus
      'End If
   End If
End Sub

Private Sub textTM16S_KeyPress(KeyAscii As Integer)
   'Modify By Cheng 2002/07/22
'   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否更新基本檔目前准駁
Private Sub textTM16S_Validate(Cancel As Boolean)
   'Modify By Cheng 2002/07/22
   '取消檢查
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'
'   If IsEmptyText(textTM16S) = False Then
'      Select Case textTM16S
'         Case "Y", "N":
'         Case Else:
'            Cancel = True
'            strTit = "資料檢核"
'            strMsg = "只可輸入Y或N"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTM16S_GotFocus
'      End Select
'   End If
End Sub

' 專用權是否存在
Private Sub textTM17_Validate(Cancel As Boolean)
   'Modify By Cheng 2002/07/22
   '取消檢查
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textTM17) = False Then
'      Select Case textTM17
'         Case "Y", "N":
'         Case Else:
'            Cancel = True
'            strTit = "資料檢核"
'            strMsg = "只可輸入Y或N"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTM17_GotFocus
'      End Select
'   End If
End Sub
' 專用期限起日
'Private Sub textTM21_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Dim strCorrDate As String
'   Dim strDate As String
'
'   Cancel = False
'   ' 原專用期限止日
'   If IsEmptyText(m_TM22) = True Then
'      GoTo ExitSub
'   End If
'   ' 未輸入專用期限起日
'   If IsEmptyText(textTM21) = True Then
'      GoTo ExitSub
'   End If
'   ' 案件性質非延展
'   If m_CP10 <> "102" Then
'      GoTo ExitSub
'   End If
'
'   ' 檢核是否為民國日期
'   If CheckIsDate(textTM21, False) = False Then
'      Cancel = True
'      strTit = "資料檢核"
'      strMsg = "請輸入正確的日期"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM21_GotFocus
'   End If
'
'   strCorrDate = ChangeWDateStringToWString(Format(DateSerial(Val(Left(m_TM22, 4)), Val(Mid(m_TM22, 5, 2)), Right(m_TM22, 2) + 1)))
'   strDate = textTM21
'   strDate = ChangeWDateStringToWString(Format(DateSerial(Val(Left(strDate, 4)), Val(Mid(strDate, 5, 2)), Right(strDate, 2) + 1)))
'   If Val(strCorrDate) <> Val(strDate) Then
'      Cancel = True
'      strTit = "資料檢核"
'      strMsg = "專用期限起日必須為原專用期限止日的後一天"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM21_GotFocus
'   End If
'
'ExitSub:
'End Sub
' 專用期限止日
'Private Sub textTM22_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Dim strCorrDate As String
'   Dim strDate As String
'   Cancel = False
'
'   ' 原專用期限止日
'   If IsEmptyText(m_TM22) = True Then
'      GoTo ExitSub
'   End If
'   ' 未輸入專用期限起日
'   If IsEmptyText(textTM22) = True Then
'      GoTo ExitSub
'   End If
'   ' 案件性質非延展
'   If m_CP10 <> "102" Then
'      GoTo ExitSub
'   End If
'
'   ' 檢核是否為民國日期
'   If CheckIsDate(textTM22, False) = False Then
'      Cancel = True
'      strTit = "資料檢核"
'      strMsg = "請輸入正確的日期"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM22_GotFocus
'   End If
'
'   strDate = DBDATE(textTM21)
'
'   Select Case m_TM08
'      Case "1", "4", "7", "8":
'         strCorrDate = ChangeWDateStringToWString(Format(DateSerial(Val(Left(m_TM22, 4)) + m_NA14, Val(Mid(m_TM22, 5, 2)), Right(m_TM22, 2))))
'      Case Else:
'         strCorrDate = textTM22
'   End Select
'   If Val(strDate) <> Val(strCorrDate) Then
'      Cancel = True
'      strTit = "資料檢核"
'      strMsg = "專用期限止日不正確"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM22_GotFocus
'   End If
'ExitSub:
'End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
'add by nickc 2005/08/04
        'Modified by Lydia 2024/03/21 + And bolJumpChgEvent = False
        If m_blnClkChgButton = False And bolJumpChgEvent = False Then
            MsgBox "請輸入變更事項!!!", vbExclamation + vbOKOnly
            Me.cmdMod.SetFocus
            GoTo EXITSUB
        End If
        
        
    'Add By Cheng 2003/06/05
    '檢查核准通知日
    If Me.textCP25.Text = "" Then
        strTit = "資料檢核"
        strMsg = "核准通知日不可為空白"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        textCP25.SetFocus
        textCP25_GotFocus
        GoTo EXITSUB
    End If
   ' 審定號不可為空白
   'If IsEmptyText(textTM15) = True Then
   '   strTit = "資料檢核"
   '   strMsg = "審定號不可為空白"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   GoTo ExitSub
   'End If
   ' 公告日
   If frm03010301_02.GetSelectResult = "2" Then
      If IsEmptyText(textTM14) = True Then
         strTit = "資料檢核"
         strMsg = "公告日不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 專用期限
   'If m_CP10 = "102" Then
   '   If IsEmptyText(textTM21) = True Or IsEmptyText(textTM22) = True Then
   '      strTit = "資料檢核"
   '      strMsg = "案件性質為延展, 專用期限不可為空白"
   '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '      GoTo ExitSub
   '   End If
   '   If Val(textTM21) > Val(textTM22) Then
   '      strTit = "資料檢核"
   '      strMsg = "專用期限的起日不可超過迄日"
   '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '      GoTo ExitSub
   '   End If
   'End If
   'Modify By Cheng 2002/07/22
'   ' 專用權是否存在
'   If IsEmptyText(textTM17) = True Then
'      strTit = "資料檢核"
'      strMsg = "專用權是否存在不可為空白"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM17.SetFocus
'      GoTo EXITSUB
'   End If
   'Modify By Cheng 2002/07/22
'   ' 是否更新基本檔目前准駁
'   If IsEmptyText(textTM16S) = True Then
'      strTit = "資料檢核"
'      strMsg = "是否更新基本檔目前准駁不可為空白"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM16S.SetFocus
'      GoTo EXITSUB
'   End If
   ' 下一程序
   If IsEmptyText(textCF15) = True Then
      If IsEmptyText(textCP06) = False Or IsEmptyText(textCP07) = False Then
         strTit = "資料檢核"
         strMsg = "無下一程序時, 本所期限及法定期限不可輸入"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   Else
      If IsEmptyText(textCP06) = True Or IsEmptyText(textCP07) = True Then
         strTit = "資料檢核"
         strMsg = "有下一程序時, 本所期限及法定期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
      'Add By Cheng 2002/03/11
      If Me.textCP06.Text <> "" Then
         'Modify By Sindy 2009/09/18
         'If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         If Val(Me.textCP06.Text) < ServerDate Then
            MsgBox "本所期限不可小於系統日期!!!", vbExclamation
            Me.textCP06.SetFocus
            textCP06_GotFocus
            GoTo EXITSUB
         Else
            If Val(Me.textCP06.Text) > Val(Me.textCP07.Text) Then
               MsgBox "本所期限不可大於法定期限!!!", vbExclamation
               Me.textCP06.SetFocus
               textCP06_GotFocus
               GoTo EXITSUB
            End If
         End If
      End If
   End If
   
   ' 承辦期限不可超過本所期限
   If IsEmptyText(textCP48) = False And IsEmptyText(textCP06) = False Then
      If Val(textCP48) > Val(textCP06) Then
         strTit = "資料檢核"
         strMsg = "承辦期限的日期不可超過本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'add by sonia 2020/10/27
   If m_TM10 = 美國國家代號 And m_CP10 = "501" And (Me.textCP30.Text = "" Or Me.textCP30.Text = "第號/第格") Then
      MsgBox "美國移轉核准請輸入登記號，此欄將記錄在案件進度檔的'大陸申請案號'欄!!!", vbExclamation + vbOKOnly
      Me.textCP30.SetFocus
      GoTo EXITSUB
   End If
   'end 2020/10/27
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textTMBM07_1_GotFocus()
   InverseTextBox textTMBM07_1
End Sub

Private Sub textTMBM07_2_GotFocus()
   InverseTextBox textTMBM07_2
End Sub

Private Sub textTM14_GotFocus()
   InverseTextBox textTM14
End Sub

'Private Sub textTM15_GotFocus()
'   InverseTextBox textTM15
'End Sub

Private Sub textTM16S_GotFocus()
   'Modify By Cheng 2002/07/22
'   InverseTextBox textTM16S
End Sub

Private Sub textTM17_GotFocus()
   'Modify By Cheng 2002/07/22
'   InverseTextBox textTM17
End Sub

'Private Sub textTM21_GotFocus()
'   InverseTextBox textTM21
'End Sub

'Private Sub textTM22_GotFocus()
'   InverseTextBox textTM22
'End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

'Private Sub textCP08_GotFocus()
'   InverseTextBox textCP08
'End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP25_GotFocus()
   InverseTextBox textCP25
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

Private Sub textPS_GotFocus()
   InverseTextBox textPS
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textFee_1_GotFocus()
   InverseTextBox textFee_1
End Sub

Private Sub textFee_2_GotFocus()
   InverseTextBox textFee_2
End Sub

Private Sub textCP18_GotFocus()
   InverseTextBox textCP18
End Sub

Private Sub textTM17_KeyPress(KeyAscii As Integer)
   'Modify By Cheng 2002/07/22
'   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
Dim strSql As String
'Add By Cheng 2003/01/01
Dim strAttach As String '附件
Dim ii As Integer
   
   'Modified by Lydia 2023/09/12
   'Select Case m_CP10
   Select Case IIf(m_strCP10 <> "", m_strCP10, m_CP10)
      '2006/1/3 ADD BY SONIA 增加跨類
      Case "101", "107": '申請
         Select Case m_TM10
            ' 日本
            Case "011":
               ' 清除定稿例外欄位檔原有資料
               EndLetter "03", m_CP09, "01", strUserNum
                'Add By Cheng 2003/01/01
                ' 公告期間
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                         "','公告期間','" & Me.Text1.Text & "')"
                cnnConnection.Execute strSql
            'Add By Sindy 2009/10/19
            ' 越南
            Case "042":
            '2009/10/19 End
            ' 德國
            Case "231":
            'Add By Cheng 2003/01/01
            ' 巴西
            Case "117":
               '若未輸入公告日, 則處理狀況為"02"
               If IsEmptyText(Me.textTM14.Text) = True Then
                    EndLetter "03", m_CP09, "02", strUserNum
                    'Add By Cheng 2003/01/01
                    ' 公告期間
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                           "','公告期間','" & Me.Text1.Text & "')"
                    cnnConnection.Execute strSql
                    'Add By Cheng 2003/01/01
                    ' 附件
                    strAttach = ""
                    For ii = 0 To Me.Check1.Count - 1
                        If Me.Check1(ii).Value = vbChecked Then strAttach = strAttach & Me.Check1(ii).Caption & "、"
                    Next ii
                    If strAttach <> "" Then
                        strAttach = Left(strAttach, Len(strAttach) - 1)
                        strAttach = "附件：" & strAttach & "。"
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                                 "','附件','" & strAttach & "')"
                        cnnConnection.Execute strSql
                    End If
               '若輸入公告日, 則處理狀況為"03"
               Else
                    EndLetter "03", m_CP09, "03", strUserNum
                    'Add By Cheng 2003/01/01
                    ' 公告期間
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                           "','公告期間','" & Me.Text1.Text & "')"
                    cnnConnection.Execute strSql
               End If
            'Add By Sindy 2012/11/28
            ' 阿曼
            Case "033":
               EndLetter "03", m_CP09, "06", strUserNum
               '本所期限
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
                      "','本所期限','" & Val(Left(DBDATE(textCP06), 4)) - 1911 & "年" & Mid(DBDATE(textCP06), 5, 2) & "月" & Right(DBDATE(textCP06), 2) & "日" & "')"
               cnnConnection.Execute strSql
               '費用
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
                      "','費用','" & Val(textFee_1) + Val(textFee_2) & "')"
               cnnConnection.Execute strSql
               '費用點數
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
                      "','費用點數','" & textCP18 & "')"
               cnnConnection.Execute strSql
            '2012/11/28 End
            'Add By Sindy 2013/2/26
            ' 香港,新加坡
            Case "013", "014":
               EndLetter "03", m_CP09, "07", strUserNum
               '公告期間
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
                      "','公告期間','" & Me.Text1.Text & "')"
               cnnConnection.Execute strSql
               '費用
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
                      "','費用','" & Val(textFee_1) + Val(textFee_2) & "')"
               cnnConnection.Execute strSql
               '費用點數
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
                      "','費用點數','" & textCP18 & "')"
               cnnConnection.Execute strSql
               'Add By Sindy 2013/3/19
               '領證通知期限
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
                      "','領證通知期限','" & DBDATE(DateAdd("m", Replace(Text1, "個月", ""), ChangeWStringToWDateString(DBDATE(textTM14)))) & "')"
               cnnConnection.Execute strSql
               '2013/3/19 End
               'Added by Lydia 2025/09/11 香港案公告期限：公告日起6週內
               If m_TM10 = "013" Then
                  strSql = IIf(Trim(textTM14) = "", DBDATE(textCP25), DBDATE(textTM14))
                  If strSql <> "" Then
                     strSql = CompWorkDay(2, CompDate(2, 42, strSql), 1)
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
                            "','香港案公告期限','" & strSql & "')"
                     cnnConnection.Execute strSql
                  End If
               End If
               'end 2025/09/11
            '2013/2/26 End
            ' 其它
            Case Else:
'               ' 清除定稿例外欄位檔原有資料
'               EndLetter "03", m_CP09, "02", strUserNum
               'Modify By Cheng 2002/07/12
               '若未輸入公告日, 則處理狀況為"02"
               If IsEmptyText(Me.textTM14.Text) = True Then
                    EndLetter "03", m_CP09, "02", strUserNum
                    'Add By Cheng 2003/01/01
                    ' 公告期間
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                           "','公告期間','" & Me.Text1.Text & "')"
                    cnnConnection.Execute strSql
                    'Add By Cheng 2003/01/01
                    ' 附件
                    strAttach = ""
                    For ii = 0 To Me.Check1.Count - 1
                        If Me.Check1(ii).Value = vbChecked Then strAttach = strAttach & Me.Check1(ii).Caption & "、"
                    Next ii
                    If strAttach <> "" Then
                        strAttach = Left(strAttach, Len(strAttach) - 1)
                        strAttach = "附件：" & strAttach & "。"
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                                 "','附件','" & strAttach & "')"
                        cnnConnection.Execute strSql
                    End If
               '若輸入公告日, 則處理狀況為"04"
               Else
                    EndLetter "03", m_CP09, "04", strUserNum
                    'Add By Cheng 2003/01/01
                    ' 公告期間
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & _
                           "','公告期間','" & Me.Text1.Text & "')"
                    cnnConnection.Execute strSql
                    'Add By Cheng 2003/01/01
                    ' 附件
                    strAttach = ""
                    For ii = 0 To Me.Check1.Count - 1
                        If Me.Check1(ii).Value = vbChecked Then strAttach = strAttach & Me.Check1(ii).Caption & "、"
                    Next ii
                    If strAttach <> "" Then
                        strAttach = Left(strAttach, Len(strAttach) - 1)
                        strAttach = "附件：" & strAttach & "。"
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & _
                                 "','附件','" & strAttach & "')"
                        cnnConnection.Execute strSql
                    End If
               End If
         End Select
      Case "105": '使用宣誓
        '若有專用期(原有發證日)
        If m_TM22 <> "" Then
            'add by nickc 2006/02/24
            Select Case m_TM10
            Case "030" '菲律賓
                '2006/02/24 may 更改定稿內容，原本共用，現在獨立出來
                EndLetter "03", m_CP09, "06", strUserNum
                ' 列印備註
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
                         "','列印備註','" & textPS & "')"
                cnnConnection.Execute strSql
                'Add By Sindy 2020/1/21
                '應呈提第五, 十, 十五年使用宣誓書
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
                         "','提使用宣誓書年度','" & PUB_Get030To105Year(m_TM01, m_TM02, m_TM03, m_TM04, m_CP07) & "')"
                cnnConnection.Execute strSql
                '2020/1/21 END
            'Add By Sindy 2009/04/20
            Case "101" '美國
                ' 清除定稿例外欄位檔原有資料
                EndLetter "03", m_CP09, "07", strUserNum
                ' 列印備註
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & _
                         "','列印備註','" & textPS & "')"
                cnnConnection.Execute strSql
            '2009/04/20 End
            Case Else
                ' 清除定稿例外欄位檔原有資料
                EndLetter "03", m_CP09, "03", strUserNum
                ' 列印備註
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                         "','列印備註','" & textPS & "')"
                cnnConnection.Execute strSql
            End Select
        '若無專用期
        Else
            'Add By Cheng 2003/02/11
            '判斷申請國家
            Select Case m_TM10
            Case "030" '菲律賓
                ' 清除定稿例外欄位檔原有資料
                EndLetter "03", m_CP09, "05", strUserNum
                ' 列印備註
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & _
                         "','列印備註','" & textPS & "')"
                cnnConnection.Execute strSql
            Case Else '其他國家
                ' 清除定稿例外欄位檔原有資料
                EndLetter "03", m_CP09, "04", strUserNum
                ' 列印備註
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & _
                         "','列印備註','" & textPS & "')"
                cnnConnection.Execute strSql
            End Select
        End If
      'add by sonia 2020/10/27
      Case "501" '移轉
         If m_TM10 = "101" And textCP30 <> "" Then '美國移轉登記號
            ii = InStr(1, textCP30, "/", 1)
            If ii > 0 Then
               EndLetter "03", m_CP09, "02", strUserNum
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                        "','補文件 V 1','" & Mid(textCP30, 1, ii - 1) & "')"
               cnnConnection.Execute strSql
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                        "','補文件 V 2','" & Right(textCP30, Val(Len(textCP30) - ii)) & "')"
               cnnConnection.Execute strSql
            End If
         End If
       'end 2020/10/27
      'Added by Lydia 2023/09/12 更正註冊證1701核准
      Case "1701"
         EndLetter "03", NewCP09, "01", strUserNum

      'Added by Lydia 2024/04/17 更正延展註冊證1713核准
      Case "1713"
         EndLetter "03", NewCP09, "03", strUserNum
         If m_TM10 = "016" Then '016紐西蘭
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                     "','各國設定1','前函通知本案延展核准證明內容有誤已交回當局請求修正，諒已知悉。今接獲代理人轉來完成修正之延展核准證明，隨函附上，敬請查收備存。')"
            cnnConnection.Execute strSql
         ElseIf m_TM10 = "015" Then '015澳洲
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                     "','各國設定1','前函通知本案延展核准通知書內容有誤已交回當局請求修正，諒已知悉。今接獲代理人轉來完成修正之延展核准通知書，隨函附上，敬請查收備存。')"
            cnnConnection.Execute strSql
         ElseIf InStr("206奧地利,101美國,239歐盟", m_TM10) > 0 Then
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                     "','各國設定1','前函通知本案延展核准通知書內容有誤已請當局修正，諒已知悉。今接獲代理人轉來正確之延展核准通知書，隨函附上，敬請查收備存。因奧地利不核發延展證書，故請妥善保留延展證書，以維權益。')"
            cnnConnection.Execute strSql
         ElseIf InStr("114哥倫比亞,115厄瓜多,116秘魯,120玻利維亞,201英國,126智利,118阿根廷", m_TM10) > 0 Then
            '電子證書
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                     "','各國設定1','前函通知本案延展證書內容有誤已請當局修正，諒已知悉。今接獲代理人轉來正確之延展證書(電子證書，" & textTM10 & "商標局已停止核發紙本證書)，隨函附上，敬請查收備存。')"
            cnnConnection.Execute strSql
         ElseIf m_TM10 = "021" Then '021沙烏地阿拉伯
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                     "','特殊年曆','9年8個月（相當於伊斯蘭曆10年）')"
            cnnConnection.Execute strSql
         ElseIf m_TM10 = "044" Then '044澳門
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                     "','例外段落1','敬請　" & PUB_GetCustomerValue(m_TM23, "CU15", "CU15") & "將註冊證續頁與原註冊證合併並妥善保存。')"
            cnnConnection.Execute strSql
         End If
      'end 2024/04/17
   End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Modified by Lydia 2023/09/12
   'Select Case m_CP10
   Select Case IIf(m_strCP10 <> "", m_strCP10, m_CP10)
      '2006/1/3 ADD BY SONIA 增加跨類
      Case "101", "107": '申請
         Select Case m_TM10
            ' 日本
            Case "011":
               ' 列印定稿
               NowPrint m_CP09, "03", "01", False, strUserNum, 0
            'Add By Sindy 2009/10/19
            ' 越南
            Case "042":
               ' 列印定稿
               NowPrint m_CP09, "03", "05", False, strUserNum, 0
            '2009/10/19 End
            ' 德國
            Case "231":
            'Add By Cheng 2003/01/01
            ' 巴西
            Case "117":
               '若未輸入公告日, 處理狀況為"02"
               If IsEmptyText(Me.textTM14.Text) = True Then
                  NowPrint m_CP09, "03", "02", False, strUserNum, 0
               '若輸入公告日, 處理狀況為"03"
               Else
                  NowPrint m_CP09, "03", "03", False, strUserNum, 0
               End If
            'Add By Sindy 2012/11/28
            ' 阿曼
            Case "033":
               NowPrint m_CP09, "03", "06", False, strUserNum, 0
            '2012/11/28 End
            'Add By Sindy 2013/2/26
            ' 香港,新加坡
            Case "013", "014":
               NowPrint m_CP09, "03", "07", False, strUserNum, 0
            '2013/2/26 End
            ' 其它
            Case Else:
               ' 列印定稿
               'Modify By Cheng 2002/07/12
'               NowPrint m_CP09, "03", "02", False, strUserNum, 0
               '若未輸入公告日, 處理狀況為"02"
               If IsEmptyText(Me.textTM14.Text) = True Then
                  NowPrint m_CP09, "03", "02", False, strUserNum, 0
               '若輸入公告日, 處理狀況為"04"
               Else
                  NowPrint m_CP09, "03", "04", False, strUserNum, 0
               End If
         End Select
      Case "105": '使用宣誓
            ' 列印定稿
            'Modify By Cheng 2003/01/01
            '若有專用期(原有發證日)
            If m_TM22 <> "" Then
                'Modify By Sindy 2009/04/17
                'NowPrint m_CP09, "03", "03", False, strUserNum, 0
                '判斷申請國家
                Select Case m_TM10
                Case "030" '菲律賓
                    NowPrint m_CP09, "03", "06", False, strUserNum, 0
                Case "101" '美國
                    NowPrint m_CP09, "03", "07", False, strUserNum, 0
                Case Else '其他國家 柬埔寨
                    NowPrint m_CP09, "03", "03", False, strUserNum, 0
                End Select
                '2009/04/17 End
                
            '若無專用期
            Else
                'Add By Cheng 2003/02/11
                '判斷申請國家
                Select Case m_TM10
                Case "030" '菲律賓
                    NowPrint m_CP09, "03", "05", False, strUserNum, 0
                Case Else '其他國家
                    NowPrint m_CP09, "03", "04", False, strUserNum, 0
                End Select
            End If
      'add by sonia 2020/10/27
      Case "501": '移轉
         If m_TM10 = "101" Then
            NowPrint m_CP09, "03", "02", False, strUserNum, 0
         Else
            NowPrint m_CP09, "03", "01", False, strUserNum, 0
         End If
      'end 2020/10/27
      'Added by Lydia 2023/09/12 更正註冊證1701核准
      Case "1701"
         NowPrint NewCP09, "03", "01", False, strUserNum, 0
         '商品及服務定稿
         NowPrint NewCP09, "03", "02", False, strUserNum, 0
      'Added by Lydia 2024/04/17 更正延展註冊證1713核准
      Case "1713"
         NowPrint NewCP09, "03", "03", False, strUserNum, 0
      'Added by Morgan 2021/12/7 CFT 沒設定都出通用定稿(不列印)
      Case Else
         If m_TM01 = "CFT" Then
            NowPrint m_CP09, "04", "00", False, strUserNum, 0, , , , , , , , , , , , , , , , , True
         End If
      'end 2021/12/7
   End Select
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCF15.Enabled = True Then
   Cancel = False
   textCF15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP06.Enabled = True Then
   Cancel = False
   textCP06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP07.Enabled = True Then
   Cancel = False
   textCP07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP14.Enabled = True Then
   Cancel = False
   textCP14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP25.Enabled = True Then
   Cancel = False
   textCP25_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP26.Enabled = True Then
   Cancel = False
   textCP26_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP48.Enabled = True Then
   Cancel = False
   textCP48_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textFee_1.Enabled = True Then
   Cancel = False
   textFee_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textFee_2.Enabled = True Then
   Cancel = False
   textFee_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPrint.Enabled = True Then
   Cancel = False
   textPrint_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM14.Enabled = True Then
   Cancel = False
   textTM14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Modify By Cheng 2002/07/22
'If Me.textTM16S.Enabled = True Then
'   Cancel = False
'   textTM16S_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If
'If Me.textTM17.Enabled = True Then
'   Cancel = False
'   textTM17_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If

If Me.textCP53.Enabled = True And Me.textCP53.Visible = True Then
   Cancel = False
   textCP53_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP54.Enabled = True And Me.textCP54.Visible = True Then
   Cancel = False
   textCP54_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   If Me.textCP53.Visible And Me.textCP54.Visible Then
      If Val(Me.textCP53.Text) > Val(Me.textCP54.Text) Then
         MsgBox "日期區間輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.textCP53.SetFocus
         textCP53_GotFocus
         Exit Function
      End If
   End If
End If

'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
     Exit Function
End If

TxtValidate = True
End Function

Private Sub textCP53_GotFocus()
InverseTextBox textCP53
End Sub

Private Sub textCP53_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

' 檢核是否為民國日期
If CheckIsDate(Me.textCP53, False) = False Then
   Cancel = True
   strTit = "資料檢核"
   strMsg = "請輸入正確的日期"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Me.textCP53.SetFocus
   textCP53_GotFocus
   Exit Sub
End If
If Val(Me.textCP53.Text) < Val(m_TM21) Or Val(Me.textCP53.Text) > Val(m_TM22) Then
   Cancel = True
   strTit = "資料檢核"
   strMsg = Replace(Me.Label4(0).Caption, "：", "") & "與專用期間不符, 是否重新輸入???" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "專用期間：" & m_TM21 & "－" & m_TM22 & Chr(10) & Chr(13) & Me.Label4(0).Caption & Me.textCP53.Text & "－" & Me.textCP54.Text
   nResponse = MsgBox(strMsg, vbYesNo, strTit)
   If nResponse = vbYes Then
      Me.textCP53.SetFocus
      textCP53_GotFocus
      Exit Sub
   End If
   Cancel = False
End If

End Sub

Private Sub textCP54_GotFocus()
InverseTextBox textCP54
End Sub

Private Sub textCP54_lostfocus()
If Me.textCP53.Visible And Me.textCP54.Visible Then
   If Val(Me.textCP53.Text) > Val(Me.textCP54.Text) Then
      MsgBox Replace(Me.Label4(0).Caption, "：", "") & "輸入錯誤!!!", vbExclamation + vbOKOnly
      Me.textCP53.SetFocus
      textCP53_GotFocus
   End If
End If
End Sub

Private Sub textCP54_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

' 檢核是否為民國日期
If CheckIsDate(Me.textCP54, False) = False Then
   Cancel = True
   strTit = "資料檢核"
   strMsg = "請輸入正確的日期"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Me.textCP54.SetFocus
   textCP54_GotFocus
   Exit Sub
End If
If Val(Me.textCP54.Text) < Val(m_TM21) Or Val(Me.textCP54.Text) > Val(m_TM22) Then
   Cancel = True
   strTit = "資料檢核"
   strMsg = Replace(Me.Label4(0).Caption, "：", "") & "與專用期間不符, 是否重新輸入???" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "專用期間：" & m_TM21 & "－" & m_TM22 & Chr(10) & Chr(13) & Me.Label4(0).Caption & Me.textCP53.Text & "－" & Me.textCP54.Text
   nResponse = MsgBox(strMsg, vbYesNo, strTit)
   If nResponse = vbYes Then
      Me.textCP54.SetFocus
      textCP54_GotFocus
      Exit Sub
   End If
   Cancel = False
End If

End Sub

