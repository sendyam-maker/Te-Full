VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_07 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(變更, 更正)"
   ClientHeight    =   6072
   ClientLeft      =   1128
   ClientTop       =   3816
   ClientWidth     =   9144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6072
   ScaleWidth      =   9144
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   4500
      MaxLength       =   4
      TabIndex        =   3
      Top             =   3780
      Width           =   540
   End
   Begin VB.TextBox txtPayToday 
      Height          =   264
      Left            =   8355
      MaxLength       =   1
      TabIndex        =   4
      Top             =   3840
      Width           =   255
   End
   Begin VB.TextBox textCP118 
      Height          =   270
      Left            =   8160
      MaxLength       =   1
      TabIndex        =   9
      Top             =   4410
      Width           =   375
   End
   Begin VB.TextBox textCP84 
      Alignment       =   1  '靠右對齊
      Height          =   285
      Left            =   3600
      TabIndex        =   1
      Top             =   3480
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5010
      TabIndex        =   59
      Top             =   4890
      Visible         =   0   'False
      Width           =   3915
      Begin VB.OptionButton Option1 
         Caption         =   "註冊變更"
         Height          =   225
         Index           =   1
         Left            =   2340
         TabIndex        =   11
         Top             =   30
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         Caption         =   "申請變更"
         Height          =   225
         Index           =   0
         Left            =   1200
         TabIndex        =   10
         Top             =   30
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "變更種類 :"
         Height          =   255
         Index           =   5
         Left            =   270
         TabIndex        =   60
         Top             =   30
         Width           =   885
      End
   End
   Begin VB.TextBox textCP22 
      Height          =   264
      Left            =   3645
      MaxLength       =   1
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1110
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5430
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   810
      Width           =   3675
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1410
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   510
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   810
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5430
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2310
      Width           =   3675
   End
   Begin VB.TextBox textTM06 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2910
      Width           =   7572
   End
   Begin VB.ComboBox textCP44 
      Height          =   300
      Left            =   1200
      TabIndex        =   6
      Top             =   4110
      Width           =   1452
   End
   Begin VB.TextBox textCF09 
      Height          =   264
      Left            =   5190
      MaxLength       =   12
      TabIndex        =   8
      Top             =   4410
      Width           =   612
   End
   Begin VB.TextBox textCP18 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   6540
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3510
      Width           =   2532
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   7
      Top             =   4410
      Width           =   372
   End
   Begin VB.TextBox textUargeDate 
      Height          =   264
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   2
      Top             =   3810
      Width           =   1092
   End
   Begin VB.TextBox textCP27 
      Height          =   264
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   0
      Top             =   3510
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   1
      Left            =   2496
      TabIndex        =   14
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   3720
      TabIndex        =   15
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H008080FF&
      Caption         =   "變更事項(&R)"
      Height          =   400
      Left            =   4968
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6996
      TabIndex        =   18
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6168
      TabIndex        =   17
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8220
      TabIndex        =   19
      Top             =   70
      Width           =   800
   End
   Begin MSForms.TextBox textTM67 
      Height          =   300
      Left            =   1200
      TabIndex        =   13
      Top             =   5730
      Width           =   7815
      VariousPropertyBits=   671105051
      MaxLength       =   200
      Size            =   "13785;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   510
      Left            =   1200
      TabIndex        =   12
      Top             =   5205
      Width           =   7815
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13785;900"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05 
      Height          =   300
      Left            =   1440
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2580
      Width           =   7572
      VariousPropertyBits=   679493661
      Size            =   "13356;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM07 
      Height          =   300
      Left            =   1440
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3240
      Width           =   7572
      VariousPropertyBits=   679493661
      Size            =   "13356;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP44_2 
      Height          =   264
      Left            =   2664
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4110
      Width           =   6324
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44 
      Height          =   264
      Left            =   5430
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   510
      Width           =   3675
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM81 
      Height          =   264
      Left            =   1200
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   2310
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM80 
      Height          =   264
      Left            =   5430
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   2010
      Width           =   3675
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM79 
      Height          =   264
      Left            =   1200
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   2010
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM78 
      Height          =   264
      Left            =   5430
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   1710
      Width           =   3675
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM05_1 
      Height          =   870
      Left            =   1440
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   2580
      Width           =   7575
      VariousPropertyBits=   679493659
      ScrollBars      =   2
      Size            =   "13356;1535"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5430
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1110
      Width           =   3675
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   264
      Left            =   5430
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1410
      Width           =   3675
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1200
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1710
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   500
      Left            =   1200
      TabIndex        =   77
      Top             =   4680
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;882"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   18
      Left            =   3600
      TabIndex        =   76
      Top             =   3825
      Width           =   765
   End
   Begin VB.Label lblPayToday 
      AutoSize        =   -1  'True
      Caption         =   "電子送件是否當日扣款:         (Y/N)"
      Height          =   180
      Left            =   6420
      TabIndex        =   75
      Top             =   3870
      Width           =   2655
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:          (Y: 是)"
      Height          =   180
      Left            =   6990
      TabIndex        =   73
      Top             =   4440
      Width           =   2085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5 :"
      Height          =   180
      Index           =   14
      Left            =   120
      TabIndex        =   68
      Top             =   2340
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4 :"
      Height          =   180
      Index           =   13
      Left            =   4410
      TabIndex        =   67
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3 :"
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   66
      Top             =   2070
      Width           =   720
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   270
      TabIndex        =   65
      Top             =   4785
      Width           =   900
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "發文規費："
      Height          =   255
      Left            =   2670
      TabIndex        =   64
      Top             =   3540
      Width           =   900
   End
   Begin VB.Label Label10 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   63
      Top             =   2580
      Width           =   1335
   End
   Begin VB.Label Label28 
      Caption         =   "進度備註 :"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   5205
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "放棄專用權 :"
      Height          =   255
      Left            =   120
      TabIndex        =   61
      Top             =   5730
      Width           =   1155
   End
   Begin VB.Label Label30 
      Caption         =   "是否出名 :"
      Height          =   255
      Left            =   2745
      TabIndex        =   58
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label31 
      Caption         =   "(N:不出名)"
      Height          =   255
      Left            =   4050
      TabIndex        =   57
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "審定號數/申請案號 :"
      Height          =   180
      Left            =   120
      TabIndex        =   56
      Top             =   1152
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4410
      TabIndex        =   55
      Top             =   1110
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   255
      Index           =   9
      Left            =   4410
      TabIndex        =   54
      Top             =   810
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   53
      Top             =   1410
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   52
      Top             =   510
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   51
      Top             =   810
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "發證日 :"
      Height          =   255
      Index           =   3
      Left            =   4410
      TabIndex        =   50
      Top             =   2310
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "FC代理人 :"
      Height          =   255
      Index           =   2
      Left            =   4410
      TabIndex        =   49
      Top             =   510
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   4410
      TabIndex        =   48
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2 :"
      Height          =   180
      Index           =   8
      Left            =   4410
      TabIndex        =   47
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人1 :"
      Height          =   180
      Left            =   120
      TabIndex        =   46
      Top             =   1770
      Width           =   720
   End
   Begin VB.Label Label9 
      Caption         =   "案件中文名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   2610
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "案件英文名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   2910
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "案件日文名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "可接獲回音"
      Height          =   255
      Left            =   5820
      TabIndex        =   42
      Top             =   4410
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "大約"
      Height          =   255
      Index           =   12
      Left            =   4650
      TabIndex        =   41
      Top             =   4410
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "點數 :"
      Height          =   255
      Index           =   10
      Left            =   5970
      TabIndex        =   40
      Top             =   3540
      Width           =   735
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
      Height          =   180
      Left            =   1680
      TabIndex        =   39
      Top             =   4410
      Width           =   2745
   End
   Begin VB.Label Label22 
      Caption         =   "列印定稿 :"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   4410
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "代理人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   4110
      Width           =   975
   End
   Begin VB.Label Label14 
      Caption         =   "催審期限 :"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   3810
      Width           =   975
   End
   Begin VB.Label Label25 
      Caption         =   "發文日 :"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   3510
      Width           =   855
   End
End
Attribute VB_Name = "frm020102_07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/23 Form2.0已修改 textTm44/textCP13/textCP14/textTM23/textTM78/textTM79/textTM80/textTM81/textTM05/textTM05_1/textTM07/textCP44_2/textCP64/lstNameAgent/textTM67(111/8/8 Lydia)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 收文號
Dim m_CP09 As String
Dim m_CP31 As String 'Add By Sindy 2011/7/12
' 申請國家
Dim m_TM10 As String
Dim m_TM09 As String 'Add By Sindy 2010/11/12 商品類別
' 案件性質代號
Dim m_CP10 As String
' 申請人
Dim m_TM23 As String
'add by nickc 2007/01/03
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存商標基本檔或服務業務基本檔檔案欄位的串列
Dim m_TMSPList() As FIELDITEM
Dim m_TMSPCount As Integer
' 儲存案件進度檔檔案欄位的串列
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer
' 宣告代理人內容結構
Private Type AGENTITEM
   aiCode As String
   aiName As String
End Type
Dim m_AgentList() As AGENTITEM
Dim m_AgentCount As Integer
'Add By Cheng 2002/12/12
Dim m_CP14 As String '原承辦人
'Add By Cheng 2003/10/06
Public m_blnClkChgButton As Boolean '是否按下變更事項按鈕
'add by nick 2004/08/12
Dim m_CP84 As String       '發文規費
'add by nick 2004/09/27
Public m_CU103 As String         '公司負責人英文名稱
'add by nick 2004/10/05
Public m_CU05 As String         '客戶英文名稱
Public m_CU88 As String         '客戶英文名稱
Public m_CU89 As String         '客戶英文名稱
Public m_CU90 As String         '客戶英文名稱
'add by nickc 2006/01/20
Public m_CU112 As String        '客戶中文地址郵遞區號
'Add By Sindy 2012/2/7
Public m_CU39 As String         '代表人1（中）
Public m_CU40 As String         '代表人1（英）
Public m_CU41 As String         '代表人1（日）
'2012/2/7 End

Dim m_TM24 As String
'add by nickc 2006/01/27
Dim m_CP110 As String
'add by nickc 2007/01/03
Dim m_textUargeDate As String
'add by nickc 2007/08/10
Dim SeekCu05(1 To 5) As String
Dim SeekCu88(1 To 5) As String
Dim SeekCu89(1 To 5) As String
Dim SeekCu90(1 To 5) As String
Dim SeekCu103(1 To 5) As String
Dim SeekCu112(1 To 5) As String
'Add By Sindy 2012/2/7
Dim SeekCu39(1 To 5) As String
Dim SeekCu40(1 To 5) As String
Dim SeekCu41(1 To 5) As String
'2012/2/7 End
'Add By Sindy 2012/10/31
Public m_CU10 As String
Dim SeekCu10(1 To 5) As String
'2012/10/31 End
'add by nickc 2008/02/22
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP09s As String, m_CP123s As String 'Add by Sindy 98/3/24 收文號,是否算發文室案件
Dim m_CP130s As String 'Add by Sindy 2009/4/24 發文-主管機關
' 智權人員
Dim m_CP13 As String
Dim m_CP12 As String
'Dim strCP45 As String
'2010/6/10 End
Dim m_CP07 As String 'Add By Sindy 2010/12/28 法定期限
Dim m_QSP As Boolean 'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
Dim m_CP16 As String 'Add By Sindy 2016/5/30 費用
Dim m_990CP09 As String 'Add By Sindy 2016/12/20
Dim strLD18 As String 'Add By Sindy 2019/12/20 信函總收文號
Dim m_strCF10 As String 'Add By Sindy 2020/8/12 取得主管機關
Dim m_AgentName As String 'Add By Amy 2021/12/23

Private Sub cmdCancel_Click()
   'Add By Sindy 2018/5/3
   If frm020102_01.bolIsEMPFlow = True Then
      frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
      frm090202_4.QueryData
   End If
   '2018/5/3 End
   frm020102_01.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
    'Add By Cheng 2004/04/08
'move by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   ' 90.10.09 modify by louis
   'Add By Sindy 2018/5/3
   If frm020102_01.bolIsEMPFlow = True Then
      frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
      frm090202_4.QueryData
   End If
   '2018/5/3 End
   Unload frm020102_01
   'frm020102_01.Show
   Unload Me
End Sub

Private Sub cmdMod_Click()
   frm020102_04.SetData 0, m_TM01, True
   frm020102_04.SetData 1, m_TM02, False
   frm020102_04.SetData 2, m_TM03, False
   frm020102_04.SetData 3, m_TM04, False
   frm020102_04.SetData 4, m_CP09, False
   ' 91.09.02 modify by louis (增加案件性質參數)
   frm020102_04.SetData 5, m_CP10, False
   frm020102_04.SetParent Me
   frm020102_04.SetParent_MainForm frm020102_01 'Add By Sindy 2018/9/25
   Me.Hide
   frm020102_04.Show
   frm020102_04.QueryData
'    m_blnClkChgButton = True
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim strNewCP64 As String 'Add by Amy 2020/02/05 進度備註
   
   'Modify By Sindy 2010/11/19 把「確定」及「同時發文」按鈕程式碼合併
   Select Case Index
      Case 0, 1
         If CheckDataValid = True Then
            'Add By Sindy 2022/2/7 T和FCT的減縮商品313發文：
            '自動新增變更事項檔並寫入CE01(總收文號)，CE45(“減縮商品請參卷宗區附件”)，若已該收文號已存在於變更事項檔則更新。
            If m_CP10 = "313" Then
               If IsChangeEventExist(m_CP09) = True Then
                  strSql = "Update CHANGEEVENT Set ce45='減縮商品請參卷宗區附件' Where ce01='" & m_CP09 & "'"
                  cnnConnection.Execute strSql
               Else
                  strSql = "insert into CHANGEEVENT(ce01,ce45) values ('" & m_CP09 & "','減縮商品請參卷宗區附件')"
                  cnnConnection.Execute strSql
               End If
            End If
            '2022/2/7 END
                        
            ' 90.08.29 檢查變更事項檔
            If IsChangeEventExist(m_CP09) = True Then
               'Add By Cheng 2002/05/23
               '重新檢查欄位有效性
               If TxtValidate = False Then Exit Sub
               'add by nick 2004/09/27
               'edit by nick 2004/10/07
               'If m_TM01 <> "FCT" Then
               If m_TM01 <> "FCT" And m_TM01 <> "TB" And m_TM01 <> "TC" And m_TM01 <> "TD" And (m_TM01 = "T" And m_TM10 <> "020") Then
                     'add by nickc 2007/08/10
                     SeekCu05(1) = "": SeekCu05(2) = "": SeekCu05(3) = "": SeekCu05(4) = "": SeekCu05(5) = ""
                     SeekCu88(1) = "": SeekCu88(2) = "": SeekCu88(3) = "": SeekCu88(4) = "": SeekCu88(5) = ""
                     SeekCu89(1) = "": SeekCu89(2) = "": SeekCu89(3) = "": SeekCu89(4) = "": SeekCu89(5) = ""
                     SeekCu90(1) = "": SeekCu90(2) = "": SeekCu90(3) = "": SeekCu90(4) = "": SeekCu90(5) = ""
                     SeekCu103(1) = "": SeekCu103(2) = "": SeekCu103(3) = "": SeekCu103(4) = "": SeekCu103(5) = ""
                     SeekCu112(1) = "": SeekCu112(2) = "": SeekCu112(3) = "": SeekCu112(4) = "": SeekCu112(5) = ""
                     'Add By Sindy 2012/2/7
                     SeekCu39(1) = "": SeekCu39(2) = "": SeekCu39(3) = "": SeekCu39(4) = "": SeekCu39(5) = ""
                     SeekCu40(1) = "": SeekCu40(2) = "": SeekCu40(3) = "": SeekCu40(4) = "": SeekCu40(5) = ""
                     SeekCu41(1) = "": SeekCu41(2) = "": SeekCu41(3) = "": SeekCu41(4) = "": SeekCu41(5) = ""
                     '2012/2/7 End
                     'Add By Sindy 2012/10/31
                     SeekCu10(1) = "": SeekCu10(2) = "": SeekCu10(3) = "": SeekCu10(4) = "": SeekCu10(5) = ""
                     '2012/10/31 End
                       'Modified by Lydia 2024/07/03 改傳入變數;
                       'GetCu103ByCustomer Me, m_TM23
                       Call Pub_GetDataFrm020102(m_TM23, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                       
                       'edit by nickc 2006/01/20
                       'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Then
                       'edit by nickc 2007/11/13 不管如何，皆要秀出來
                       'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                              'Modified by Lydia 2024/07/03
                              'Set frm020102_22.oNextForm = Me
                              Call frm020102_22.SetParent(Me, m_TM23)
                              frm020102_22.Label4.Caption = m_TM23 & " " & textTM23 'Add By Sindy 2014/7/30
                              frm020102_22.Show vbModal
                              'add by nickc 2007/08/10
                              SeekCu05(1) = m_CU05
                              SeekCu88(1) = m_CU88
                              SeekCu89(1) = m_CU89
                              SeekCu90(1) = m_CU90
                              SeekCu103(1) = m_CU103
                              SeekCu112(1) = m_CU112
                              'Add By Sindy 2012/2/27
                              SeekCu39(1) = m_CU39
                              SeekCu40(1) = m_CU40
                              SeekCu41(1) = m_CU41
                              '2012/2/27 End
                              'Add By Sindy 2012/10/31
                              SeekCu10(1) = m_CU10
                              '2012/10/31 End
                       'edit by nickc 2007/11/13 不管如何，皆要秀出來
                       'End If
                     'add by nickc 2007/08/10 多申請人也要
                     If m_TM78 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                     'Modified by Lydia 2024/07/03 改傳入變數;
                     'GetCu103ByCustomer Me, m_TM78
                     Call Pub_GetDataFrm020102(m_TM78, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                     
                     'edit by nickc 2007/11/13 不管如何，皆要秀出來
                     'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                           'Modified by Lydia 2024/07/03
                           'Set frm020102_22.oNextForm = Me
                           Call frm020102_22.SetParent(Me, m_TM78)
                           frm020102_22.Label4.Caption = m_TM78 & " " & textTM78 'Add By Sindy 2014/7/30
                           frm020102_22.Show vbModal
                           SeekCu05(2) = m_CU05
                           SeekCu88(2) = m_CU88
                           SeekCu89(2) = m_CU89
                           SeekCu90(2) = m_CU90
                           SeekCu103(2) = m_CU103
                           SeekCu112(2) = m_CU112
                           'Add By Sindy 2012/2/27
                           SeekCu39(2) = m_CU39
                           SeekCu40(2) = m_CU40
                           SeekCu41(2) = m_CU41
                           '2012/2/27 End
                           'Add By Sindy 2012/10/31
                           SeekCu10(2) = m_CU10
                           '2012/10/31 End
                     'edit by nickc 2007/11/13 不管如何，皆要秀出來
                     'End If
                     End If
                     If m_TM79 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                     'Modified by Lydia 2024/07/03 改傳入變數;
                     'GetCu103ByCustomer Me, m_TM79
                     Call Pub_GetDataFrm020102(m_TM79, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                     
                     'edit by nickc 2007/11/13 不管如何，皆要秀出來
                     'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                           'Modified by Lydia 2024/07/03
                           'Set frm020102_22.oNextForm = Me
                           Call frm020102_22.SetParent(Me, m_TM79)
                           frm020102_22.Label4.Caption = m_TM79 & " " & textTM79 'Add By Sindy 2014/7/30
                           frm020102_22.Show vbModal
                           SeekCu05(3) = m_CU05
                           SeekCu88(3) = m_CU88
                           SeekCu89(3) = m_CU89
                           SeekCu90(3) = m_CU90
                           SeekCu103(3) = m_CU103
                           SeekCu112(3) = m_CU112
                           'Add By Sindy 2012/2/27
                           SeekCu39(3) = m_CU39
                           SeekCu40(3) = m_CU40
                           SeekCu41(3) = m_CU41
                           '2012/2/27 End
                           'Add By Sindy 2012/10/31
                           SeekCu10(3) = m_CU10
                           '2012/10/31 End
                     'edit by nickc 2007/11/13 不管如何，皆要秀出來
                     'End If
                     End If
                     If m_TM80 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                     'Modified by Lydia 2024/07/03 改傳入變數;
                     'GetCu103ByCustomer Me, m_TM80
                     Call Pub_GetDataFrm020102(m_TM80, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                     
                     'edit by nickc 2007/11/13 不管如何，皆要秀出來
                     'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                           'Modified by Lydia 2024/07/03
                           'Set frm020102_22.oNextForm = Me
                           Call frm020102_22.SetParent(Me, m_TM80)
                           frm020102_22.Label4.Caption = m_TM80 & " " & textTM80 'Add By Sindy 2014/7/30
                           frm020102_22.Show vbModal
                           SeekCu05(4) = m_CU05
                           SeekCu88(4) = m_CU88
                           SeekCu89(4) = m_CU89
                           SeekCu90(4) = m_CU90
                           SeekCu103(4) = m_CU103
                           SeekCu112(4) = m_CU112
                           'Add By Sindy 2012/2/27
                           SeekCu39(4) = m_CU39
                           SeekCu40(4) = m_CU40
                           SeekCu41(4) = m_CU41
                           '2012/2/27 End
                           'Add By Sindy 2012/10/31
                           SeekCu10(4) = m_CU10
                           '2012/10/31 End
                     'edit by nickc 2007/11/13 不管如何，皆要秀出來
                     'End If
                     End If
                     If m_TM81 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                     'Modified by Lydia 2024/07/03 改傳入變數;
                     'GetCu103ByCustomer Me, m_TM81
                     Call Pub_GetDataFrm020102(m_TM81, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                     
                     'edit by nickc 2007/11/13 不管如何，皆要秀出來
                     'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                           'Modified by Lydia 2024/07/03
                           'Set frm020102_22.oNextForm = Me
                           Call frm020102_22.SetParent(Me, m_TM81)
                           frm020102_22.Label4.Caption = m_TM81 & " " & textTM81 'Add By Sindy 2014/7/30
                           frm020102_22.Show vbModal
                           SeekCu05(5) = m_CU05
                           SeekCu88(5) = m_CU88
                           SeekCu89(5) = m_CU89
                           SeekCu90(5) = m_CU90
                           SeekCu103(5) = m_CU103
                           SeekCu112(5) = m_CU112
                           'Add By Sindy 2012/2/27
                           SeekCu39(5) = m_CU39
                           SeekCu40(5) = m_CU40
                           SeekCu41(5) = m_CU41
                           '2012/2/27 End
                           'Add By Sindy 2012/10/31
                           SeekCu10(5) = m_CU10
                           '2012/10/31 End
                     'edit by nickc 2007/11/13 不管如何，皆要秀出來
                     'End If
                     End If
                 End If
                 
                 strNewCP64 = textCP64 'Add by Amy 2020/02/05
                 
                 'Modify By Sindy 2011/3/9 若為電子送件則不經發文室
                 'Modify By Sindy 2023/8/1 電子送件欄位值不是空白者,即為電子送件
                 If (textCP118.Visible = True And textCP118 <> "") Then
                     'Added by Morgan 2016/5/16 電子送件也要記錄主管機關
                     If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27, , True) = False Then
                        Exit Sub
                     End If
                     'end 2016/5/16
                     
                     'Add by Amy 2020/02/05 +輸入收文文號
                     If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
                        'Add By Sindy 2020/8/12 主管機關為經濟部智慧財產局,才做自動扣款
                        If m_CP130s = "經濟部智慧財產局" Then
                        '2020/8/12 END
                           'Add by Amy 2020/01/13
                           'If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
                              'If textCP118 = "Y" And Val(textCP84) > 0 Then
                              If Val(textCP84) > 0 Then
                                 If txtPayToday.Visible = True And txtPayToday = "" Then
                                    MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
                                    txtPayToday.SetFocus
                                    Exit Sub
                                 End If
                                 strExc(0) = InputBox("請輸入智慧局收文文號!!")
                                 If strExc(0) = "" Then
                                    Exit Sub
                                 Else
                                    strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & textCP64 '先保留進度備註，等檢查完後更新欄位
                                 End If
                              End If
                           'End If
                           'end 2020/01/13
                        'Add By Sindy 2020/8/12
                        ElseIf txtPayToday.Visible = True And txtPayToday <> "" Then
                           txtPayToday = ""
                        End If
                        '2020/8/12 END
                     End If
                     'end 2020/02/05
                 Else
                     'Add by Sindy 98/3/24
                     If m_TM10 = "000" Then
                        m_CP09s = m_CP09
                        'Add by Sindy 2009/4/24
                        If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
                           Exit Sub
                  '      Else
                  '         m_CP123s = GetCPMSendYn(m_TM01, m_CP10, 1)
                        End If
                     End If
                 End If
               
               'Added by Lydia 2019/12/09 T台灣案在申請意見書202之後收文放棄專用權206或減縮商品313，於發文時提醒修改預估准駁。
               If m_TM01 = "T" And m_TM10 = "000" And m_CP10 = "313" Then
                   strExc(1) = m_TM01: strExc(2) = m_TM02: strExc(3) = m_TM03: strExc(4) = m_TM04
                   If PUB_ChkCPExist(strExc, "202", 2) = True Then
                       MsgBox "此案已有申請意見書發文，請自行判斷是否修改預估准駁！", vbInformation, "案件提醒"
                   End If
               End If
               'end 2019/12/09
               
               textCP64 = strNewCP64 'Add by Amy 2020/02/05
               
               ' 設定滑鼠游標為等待狀態
               Screen.MousePointer = vbHourglass
               ' 更新欄位輸入的內容
               OnUpdateField
               ' 存檔
               'Modify By Cheng 2002/11/06
         '      'OnSaveData
               If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
               ' 列印定稿
               If textPrint <> "N" Then
                  PrintLetter
               'Add By Sindy 2021/2/25
               End If
               If textPrint = "N" Then
                  If strLD18 <> "" Then
                     Call PUB_TCaseAskIsPost(strLD18)
                  End If
               '2021/2/25 END
               End If
               
               '2012/7/23 add by sonia
               '台灣案發文規費與收文規費不符時,mail給智權人員
               If textCP84.Enabled = True And m_TM10 = "000" And Val(Me.textCP84.Text) <> Val(m_CP84) Then
                    'Add by Lydia 2014/10/13 內商服務業務(TC)之台灣案發文-規費與收文規費不符時,請加同時發給特殊設定人員"財務處總帳人員"
                    If m_QSP = True Then
                      PUB_ChkOfficialFee m_CP09, Me.textCP84.Text, "A"
                    Else
                      '2020/01/13 Modify by Amy +if 傳strCP118參數
                      If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
                        PUB_ChkOfficialFee m_CP09, Me.textCP84.Text, IIf(textCP118 = "Y", "A", "")
                      Else
                        PUB_ChkOfficialFee m_CP09, Me.textCP84.Text
                      End If
                    End If
               End If
               '2012/7/23 end
            
               ' 設定滑鼠游標為預設
               Screen.MousePointer = vbDefault
               
               'Add By Sindy 2018/5/3
               If frm020102_01.bolIsEMPFlow = True Then
                  frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
                  frm090202_4.QueryData
               End If
               '2018/5/3 End
               
               'Add By Sindy 2025/7/11 外商發文時,增加發Mail通知承辦人及副本給判發主管
               If Left(m_CP12, 1) = "F" Then
                  Call PUB_FCTSendRecvMail(m_CP09)
               End If
               '2025/7/11 END
               
               If Index = 0 Then '確定鍵
                  '********* 901123 nick   清畫面
                  'frm020102_01.radio(0).Value = True
                  'frm020102_01.textCP09.Enabled = True
                  'frm020102_01.textCP09.Text = ""
                  'frm020102_01.textTM01.Enabled = False
                  'frm020102_01.textTM01.Text = "" modify by sonia
                  'frm020102_01.textTM02.Enabled = False
                  'frm020102_01.textTM02.Text = ""
                  'frm020102_01.textTM02_2.Enabled = False
                  'frm020102_01.textTM02_2.Text = ""
                  'frm020102_01.textTM03.Enabled = False
                  'frm020102_01.textTM03.Text = ""
                  'frm020102_01.textTM04.Enabled = False
                  'frm020102_01.textTM04.Text = ""
                  'frm020102_01.grdList.Clear
                  'frm020102_01.grdList.Rows = 2
                  '*********************************
                  'frm020102_01.RefreshData
                  'Add By Cheng 2002/04/30
                  '若有未發文資料顯示警告
                  If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = False Then
                     'Add By Sindy 2018/5/3
                     If frm020102_01.bolIsEMPFlow = True Then
                        Unload frm020102_01
                        frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
                        frm090202_4.Show
                        Unload Me
                        Exit Sub
                     End If
                     '2018/5/3 End
                  End If
                  
                  frm020102_01.Show
                  ' 90.12.07 modify by louis
         '         frm020102_01.Clear
         
                  'Add By Cheng 2002/01/10
                  frm020102_01.Clear1
                  
                  Unload Me
               ElseIf Index = 1 Then '同時發文鍵
                  ' 呼叫第一個畫面
                  frm020102_01.SetData 0, m_TM01, True
                  frm020102_01.SetData 1, m_TM02, False
                  frm020102_01.SetData 2, m_TM03, False
                  frm020102_01.SetData 3, m_TM04, False
                  frm020102_01.SetQueryFromTM
                  Unload Me
                  frm020102_01.Show
                  frm020102_01.radio(1).Value = True
                  frm020102_01.radio_Click 1
                  frm020102_01.QueryData
               End If
            Else
               MsgBox "請先輸入變更事項資料!", vbCritical + vbOKOnly, "檢核資料"
            End If
         End If
      Case Else
   End Select
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

'Private Sub cmdTogether_Click()
'   If CheckDataValid = True Then
'      ' 90.08.29 檢查變更事項檔
'      If IsChangeEventExist(m_CP09) = True Then
'         'Add By Cheng 2002/07/15
'         '重新檢查欄位有效性
'         If TxtValidate = False Then Exit Sub
'
'         ' 設定滑鼠游標為等待狀態
'         Screen.MousePointer = vbHourglass
'         ' 更新欄位輸入的內容
'         OnUpdateField
'         ' 存檔
'        'Modify By Cheng 2002/11/06
''      'OnSaveData
'      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
'        ' 列印定稿
'        If textPrint <> "N" Then
'           PrintLetter
'        End If
'
'         ' 設定滑鼠游標為預設
'         Screen.MousePointer = vbDefault
'
'         ' 呼叫第一個畫面
'         frm020102_01.SetData 0, m_TM01, True
'         frm020102_01.SetData 1, m_TM02, False
'         frm020102_01.SetData 2, m_TM03, False
'         frm020102_01.SetData 3, m_TM04, False
'         frm020102_01.SetQueryFromTM
'         Unload Me
'         frm020102_01.Show
'         frm020102_01.radio(1).Value = True
'         frm020102_01.radio_Click 1
'         frm020102_01.QueryData
'      Else
'         MsgBox "請先輸入變更事項資料!", vbCritical + vbOKOnly, "檢核資料"
'      End If
'   End If
'End Sub

'Private Sub Form_Activate()
'    'Add By Cheng 2003/10/06
'    '若有按下變更事項按鈕, 則重新讀取資料
'    'edit by nickc 2005/08/23
'    'If m_blnClkChgButton = True  Then
'    If m_blnClkChgButton = True Or (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
'        pub_ModifyCaseNum = ""
'        QueryData
''        m_blnClkChgButton = False
'    End If
'End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM05.BackColor = &H8000000F
   textTM05_1.BackColor = &H8000000F
   textTM06.BackColor = &H8000000F
   textTM07.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textTM44.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   
   'add by nickc 2007/01/24
   textTM78.BackColor = &H8000000F
   textTM79.BackColor = &H8000000F
   textTM80.BackColor = &H8000000F
   textTM81.BackColor = &H8000000F
   
   MoveFormToCenter Me
'    m_blnClkChgButton = False
   'Add by nickc 2006/01/27
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Add by Amy 2021/12/23一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 500
   lstNameAgent.Width = 1300

End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2012/4/17
   
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 收文號
      Case 0: m_CP09 = strData
         'Add By Sindy 2012/4/17
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

Private Sub ClearAgentList()
   If m_AgentCount > 0 Then
      Erase m_AgentList
   End If
   m_AgentCount = 0
End Sub

Private Sub AddAgent(ByVal strAgentCode As String, ByVal strAgentName As String)
   Dim nIndex As Integer
   Dim bFind As Boolean
   bFind = False
   For nIndex = 0 To m_AgentCount - 1
      If m_AgentList(nIndex).aiCode = strAgentCode Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_AgentList(m_AgentCount + 1)
      m_AgentList(m_AgentCount).aiCode = strAgentCode
      m_AgentList(m_AgentCount).aiName = strAgentName
      m_AgentCount = m_AgentCount + 1
   End If
End Sub

' 清除商標基本檔檔案欄位串列
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiOldData = strFieldData
         m_TMSPList(nPos).fiNewData = strFieldData
         m_TMSPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_TMSPList(m_TMSPCount + 1)
      m_TMSPList(m_TMSPCount).fiName = strFieldName
      m_TMSPList(m_TMSPCount).fiOldData = strFieldData
      m_TMSPList(m_TMSPCount).fiNewData = strFieldData
      m_TMSPList(m_TMSPCount).fiType = nFieldType
      m_TMSPCount = m_TMSPCount + 1
   End If
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 清除案件進度檔檔案欄位串列
Private Sub ClearCPFieldList()
   If m_CPCount > 0 Then
      Erase m_CPList
   End If
   m_CPCount = 0
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiOldData = strFieldData
         m_CPList(nPos).fiNewData = strFieldData
         m_CPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_CPList(m_CPCount + 1)
      m_CPList(m_CPCount).fiName = strFieldName
      m_CPList(m_CPCount).fiOldData = strFieldData
      m_CPList(m_CPCount).fiNewData = strFieldData
      m_CPList(m_CPCount).fiType = nFieldType
      m_CPCount = m_CPCount + 1
   End If
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 取得商標基本檔的欄位內容
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSubSQL As String
   Dim rsSubTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_TM44 = CheckStr(rsTmp.Fields("TM44"))
      'Add By Sindy 2013/1/31
      If m_TM44 <> "" Then
         textTM44 = m_TM44 & "  " & GetPrjName1(m_TM44)
      Else
         textTM44 = ""
      End If
      '2013/1/31 End
      m_TM119 = CheckStr(rsTmp.Fields("TM119"))
      m_TM120 = CheckStr(rsTmp.Fields("TM120"))
      ' 審定號數
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      Else
         'Add By Sindy 2009/06/29
         ' 申請案號
         If IsNull(rsTmp.Fields("TM12")) = False Then
            textTM15 = rsTmp.Fields("TM12")
         End If
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         'textTM12 = rsTmp.Fields("TM12")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
      End If
      ' 案件中文名稱
'      textTM05 = Empty
'      If IsNull(rsTmp.Fields("TM05")) = False Then
'         textTM05 = rsTmp.Fields("TM05")
'      End If
      textTM05_1 = Empty
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05_1 = rsTmp.Fields("TM05")
      End If
'      ' 案件英文名稱
'      textTM06 = Empty
'      If IsNull(rsTmp.Fields("TM06")) = False Then
'         textTM06 = rsTmp.Fields("TM06")
'      End If
'      ' 案件日文名稱
'      textTM07 = Empty
'      If IsNull(rsTmp.Fields("TM07")) = False Then
'         textTM07 = rsTmp.Fields("TM07")
'      End If
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         'textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      'Add By Sindy 2010/11/12 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         m_TM09 = rsTmp.Fields("TM09")
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
      End If
      'add by nickc 2007/01/03
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
      ' 正商標號數
      If IsNull(rsTmp.Fields("TM27")) = False Then
         'textTM27 = rsTmp.Fields("TM27")
      End If
        'Add By Cheng 2003/05/30
        '放棄專用權
        If IsNull(rsTmp.Fields("TM67")) = False Then
            textTM67 = rsTmp.Fields("TM67")
        End If
      'add by nickc 2006/01/26
      m_TM24 = CheckStr(rsTmp.Fields("tm24"))
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("tm77"))
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 取得服務業務基本檔的欄位內容
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      'add by nickc 2008/02/22
      m_TM44 = CheckStr(rsTmp.Fields("SP26"))
      'Add By Sindy 2013/1/31
      If m_TM44 <> "" Then
         textTM44 = m_TM44 & "  " & GetPrjName1(m_TM44)
      Else
         textTM44 = ""
      End If
      '2013/1/31 End
      Select Case m_TM01
      Case "TS"
          textTM05_1 = "" & rsTmp.Fields("SP05")
      Case Else
          ' 案件中文名稱
          If IsNull(rsTmp.Fields("SP05")) = False Then
             textTM05 = rsTmp.Fields("SP05")
          End If
      End Select
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textTM06 = rsTmp.Fields("SP06")
      End If
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textTM07 = rsTmp.Fields("SP07")
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
      End If
      'add by nickc 2007/01/03
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("SP58")) = False Then
         m_TM78 = rsTmp.Fields("SP58")
      End If
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("SP59")) = False Then
         m_TM79 = rsTmp.Fields("SP59")
      End If
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("SP65")) = False Then
         m_TM80 = rsTmp.Fields("SP65")
      End If
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("SP66")) = False Then
         m_TM81 = rsTmp.Fields("SP66")
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         'textTM12 = rsTmp.Fields("SP11")
         'Add By Sindy 2009/06/29
         textTM15 = rsTmp.Fields("SP11")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("SP12")) = False Then
         textTM20 = TAIWANDATE(rsTmp.Fields("SP12"))
      End If
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("sp72"))
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgress()
Dim strSql As String
Dim strSubSQL As String
Dim rsTmp As New ADODB.Recordset
Dim rsSubTmp As New ADODB.Recordset
Dim strCP27 As String
Dim strCP44 As String
Dim nIndex As Integer
Dim bFind As Boolean
'Add By Cheng 2002/07/09
Dim strTempName As String
Dim m_Fee As String         '銷帳服務費 2012/8/3 add by sonia
Dim m_Official As String    '銷帳規費   2012/8/3 add by sonia
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      ' 案件性質
      'Add By Cheng 2002/07/17
      m_CP10 = ""
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      
      'Add By Sindy 2016/5/30 費用
      m_CP16 = Empty
      If IsNull(rsTmp.Fields("CP16")) = False Then
         m_CP16 = rsTmp.Fields("CP16")
      End If
      '2016/5/30 END
      
      ' 業務區別
      m_CP12 = ""
      If IsNull(rsTmp.Fields("CP12")) = False Then
         '91.6.11 MODIFY BY SONIA
         'textCP12 = GetStaffDepartment(rsTmp.Fields("CP12"))
         'textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      m_CP13 = ""
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
         m_CP13 = rsTmp.Fields("CP13")
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      'Add By Cheng 2002/12/12
      m_CP14 = "" & rsTmp.Fields("CP14")
      
      'Add By Sindy 2010/12/28 法定期限
      m_CP07 = ""
      If IsNull(rsTmp.Fields("CP07")) = False Then
         m_CP07 = rsTmp.Fields("CP07")
      End If
      '2010/12/28 End
      
      'Add By Sindy 2011/7/12
      m_CP31 = Empty
      If IsNull(rsTmp.Fields("CP31")) = False Then
         m_CP31 = rsTmp.Fields("CP31")
      End If
      'Add By Sindy 2011/3/9
      ' 是否電子送件
      textCP118 = Empty
      If IsNull(rsTmp.Fields("CP118")) = False Then
         textCP118 = rsTmp.Fields("CP118")
      End If
      SetCPFieldOldData "CP118", textCP118, 0
      
      ' 是否出名
      textCP22 = Empty
      If IsNull(rsTmp.Fields("CP22")) = False Then
         textCP22 = rsTmp.Fields("CP22")
      End If
      SetCPFieldOldData "CP22", textCP22, 0
      
      ' 發文日(預設為系統日)
        'Modify By Cheng 2004/02/04
        '若發文日欄位無值時, 預設系統日
'      textCP27 = TAIWANDATE(SystemDate())
      If Me.textCP27.Text = "" Then
         Me.textCP27.Text = strSrvDate(2)
      End If
      'End
      strCP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         strCP27 = DBDATE(rsTmp.Fields("CP27"))
      End If
      SetCPFieldOldData "CP27", strCP27, 1
      'ADD BY SONIA 2014/11/6 電子送件案預設發文日為承辦人發文日CP85
      If textCP118 = "Y" Then
         textCP27 = TAIWANDATE(rsTmp.Fields("CP85"))
      End If
      'END  2014/11/6
      
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = rsTmp.Fields("CP44")
      End If
      SetCPFieldOldData "CP44", textCP44, 0
      ' 彼所案號
'      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         textTM45 = rsTmp.Fields("CP45")
'         strCP45 = rsTmp.Fields("CP45")
      End If
'      SetCPFieldOldData "CP45", strCP45, 0
      SetCPFieldOldData "CP45", textTM45, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      'add by nickc 2006/01/27
      'm_CP110 = CheckStr(rsTmp.Fields("cp110"))
      'SetCPFieldOldData "CP110", m_CP110, 0
      'Modify By Sindy 2010/9/20
      If m_CP110 = "" Then m_CP110 = CheckStr(rsTmp.Fields("cp110"))
      If m_CP110 = "" And m_CP10 = "301" And m_TM10 = "000" Then m_CP110 = "94007,81040" 'Add By Sindy 2016/8/31 變更(301)時,出名代理人預設為94007.林景郁和81040.閻啟泰
      SetCPFieldOldData "CP110", CheckStr(rsTmp.Fields("cp110")), 0
      '2010/9/20 End
      
      ' 代理人
      ClearAgentList
      'add by nickc 2008/03/26 若是原先有，也要加入
      If textCP44.Text <> "" Then
         If PUB_GetAgentName(m_TM01, textCP44, strTempName) Then
            strCP44 = strTempName
         Else
            strCP44 = ""
         End If
         AddAgent textCP44, strCP44
      End If
      
        'Modify By Cheng 2004/02/20
'      strSubSQL = "SELECT * FROM CaseProgress " & _
'                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
'                        "CP02 = '" & m_TM02 & "' AND " & _
'                        "CP03 = '" & m_TM03 & "' AND " & _
'                        "CP04 = '" & m_TM04 & "' AND " & _
'                        "CP09 <> '" & m_CP09 & "' "
      strSubSQL = "SELECT CP44, Max(CP27||CP09) FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null Group By CP44 Order By 2 Desc, 1 "
        'End
      rsSubTmp.CursorLocation = adUseClient
      rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSubTmp.RecordCount > 0 Then
         rsSubTmp.MoveFirst
         ' 依序將代理人加入到系統串列中
         Do While rsSubTmp.EOF = False
            If IsNull(rsSubTmp.Fields("CP44")) = False Then
               'Modify By Cheng 2002/07/09
'               strCP44 = GetFAgentName(rsSubTmp.Fields("CP44"))
'               AddAgent rsSubTmp.Fields("CP44"), GetFAgentName(rsSubTmp.Fields("CP44"))
               If PUB_GetAgentName(m_TM01, rsSubTmp.Fields("CP44"), strTempName) Then
                  strCP44 = strTempName
               Else
                  strCP44 = ""
               End If
               AddAgent rsSubTmp.Fields("CP44"), strTempName
            End If
            rsSubTmp.MoveNext
         Loop
      End If
      rsSubTmp.Close
        ' 從系統串列中取得所有代理人並放入Combo Box中
        For nIndex = 0 To m_AgentCount - 1
           'Modify By Cheng 2002/09/18
'            textCP44.AddItem m_AgentList(nIndex).aiName
           textCP44.AddItem m_AgentList(nIndex).aiCode
        Next nIndex
        ' 設定顯示為第一筆
        If textCP44.ListCount > 0 Then
           textCP44.ListIndex = 0
           textCP44_Validate False
        End If
    'add by nick 2004/08/12 發文規費
     If IsNull(rsTmp.Fields("CP17")) = False And textCP84.Enabled = True Then
         m_CP84 = CheckStr(rsTmp.Fields("CP17"))
         '2012/8/3 add by sonia 若有銷帳則要扣除銷帳規費
         If Val("" & rsTmp.Fields("CP77")) <> 0 Then
            If GetCP77Detail(m_CP09, m_Fee, m_Official) = True Then
               m_CP84 = m_CP84 - m_Official
            End If
         End If
         '2012/8/3 end
         textCP84.Text = m_CP84
     End If
     
     'Added by Morgan 2012/9/6 電子送件發文規費預設為承辦人已輸入的金額
      If rsTmp.Fields("cp118") = "Y" Then
         textCP84 = Val("" & rsTmp.Fields("cp84"))
      End If
      'end 2012/9/6
      'Add by Amy 2020/01/13 電子送件一率自動扣款(A)若超過3點半發文則須人工輸入是否當日扣款
      If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
         txtPayToday = ""
         If textCP118 = "Y" Then
            'Modify by Amy 2020/08/11 發文日小於系統日,電子送件是否當日扣款設N;發文日為當天且3點半前才設Y(原只判斷3點半)
            If Val(textCP27) < strSrvDate(2) Then
               txtPayToday = "N"
            ElseIf Val(textCP27) = strSrvDate(2) And Val(ServerTime) <= 153000 Then
               txtPayToday = "Y"
            End If
            'end 2020/08/11
         End If
      End If
      'end 2020/01/13
      textCP27.Tag = textCP27.Text 'Add By Sindy 2020/8/12
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    'Add By Cheng 2003/01/14
    Dim strTM23Nation As String
    'add by nickc 2006/01/27
   Dim tm(1 To 4) As String
   
   ' 先清除商標基本檔或服務業務基本檔欄位串列
   ClearTMSPFieldList
   ' 先清除案件進度檔欄位串列
   ClearCPFieldList
   
   ' 先取得本所案號
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 本所案號
      If IsNull(rsTmp.Fields("CP01")) = False Then: m_TM01 = rsTmp.Fields("CP01")
      If IsNull(rsTmp.Fields("CP02")) = False Then: m_TM02 = rsTmp.Fields("CP02")
      If IsNull(rsTmp.Fields("CP03")) = False Then: m_TM03 = rsTmp.Fields("CP03")
      If IsNull(rsTmp.Fields("CP04")) = False Then: m_TM04 = rsTmp.Fields("CP04")
   End If
   rsTmp.Close
   
   ' 取得國家代碼
   m_TM10 = GetNationNo(m_TM01, m_TM02, m_TM03, m_TM04)

   ' 本所案號
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   
   ' 收文號
   textCP09 = m_CP09
    Select Case m_TM01
    Case "T", "FCT", "TF", "TS"
        Me.Label10.Visible = True
        Me.textTM05_1.Visible = True
        Me.Label9.Visible = False
        Me.textTM05.Visible = False
        Me.Label8.Visible = False
        Me.textTM06.Visible = False
        Me.Label7.Visible = False
        Me.textTM07.Visible = False
    Case Else
        Me.Label10.Visible = False
        Me.textTM05_1.Visible = False
        Me.Label9.Visible = True
        Me.textTM05.Visible = True
        Me.Label8.Visible = True
        Me.textTM06.Visible = True
        Me.Label7.Visible = True
        Me.textTM07.Visible = True
    End Select
   ' 取得案件進度檔的欄位
   QueryCaseProgress
   
   'add by nickc 2006/01/27
   tm(1) = m_TM01
   tm(2) = m_TM02
   tm(3) = m_TM03
   tm(4) = m_TM04
   'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
   m_QSP = False
   
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "T", "TF", "FCT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
        'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
         m_QSP = True
   End Select
   
   'Add By Sindy 2021/1/15 T發文所有程式,台灣案鎖住畫面上之CP44,不可輸入
   If m_TM10 = "000" Then
      textCP44.Enabled = False
   End If
   '2021/1/15 END
   
   'Modify By Sindy 2012/7/26
   '台灣案才需顯示出名代理人
   lstNameAgent.Clear
   If m_TM10 = "000" Then
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
      'Modify by Amy 2021/12/23 改Form2.0,bForm2設True
      PUB_SetOurAgent lstNameAgent, tm(), m_CP110, , True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2012/7/26 End
   
   '910723 Sieg
   '申請人
   If IsEmptyText(m_TM23) = False Then
      strExc(0) = ""
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetCustomer(m_TM23, strExc(0)) Then
      If ClsPDGetCustomer(m_TM23, strExc(0)) Then
         textTM23 = strExc(0)
      End If
   End If
   
   'add by nickc 2007/01/03
   If IsEmptyText(m_TM78) = False Then
      strExc(0) = ""
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetCustomer(m_TM78, strExc(0)) Then
      If ClsPDGetCustomer(m_TM78, strExc(0)) Then
         textTM78 = strExc(0)
      End If
   End If
   If IsEmptyText(m_TM79) = False Then
      strExc(0) = ""
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetCustomer(m_TM79, strExc(0)) Then
      If ClsPDGetCustomer(m_TM79, strExc(0)) Then
         textTM79 = strExc(0)
      End If
   End If
   If IsEmptyText(m_TM80) = False Then
      strExc(0) = ""
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetCustomer(m_TM80, strExc(0)) Then
      If ClsPDGetCustomer(m_TM80, strExc(0)) Then
         textTM80 = strExc(0)
      End If
   End If
   If IsEmptyText(m_TM81) = False Then
      strExc(0) = ""
      'edit by nickc 2007/02/06 不用 dll 了
      'If objPublicData.GetCustomer(m_TM81, strExc(0)) Then
      If ClsPDGetCustomer(m_TM81, strExc(0)) Then
         textTM81 = strExc(0)
      End If
   End If
      
   'add by nickc 2007/01/03  取得催審期限的日期
   textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
   m_textUargeDate = textUargeDate
   textCP27.Tag = textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   
   ' 大約?可接獲回音(欄位)
   textCF09 = Empty
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF09")) = False Then
         textCF09 = rsTmp.Fields("CF09")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
    'Add By Cheng 2003/01/14
    '若案件性質為變更(301)且台灣案申請人國籍非台灣時, 顯示變更種類選項
    '2007/6/7 加減縮商品
    '2008/11/18 modify by sonia 取消減縮商品313,T-156581
    If m_CP10 = "301" Then
        strTM23Nation = Empty
        If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
        'Modify By Cheng 2003/02/27
        '申請國家為台灣
'        If m_TM23 < "010" And strTM23Nation >= "010" Then
        'Modify By Cheng 2004/02/19
        '系統類別為"T"的案件
'        If m_TM10 < "010" And strTM23Nation >= "010" Then
        If m_TM01 = "T" And m_TM10 < "010" And strTM23Nation >= "010" Then
        'End
            Me.Frame1.Visible = True
        End If
    End If
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/17 若已經從基本檔抓出來，就不重抓
   If Trim(textPrint) = "" Then
        textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   'Add By Sindy 2025/8/11 檢查卷宗區是否已有承辦放入之CUS,若有,系統不產出定稿
   If PUB_CPPChkFileExists(m_CP09, "cus") = True Then
      textPrint = "N"
   End If
   '2025/8/11 END
   
   'Add By Sindy 2011/10/28 T內商000台灣案所有案件性質加電子送件功能
   'Modify by Amy 2020/01/23 +是否電子送件
   lblPayToday.Visible = False
   txtPayToday.Visible = False
   If m_TM01 = "T" And m_TM10 = "000" Then
      Label43.Visible = True
      textCP118.Visible = True
      If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
        lblPayToday.Visible = True
        txtPayToday.Visible = True
      End If
   'end 2020/01/13
   Else
      Label43.Visible = False
      textCP118.Visible = False
   End If
   '2011/10/28 End
   
   Call PUB_TCaseEFeeRemind(m_CP09) 'Add By Sindy 2016/5/9 內商電子收文請款提醒訊息
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
    'Add By Cheng 2002/07/18
   Set frm020102_07 = Nothing
End Sub

'add by nickc 2006/01/27
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/5 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify by Amy 2021/12/23 改Form2.0,使用PUB_Num2Id會錯
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         'end 2021/12/23

         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      textCP22 = ""
   Else
      textCP22 = "N"
   End If
   'Add By Sindy 2015/7/22
   If textCP118 = "Y" And textCP22 = "N" Then
      Cancel = True
      MsgBox "電子送件時不可為不出名!!!", vbExclamation, "資料檢核"
      lstNameAgent.SetFocus
   End If
   '2015/7/22 END
End Sub

'Add By Sindy 2011/10/28
Private Sub textCP118_GotFocus()
   TextInverse textCP118
   CloseIme
End Sub

'Add By Sindy 2011/10/28
Private Sub textCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

' 是否出名
Private Sub textCP22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'edit by nickc 2006/01/27
' 是否出名
'Private Sub textCP22_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP22) = False Then
'      Select Case textCP22
'         Case " ", "N":
'         Case Else
'            Cancel = True
'            strTit = "資料檢核"
'            strMsg = "只可輸入空白或N"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textCP22_GotFocus
'      End Select
'   End If
'End Sub

' 發文日
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP27) = False Then
      ' 發文日日期不正確
      If CheckIsTaiwanDate(textCP27, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' 發文日日期不可超過系統日
      'edit by nick 2006/06/22 系統日加一天
      'If Val(DBDATE(textCP27)) > Val(DBDATE(SystemDate())) Then
      If Val(DBDATE(textCP27)) > Val(DBDATE(PUB_GetWorkDay(2))) Then
         Cancel = True
         strTit = "資料檢核"
         'edit by nick 2006/06/22
         'strMsg = "發文日不可超過系統日"
         strMsg = "發文日不可超過系統日加一天"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      'Add by Amy 2020/01/13 當發文日有改時,電子送件案要人工輸入是否當日扣款
      If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
        If textCP27.Tag <> textCP27.Text Then
            textCP27.Tag = textCP27.Text
            If textCP118 = "Y" Then
                txtPayToday.Text = ""
            End If
        End If
      End If
      'end 2020/01/13
      'add by nickc 2007/01/03 重新算催審期限
      'Modified by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
      'If textUargeDate = m_textUargeDate Then
      If textCP27.Tag <> textCP27.Text Then
        textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
        m_textUargeDate = textUargeDate
      End If
      textCP27.Tag = textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   End If
EXITSUB:
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCP44_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCP44.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 代理人
Private Sub textCP44_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Cheng 2002/07/09
   Dim strTempName As String
   
   Cancel = False
   'Add By Cheng 2002/03/08
   'Modify By Sindy 2012/3/29 TD申請時皆在台灣申請不須控管CF代理人
   If m_TM10 <> 台灣國家代號 And m_TM01 <> "TD" Then
      If Len(Me.textCP44.Text) <= 0 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Cancel = True
         Exit Sub
      End If
   End If
   
   If textCP44.ListIndex >= 0 Then
      textCP44 = m_AgentList(textCP44.ListIndex).aiCode
   End If
   'Add By Cheng 2002/12/03
   '若有輸入代理人則將代碼補滿9碼
   If Me.textCP44.Text <> "" Then Me.textCP44.Text = Left(Me.textCP44.Text & "000000000", 9)
   
   If IsEmptyText(textCP44) = False Then
      'Modify By Cheng 2002/07/09
'      textCP44_2 = GetFAgentName(textCP44)
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      'If PUB_GetAgentName(m_TM01, Me.textCP44.Text, strTempName) Then
      If PUB_GetAgentNameAndState(m_TM01, Me.textCP44.Text, strTempName) Then
         textCP44_2 = strTempName
      Else
         textCP44_2 = ""
        'edit by nick 2004/07/22
        If strTempName <> "" Then
             Cancel = True
             Exit Sub
        End If
      End If
      If IsEmptyText(textCP44_2) = True Then
         'Modify By Sindy 2012/3/29 TD申請時皆在台灣申請不須控管CF代理人
         If m_TM01 <> "TD" Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "代理人不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP44_GotFocus
         End If
      Else
         ' 依所選擇的代理人找出案件進度檔中其收文日最大的一筆其彼所案號更新到畫面上的彼所案號欄位
         strSql = "SELECT CP45 FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP44 = '" & textCP44 & "' AND " & _
                        "CP05 IN (SELECT MAX(CP05) FROM CASEPROGRESS " & _
                                 "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                       "CP02 = '" & m_TM02 & "' AND " & _
                                       "CP03 = '" & m_TM03 & "' AND " & _
                                       "CP04 = '" & m_TM04 & "' AND " & _
                                       "CP44 = '" & textCP44 & "')"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CP45")) = False Then
               textTM45 = rsTmp.Fields("CP45")
            End If
         End If
         rsTmp.Close
      End If
   End If
   Set rsTmp = Nothing
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

Private Sub textCP84_GotFocus()
   InverseTextBox textCP84
End Sub

Private Sub textCP84_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
If IsEmptyText(textCP84) = False Then
    If IsNumeric(textCP84) = False Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "請輸入數字"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP84_GotFocus
    Else
        textCP84.Text = Trim(Val(textCP84.Text))
    End If
End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
  
End Sub

' 列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         'edit by nickc 2006/06/29
         'Case " ", "N":
         Case "N", "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            'edit by nickc 2006/06/29
            'strMsg = "只可輸入空白或N"
            strMsg = "只可輸入 N 或 1-3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
   Dim strCP64 As String
   
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' 代理人
   If IsEmptyText(textCP44) = False Then
      SetCPFieldNewData "CP44", textCP44 & String(9 - Len(textCP44), "0")
      'add by nickc 2008/02/22
      m_CP44New = textCP44 & String(9 - Len(textCP44), "0")
   Else
      SetCPFieldNewData "CP44", textCP44
      'add by nickc 2008/02/22
      m_CP44New = textCP44
   End If
   ' 彼所案號
   SetCPFieldNewData "CP45", textTM45
   ' 91.09.02 modify by louis
   ' 進度備註
   'SetCPFieldNewData "CP64", textCP64
   strCP64 = textCP64
'edit by nickc 2006/01/27
'   If IsEmptyText(textAgName) = False Then
'      strCP64 = strCP64 & "," & "本所出名代理人:" & textAgName
'   End If
   SetCPFieldNewData "CP64", strCP64
   
   ' 是否出名
   SetCPFieldNewData "CP22", textCP22
   'add by nickc 2006/01/27
   SetCPFieldNewData "CP110", m_CP110
   'Add By Sindy 2011/3/9
   ' 是否電子送件
   SetCPFieldNewData "CP118", textCP118
   
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
End Sub

' 更新案件進度檔
'Modify By Cheng 2002/11/06
'Private Sub OnUpdateCaseProgress()
Private Function OnUpdateCaseProgress() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnUpdateCaseProgress = True

   ' 更新案件進度檔
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = NULL "
            Else
               ' 91.03.25 modify by louis (單引號)
               'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
               strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_CPList(nIndex).fiName & " = " & m_CPList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateCaseProgress = False
End Function

'Modify By Cheng 2002/11/06
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nIndex As Integer
   Dim strNP08 As String
   Dim strNP07 As String
   Dim strNP22 As String
   Dim objCopyCP As ClsCopyCP
   Dim strCP09 As String 'Add By Sindy 2010/6/10
   Dim strCP44 As String 'Add By Sindy 2010/11/5
   Dim bolSysDt As Boolean 'Add By Sindy 2010/12/28
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler

cnnConnection.BeginTrans
   
   'Add By Sindy 2010/12/28
   '非台灣案發文, 法定期限有值且為系統日或者過期時, 顯示訊息, 但仍可發文
   '上述情形的收達期限或提申期限都管制為系統日期
   bolSysDt = False
   If m_TM10 >= "010" Then
      If Trim(m_CP07) <> "" Then
         If Val(m_CP07) = Val(strSrvDate(1)) Then
            MsgBox "此案件已屆法定期限, 請注意！", vbExclamation + vbOKOnly
            bolSysDt = True
         ElseIf Val(m_CP07) < Val(strSrvDate(1)) Then
            MsgBox "此案件已逾法定期限, 請注意！", vbExclamation + vbOKOnly
            bolSysDt = True
         End If
      End If
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Add By Cheng 2003/05/30
    '更新商標基本檔
    If m_TM01 = "T" Or m_TM01 = "TF" Or m_TM01 = "FCT" Then
        strSql = "Update Trademark Set TM67='" & ChgSQL(Me.textTM67.Text) & "' Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
        cnnConnection.Execute strSql
        'add by nickc 2006/11/17
        If textPrint <> "N" Then
            strSql = "Update Trademark Set TM77='" & textPrint & "' Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute strSql
        End If
    Else
        'add by nickc 2006/11/17
        If textPrint <> "N" Then
            strSql = "Update ServicePractice Set SP72='" & textPrint & "' Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute strSql
        End If
    End If
    'add by nickc 2006/01/26
    If m_CU112 <> "" Then
        'Modify By Sindy 2011/2/22
        'strSql = "update trademark set tm24='" & ChgSQL(Pub_RplCu112(m_TM24, m_CU112)) & " where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
        strSql = "update trademark set tm24='" & ChgSQL(Pub_RplCu112(m_TM24, m_CU112, m_TM23)) & " where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
    End If
   ' 更新案件進度檔
    'Modify By Cheng 2002/11/06
'   OnUpdateCaseProgress
   If OnUpdateCaseProgress = False Then GoTo ErrorHandler
   ' 若有輸入催審期限時, 新增一筆催審的記錄到下一程序檔
   If IsEmptyText(textUargeDate) = False Then
      strNP07 = "305"
      strNP22 = GetNextProgressNo()
        'Modify By Cheng 2002/12/12
        '下一程序智權人員欄位應掛承辦人員非程序人員
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                        DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strUserNum & "'," & strNP22 & ")"
      'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                        DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & m_CP14 & "'," & strNP22 & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                        PUB_GetWorkDay1(textUargeDate, True) & "," & DBDATE(textUargeDate) & ",'" & m_CP14 & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 收達不印接洽結案單
'      '92.6.8 SONIA 加 言詞辯論, 準備程序
      Select Case strNP07
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
         Case "102", "105", "702", "708", "305", "998", "997"
         Case Else:
            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有定義代理人收達天數時, 新增一筆收達的記錄到下一程序檔
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CF23")) = False Then
         strNP07 = "997"
         'Add By Sindy 2010/12/28
         '非台灣案發文, 法定期限有值且為系統日或者過期時, 收達期限或提申期限都管制為系統日期
         If bolSysDt = True Then
            strNP08 = strSrvDate(1)
         Else
         '2010/12/28 End
            strNP08 = DBDATE(textCP27)
           'Modify By Cheng 2003/09/01
   '         strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP08)), Val(DBMONTH(strNP08)), Val(DBDAY(strNP08)) + Val(rsTmp.Fields("CF23")))))
            strNP08 = DBDATE(DateAdd("d", Val(rsTmp.Fields("CF23")), ChangeWStringToWDateString(DBDATE(strNP08))))
            'Add By Sindy 2019/6/11 檢查期限是否正確
            strNP08 = PUB_T997998LimitDate(strNP08, m_CP07, 1)
            '2019/6/11 END
         End If
         strNP22 = GetNextProgressNo()
         'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            strNP08 & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
         cnnConnection.Execute strSql
         
         'Add By Sindy 2022/6/7 判斷案件國家收費表內有設定提申期限(天)CF11，要加掛提申(998)期限
         If IsNull(rsTmp.Fields("CF11")) = False Then
            strNP07 = "998"
            '非台灣案發文, 法定期限有值且為系統日或者過期時, 收達期限或提申期限都管制為系統日期
            If bolSysDt = True Then
               strNP08 = strSrvDate(1)
            Else
               strNP08 = DBDATE(textCP27)
               strNP08 = DBDATE(DateAdd("d", Val(rsTmp.Fields("CF11")), ChangeWStringToWDateString(DBDATE(strNP08))))
               '檢查期限是否正確
               strNP08 = PUB_T997998LimitDate(strNP08, m_CP07, 1)
            End If
            strNP22 = GetNextProgressNo()
            '本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                               PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
            cnnConnection.Execute strSql
         End If
         '2022/6/7 END
         
         ' 延展, 使用宣誓, 刊登廣告, 繳年費, 收達不印接洽結案單
'         '92.6.8 SONIA 加 言詞辯論, 準備程序
         Select Case strNP07
'            Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
            Case "102", "105", "702", "708", "305", "998", "997"
            Case Else:
               ' 列印國內案件接洽及結案記錄單
'               g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
                'Add By Cheng 2004/04/08
                '新增列印接洽結案單資料
                pub_AddressListSN = pub_AddressListSN + 1
                PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
         End Select
      End If
   End If
   rsTmp.Close
   
   'add by nick 2004/08/12 更新實際發文規費
   If textCP84.Enabled = True Then
      strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若該筆記錄是母案時, 同時對所有的子案做新增案件進度檔的工作
   If m_TM01 = "TF" And m_TM03 = "0" And m_TM04 = "00" Then
      Set objCopyCP = New ClsCopyCP
        'Modify By Cheng 2002/11/06
'      objCopyCP.CopyCaseProgress m_CP09
      If objCopyCP.CopyCaseProgress(m_CP09) = False Then GoTo ErrorHandler
      Set objCopyCP = Nothing
   End If
   
   'add by nick 2004/09/27 存公司負責人英文名稱
   'edit by nick 2004/10/07
   'If m_CU103 <> "" And m_TM01 <> "FCT" Then
   'edit by nickc 2006/01/20
   'If (m_CU103 <> "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) <> "") And m_TM01 <> "FCT" Then
   'edit by nickc 2007/08/10
   'If (m_CU103 <> "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) <> "" Or m_CU112 <> "") And m_TM01 <> "FCT" Then
   'Modify By Sindy 2012/10/31 +SeekCu10(1),SeekCu10(2),SeekCu10(3),SeekCu10(4),SeekCu10(5)
   If (SeekCu103(1) <> "" Or (SeekCu05(1) & SeekCu88(1) & SeekCu89(1) & SeekCu90(1)) <> "" Or SeekCu112(1) <> "" Or (SeekCu39(1) & SeekCu40(1) & SeekCu41(1)) <> "" Or SeekCu10(1) <> "") And m_TM01 <> "FCT" Then
            'edit by nickc 2006/01/20
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "'  Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' "
            'edit by nickc 2007/08/10
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "',cu112='" & ChgSQL(m_CU112) & "'  Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' "
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(1)) & "',cu05='" & ChgSQL(SeekCu05(1)) & "',cu88='" & ChgSQL(SeekCu88(1)) & "',cu89='" & ChgSQL(SeekCu89(1)) & "',cu90='" & ChgSQL(SeekCu90(1)) & "',cu112='" & ChgSQL(SeekCu112(1)) & "',cu39='" & ChgSQL(SeekCu39(1)) & "',cu40='" & ChgSQL(SeekCu40(1)) & "',cu41='" & ChgSQL(SeekCu41(1)) & "',cu10='" & ChgSQL(SeekCu10(1)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(1)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   'add by nickc 2007/08/10 加多申請人也要
   If (SeekCu103(2) <> "" Or (SeekCu05(2) & SeekCu88(2) & SeekCu89(2) & SeekCu90(2)) <> "" Or SeekCu112(2) <> "" Or (SeekCu39(2) & SeekCu40(2) & SeekCu41(2)) <> "" Or SeekCu10(2) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(2)) & "',cu05='" & ChgSQL(SeekCu05(2)) & "',cu88='" & ChgSQL(SeekCu88(2)) & "',cu89='" & ChgSQL(SeekCu89(2)) & "',cu90='" & ChgSQL(SeekCu90(2)) & "',cu112='" & ChgSQL(SeekCu112(2)) & "',cu39='" & ChgSQL(SeekCu39(2)) & "',cu40='" & ChgSQL(SeekCu40(2)) & "',cu41='" & ChgSQL(SeekCu41(2)) & "',cu10='" & ChgSQL(SeekCu10(2)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM78), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM78), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(2)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM78), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM78), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(3) <> "" Or (SeekCu05(3) & SeekCu88(3) & SeekCu89(3) & SeekCu90(3)) <> "" Or SeekCu112(3) <> "" Or (SeekCu39(3) & SeekCu40(3) & SeekCu41(3)) <> "" Or SeekCu10(3) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(3)) & "',cu05='" & ChgSQL(SeekCu05(3)) & "',cu88='" & ChgSQL(SeekCu88(3)) & "',cu89='" & ChgSQL(SeekCu89(3)) & "',cu90='" & ChgSQL(SeekCu90(3)) & "',cu112='" & ChgSQL(SeekCu112(3)) & "',cu39='" & ChgSQL(SeekCu39(3)) & "',cu40='" & ChgSQL(SeekCu40(3)) & "',cu41='" & ChgSQL(SeekCu41(3)) & "',cu10='" & ChgSQL(SeekCu10(3)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM79), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM79), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(3)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM79), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM79), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(4) <> "" Or (SeekCu05(4) & SeekCu88(4) & SeekCu89(4) & SeekCu90(4)) <> "" Or SeekCu112(4) <> "" Or (SeekCu39(4) & SeekCu40(4) & SeekCu41(4)) <> "" Or SeekCu10(4) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(4)) & "',cu05='" & ChgSQL(SeekCu05(4)) & "',cu88='" & ChgSQL(SeekCu88(4)) & "',cu89='" & ChgSQL(SeekCu89(4)) & "',cu90='" & ChgSQL(SeekCu90(4)) & "',cu112='" & ChgSQL(SeekCu112(4)) & "',cu39='" & ChgSQL(SeekCu39(4)) & "',cu40='" & ChgSQL(SeekCu40(4)) & "',cu41='" & ChgSQL(SeekCu41(4)) & "',cu10='" & ChgSQL(SeekCu10(4)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM80), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM80), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(4)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM80), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM80), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(5) <> "" Or (SeekCu05(5) & SeekCu88(5) & SeekCu89(5) & SeekCu90(5)) <> "" Or SeekCu112(5) <> "" Or (SeekCu39(5) & SeekCu40(5) & SeekCu41(5)) <> "" Or SeekCu10(5) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(5)) & "',cu05='" & ChgSQL(SeekCu05(5)) & "',cu88='" & ChgSQL(SeekCu88(5)) & "',cu89='" & ChgSQL(SeekCu89(5)) & "',cu90='" & ChgSQL(SeekCu90(5)) & "',cu112='" & ChgSQL(SeekCu112(5)) & "',cu39='" & ChgSQL(SeekCu39(5)) & "',cu40='" & ChgSQL(SeekCu40(5)) & "',cu41='" & ChgSQL(SeekCu41(5)) & "',cu10='" & ChgSQL(SeekCu10(5)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM81), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM81), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(5)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM81), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM81), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If
   
'   'Add By Sindy 2010/6/10
'   '由巨京代理之變更案，發文變更時，請同時假收文變更代理人
'   '變更案的收文與發文時間同變更案
'   '前次發文代理人不是Y52269
'   strSql = "SELECT CP44, Max(CP27||CP09) FROM CaseProgress " & _
'                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
'                        "CP02 = '" & m_TM02 & "' AND " & _
'                        "CP03 = '" & m_TM03 & "' AND " & _
'                        "CP04 = '" & m_TM04 & "' AND " & _
'                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null And CP27 Is Not Null Group By CP44 Order By 2 Desc, 1 "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 0 Then strCP44 = ""
'   If intI = 1 Then strCP44 = Trim(RsTemp.Fields("CP44"))
'   If Left(strCP44, 6) <> "Y52269" And Left(Trim(textCP44.Text), 6) = "Y52269" Then
'      If m_TM01 = "T" And m_TM10 = "020" And m_CP10 = "301" Then
'         strCP09 = AutoNo("B", 6)
'         '新增一筆B類
'         strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp43,cp44,cp45,cp64) " & _
'                        "values (" & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & _
'                        "," & CNULL(m_TM04) & "," & CNULL(DBDATE(textCP27)) & "," & CNULL(strCP09) & ",301," & _
'                        CNULL(m_CP12) & "," & CNULL(m_CP13) & "," & CNULL(strUserNum) & ",'N','N'," & CNULL(DBDATE(textCP27)) & ",'N'," & _
'                        CNULL(m_CP09) & ",'" & textCP44.Text & "','" & strCP45 & "','變更代理人')"
'         cnnConnection.Execute strSql
'         '新增變更事項檔
'         strSql = "insert into ChangeEvent(CE01,CE55) values('" & strCP09 & "','V')"
'         cnnConnection.Execute strSql
'      End If
'   End If
   'Modify By Sindy 2012/3/23
   Call PUB_T020InsB301(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textCP44, m_TM10, m_CP10, textCP27, m_CP12, m_CP13, textTM45)
   
   'Add by Amy 2020/01/13
   If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
        strSql = ""
        If textCP118 = "Y" And Val(textCP84) > 0 Then
           If txtPayToday <> "" Then
              strSql = ",CP118 = 'A' "
              If txtPayToday = "Y" Then
                  strSql = strSql & ",CP152 = " & CompWorkDay(2, DBDATE(textCP27))
              Else
                  strSql = strSql & ",CP152 =" & CompWorkDay(3, DBDATE(textCP27))
              End If
              strSql = "Update CaseProgress Set " & Mid(strSql, 2) & " Where CP09 = '" & m_CP09 & "' "
              cnnConnection.Execute strSql
           End If
        End If
   End If
   'end 2020/01/13
   'Add By Sindy 2011/3/9 若為電子送件則自動設定為不經發文室
   '以防動作為重新發文, 所以一併把發文室相關欄位清空
   If textCP118.Visible = True And textCP118 = "Y" Then
      strSql = "Update CaseProgress Set CP123=null" & _
                                                          ",CP124=null" & _
                                                          ",CP125=null" & _
                                                          ",CP28=null" & _
                                                          ",CP131=null" & _
                                                          ",CP132=null" & _
                   " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
      
   'Add by Sindy 98/3/24
   If m_TM10 = "000" Then
      'Modify By Sindy 2009/04/24
      'PUB_UpdateDispatch m_CP09s, m_CP123s
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
   End If
   
   'Add by Sindy 2012/10/4 外->台,智權人員是葉雪貞及巨京,發文規費和收文規費不相同時,系統自動更改進度檔內規費費用及計算點數
   'Modified by Lydia 2015/10/16 + m_CP84
   Call PUB_TSendUpdateCP1718(m_CP09, textCP84, textPrint, m_TM10, m_CP13, m_CP84)
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Add By Sindy 2019/12/20 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = m_CP09
      PUB_AddLetterProgress strLD18, 0, IIf(textPrint = "N", False, True), "", False, m_TM23, m_CP10, m_TM44
   End If
   '2019/12/20 END
   Call PUB_UpdateLP19_T(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textCP27) 'Add by Sindy 2020/2/12 收據/回執設定
   
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
   
   Set rsTmp = Nothing
'Add By Cheng 2002/11/06
cnnConnection.CommitTrans

     'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
    PUB_CheckEMail m_CP44New, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22
    
OnSaveData = True
Exit Function

ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

' 催審期限
Private Sub textUargeDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textUargeDate) = False Then
      If CheckIsTaiwanDate(textUargeDate, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "催審期限日期不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textUargeDate_GotFocus
      End If
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/23檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
    End If
   'add by nickc 2008/05/01
   If IsDebt(m_TM10, textCP09) Then
        strTit = "警告！禁止發文！"
        strMsg = "未收款且無 預定收款日 請轉告智權同仁！！"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        GoTo EXITSUB
   End If
   
   'Add By Sindy 2012/4/17
   If m_CP10 = "301" Then
      ' 檢查變更事項檔是否有資料
      If IsChangeEventExist(m_CP09) = False Or m_blnClkChgButton = False Then
         strTit = "檢核資料"
         strMsg = "請輸入變更事項!!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         cmdMod.SetFocus
         GoTo EXITSUB
      End If
   Else
      'Modify By Sindy 2022/2/7
      'T和FCT的減縮商品313發文：不必再檢查一定要按變更事項的按鈕。
      If m_CP10 <> "313" Then
      '2022/2/7 END
         If m_blnClkChgButton = False Then
            MsgBox "請輸入變更事項!!!", vbExclamation + vbOKOnly
            Me.cmdMod.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   ' 發文日
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' 90.06.18 台灣不需輸入代理人
   'Modify By Sindy 2012/3/29 TD申請時皆在台灣申請不須控管CF代理人
   If m_TM10 > "010" And m_TM01 <> "TD" Then
      ' 代理人
      If IsEmptyText(textCP44) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入代理人"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44.SetFocus
         GoTo EXITSUB
      End If
   End If
    'Add By Cheng 2003/01/14
    '若有顯示變更種類, 一定要點選
    If Me.Frame1.Visible = True Then
        '若沒點選變更種類
        'Modify By Cheng 2003/05/23
'        If Me.Option1(0).Value = False And Me.Option1(1).Value = False Then
        'Modify By Sindy 2015/8/10
        'If Me.Option1(0).Value = False And Me.Option1(1).Value = False And Me.Option1(2).Value = False Then
        If Me.Option1(0).Value = False And Me.Option1(1).Value = False Then
        '2015/8/10 END
            strTit = "檢核資料"
            strMsg = "請點選變更種類!!!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
        End If
    End If
    
   'Add By Sindy 2011/01/06
   '內商(TS)申請人1或FC代理人至少要輸入一個
   '其他的一定要輸入申請人1
   If m_TM01 = "TS" Then
        If textTM23 = "" And m_TM44 = "" Then
            MsgBox "申請人1或FC代理人至少要輸入一個!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   Else
        If textTM23 = "" Then
            MsgBox "申請人1不可空白!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   End If
    
    CheckDataValid = True
EXITSUB:
End Function

Private Sub textUargeDate_GotFocus()
   InverseTextBox textUargeDate
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP22_GotFocus()
   InverseTextBox textCP22
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP44_GotFocus()
   InverseTextBox textCP44
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
End Sub

'Add by Amy 2020/01/13
Private Sub txtPayToday_GotFocus()
    TextInverse txtPayToday
    CloseIme
End Sub

Private Sub txtPayToday_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
        KeyAscii = 0
        Beep
    End If
End Sub
'end 2020/01/13

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strTM23Nation As String
   Dim strSql As String
   Dim arrTM09 As Variant, strGoodsKind As String 'Add By Sindy 2010/11/12
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 案件性質為變更 2007/6/7 加減縮商品
   If m_CP10 = "301" Or m_CP10 = "313" Then
      Select Case m_TM10
         ' 申請國家為台灣
         'edit by nickc 2005/09/07
         'Case "010"
         Case "000"
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "02", strUserNum
               ' 回音
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
                        "'回音','" & "大約" & textCF09 & "後可接獲回音。')"
               '         "'" & "回音" & "','" & textCF09 & "')"
               cnnConnection.Execute strSql
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            'ElseIf textPrint = "2" Then
            ElseIf textPrint = "2" And m_CP10 = "301" Then
               ' 清除定稿例外欄位檔原有資料
               'Modify By Cheng 2003/01/14
   '            EndLetter "01", m_CP09, "06", strUserNum
               'Modify By Cheng 2003/05/23
   '            EndLetter "01", m_CP09, IIf(Me.Option1(0).Value = True, "06", "07"), strUserNum
'Modify By Sindy 2015/8/10 取消1.核駁變更
               'EndLetter "01", m_CP09, IIf(Me.Option1(0).Value = True, "06", IIf(Me.Option1(1).Value = True, "07", "08")), strUserNum
               EndLetter "01", m_CP09, IIf(Me.Option1(0).Value = True, "07", "08"), strUserNum
'               'Add By Sindy 2010/11/12
'               If Me.Option1(0).Value = True Then
'                  '1-34商品 35-45服務
'                  strGoodsKind = "商品"
'                  If Trim(m_TM09) > "" Then
'                     arrTM09 = Split(m_TM09, ",")
'                     If Val(arrTM09(0)) >= 35 And Val(arrTM09(0)) <= 45 Then
'                        strGoodsKind = "服務"
'                     End If
'                  End If
'                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
'                        "'商品或服務','" & strGoodsKind & "')"
'                  cnnConnection.Execute strSql
'               End If
'               '2010/11/12 End
'2015/8/10 END
            'add by sonia 2020/2/15
            ElseIf textPrint = "2" And m_CP10 = "313" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "06", strUserNum
               '1-34商品 35-45服務
               strGoodsKind = "商品"
               If Trim(m_TM09) > "" Then
                  arrTM09 = Split(m_TM09, ",")
                  If Val(arrTM09(0)) >= 35 And Val(arrTM09(0)) <= 45 Then
                     strGoodsKind = "服務"
                  End If
               End If
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                     "'商品或服務','" & strGoodsKind & "')"
               cnnConnection.Execute strSql
            'end 2020/2/15
            End If
         ' 申請國家為大陸, 2005/8/31 SONIA 加 馬德里
         Case "020", "238"
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                ' 清除定稿例外欄位檔原有資料
                EndLetter "01", m_CP09, "00", strUserNum
                ' 案件性質分類
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                         "'" & "案件性質分類" & "','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
                cnnConnection.Execute strSql
            End If
         Case Else
      End Select
   'Add By Sindy 2015/8/5
   '更正
   ElseIf m_CP10 = "302" Then
      Select Case m_TM10
         ' 申請國家為台灣
         Case "000"
            ' 申請人國籍非台灣
            If textPrint = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "01", strUserNum
               'Add By Sindy 2016/5/30 有費用
               If Val(m_CP16) > 0 Then
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                           "'" & "有費用" & "','及本所收費通知各乙紙')"
                  cnnConnection.Execute strSql
               End If
               '2016/5/30 END
            End If
      End Select
   '2015/8/5 END
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strTM23Nation As String
'Add By Sindy 2012/1/12
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/12 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/12
   ET01 = "01"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/12 End
   
   ' 案件性質為變更  2007/6/7 加減縮商品
   If m_CP10 = "301" Or m_CP10 = "313" Then
      Select Case m_TM10
         ' 申請國家為台灣
         'edit by nickc 2005/09/07
         'Case "010"
         Case "000"
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
'               NowPrint m_CP09, "01", "02", False, strUserNum, 0
               ET03 = "02" 'Modify By Sindy 2012/1/12
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
'Modify By Sindy 2015/8/10 取消1.核駁變更
            'ElseIf textPrint = "2" Then
            ElseIf textPrint = "2" And m_CP10 = "301" Then
'               If m_CP10 = "313" Then Me.Option1(0).Value = True   '2008/11/18 add by sonia
'2015/8/10 END
               '2005/9/23 MODIFY BY SONIA 不開 WORD
               'NowPrint m_CP09, "01", IIf(Me.Option1(0).Value = True, "06", IIf(Me.Option1(1).Value = True, "07", "08")), True, strUserNum, 0
'               NowPrint m_CP09, "01", IIf(Me.Option1(0).Value = True, "06", IIf(Me.Option1(1).Value = True, "07", "08")), False, strUserNum, 0
               'ET03 = IIf(Me.Option1(0).Value = True, "06", IIf(Me.Option1(1).Value = True, "07", "08")) 'Modify By Sindy 2012/1/12
               ET03 = IIf(Me.Option1(0).Value = True, "07", "08") 'Modify By Sindy 2015/8/10
            'add by sonia 2020/2/15
            ElseIf textPrint = "2" And m_CP10 = "313" Then
               ET03 = "06"
            'end 2020/2/15
            End If
         ' 申請國家為大陸, 2005/8/31 SONIA 加 馬德里
         Case "020", "238"
            'add by nickc 2006/06/30
            If textPrint = "1" Then
'                NowPrint m_CP09, "01", "00", False, strUserNum, 0
               ET03 = "00" 'Modify By Sindy 2012/1/12
            End If
         Case Else
      End Select
   'Add By Sindy 2015/8/5
   '更正
   ElseIf m_CP10 = "302" Then
      Select Case m_TM10
         ' 申請國家為台灣
         Case "000"
            ' 申請人國籍非台灣
            If textPrint = "2" Then
               ET03 = "01"
            'Added by Lydia 2020/03/06 申請人國籍=台灣
            ElseIf textPrint = "1" Then
               ET03 = "02"
            End If
      End Select
   '2015/8/5 END
   End If
   
   'Add By Sindy 2012/1/12
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2020/1/7 + 信函總收文號
         If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
         Else
         '2020/1/7 END
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
         End If
      Else
         'Add By Sindy 2019/12/20 + strLD18.信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
      End If
   'Add By Sindy 2021/1/5 沒有系統產出的定稿
   Else
      If strLD18 <> "" Then
         'Modify By Sindy 2025/8/15
         'Call PUB_TCaseAskIsPost(strLD18)
         textPrint = "N"
         '2025/8/15 END
      End If
   '2021/1/5 EMD
   End If
   '2012/1/12 End
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
   'add by nick 2004/08/12 發文規費，申請國家台灣才檢查
   If Me.textCP84.Enabled = True Then
      Cancel = False
      textCP84_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP84.Enabled = True And m_TM10 = "000" Then
       If Val(textCP84.Text) <> Val(m_CP84) Then
           If MsgBox("收文規費[" & Trim(Val(m_CP84)) & "] 與實際發文規費[" & Trim(Val(textCP84.Text)) & "]不同", vbOKCancel) = vbCancel Then
               textCP84_GotFocus
               Exit Function
           End If
       End If
   End If
   'edit by nickc 2006/01/27
   'If Me.textCP22.Enabled = True Then
   '   Cancel = False
   '   textCP22_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
   If Me.textCP27.Enabled = True Then
      Cancel = False
      textCP27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP44.Enabled = True Then
      Cancel = False
      textCP44_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
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
   
   If Me.textUargeDate.Enabled = True Then
      Cancel = False
      textUargeDate_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'add by nickc 2006/03/07
   'Modify By Sindy 2012/7/26
   'If lstNameAgent.Enabled = True Then
   If lstNameAgent.Visible = True Then
   '2012/7/26 End
       Cancel = False
       lstNameAgent_Validate Cancel
       If Cancel = True Then
           Exit Function
       End If
   End If

   'Add By Sindy 2016/12/20
   '檢查有設定副本收受人需提醒並新增信函副本B類收文
   m_990CP09 = ""
   If textPrint = "N" Then '不印定稿
      If PUB_ChkCC(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_990CP09) = False Then
         Exit Function
      End If
   End If
   '2016/12/20 END

    'Added by Lydia 2021/06/04 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
    If Pub_ChkACS112isNull(m_TM01, m_TM02, m_TM03, m_TM04, txtCP113) = True Then
        txtCP113.SetFocus
        txtCP113_GotFocus
        Exit Function
    End If
    'end 2021/06/04
    
   TxtValidate = True
End Function

'edit by nickc 2006/01/27
' 91.09.02 modify by louis
'Private Sub textAgName_GotFocus()
'   InverseTextBox textAgName
'   textAgName.IMEMode = 1
'End Sub
'
'' 91.09.02 modify by louis
'' 本所出名代理人
'Private Sub textAgName_Validate(Cancel As Boolean)
'   Cancel = False
'   If CheckLengthIsOK(textAgName, 10) = False Then
'      Cancel = True
'   End If
'   If Cancel = False Then: textAgName.IMEMode = 2
'End Sub

'Added by Lydia 2021/06/04
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/06/04
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

