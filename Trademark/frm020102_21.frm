VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_21 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(其它服務業務)"
   ClientHeight    =   5568
   ClientLeft      =   5736
   ClientTop       =   1968
   ClientWidth     =   9132
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5568
   ScaleWidth      =   9132
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   8490
      MaxLength       =   4
      TabIndex        =   13
      Top             =   4110
      Width           =   540
   End
   Begin VB.TextBox textCP84 
      Alignment       =   1  '靠右對齊
      Height          =   264
      Left            =   3450
      TabIndex        =   1
      Top             =   1815
      Width           =   1425
   End
   Begin VB.TextBox textSP66 
      Height          =   264
      Left            =   990
      MaxLength       =   9
      TabIndex        =   9
      Top             =   3840
      Width           =   1092
   End
   Begin VB.TextBox textSP65 
      Height          =   264
      Left            =   5220
      MaxLength       =   9
      TabIndex        =   8
      Top             =   3570
      Width           =   1092
   End
   Begin VB.TextBox textSP59 
      Height          =   264
      Left            =   990
      MaxLength       =   9
      TabIndex        =   7
      Top             =   3570
      Width           =   1092
   End
   Begin VB.TextBox textSP58 
      Height          =   264
      Left            =   5220
      MaxLength       =   9
      TabIndex        =   6
      Top             =   3300
      Width           =   1092
   End
   Begin VB.TextBox textSP64 
      Height          =   264
      Left            =   5070
      MaxLength       =   10
      TabIndex        =   22
      Top             =   1560
      Width           =   1908
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   11
      Top             =   4110
      Width           =   372
   End
   Begin VB.TextBox textWord 
      Height          =   264
      Left            =   6180
      MaxLength       =   1
      TabIndex        =   12
      Top             =   4110
      Width           =   372
   End
   Begin VB.TextBox textSP51 
      Height          =   264
      Left            =   5220
      MaxLength       =   30
      TabIndex        =   10
      Top             =   3840
      Width           =   3795
   End
   Begin VB.TextBox textSP11 
      Height          =   300
      Left            =   5640
      MaxLength       =   20
      TabIndex        =   16
      Top             =   4650
      Width           =   3375
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   4968
      TabIndex        =   24
      Top             =   24
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8244
      TabIndex        =   27
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6168
      TabIndex        =   25
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7020
      TabIndex        =   26
      Top             =   24
      Width           =   1200
   End
   Begin VB.TextBox textSP49 
      Height          =   264
      Left            =   1200
      MaxLength       =   25
      TabIndex        =   14
      Top             =   4380
      Width           =   7815
   End
   Begin VB.TextBox textSP20 
      Height          =   300
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   15
      Top             =   4650
      Width           =   1092
   End
   Begin VB.TextBox textSP21 
      Height          =   300
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   23
      Top             =   4650
      Width           =   1092
   End
   Begin VB.TextBox textSP08 
      Height          =   264
      Left            =   990
      MaxLength       =   9
      TabIndex        =   5
      Top             =   3300
      Width           =   1092
   End
   Begin VB.TextBox textCP18 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   8085
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   1815
      Width           =   852
   End
   Begin VB.TextBox textCP22 
      Height          =   264
      Left            =   5985
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1815
      Width           =   372
   End
   Begin VB.ComboBox textCP44 
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   2085
      Width           =   1716
   End
   Begin VB.TextBox textSP06 
      Height          =   270
      Left            =   1560
      MaxLength       =   60
      TabIndex        =   20
      Top             =   2685
      Width           =   7392
   End
   Begin VB.TextBox textCP27 
      Height          =   264
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1815
      Width           =   1092
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5070
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1005
      Width           =   4035
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1275
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   465
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   735
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5070
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   735
      Width           =   4035
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1005
      Width           =   2532
   End
   Begin MSForms.TextBox textSP05_1 
      Height          =   900
      Left            =   1560
      TabIndex        =   4
      Top             =   2370
      Width           =   7395
      VariousPropertyBits=   -1466941413
      MaxLength       =   140
      ScrollBars      =   2
      Size            =   "13044;1587"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP05 
      Height          =   315
      Left            =   1560
      TabIndex        =   19
      Top             =   2385
      Width           =   7395
      VariousPropertyBits=   671105051
      MaxLength       =   140
      Size            =   "13044;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP07 
      Height          =   315
      Left            =   1560
      TabIndex        =   21
      Top             =   2955
      Width           =   7395
      VariousPropertyBits=   671105051
      MaxLength       =   60
      Size            =   "13044;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP66_2 
      Height          =   285
      Left            =   2100
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   3840
      Width           =   2235
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3942;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP65_2 
      Height          =   285
      Left            =   6360
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   3570
      Width           =   2235
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3942;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP59_2 
      Height          =   285
      Left            =   2100
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   3570
      Width           =   2235
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3942;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP58_2 
      Height          =   285
      Left            =   6360
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   3300
      Width           =   2235
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3942;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP08_2 
      Height          =   285
      Left            =   2100
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3300
      Width           =   2235
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3942;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   315
      Left            =   1200
      TabIndex        =   17
      Top             =   4950
      Width           =   7815
      VariousPropertyBits=   671105051
      MaxLength       =   2000
      Size            =   "13785;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP18 
      Height          =   315
      Left            =   1200
      TabIndex        =   18
      Top             =   5250
      Width           =   7815
      VariousPropertyBits=   671105051
      MaxLength       =   2000
      Size            =   "13785;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP44_2 
      Height          =   315
      Left            =   2895
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2085
      Width           =   6045
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "10663;556"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5070
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1275
      Width           =   4035
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "7223;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44 
      Height          =   285
      Left            =   5070
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   465
      Width           =   4035
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "7223;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1200
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1545
      Width           =   2532
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "7223;529"
      SpecialEffect   =   0
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
      Left            =   7680
      TabIndex        =   79
      Top             =   4155
      Width           =   765
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "發文規費："
      Height          =   180
      Left            =   2490
      TabIndex        =   78
      Top             =   1845
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2 :"
      Height          =   180
      Index           =   17
      Left            =   4440
      TabIndex        =   73
      Top             =   3345
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3 :"
      Height          =   180
      Index           =   16
      Left            =   120
      TabIndex        =   72
      Top             =   3612
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4 :"
      Height          =   180
      Index           =   15
      Left            =   4440
      TabIndex        =   71
      Top             =   3615
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5 :"
      Height          =   180
      Index           =   14
      Left            =   120
      TabIndex        =   70
      Top             =   3882
      Width           =   720
   End
   Begin VB.Label Label14 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   69
      Top             =   2370
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "申請者代號 :"
      Height          =   255
      Left            =   3930
      TabIndex        =   68
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
      Height          =   180
      Left            =   1620
      TabIndex        =   67
      Top             =   4170
      Width           =   2745
   End
   Begin VB.Label Label34 
      Caption         =   "列印定稿 :"
      Height          =   225
      Left            =   120
      TabIndex        =   66
      Top             =   4155
      Width           =   975
   End
   Begin VB.Label Label35 
      Caption         =   "(Y:修改)"
      Height          =   195
      Left            =   6570
      TabIndex        =   65
      Top             =   4140
      Width           =   825
   End
   Begin VB.Label Label36 
      Caption         =   "是否修改定稿內容 :"
      Height          =   195
      Left            =   4590
      TabIndex        =   64
      Top             =   4140
      Width           =   1575
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "主管機關 :"
      Height          =   180
      Left            =   4380
      TabIndex        =   63
      Top             =   3885
      Width           =   810
   End
   Begin VB.Label Label11 
      Caption         =   "TD序號 :"
      Height          =   252
      Left            =   4680
      TabIndex        =   62
      Top             =   4680
      Width           =   852
   End
   Begin VB.Line Line1 
      X1              =   2940
      X2              =   3060
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label29 
      Caption         =   "進度備註 :"
      Height          =   252
      Left            =   120
      TabIndex        =   60
      Top             =   4980
      Width           =   972
   End
   Begin VB.Label Label32 
      Caption         =   "案件備註 :"
      Height          =   252
      Left            =   120
      TabIndex        =   59
      Top             =   5280
      Width           =   972
   End
   Begin VB.Label Label7 
      Caption         =   "TD繳費新使用時間 :"
      Height          =   252
      Left            =   120
      TabIndex        =   58
      Top             =   4680
      Width           =   1692
   End
   Begin VB.Label Label6 
      Caption         =   "TD密碼 :"
      Height          =   252
      Left            =   120
      TabIndex        =   57
      Top             =   4380
      Width           =   852
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請人1 :"
      Height          =   180
      Left            =   120
      TabIndex        =   56
      Top             =   3342
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點數 :"
      Height          =   180
      Index           =   10
      Left            =   7620
      TabIndex        =   55
      Top             =   1845
      Width           =   450
   End
   Begin VB.Label Label25 
      Caption         =   "發文日 :"
      Height          =   180
      Left            =   90
      TabIndex        =   53
      Top             =   1845
      Width           =   855
   End
   Begin VB.Label Label31 
      Caption         =   "(N:不出名)"
      Height          =   180
      Left            =   6360
      TabIndex        =   52
      Top             =   1845
      Width           =   975
   End
   Begin VB.Label Label30 
      Caption         =   "是否出名 :"
      Height          =   180
      Left            =   5070
      TabIndex        =   51
      Top             =   1845
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "代理人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   50
      Top             =   2085
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "案件日文名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   2985
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "案件英文名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   48
      Top             =   2685
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "案件中文名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   2385
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員 :"
      Height          =   180
      Index           =   11
      Left            =   3930
      TabIndex        =   45
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   255
      Index           =   9
      Left            =   3930
      TabIndex        =   44
      Top             =   1005
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   43
      Top             =   1275
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   42
      Top             =   465
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   41
      Top             =   735
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   255
      Left            =   3930
      TabIndex        =   40
      Top             =   735
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "FC代理人 :"
      Height          =   255
      Index           =   2
      Left            =   3930
      TabIndex        =   39
      Top             =   465
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   1545
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "申請國家 :"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   1005
      Width           =   855
   End
End
Attribute VB_Name = "frm020102_21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/2/21 Form2.0已修改
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
' 申請人
Dim m_TM23 As String
'add by nickc 2007/02/01
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String

'Add By Sindy 2009/04/30
Dim m_CP84 As String       '發文規費

' 案件性質代號
Dim m_CP10 As String

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
'Add By Cheng 2002/06/14
Dim m_CP12 As String '業務區別代號
Dim m_CP13 As String '智權人員代號
Dim m_CP14 As String '承辦員代號
'Add By Cheng 2002/08/23
Dim m_strCust1 As String '申請人1
'add by nickc 2007/02/01
Dim m_strCust2 As String '申請人2
Dim m_strCust3 As String '申請人3
Dim m_strCust4 As String '申請人4
Dim m_strCust5 As String '申請人5

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
'add by nickc 2006/11/17
Dim m_textPrint As String
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
Dim m_CP07 As String 'Add By Sindy 2010/12/28 法定期限
Dim m_QSP As Boolean 'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
Dim m_990CP09 As String 'Add By Sindy 2016/12/20
Dim strLD18 As String 'Add By Sindy 2019/12/25 信函總收文號


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
'edit by nickc 2008/04/25 改整批印
'    'Add By Cheng 2004/04/08
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

Private Sub cmdok_Click()
Dim ii As Integer
   
   If CheckDataValid = True Then
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
            'GetCu103ByCustomer Me, textSP08.Text
            Call Pub_GetDataFrm020102(textSP08.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                        
            'edit by nickc 2006/01/20
            'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Then
            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  'Modified by Lydia 2024/07/03
                  'Set frm020102_22.oNextForm = Me
                  Call frm020102_22.SetParent(Me, textSP08.Text)
                  frm020102_22.Label4.Caption = textSP08.Text & " " & textSP08_2 'Add By Sindy 2014/7/30
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
            End If
            'add by nickc 2007/08/10 多申請人也要
            If textSP58.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, textSP58.Text
            Call Pub_GetDataFrm020102(textSP58.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                                    
            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  'Modified by Lydia 2024/07/03
                  'Set frm020102_22.oNextForm = Me
                  Call frm020102_22.SetParent(Me, textSP58.Text)
                  frm020102_22.Label4.Caption = textSP58.Text & " " & textSP58_2 'Add By Sindy 2014/7/30
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
            End If
            End If
            If textSP59.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, textSP59.Text
            Call Pub_GetDataFrm020102(textSP59.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                                    
            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  'Modified by Lydia 2024/07/03
                  'Set frm020102_22.oNextForm = Me
                  Call frm020102_22.SetParent(Me, textSP59.Text)
                  frm020102_22.Label4.Caption = textSP59.Text & " " & textSP59_2 'Add By Sindy 2014/7/30
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
            End If
            End If
            If textSP65.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, textSP65.Text
            Call Pub_GetDataFrm020102(textSP65.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)

            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  'Modified by Lydia 2024/07/03
                  'Set frm020102_22.oNextForm = Me
                  Call frm020102_22.SetParent(Me, textSP65.Text)
                  frm020102_22.Label4.Caption = textSP65.Text & " " & textSP65_2 'Add By Sindy 2014/7/30
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
            End If
            End If
            If textSP66.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, textSP66.Text
            Call Pub_GetDataFrm020102(textSP66.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)

            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  'Modified by Lydia 2024/07/03
                  'Set frm020102_22.oNextForm = Me
                  Call frm020102_22.SetParent(Me, textSP66.Text)
                  frm020102_22.Label4.Caption = textSP66.Text & " " & textSP66_2 'Add By Sindy 2014/7/30
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
            End If
            End If
      End If
      
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
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 更新欄位輸入的內容
      OnUpdateField
      ' 存檔
        'Modify By Cheng 2002/11/07
'      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      'Add By Cheng 2002/11/08
      If Me.textPrint.Text <> "N" Then
         PrintLetter
      'Add By Sindy 2021/3/31
      End If
      If textPrint = "N" Then
         If strLD18 <> "" Then
            Call PUB_TCaseAskIsPost(strLD18)
         End If
      '2021/3/31 END
      End If
      
      '2012/7/23 add by sonia
      '台灣案發文規費與收文規費不符時,mail給智權人員
      If textCP84.Enabled = True And m_TM10 = "000" And Val(Me.textCP84.Text) <> Val(m_CP84) Then
        'Add by Lydia 2014/10/13 內商服務業務(TC)之台灣案發文-規費與收文規費不符時,請加同時發給特殊設定人員"財務處總帳人員"
        If m_QSP = True Then
          PUB_ChkOfficialFee m_CP09, Me.textCP84.Text, "A"
        Else
          PUB_ChkOfficialFee m_CP09, Me.textCP84.Text
        End If
      End If
      '2012/7/23 end
      
      'Add By Sindy 2018/5/3
      If frm020102_01.bolIsEMPFlow = True Then
         frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
         frm090202_4.QueryData
      End If
      '2018/5/3 End
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
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
'      frm020102_01.Clear
      
      'Add By Cheng 2002/01/10
      frm020102_01.Clear1
      
      Unload Me
   End If
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

Private Sub Form_Activate()
'add by nickc 2005/08/23
If (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
   pub_ModifyCaseNum = ""
   QueryData
End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   textSP08_2.BackColor = &H8000000F
   
   'add by nickc 2007/02/15
   textSP58_2.BackColor = &H8000000F
   textSP59_2.BackColor = &H8000000F
   textSP65_2.BackColor = &H8000000F
   textSP66_2.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textTM44.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
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
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         'textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
      End If
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textSP05 = rsTmp.Fields("TM05")
      End If
      SetTMSPFieldOldData "TM05", textSP05, 0
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then
         textSP06 = rsTmp.Fields("TM06")
      End If
      SetTMSPFieldOldData "TM05", textSP06, 0
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then
         textSP07 = rsTmp.Fields("TM07")
      End If
      SetTMSPFieldOldData "TM07", textSP07, 0
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         'add by nickc 2007/02/01
         textSP08 = "" & rsTmp.Fields("TM23")
      End If
      'Add By Cheng 2002/08/23
      m_strCust1 = "" & Me.textSP08.Text
      'add by nickc 2007/02/01
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = rsTmp.Fields("TM78")
         textSP58 = "" & rsTmp.Fields("TM78")
      End If
      m_strCust2 = "" & Me.textSP58.Text
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = rsTmp.Fields("TM79")
         textSP59 = "" & rsTmp.Fields("TM79")
      End If
      m_strCust3 = "" & Me.textSP59.Text
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = rsTmp.Fields("TM80")
         textSP65 = "" & rsTmp.Fields("TM80")
      End If
      m_strCust4 = "" & Me.textSP65.Text
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = rsTmp.Fields("TM81")
         textSP66 = "" & rsTmp.Fields("TM81")
      End If
      m_strCust5 = "" & Me.textSP66.Text
      
      ' 案件備註
      If IsNull(rsTmp.Fields("TM58")) = False Then
         textSP18 = rsTmp.Fields("TM58")
      End If
      SetTMSPFieldOldData "TM58", textSP18, 0
      'add by nickc 2006/01/26
      m_TM24 = CheckStr(rsTmp.Fields("tm24"))
      SetTMSPFieldOldData "TM24", m_TM24, 0
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("tm77"))
      m_textPrint = textPrint
      SetTMSPFieldOldData "TM77", textPrint, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 取得服務業務基本檔的欄位內容
Private Sub QueryServicePractice()
   Dim strSP62 As String
   Dim strSql As String
   Dim strTemp As String
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
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textSP08 = rsTmp.Fields("SP08")
      End If
      SetTMSPFieldOldData "SP08", textSP08, 0
      'Add By Cheng 2002/08/23
      m_strCust1 = "" & Me.textSP08.Text
        Select Case m_TM01
        Case "TS"
            textSP05_1 = "" & rsTmp.Fields("SP05")
            SetTMSPFieldOldData "SP05", textSP05_1, 0
        Case Else
            ' 案件中文名稱
            If IsNull(rsTmp.Fields("SP05")) = False Then
               textSP05 = rsTmp.Fields("SP05")
            End If
            SetTMSPFieldOldData "SP05", textSP05, 0
        End Select
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textSP06 = rsTmp.Fields("SP06")
      End If
      SetTMSPFieldOldData "SP06", textSP06, 0
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textSP07 = rsTmp.Fields("SP07")
      End If
      SetTMSPFieldOldData "SP07", textSP07, 0
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
      End If
      'add by nickc 2007/08/10
      ' 申請人 2-5
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("SP58")) = False Then
         textSP58 = rsTmp.Fields("SP58")
         m_TM78 = rsTmp.Fields("SP58")
      End If
      SetTMSPFieldOldData "SP58", textSP58, 0
      m_strCust2 = "" & Me.textSP58.Text
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("SP59")) = False Then
         textSP59 = rsTmp.Fields("SP59")
         m_TM79 = rsTmp.Fields("SP59")
      End If
      SetTMSPFieldOldData "SP59", textSP59, 0
      m_strCust3 = "" & Me.textSP59.Text
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("SP65")) = False Then
         textSP65 = rsTmp.Fields("SP65")
         m_TM80 = rsTmp.Fields("SP65")
      End If
      SetTMSPFieldOldData "SP65", textSP65, 0
      m_strCust4 = "" & Me.textSP65.Text
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("SP66")) = False Then
         textSP66 = rsTmp.Fields("SP66")
         m_TM81 = rsTmp.Fields("SP66")
      End If
      SetTMSPFieldOldData "SP66", textSP66, 0
      m_strCust5 = "" & Me.textSP66.Text
      
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
         textTM10 = GetNationName(rsTmp.Fields("SP09"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("SP12")) = False Then
         'textTM20 = TAIWANDATE(rsTmp.Fields("SP12"))
      End If
      textSP20 = Empty
      textSP21 = Empty
      textSP49 = Empty
      If m_TM01 = "TD" Then
         EnableTextBox textSP20, True
         EnableTextBox textSP21, True
         EnableTextBox textSP49, True
         ' 使用期間(起)
         strTemp = Empty
         If IsNull(rsTmp.Fields("SP20")) = False Then
            strTemp = rsTmp.Fields("SP20")
            textSP20 = TAIWANDATE(rsTmp.Fields("SP20"))
         End If
         SetTMSPFieldOldData "SP20", strTemp, 1
         ' 使用期間(迄)
         strTemp = Empty
         If IsNull(rsTmp.Fields("SP21")) = False Then
            strTemp = rsTmp.Fields("SP21")
            textSP21 = TAIWANDATE(rsTmp.Fields("SP21"))
         End If
         SetTMSPFieldOldData "SP21", strTemp, 1
         ' 網域密碼
         If IsNull(rsTmp.Fields("SP49")) = False Then
            textSP49 = rsTmp.Fields("SP49")
         End If
         SetTMSPFieldOldData "SP49", textSP49, 0
      Else
         EnableTextBox textSP20, False
         EnableTextBox textSP21, False
         EnableTextBox textSP49, False
      End If
      ' 案件備註
      If IsNull(rsTmp.Fields("SP18")) = False Then
         textSP18 = rsTmp.Fields("SP18")
      End If
      SetTMSPFieldOldData "SP18", textSP18, 0
      ' TD序號
      textSP11 = Empty
      If m_TM01 = "TD" Then
         EnableTextBox textSP11, True
         If IsNull(rsTmp.Fields("SP11")) = False Then
            textSP11 = rsTmp.Fields("SP11")
         End If
         SetTMSPFieldOldData "SP11", textSP11, 0
      Else
         EnableTextBox textSP11, False
      End If
      ' 主管機關
      textSP51 = Empty
      If IsNull(rsTmp.Fields("SP51")) = False Then
         textSP51 = rsTmp.Fields("SP51")
      End If
      SetTMSPFieldOldData "SP51", textSP51, 0
      ' 91.09.02 modify by louis
      ' 增加申請者代號欄位
      ' 申請者代號
      If m_TM01 = "TD" And m_CP10 = "805" Then
         textSP64 = Empty
         If IsNull(rsTmp.Fields("SP64")) = False Then
            textSP64 = rsTmp.Fields("SP64")
         End If
         SetTMSPFieldOldData "SP64", textSP64, 0
      End If
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("SP72"))
      m_textPrint = textPrint
      SetTMSPFieldOldData "SP72", textPrint, 0
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
Dim strCP43 As String
Dim strCP44 As String
Dim strCP45 As String
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
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 業務區別
      If IsNull(rsTmp.Fields("CP12")) = False Then
         '91.6.11 MODIFY BY SONIA
         'textCP12 = GetStaffDepartment(rsTmp.Fields("CP12"))
         'textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      'Add By Cheng 2002/06/14
      m_CP12 = "" & rsTmp.Fields("CP12").Value
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      'Add By Cheng 2002/06/14
      m_CP13 = "" & rsTmp.Fields("CP13").Value
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      'Add By Cheng 2002/06/14
      m_CP14 = "" & rsTmp.Fields("CP14").Value
      
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
      
      ' 是否出名
      textCP22 = Empty
      If IsNull(rsTmp.Fields("CP22")) = False Then
         textCP22 = rsTmp.Fields("CP22")
      End If
      SetCPFieldOldData "CP22", textCP22, 0
      ' 發文日(預設為系統日)
      textCP27 = TAIWANDATE(SystemDate())
      strCP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         strCP27 = rsTmp.Fields("CP27")
      End If
      SetCPFieldOldData "CP27", strCP27, 1
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then: textCP44 = rsTmp.Fields("CP44")
      SetCPFieldOldData "CP44", textCP44, 0
      ' 彼所案號
      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then: strCP45 = rsTmp.Fields("CP45")
      SetCPFieldOldData "CP45", strCP45, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then: textCP18 = rsTmp.Fields("CP18")
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then: textCP64 = rsTmp.Fields("CP64")
      SetCPFieldOldData "CP64", textCP64, 0
      'Add By Sindy 2009/04/30 發文規費
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
      
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
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
'      strSubSQL = "SELECT DISTINCT CP44 FROM CaseProgress " & _
'                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
'                        "CP02 = '" & m_TM02 & "' AND " & _
'                        "CP03 = '" & m_TM03 & "' AND " & _
'                        "CP04 = '" & m_TM04 & "' AND " & _
'                        "CP09 <> '" & m_CP09 & "' "
      strSubSQL = "SELECT CP44, Max(CP05||CP09) FROM CaseProgress " & _
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
       'Modify By Cheng 2002/09/19
'            textCP44.AddItem m_AgentList(nIndex).aiName
       textCP44.AddItem m_AgentList(nIndex).aiCode
    Next nIndex
    ' 設定顯示為第一筆
    If textCP44.ListCount > 0 Then
       textCP44.ListIndex = 0
       textCP44_Validate False
    End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   EnableTextBox textSP20, False
   EnableTextBox textSP21, False
   EnableTextBox textSP49, False
   
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
    Case "TS"
        Me.Label14.Visible = True
        Me.textSP05_1.Visible = True
        Me.textSP05_1.Enabled = True
        Me.Label10.Visible = False
        Me.textSP05.Visible = False
        Me.textSP05.Enabled = False
        Me.Label9.Visible = False
        Me.textSP06.Visible = False
        Me.textSP06.Enabled = False
        Me.Label8.Visible = False
        Me.textSP07.Visible = False
        Me.textSP07.Enabled = False
    Case Else
        Me.Label14.Visible = False
        Me.textSP05_1.Visible = False
        Me.textSP05_1.Enabled = False
        Me.Label10.Visible = True
        Me.textSP05.Visible = True
        Me.textSP05.Enabled = True
        Me.Label9.Visible = True
        Me.textSP06.Visible = True
        Me.textSP06.Enabled = True
        Me.Label8.Visible = True
        Me.textSP07.Visible = True
        Me.textSP07.Enabled = True
    End Select
   ' 取得案件進度檔的欄位
   QueryCaseProgress
   
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
   
   textSP08_2 = GetCustomerName(textSP08, 0)
   ' 91.09.02 modify by louis
   ' 系統類別為TD, 案件性質為申請805時才可輸入申請者代號欄位
   If m_TM01 = "TD" And m_CP10 = "805" Then
      EnableTextBox textSP64, True
   Else
      EnableTextBox textSP64, False
   End If
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/17 若已經從基本檔抓出來，就不重抓
   If Trim(textPrint) = "" Then
       textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   
   'Add By Sindy 2021/3/31 案件性質為706(其它),定稿列印請自動上 "N"
   If m_CP10 = "706" Then
      textPrint = "N"
   End If
   '2021/3/31 END
   'Add By Sindy 2025/8/11 檢查卷宗區是否已有承辦放入之CUS,若有,系統不產出定稿
   If PUB_CPPChkFileExists(m_CP09, "cus") = True Then
      textPrint = "N"
   End If
   '2025/8/11 END
   'Added by Lydia 2025/10/27 各專業部的智財協作發文，都預設不出定稿
   If m_TM01 = "TT" And m_CP10 = "737" Then
      textPrint = "N"
   End If
   'end 2025/10/27
   
   Call PUB_TCaseEFeeRemind(m_CP09) 'Add By Sindy 2016/5/9 內商電子收文請款提醒訊息
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache  'Added by Lydia 2025/02/24
   
   'Add By Cheng 2002/07/18
   Set frm020102_21 = Nothing
End Sub

' 是否出名
Private Sub textCP22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否出名
Private Sub textCP22_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP22) = False Then
      Select Case textCP22
         Case " ", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP22_GotFocus
      End Select
   End If
End Sub

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
    'Add By Cheng 2002/12/03
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
   If CheckLengthIsOK(textCP64, textCP64.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

'Add By Sindy 2009/04/30
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
'2009/04/30 End

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
'edit by nickc 2006/06/29
'   If KeyAscii <> 78 And KeyAscii <> 8 Then
'      KeyAscii = 0
'   End If
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
   Dim nIndex As Integer
   Dim strSP62 As String
   
   ' 是否出名7
   SetCPFieldNewData "CP22", textCP22
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' 代理人代號
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
   ' 進度備註
   SetCPFieldNewData "CP64", textCP64
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
   Select Case m_TM01
      Case "T", "TF", "FCT":
         ' 案件名稱(中)
         SetTMSPFieldNewData "TM05", textSP05
         ' 案件名稱(英)
         SetTMSPFieldNewData "TM06", textSP06
         ' 案件名稱(日)
         SetTMSPFieldNewData "TM07", textSP07
         ' 案件備註
         SetTMSPFieldNewData "TM58", textSP18
         ' 申請人 2007/02/15
         If IsEmptyText(textSP08) = False Then
            SetTMSPFieldNewData "TM23", textSP08 & String(9 - Len(textSP08), "0")
         Else
            SetTMSPFieldNewData "TM23", textSP08
         End If
         'add by nickc 2007/02/01
         If IsEmptyText(textSP58) = False Then
            SetTMSPFieldNewData "TM78", textSP58 & String(9 - Len(textSP58), "0")
         Else
            SetTMSPFieldNewData "TM78", textSP58
         End If
         If IsEmptyText(textSP59) = False Then
            SetTMSPFieldNewData "TM79", textSP59 & String(9 - Len(textSP59), "0")
         Else
            SetTMSPFieldNewData "TM79", textSP59
         End If
         If IsEmptyText(textSP65) = False Then
            SetTMSPFieldNewData "TM80", textSP65 & String(9 - Len(textSP65), "0")
         Else
            SetTMSPFieldNewData "TM80", textSP65
         End If
         If IsEmptyText(textSP66) = False Then
            SetTMSPFieldNewData "TM81", textSP66 & String(9 - Len(textSP66), "0")
         Else
            SetTMSPFieldNewData "TM81", textSP66
         End If
         'add by nickc 2006/01/26
         If m_CU112 <> "" Then
            'Modify By Sindy 2011/2/22
            'SetTMSPFieldNewData "TM24", Pub_RplCu112(m_TM24, m_CU112)
            SetTMSPFieldNewData "TM24", Pub_RplCu112(m_TM24, m_CU112, m_TM23)
         Else
            SetTMSPFieldNewData "TM24", m_TM24
         End If
         'add by nickc 2006/11/17
         If textPrint <> "N" Then
            SetTMSPFieldNewData "TM77", textPrint
         Else
            SetTMSPFieldNewData "TM77", m_textPrint
         End If
      Case Else:
        Select Case m_TM01
        Case "TS"
            ' 案件名稱(中)
            SetTMSPFieldNewData "SP05", textSP05_1
        Case Else
            ' 案件名稱(中)
            SetTMSPFieldNewData "SP05", textSP05
        End Select
         ' 案件名稱(英)
         SetTMSPFieldNewData "SP06", textSP06
         ' 案件名稱(日)
         SetTMSPFieldNewData "SP07", textSP07
         ' 申請人
         If IsEmptyText(textSP08) = False Then
            SetTMSPFieldNewData "SP08", textSP08 & String(9 - Len(textSP08), "0")
         Else
            SetTMSPFieldNewData "SP08", textSP08
         End If
         'add by nickc 2007/02/01
         If IsEmptyText(textSP58) = False Then
            SetTMSPFieldNewData "SP58", textSP58 & String(9 - Len(textSP58), "0")
         Else
            SetTMSPFieldNewData "SP58", textSP58
         End If
         If IsEmptyText(textSP59) = False Then
            SetTMSPFieldNewData "SP59", textSP59 & String(9 - Len(textSP59), "0")
         Else
            SetTMSPFieldNewData "SP59", textSP59
         End If
         If IsEmptyText(textSP65) = False Then
            SetTMSPFieldNewData "SP65", textSP65 & String(9 - Len(textSP65), "0")
         Else
            SetTMSPFieldNewData "SP65", textSP65
         End If
         If IsEmptyText(textSP66) = False Then
            SetTMSPFieldNewData "SP66", textSP66 & String(9 - Len(textSP66), "0")
         Else
            SetTMSPFieldNewData "SP66", textSP66
         End If
         
         
         ' 案件備註
         SetTMSPFieldNewData "SP18", textSP18
         ' 主管機關
         SetTMSPFieldNewData "SP51", textSP51
         ' 系統類別為TD時
         If m_TM01 = "TD" Then
            ' TD密碼
            SetTMSPFieldNewData "SP49", textSP49
            ' TD繳年費新使用時間(起)
            SetTMSPFieldNewData "SP20", DBDATE(textSP20)
            ' TD繳年費新使用時間(迄)
            SetTMSPFieldNewData "SP21", DBDATE(textSP21)
            ' TD序號
            SetTMSPFieldNewData "SP11", textSP11
            ' 91.09.02 modify by louis
            ' 申請者代號
            If m_CP10 = "805" Then
               SetTMSPFieldNewData "SP64", textSP64
            End If
         End If
         'add by nickc 2006/11/17
         If textPrint <> "N" Then
            SetTMSPFieldNewData "SP72", textPrint
         Else
            SetTMSPFieldNewData "SP72", m_textPrint
         End If
   End Select
End Sub

' 更新商標基本檔的相關欄位
'Modify By Cheng 2002/11/07
'Private Sub OnUpdateTradeMark()
Private Function OnUpdateTradeMark() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnUpdateTradeMark = True

   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
            Else
               ' 91.03.25 modify by louis
               'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
               strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
'Add By Cheng 2002/11/07
Exit Function
ErrorHandler:
    OnUpdateTradeMark = False
End Function

' 更新服務業務基本檔的相關欄位
'Mlodify By Cheng 2002/11/07
'Private Sub OnUpdateServicePractice()
Private Function OnUpdateServicePractice() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnUpdateServicePractice = True

   ' 更新案件進度檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
            Else
               ' 91.03.25 modify by louis
               'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
               strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
   
   'Add By Cheng 2002/06/14
   '若案件性質為"網路名稱申請"(805), 且申請國家為台灣時, 發文日同時更新服務基本檔的申請日
   '2012/5/11 MODIFY BY SONIA 取申請國家限制 TD-000153
   'If m_CP10 = "805" And m_TM10 = 台灣國家代號 Then
   If m_CP10 = "805" Then
      strSql = " Update ServicePractice Set SP10=" & DBDATE(Me.textCP27.Text) & _
               " WHERE SP01 = '" & m_TM01 & "' AND " & _
                      "SP02 = '" & m_TM02 & "' AND " & _
                      "SP03 = '" & m_TM03 & "' AND " & _
                      "SP04 = '" & m_TM04 & "' "
      cnnConnection.Execute strSql
   End If
'Add By Cheng 2002/11/07
Exit Function
ErrorHandler:
    OnUpdateServicePractice = False
End Function

' 更新案件進度檔
'Modify By Cheng 2002/11/07
'Private Sub OnUpdateCaseProgress()
Private Function OnUpdateCaseProgress() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   
'Add By Cheng 2002/11/07
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
               ' 91.03.25 modify by louis
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
   
'Add By Cheng 2002/11/07
Exit Function
ErrorHandler:
    OnUpdateCaseProgress = False
End Function

'Modify By Cheng 2002/11/0
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strNP08 As String
   Dim strNP07 As String
   Dim strNP22 As String
   Dim bolSysDt As Boolean 'Add By Sindy 2010/12/28
   'Add By Cheng 2002/06/14
   Dim strCP09  As String '收文號
   Dim strCP05  As String '收文日
   Dim strCP10  As String '案件性質
   Dim strCP06  As String '本所期限
   Dim strCP07  As String '法定期限
   Dim strCP64  As String '進度備註
   Dim strCP20  As String '是否向客戶請款
   
'Add By Cheng 2002/11/07
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
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新案件進度檔
    'Modify By Cheng 2002/11/07
'   OnUpdateCaseProgress
   If OnUpdateCaseProgress = False Then GoTo ErrorHandler
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   'Added by Lydia 2025/02/24 TIPS分配比例管制：與ACS案有關之智財協作發文時一併產生TIPS案請款階段分配比例
   If m_TM01 = "TT" Then
      Call PUB_InsertACS_TIPS_Rate(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_CP10)
   End If
   'end 2025/02/24
   
   ' 若案件性質為"網域名稱申請"且申請國家為台灣時, 新增一筆資料到案件進度檔中
   If m_CP10 = "805" And m_TM10 = 台灣國家代號 Then
      ' 收文號
      strCP09 = Empty
      strCP09 = AutoNo("B", 6)
      ' 收文日
      strCP05 = DBDATE(SystemDate())
      ' 案件性質
      strCP10 = "706"
      '本所期限
    'Modify By Cheng 2003/09/01
'      strCP06 = DBDATE(DateSerial(DBYEAR(Me.textCP27.Text), DBMONTH(Me.textCP27.Text), DBDAY(Me.textCP27.Text) + 7))
      strCP06 = DBDATE(DateAdd("d", 7, ChangeWStringToWDateString(DBDATE(Me.textCP27.Text))))
      strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      '法定期限
      strCP07 = strCP06
      '進度備註
      strCP64 = "網域申請繳款"
      '是否向客戶請款
      strCP20 = "N"
      
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP20,CP22,CP43,CP44,CP64) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                       strCP06 & "," & strCP07 & ",'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & m_CP14 & "','" & strCP20 & "'," & _
                       "'" & Me.textCP22.Text & "','" & m_CP09 & "','" & Me.textCP44.Text & "','" & strCP64 & "') "
      cnnConnection.Execute strSql
   End If
   
   ' 更新基本檔
   Select Case m_TM01
      Case "T", "TF", "FCT":
        'Modify By Cheng 2002/11/07
'         OnUpdateTradeMark
         If OnUpdateTradeMark = False Then GoTo ErrorHandler
      Case Else:
        'Modify By Cheng 2002/11/07
'         OnUpdateServicePractice
         If OnUpdateServicePractice = False Then GoTo ErrorHandler
   End Select
      
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
         
         ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
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
      'Add By Sindy 2012/9/10
      If IsNull(rsTmp.Fields("CF05")) = False Then
         strNP07 = "305"
         strNP08 = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
         strNP22 = GetNextProgressNo()
         'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            strNP08 & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
         cnnConnection.Execute strSql
      End If
      '2012/9/10 End
   End If
   rsTmp.Close
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
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textSP08.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP08.Text), 9, 1) & "' "
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "',cu112='" & ChgSQL(m_CU112) & "'  Where Cu01 = '" & Mid(ChangeCustomerL(textSP08.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP08.Text), 9, 1) & "' "
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(1)) & "',cu05='" & ChgSQL(SeekCu05(1)) & "',cu88='" & ChgSQL(SeekCu88(1)) & "',cu89='" & ChgSQL(SeekCu89(1)) & "',cu90='" & ChgSQL(SeekCu90(1)) & "',cu112='" & ChgSQL(SeekCu112(1)) & "',cu39='" & ChgSQL(SeekCu39(1)) & "',cu40='" & ChgSQL(SeekCu40(1)) & "',cu41='" & ChgSQL(SeekCu41(1)) & "',cu10='" & ChgSQL(SeekCu10(1)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textSP08.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP08.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(1)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textSP08.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP08.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   'add by nickc 2007/08/10 加多申請人也要
   If (SeekCu103(2) <> "" Or (SeekCu05(2) & SeekCu88(2) & SeekCu89(2) & SeekCu90(2)) <> "" Or SeekCu112(2) <> "" Or (SeekCu39(2) & SeekCu40(2) & SeekCu41(2)) <> "" Or SeekCu10(2) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(2)) & "',cu05='" & ChgSQL(SeekCu05(2)) & "',cu88='" & ChgSQL(SeekCu88(2)) & "',cu89='" & ChgSQL(SeekCu89(2)) & "',cu90='" & ChgSQL(SeekCu90(2)) & "',cu112='" & ChgSQL(SeekCu112(2)) & "',cu39='" & ChgSQL(SeekCu39(2)) & "',cu40='" & ChgSQL(SeekCu40(2)) & "',cu41='" & ChgSQL(SeekCu41(2)) & "',cu10='" & ChgSQL(SeekCu10(2)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textSP58.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP58.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(2)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textSP58.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP58.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(3) <> "" Or (SeekCu05(3) & SeekCu88(3) & SeekCu89(3) & SeekCu90(3)) <> "" Or SeekCu112(3) <> "" Or (SeekCu39(3) & SeekCu40(3) & SeekCu41(3)) <> "" Or SeekCu10(3) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(3)) & "',cu05='" & ChgSQL(SeekCu05(3)) & "',cu88='" & ChgSQL(SeekCu88(3)) & "',cu89='" & ChgSQL(SeekCu89(3)) & "',cu90='" & ChgSQL(SeekCu90(3)) & "',cu112='" & ChgSQL(SeekCu112(3)) & "',cu39='" & ChgSQL(SeekCu39(3)) & "',cu40='" & ChgSQL(SeekCu40(3)) & "',cu41='" & ChgSQL(SeekCu41(3)) & "',cu10='" & ChgSQL(SeekCu10(3)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textSP59.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP59.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(3)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textSP59.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP59.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(4) <> "" Or (SeekCu05(4) & SeekCu88(4) & SeekCu89(4) & SeekCu90(4)) <> "" Or SeekCu112(4) <> "" Or (SeekCu39(4) & SeekCu40(4) & SeekCu41(4)) <> "" Or SeekCu10(4) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(4)) & "',cu05='" & ChgSQL(SeekCu05(4)) & "',cu88='" & ChgSQL(SeekCu88(4)) & "',cu89='" & ChgSQL(SeekCu89(4)) & "',cu90='" & ChgSQL(SeekCu90(4)) & "',cu112='" & ChgSQL(SeekCu112(4)) & "',cu39='" & ChgSQL(SeekCu39(4)) & "',cu40='" & ChgSQL(SeekCu40(4)) & "',cu41='" & ChgSQL(SeekCu41(4)) & "',cu10='" & ChgSQL(SeekCu10(4)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textSP65.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP65.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(4)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textSP65.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP65.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(5) <> "" Or (SeekCu05(5) & SeekCu88(5) & SeekCu89(5) & SeekCu90(5)) <> "" Or SeekCu112(5) <> "" Or (SeekCu39(5) & SeekCu40(5) & SeekCu41(5)) <> "" Or SeekCu10(5) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(5)) & "',cu05='" & ChgSQL(SeekCu05(5)) & "',cu88='" & ChgSQL(SeekCu88(5)) & "',cu89='" & ChgSQL(SeekCu89(5)) & "',cu90='" & ChgSQL(SeekCu90(5)) & "',cu112='" & ChgSQL(SeekCu112(5)) & "',cu39='" & ChgSQL(SeekCu39(5)) & "',cu40='" & ChgSQL(SeekCu40(5)) & "',cu41='" & ChgSQL(SeekCu41(5)) & "',cu10='" & ChgSQL(SeekCu10(5)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textSP66.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP66.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(5)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textSP66.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textSP66.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Modify By Cheng 2002/11/08
'   'Modify By Cheng 2002/06/14
'   ' 直接列印定稿
'   If Me.textPrint.Text <> "N" Then
'      PrintLetter
'   End If
   
   'Add by Sindy 98/3/24
   If m_TM10 = "000" Then
      'Modify By Sindy 2009/04/24
      'PUB_UpdateDispatch m_CP09s, m_CP123s
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
   End If
   
   'Add By Sindy 2009/04/30 更新實際發文規費
   If textCP84.Enabled = True Then
      strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Add by Sindy 2012/10/4 外->台,智權人員是葉雪貞及巨京,發文規費和收文規費不相同時,系統自動更改進度檔內規費費用及計算點數
   'Modified by Lydia 2015/10/16 + m_CP84
   Call PUB_TSendUpdateCP1718(m_CP09, textCP84, textPrint, m_TM10, m_CP13, m_CP84)
   
   'Add By Sindy 2019/12/25 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = m_CP09
      PUB_AddLetterProgress strLD18, 0, IIf(textPrint = "N", False, True), "", False, m_TM23, m_CP10, m_TM44
   End If
   '2019/12/25 END
   Call PUB_UpdateLP19_T(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textCP27) 'Add by Sindy 2020/2/12 收據/回執設定
   
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
   
   Set rsTmp = Nothing
'Add By Cheng 2002/11/07
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

' 檢查欄位是否都已輸入或是輸入的值是否正確
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'add by nickc 2008/05/01
   If IsDebt(m_TM10, textCP09) Then
        strTit = "警告！禁止發文！"
        strMsg = "未收款且無 預定收款日 請轉告智權同仁！！"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        GoTo EXITSUB
   End If
   ' 發文日
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   ' 申請國家非台灣時代理人不可空白
   'Modify By Sindy 2012/3/29 TD申請時皆在台灣申請不須控管CF代理人
   If m_TM10 >= "010" And m_TM01 <> "TD" Then
      If IsEmptyText(textCP44) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入代理人"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   ' 案件名稱不可同時空白
    Select Case m_TM01
    Case "TS"
        If IsEmptyText("textSP05_1") = True Then
           strTit = "檢核資料"
           strMsg = "請輸入案件名稱"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           GoTo EXITSUB
        End If
    Case Else
        If IsEmptyText("textSP05") = True And IsEmptyText("textSP06") = True And IsEmptyText("textSP07") = True Then
           strTit = "檢核資料"
           strMsg = "請輸入案件名稱"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           GoTo EXITSUB
        End If
    End Select
   If m_TM01 = "TD" Then
      ' 90.07.2 modify (繳年費期間不需要一定輸入)
      '' TD繳年費新使用時間(起)
      'If IsEmptyText(textSP20) = True Then
      '   strTit = "檢核資料"
      '   strMsg = "請輸入TD繳年費新使用時間(起)"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   GoTo EXITSUB
      'End If
      '' TD繳年費新使用時間(迄)
      'If IsEmptyText(textSP21) = True Then
      '   strTit = "檢核資料"
      '   strMsg = "請輸入TD繳年費新使用時間(迄)"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   GoTo EXITSUB
      'End If
      If IsEmptyText(textSP20) = False And IsEmptyText(textSP21) = False Then
         If DBDATE(textSP20) > DBDATE(textSP21) Then
            strTit = "檢核資料"
            strMsg = "繳年費期間範圍不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
      End If
      ' TD密碼
      'If IsEmptyText(textSP49) = True Then
      '   strTit = "檢核資料"
      '   strMsg = "請輸入TD密碼"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   GoTo EXITSUB
      'End If
   End If
   
   'Add By Sindy 2011/01/06
   '內商(TS)申請人1或FC代理人至少要輸入一個
   '其他的一定要輸入申請人1
   'MODIFY BY SONIA 2015/11/5 +TT
   If m_TM01 = "TS" Or m_TM01 = "TT" Then
        If textSP08 = "" And m_TM44 = "" Then
            MsgBox "申請人1或FC代理人至少要輸入一個!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   Else
        If textSP08 = "" Then
            MsgBox "申請人1不可空白!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

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

Private Sub textSP05_1_GotFocus()
    TextInverse Me.textSP05_1
End Sub

Private Sub textSP05_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP05_1, textSP05_1.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件名稱內容太長"
      textSP05_1_GotFocus
   End If
End Sub

' 案件中文名稱
Private Sub textSP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP05, textSP05.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件中文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP05_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textSP05.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 案件英文名稱
Private Sub textSP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP06, textSP06.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件英文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP06_GotFocus
   End If
End Sub

' 案件日文名稱
Private Sub textSP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP07, textSP07.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件日文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP07_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textSP07.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
'add by nickc 2007/02/01
Private Sub textSP08_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 申請人
Private Sub textSP08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textSP08_2 = Empty
   If IsEmptyText(textSP08) = False Then
        Me.textSP08.Text = ChangeCustomerL(Me.textSP08.Text)
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      oState = True
      'textSP08_2 = GetCustomerName(textSP08, 0)
      textSP08_2 = GetCustomerNameAndState(textSP08, 0, oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If IsEmptyText(textSP08_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人1代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP08_GotFocus
      End If
   End If
   'Add By Cheng 2002/08/22
   If Cancel = False Then
      If Me.textSP08.Text <> m_strCust1 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textSP08_GotFocus
   
End Sub

' TD繳費新使用時間(起)
Private Sub textSP20_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textSP20) = False Then
      If CheckIsTaiwanDate(textSP20, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請輸入正確的TD繳費新使用時間(起)"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP20_GotFocus
      End If
   End If
End Sub

' TD繳費新使用時間(迄)
Private Sub textSP21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textSP21) = False Then
      If CheckIsTaiwanDate(textSP21, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請輸入正確的TD繳費新使用時間(迄)"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP21_GotFocus
      End If
   End If
End Sub

' TD密碼
Private Sub textSP49_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP49, textSP49.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "TD密碼內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP49_GotFocus
   End If
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

Private Sub textSP05_GotFocus()
   InverseTextBox textSP05
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textSP05.IMEMode = 1
   OpenIme
End Sub

Private Sub textSP06_GotFocus()
   InverseTextBox textSP06
End Sub

Private Sub textSP07_GotFocus()
   InverseTextBox textSP07
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textSP07.IMEMode = 1
   OpenIme
End Sub

Private Sub textSP08_GotFocus()
   InverseTextBox textSP08
End Sub

Private Sub textSP20_GotFocus()
   InverseTextBox textSP20
End Sub

Private Sub textSP21_GotFocus()
   InverseTextBox textSP21
End Sub

Private Sub textSP49_GotFocus()
   InverseTextBox textSP49
End Sub
'add by nickc 2007/02/01
Private Sub textSP58_GotFocus()
InverseTextBox textSP58
End Sub
Private Sub textSP58_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textSP58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textSP58_2 = Empty
   If IsEmptyText(textSP58) = False Then
        Me.textSP58.Text = ChangeCustomerL(Me.textSP58.Text)
      Dim oState As Boolean
      oState = True
      textSP58_2 = GetCustomerNameAndState(textSP58, 0, oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If IsEmptyText(textSP58_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人2代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP58_GotFocus
      End If
   End If
   If Cancel = False Then
      If Me.textSP58.Text <> m_strCust2 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textSP58_GotFocus
   
End Sub
Private Sub textSP59_GotFocus()
InverseTextBox textSP59
End Sub
Private Sub textSP59_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textSP59_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textSP59_2 = Empty
   If IsEmptyText(textSP59) = False Then
        Me.textSP59.Text = ChangeCustomerL(Me.textSP59.Text)
      Dim oState As Boolean
      oState = True
      textSP59_2 = GetCustomerNameAndState(textSP59, 0, oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If IsEmptyText(textSP59_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人3代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP59_GotFocus
      End If
   End If
   If Cancel = False Then
      If Me.textSP59.Text <> m_strCust3 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textSP59_GotFocus
   
End Sub

Private Sub textSP64_GotFocus()
   InverseTextBox textSP64
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
Dim strTM23Nation As String
Dim strSql As String
Dim strTmp As String               '2011/5/25 add by sonia
Dim rsTmp As New ADODB.Recordset   '2011/5/25 add by sonia
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 系統類別TD
   If m_TM01 = "TD" Then
      ' 案件性質為網域名稱申請
      If m_CP10 = "805" Then
         'add by nickc 2006/06/29
         If textPrint = "1" Then
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "00", strUserNum
         End If
      '2011/5/25 add by sonia 增加通用發文定稿
      Else
         If textPrint = "1" Then
            EndLetter "01", m_CP09, "00", strUserNum
            strTmp = Empty
            strSql = "SELECT * FROM CaseFee WHERE CF01 = '" & m_TM01 & "' AND " & _
                           "CF02 = '" & m_TM10 & "' AND CF03 = '" & m_CP10 & "' "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               If IsNull(rsTmp.Fields("CF09")) = False Then
                  strTmp = rsTmp.Fields("CF09")
               End If
            End If
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                     "'回音'," & CNULL(strTmp) & ")"
            cnnConnection.Execute strSql
         End If
      '2011/5/25 end
      End If
   'Add By Sindy 2016/10/17
   ElseIf m_TM01 = "TT" Then
      If textPrint = "2" Then '大->台
         ' 清除定稿例外欄位檔原有資料
         EndLetter "01", m_CP09, "01", strUserNum
      End If
   '2016/10/17 END
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
   bolEdit = IIf(Me.textWord.Text = "Y", True, False)
   '2012/1/12 End
   
   ' 系統類別TD
   If m_TM01 = "TD" Then
      ' 案件性質為網域名稱申請
      If m_CP10 = "805" Then
         ' 列印定稿
         'Modify By Cheng 2002/06/14
'         NowPrint m_CP09, "01", "31", False, strUserNum, 0
         'add by nickc 2006/06/29
         If textPrint = "1" Then
'            NowPrint m_CP09, "01", "00", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
            ET03 = "00" 'Modify By Sindy 2012/1/12
         End If
      '2011/5/25 add by sonia 增加通用發文定稿
      Else
         If textPrint = "1" Then
'            NowPrint m_CP09, "01", "00", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
            ET03 = "00" 'Modify By Sindy 2012/1/12
         End If
      '2011/5/25 end
      End If
   'Add By Sindy 2016/10/17
   ElseIf m_TM01 = "TT" Then
      If textPrint = "2" Then '大->台
         ET03 = "01" '公證
      End If
   '2016/10/17 END
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
         'Add By Sindy 2019/12/25 + strLD18.信函總收文號
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
   
   'Add By Sindy 2009/04/30
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
   '2009/04/30 End
   
   If Me.textCP22.Enabled = True Then
      Cancel = False
      textCP22_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
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
   
   If Me.textSP05.Enabled = True Then
      Cancel = False
      textSP05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP06.Enabled = True Then
      Cancel = False
      textSP06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP07.Enabled = True Then
      Cancel = False
      textSP07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP08.Enabled = True Then
      Cancel = False
      textSP08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP20.Enabled = True Then
      Cancel = False
      textSP20_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP21.Enabled = True Then
      Cancel = False
      textSP21_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP49.Enabled = True Then
      Cancel = False
      textSP49_Validate Cancel
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
   
Private Sub textSP64_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'add by nickc 2007/02/01
Private Sub textSP65_GotFocus()
InverseTextBox textSP65
End Sub
Private Sub textSP65_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textSP65_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textSP65_2 = Empty
   If IsEmptyText(textSP65) = False Then
        Me.textSP65.Text = ChangeCustomerL(Me.textSP65.Text)
      Dim oState As Boolean
      oState = True
      textSP65_2 = GetCustomerNameAndState(textSP65, 0, oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If IsEmptyText(textSP65_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人4代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP65_GotFocus
      End If
   End If
   If Cancel = False Then
      If Me.textSP65.Text <> m_strCust4 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textSP65_GotFocus
   
End Sub
Private Sub textSP66_GotFocus()
InverseTextBox textSP66
End Sub
Private Sub textSP66_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textSP66_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textSP66_2 = Empty
   If IsEmptyText(textSP66) = False Then
        Me.textSP66.Text = ChangeCustomerL(Me.textSP66.Text)
      Dim oState As Boolean
      oState = True
      textSP66_2 = GetCustomerNameAndState(textSP66, 0, oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If IsEmptyText(textSP66_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人5代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP66_GotFocus
      End If
   End If
   If Cancel = False Then
      If Me.textSP66.Text <> m_strCust5 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textSP66_GotFocus
   
End Sub

Private Sub textWord_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
   End If
End Sub

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
