VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_7 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文-領證及繳年費"
   ClientHeight    =   5412
   ClientLeft      =   492
   ClientTop       =   1716
   ClientWidth     =   8820
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5412
   ScaleWidth      =   8820
   Begin VB.TextBox txtEmail 
      Height          =   270
      Left            =   7380
      MaxLength       =   1
      TabIndex        =   14
      Top             =   4980
      Width           =   375
   End
   Begin VB.TextBox txtRecDate 
      Height          =   270
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   13
      Top             =   4980
      Width           =   375
   End
   Begin VB.TextBox txtPAID 
      Height          =   270
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   12
      Top             =   4980
      Width           =   375
   End
   Begin VB.TextBox txtPayToday 
      Height          =   270
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   11
      Top             =   4650
      Width           =   375
   End
   Begin VB.TextBox txtCP118 
      Height          =   270
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   10
      Top             =   4650
      Width           =   375
   End
   Begin VB.TextBox Text13 
      Height          =   270
      Left            =   8370
      MaxLength       =   8
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   990
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TextCP148 
      Height          =   270
      Left            =   7380
      MaxLength       =   1
      TabIndex        =   5
      Top             =   3090
      Width           =   375
   End
   Begin VB.CheckBox chk412 
      Enabled         =   0   'False
      Height          =   195
      Left            =   2985
      TabIndex        =   6
      Top             =   3435
      Value           =   1  '核取
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.TextBox txtCP71 
      Enabled         =   0   'False
      Height          =   270
      Left            =   5010
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "3"
      Top             =   3390
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   270
      Index           =   7
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   18
      Top             =   3390
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Index           =   1
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   1
      Top             =   2790
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060104_7.frx":0000
      Left            =   1080
      List            =   "frm060104_7.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   24
      Top             =   2385
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7704
      TabIndex        =   23
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "變更事項(&R)"
      Height          =   400
      Index           =   4
      Left            =   4428
      TabIndex        =   20
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5652
      TabIndex        =   21
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6480
      TabIndex        =   22
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   3
      Left            =   3204
      TabIndex        =   19
      Top             =   70
      Width           =   1200
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   4920
      MaxLength       =   9
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   2040
      MaxLength       =   9
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   4080
      MaxLength       =   7
      TabIndex        =   4
      Top             =   3090
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   1590
      MaxLength       =   1
      TabIndex        =   3
      Top             =   3090
      Width           =   375
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2790
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   0
      Text            =   "1"
      Top             =   2790
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   28
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   27
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   26
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   25
      Top             =   1080
      Width           =   375
   End
   Begin MSForms.TextBox Text12 
      Height          =   825
      Left            =   1200
      TabIndex        =   9
      Top             =   3690
      Width           =   5895
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "10557;1244"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7200
      TabIndex        =   8
      Top             =   3390
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;556"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "Email維護:             (Y:是)"
      Height          =   180
      Left            =   6480
      TabIndex        =   71
      Top             =   5025
      Width           =   1860
   End
   Begin VB.Label lblRecDate 
      AutoSize        =   -1  'True
      Caption         =   "當天請款:             (Y:是)"
      Height          =   210
      Left            =   4200
      TabIndex        =   70
      Top             =   5010
      Width           =   1815
   End
   Begin VB.Label lblPAID 
      AutoSize        =   -1  'True
      Caption         =   "已收款:           (1-不寄D/N, 2-寄D/N)"
      Height          =   180
      Left            =   750
      TabIndex        =   69
      Top             =   5025
      Width           =   2700
   End
   Begin VB.Label lblPayToday 
      AutoSize        =   -1  'True
      Caption         =   "電子送件是否當日扣款:           (Y/N)"
      Height          =   180
      Left            =   3240
      TabIndex        =   68
      Top             =   4695
      Width           =   2745
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:          (Y: 是)"
      Height          =   180
      Left            =   240
      TabIndex        =   67
      Top             =   4695
      Width           =   2085
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Instruction No. "
      Height          =   180
      Left            =   7230
      TabIndex        =   66
      Top             =   1140
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "是否特殊請款:"
      Height          =   180
      Left            =   6210
      TabIndex        =   65
      Top             =   3120
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(Y:是)"
      Height          =   180
      Left            =   7800
      TabIndex        =   64
      Top             =   3120
      Width           =   465
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6300
      TabIndex        =   63
      Top             =   3450
      Width           =   900
   End
   Begin VB.Label lblCP71 
      AutoSize        =   -1  'True
      Caption         =   "延緩公告發文：延緩　　　個月"
      Height          =   180
      Left            =   3255
      TabIndex        =   62
      Top             =   3435
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   240
      X2              =   8520
      Y1              =   2745
      Y2              =   2745
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   8520
      Y1              =   2775
      Y2              =   2775
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "下次繳費日:"
      Height          =   180
      Left            =   240
      TabIndex        =   61
      Top             =   3420
      Width           =   945
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   10
      Left            =   6060
      TabIndex        =   60
      Top             =   390
      Visible         =   0   'False
      Width           =   2610
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "4604;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   9
      Left            =   3240
      TabIndex        =   59
      Top             =   390
      Visible         =   0   'False
      Width           =   480
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5741;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   8
      Left            =   1740
      TabIndex        =   58
      Top             =   2415
      Width           =   6900
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "12171;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   7
      Left            =   1080
      TabIndex        =   57
      Top             =   2070
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   6
      Left            =   4920
      TabIndex        =   56
      Top             =   1740
      Width           =   3750
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6615;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   5
      Left            =   1080
      TabIndex        =   55
      Top             =   1740
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   54
      Top             =   1410
      Width           =   3750
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "6615;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   53
      Top             =   1410
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   52
      Top             =   1080
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   51
      Top             =   750
      Width           =   1410
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2487;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   50
      Top             =   750
      Width           =   2130
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "3757;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   240
      TabIndex        =   49
      Top             =   3690
      Width           =   765
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "年費通知人:"
      Height          =   180
      Left            =   3900
      TabIndex        =   48
      Top             =   390
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   1080
      TabIndex        =   47
      Top             =   390
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   3420
      TabIndex        =   46
      Top             =   3120
      Width           =   585
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "(Y:Word)"
      Height          =   180
      Left            =   2010
      TabIndex        =   45
      Top             =   3120
      Width           =   690
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書:"
      Height          =   180
      Left            =   240
      TabIndex        =   44
      Top             =   3120
      Width           =   1305
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "(Y:雙倍)"
      Height          =   180
      Left            =   6120
      TabIndex        =   43
      Top             =   2835
      Width           =   645
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "費用是否要雙倍:"
      Height          =   180
      Left            =   4260
      TabIndex        =   42
      Top             =   2835
      Width           =   1305
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "年 年費"
      Height          =   180
      Left            =   2640
      TabIndex        =   41
      Top             =   2835
      Width           =   585
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "至"
      Height          =   180
      Left            =   1800
      TabIndex        =   40
      Top             =   2835
      Width           =   180
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "繳納第:"
      Height          =   180
      Left            =   240
      TabIndex        =   39
      Top             =   2835
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   240
      TabIndex        =   38
      Top             =   750
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   4080
      TabIndex        =   37
      Top             =   750
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   36
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   4080
      TabIndex        =   35
      Top             =   1080
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Left            =   240
      TabIndex        =   34
      Top             =   1410
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Left            =   4080
      TabIndex        =   33
      Top             =   1410
      Width           =   675
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Left            =   240
      TabIndex        =   32
      Top             =   1740
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Left            =   4080
      TabIndex        =   31
      Top             =   1740
      Width           =   675
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Left            =   240
      TabIndex        =   30
      Top             =   2070
      Width           =   675
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   29
      Top             =   2385
      Width           =   765
   End
End
Attribute VB_Name = "frm060104_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/16 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/8/4 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, m_CP110 As String, m_AgentName As String

Dim intWhere As Integer, strSales As String
'Dim strCaseFee(1 To 2) As String 'strCaseFee(1) 國家檔中繳費年度，strCaseFee(2) 國家檔中起算日
Public strCaseFee1 As String 'strCaseFee1 國家檔中繳費年度
Public strCaseFee2 As String 'strCaseFee2 國家檔中起算日
'Add By Cheng 2002/07/04
Dim m_CP07 As String '法定期限
Dim m_CP10 As String '案件性質
Dim m_CP14 As String 'Add By Sindy 2016/11/16
'Add By Cheng 2003/01/01
Public m_strOfficalFee As String '規費
Dim m_strServiceFee As String '服務費
Dim m_strPoints As String '點數
'92.7.7 ADD BY SONIA
Dim m_EndDate As String     '專用期止日
Dim m_Pay As String         '是否有下次繳費日
'Add By Cheng 2003/09/01
Dim m_strNP09 As String
Dim m_strNP09_1 As String
'Add By Cheng 2003/10/06
Dim m_blnClkChgEvnBtn As Boolean '是否按下變更事項按鈕
'Add by Morgan 2004/6/24   延緩公告收文號,延緩月數
Dim m_str412CP09 As String, m_str412CP71 As String
'Modified by Lydia 2021/08/18 配合PUB_Get605NP模組
'Dim m_strDate(1 To 3) As String
Dim m_strDate(1 To 4) As String
Dim m_bolNew As Boolean '是否用新法
Dim m_bol412 As Boolean '是否有收延緩公告
'Dim m_bolSub As Boolean '新型是否可扣繳 'Removed by Morgan 2013/1/3 檢查統已無符合條件案件可取消
Public m_CP81 As String '可否減免
Public m_lngDisc As Long '減免金額
Public m_lngDisc1Year As Long '第一年減免金額 Add By Sindy 2020/3/31
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/20 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
'Add by Morgan 2010/3/25
Public m_lngFee1  As Long  '證書費+第一年年費(未減免)
Public m_lngFee2  As Long  '第二年以後年費(未減免)
Dim m_lngSub As Long  '抵減金額
Dim m_bolChanged As Boolean '是否有變更事項
Public m_DiscType As String 'Add by Morgan 2010/6/28 減免身分
Dim m_CP142 As String 'Add By Sindy 2015/12/17
Dim m_CP164 As String 'Add By Sindy 2021/4/20
'Added by Lydia 2018/09/11
Dim m_CP118 As String '是否電子送件
Dim m_CP82 As String '發文時間
Dim m_PrevForm As Form '前一畫面 Add By Sindy 2018/11/8
Public m_lngOfficalFee1 As Long '領證費
Public m_lngOfficalFee1Year As Long '第一年年費 Add By Sindy 2018/11/9
Dim m_strNA81Appl As String 'Add By Sindy 2019/1/22
Dim m_CP60 As String 'Added by Lydia 2019/10/01 請款編號
Dim m_AddMcRecord As String 'Added by Lydia 2020/08/17 人工Email維護(語法)
Dim m_eFlag As String 'Added by Lydia 2020/08/17 是否e/E化
Dim m_str412AddCP64 As String 'Added by Morgan 2022/12/27

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 10) As String, strTmp As String
'Add By Cheng 2002/10/24
Dim ii As Integer

   'Add By Cheng 2002/10/24
   ii = 1
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   If Text7(0).Text = Text7(1).Text Then
      strTmp = "第 " & Text7(0) & " 年年費"
   Else
      strTmp = "第 " & Text7(0) & " 至 " & Text7(1) & " 年年費"
   End If
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','第幾年至幾年費','" & strTmp & "')"
   ii = ii + 1
   
   If Text6 = "Y" Then
      '20140325START MODIFY By eric
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','列印備註','(逾期補繳)')"
      'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      '   "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      '   "','列印備註','逾期補繳')"
      '20140325END
      ii = ii + 1
      
      'Added by Morgan 2013/4/1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','申請回復領證','♀')"
      ii = ii + 1
   End If
   ' 91.10.25 MODIFY BY SONIA
   'Add By Cheng 2002/10/24
   'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '   "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
   '   "','規費','" & Val(PUB_GetYF07(pa(9), pa(8), "Y00000000", m_CP10, Me.Text7(0).Text, Me.Text7(1).Text)) / 2 & "')"
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','規費','" & Val(PUB_GetYF07(pa(9), pa(8), "Y00000000", m_CP10, Me.Text7(0).Text, Me.Text7(1).Text)) & "')"
   ' 91.10.25 END
   ii = ii + 1
   
'cancel by sonia 2019/2/20 敏莉:取消加註
'   'Add By Sindy 2019/1/22
'   If m_strNA81Appl <> "" Then
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'          "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'          "','外商國名申請人','" & m_strNA81Appl & "')"
'      ii = ii + 1
'   End If
'   '2019/1/22 END
'end 2019/2/20
   
   'Add by Morgan 2008/1/28
   If m_lngDisc > 0 Then
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','減免金額','" & m_lngDisc & "')"
      ii = ii + 1
      
      'Add by Morgan 2010/6/28
      strExc(1) = ""
      'Modified by Morgan 2012/7/9 中小企業要印條款
      'Modify By Sindy 2015/12/22 調整法規條款顯示字樣
      '   認定標準第2條第1項第1款之減免年費規定 ==> 認定標準第2條第1款之減免年費規定
      'Modified by Morgan 2018/1/19 減免說明再調整--吳若芬,王�E娟
      If InStr(m_DiscType, "1") > 0 And InStr(m_DiscType, "2") > 0 And InStr(m_DiscType, "3") > 0 Then
         'strExc(1) = "為自然人、學校及中小企業且符合中小企業認定標準第2條第1項第1款之減免年費規定"
         'strExc(1) = "為自然人、學校及中小企業且符合中小企業認定標準第2條第1款之減免年費規定"
         'Modify By Sindy 2019/5/30 修改為 第2條第1項第1款 => 第2條第1款
         strExc(1) = "為自然人、學校及中小企業且資格符合中小企業認定標準第2條第1款之規定"
      Else
         If InStr(m_DiscType, "1") > 0 Then
            
            strExc(1) = "為自然人"
            If InStr(m_DiscType, "2") > 0 Then
               'strExc(1) = strExc(1) & "及學校符合減免年費規定"
               strExc(1) = strExc(1) & "及學校"
            ElseIf InStr(m_DiscType, "3") > 0 Then
               'strExc(1) = strExc(1) & "及中小企業且符合中小企業認定標準第2條第1項第1款之減免年費規定"
               'strExc(1) = strExc(1) & "及中小企業且符合中小企業認定標準第2條第1款之減免年費規定"
               'Modify By Sindy 2019/5/30 修改為 第2條第1項第1款 => 第2條第1款
               strExc(1) = strExc(1) & "及中小企業且資格符合中小企業認定標準第2條第1款之規定"
            'Removed by Morgan 2018/1/19
            'Else
            '   strExc(1) = strExc(1) & "符合減免年費規定"
            'end 2018/1/19
            End If
            
         ElseIf InStr(m_DiscType, "2") > 0 Then
            
            strExc(1) = "為學校"
            If InStr(m_DiscType, "3") > 0 Then
               'strExc(1) = strExc(1) & "及中小企業且符合中小企業認定標準第2條第1項第1款之減免年費規定"
               'strExc(1) = strExc(1) & "及中小企業且符合中小企業認定標準第2條第1款之減免年費規定"
               'Modify By Sindy 2019/5/30 修改為 第2條第1項第1款 => 第2條第1款
               strExc(1) = strExc(1) & "及中小企業且資格符合中小企業認定標準第2條第1款之規定"
            'Removed by Morgan 2018/1/19
            'Else
            '   strExc(1) = strExc(1) & "符合減免年費規定"
            'end 2018/1/19
            End If
         ElseIf InStr(m_DiscType, "3") > 0 Then
            'strExc(1) = "符合中小企業認定標準第2條第1項第1款之減免年費規定"
            'strExc(1) = "符合中小企業認定標準第2條第1款之減免年費規定"
            'Modify By Sindy 2019/5/30 修改為 第2條第1項第1款 => 第2條第1款
            strExc(1) = "為中小企業且資格符合中小企業認定標準第2條第1款之規定"
         End If
      End If
      If strExc(1) <> "" Then strExc(1) = strExc(1) & "，依據專利年費減免辦法規定" 'Added by Morgan 2018/1/19
      
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','減免說明','" & strExc(1) & "')"
      ii = ii + 1
   End If
   
   'Add by Morgan 2004/6/24
   If m_bolNew Then
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
        "','機關文書','" & IIf(pa(8) = "2", "處分書", "審定書") & "')"
      ii = ii + 1
      If m_bol412 = True Then
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
           "','延緩月數','" & PUB_ChgNumber2Chinese(txtCP71) & "')"
         ii = ii + 1
      End If
   End If
   
   'edit by nickc 2007/02/05 不用 dll 了
    'If Not objLawDll.ExecSQL(ii - 1, strTxt) Then
    If Not ClsLawExecSQL(ii - 1, strTxt) Then
       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
    End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean
Dim strTmp As String
'Added by Lydia 2018/09/11
Dim strFilePath As String '記錄智慧局收文文號
Dim strNewCP64 As String '保留進度備註
'end 2018/09/11
Dim ii As Integer
Dim nFrm As Form  'Added by Lydia 2019/10/01
Dim str412FilePath As String 'Added by Morgan 2022/12/27

   Select Case Index
   
      'Modify by Morgan 2009/3/26 將同時發文併入
      'Case 0 '確定
      Case 0, 3
         'Add By Cheng 2002/01/29
         '檢查下次繳費日是否空白
         If Len(Trim(Text5(7).Text)) <= 0 And m_Pay <> "N" Then
            MsgBox "未計算出下次繳費日, 請檢查專利基本檔資料是否正確!!!", vbExclamation
            Exit Sub
         End If
         'Add By Cheng 2002/05/21
         If CheckDataValid = False Then
            Exit Sub
         End If
         
         'Added by Lydia 2020/08/17 Email維護; 從「上傳檔案到卷宗區」的下面搬上來
         m_AddMcRecord = ""
         If txtEmail.Visible = True And txtEmail.Text = "Y" Then
            strExc(5) = Pub_FcpSetPayToday("2", Text9.Text, txtPayToday.Text) '扣款日
            '開啟Email畫面
            Call PUB_GetFCPEmpMail("2", strReceiveNo, m_eFlag, textCP148, txtPAID, txtRecDate, IIf(Text6.Text = "Y", "逾期補繳", ""), strExc(5), strExc(1), strExc(2), strExc(3), strExc(4))
            If strExc(1) <> "" And strExc(2) <> "" Then
               frm880019.txtReceiver = strExc(1)
               frm880019.txtSubject = strExc(2)
               frm880019.txtContent = strExc(3)
               frm880019.txtCopy = strExc(4)
               frm880019.m_AddMailCache = "Y"
               frm880019.SetParent Me
               frm880019.Show vbModal
               m_AddMcRecord = frm880019.m_AddMailCache
               Unload frm880019
               'Modified by Lydia 2020/09/11
               'If m_AddMcRecord = "Y" Then
               If m_AddMcRecord = "" Or m_AddMcRecord = "Y" Then
                   MsgBox "Email維護未確認，請重新確認Email !", vbCritical, "檢核資料"
                   Exit Sub
               End If
            End If
         End If
         'end 2020/08/17
         
         'Added by Lydia 2018/09/11 是否電子送件
         strNewCP64 = Text12
         If txtCP118 = "Y" Then
            '電子送件也要記錄主管機關
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text9, , True) = False Then
               Exit Sub
            End If
            strExc(0) = InputBox("請輸入智慧局收文文號!!")
            If strExc(0) = "" Then
               Exit Sub
            Else
               strFilePath = strExc(0)  '記錄智慧局收文文號
               strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & Text12
            End If
            
            'Added by Morgan 2022/12/27
            If strSrvDate(1) >= "20230101" And m_str412CP09 <> "" Then
               strExc(0) = InputBox("請再輸入【延緩公告】智慧局收文文號!!")
               If strExc(0) = "" Then
                  Exit Sub
               Else
                  str412FilePath = strExc(0)
                  m_str412AddCP64 = "智慧局收文文號:" & strExc(0) & ";"
               End If
            End If
            'end 2022/12/27
         Else
         'end 2018/09/11
            'Add by Morgan 2009/4/28
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text9) = False Then
               Exit Sub
            End If
            If m_CP123s = "Y" Then
            'end 2009/4/28
               'Add by Morgan 2009/3/20 設定是否算發文室案件
               'modify by sonia 2014/6/23 加傳發文規費,一定有規費先用1, P-108903
               If ModifyDispatch(strReceiveNo, m_CP09s, m_CP123s, 1, Text9) = False Then
                   Exit Sub
               End If
               'end 2009/3/20
            End If
         End If 'end 2018/09/11
         
         'Added by Lydia 2018/09/11 依據輸入的智慧局收文號(受理號,ex: 1073066637-0)，將本機C:\E-SET\RdcDocDir\(收文號ex: 1073066637-0)的pdf檔自動搬移到卷宗區(by Phoebe);
         If txtCP118.Text = "Y" And strFilePath <> "" Then
             strExc(1) = m_CP82
             If Val(m_CP82) > 0 Then
                 If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                      strExc(1) = ""
                 End If
             End If
             If Val(strExc(1)) = 0 Then
                'Modified by Lydia 2019/03/22 +傳入發文日
                If Pub_AutoEsetToCpp(True, pa(1), pa(2), pa(3), pa(4), pa(8), Label2(0).Caption, m_CP10, strFilePath, Text9.Text) = False Then
                      Exit Sub
                End If
                
                'Added by Morgan 2022/12/27
                '延緩公告申請書
                If str412FilePath <> "" Then
                  If Pub_AutoEsetToCpp(True, pa(1), pa(2), pa(3), pa(4), pa(8), m_str412CP09, "412", str412FilePath, Text9.Text) = False Then
                      Exit Sub
                  End If
                End If
                'end 2022/12/27
             End If
         End If
         'end 2018/09/11
         
         'Added by Lydia 2018/09/11 檢查完畢，更新備註欄位
         Text12.Text = strNewCP64
         
         'Add by Sindy 2021/11/16 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         
         ' 設定滑鼠游標為等待狀態
         Screen.MousePointer = vbHourglass
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
         
         'Add by Morgan 2008/2/20 檢查代理人Email
         PUB_CheckEMail pa(75), pa(144)
         If pa(145) <> "" Then
            PUB_CheckEMail pa(75), pa(145)
         End If
         'end 2008/2/20
            
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
         'Modify By Sindy 2018/11/19
         '紙本送件,才需要在此處產生申請書
         If txtCP118.Text = "" Then
         '2018/11/19 END
            If Text8.Text = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            
            'Add By Sindy 2019/1/22 為取值,設假的定稿號碼
            ii = 0
            EndLetter "99", strReceiveNo, "99", strUserNum
            Call PUB_GetApplPA_EData("99", "99", strReceiveNo, pa(), , , , , m_strNA81Appl)
            '2019/1/22 EMD
            
            strTmp = "00"
            'Modify by Morgan 2004/6/24  新法無公告日定稿
            If m_bolNew Then
               If m_bol412 Then
                  'Modify by Morgan 2010/3/26 有延緩公告時改用新定稿
                  'strTmp = "09"
                  strTmp = "01"
               Else
                  'Modified by Morgan 2012/7/9 中小企業要印條款改新定稿
                  'strTmp = "08"
                  strTmp = "12"
               End If
                  
   'Removed by Morgan 2013/1/3 檢查統已無符合條件案件可取消
   '
   '            'Add by Morgan 2004/7/8
   '            '台灣新型93.7.1以前申請,93.7.1(含)以後核准的規費可減免1500
   '            'Modify by Morgan 2006/6/30 改抓存檔時的判斷
   '            'If (pA(8) = "2" And pA(9) = "000" And Val(pA(10)) < 930701 And Val(pA(20)) >= 930701) Then
   '            If m_bolSub = True Then
   '               strTmp = Format(Val(strTmp) + 2, "00")
   '            End If
   '
   'end 2013/1/3
               
               'Modify by Morgan 2010/3/26 有延緩公告時改用新定稿
               If m_bol412 Then
                  StartLetter1 "01", strTmp
               Else
                  StartLetter "01", strTmp
               End If
            Else
               StartLetter "01", strTmp
            End If
            NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum, 0
         End If
         
         'Added by Lydia 2019/10/01  領證/年費發文直接產生:承辦單+請款定稿+帳單(請款單)
         'Modified by Lydia 2020/08/17 重新發文不產生請款單
         'If m_CP60 = "" Then
         If m_CP60 = "" And Val(m_CP82) = 0 Then
            '檢查表單是否已開啟，若是，則關閉
            For Each nFrm In Forms
               If StrComp(nFrm.Name, "frm060307", vbTextCompare) = 0 Then
                  Unload frm060307
                  Exit For
               End If
            Next
            frm060307.m_KeyCP09 = strReceiveNo
            frm060307.m_KeyCP10 = m_CP10
            'Added by Lydia 2020/08/17
            Call frm060307.SetData(0, "1", True) '外部呼叫,預設類別
            Call frm060307.SetData(1, txtPAID.Text)  '已收款
            Call frm060307.SetData(2, txtRecDate.Text)  '當天請款
            Call frm060307.SetData(3, IIf(Text6.Text = "Y", "逾期補繳", "")) '逾期補繳(來源表單的設定之描述)
            If m_AddMcRecord <> "" Then
                Call frm060307.SetData(4, m_AddMcRecord)
            End If
            'end 2020/08/17
            frm060307.Show
            Call frm060307.cmdok_Click(0)
            'Modified by Lydia 2020/08/17 Transaction將發文和年證費請款函一併包入; 因為5/20的FCP762756發文程式失敗,經電腦中心測試m51存dn.pdf於typing2,造成有發文日卻無實際請款單
            'Unload frm060307
            If frm060307.m_bTransOK = True Then
                 cnnConnection.CommitTrans
                 Unload frm060307
            Else
                 cnnConnection.RollbackTrans
                 Unload frm060307
                 MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
            End If
            'end 2020/08/17
            
         'Added by Lydia 2020/08/17
         Else
             cnnConnection.CommitTrans
         'end 2020/08/17
         End If
         'end 2019/10/01
         
         If pa(1) = "FCP" Then
            'Add By Sindy 2016/11/16 特殊代理人彈訊息提醒
            If (PUB_GetST03(m_CP14) = "F21" Or PUB_GetST03(m_CP14) = "F51" Or PUB_GetST03(m_CP14) = "F52") And _
               Not (m_CP10 = "901" And m_CP10 = "902" And m_CP10 = "1202" And m_CP10 = "1002") Then
               strExc(0) = "select cp130 from caseprogress where cp09='" & strReceiveNo & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If "" & RsTemp.Fields(0) <> "" Then
                     If ChangeCustomerL(pa(75)) = "Y34440B30" Then
                        MsgBox "請當日優先請款報告!!"
                     End If
                  End If
               End If
            End If
         End If
         '2016/11/16 END
         
         'Add By Cheng 2002/04/30
         If Index = 0 Then
            'Add By Sindy 2023/11/9
            If frm060104_1.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2023/11/9 End
            '若有未發文資料顯示警告
            'Modify By Sindy 2023/11/9
            If PUB_GetCPunIssueDatas("" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)) Then
               frm060104_1.Show
               frm060104_1.ReQuery
            Else
               'Add By Sindy 2023/11/9
               If frm060104_1.bolIsEMPFlow = True Then
                  Unload frm060104_1
               Else
               '2023/11/9 End
                  frm060104_1.Show
                  frm060104_1.Clear
               End If
            End If
         Else
            frm060104_1.Show
            frm060104_1.ReQuery
         End If
         Unload Me
      Case 1
         frm060104_1.Show
         Unload Me
      Case 2
         'Add By Sindy 2018/11/8
         If UCase(TypeName(m_PrevForm)) <> UCase("frm06010310_1") Then
            Unload frm060104_1
         End If
         '2018/11/8 END
         Unload Me

      Case 4
         Me.Hide
         frm060104_5.LoadMe strReceiveNo, pa(1), pa(2), pa(3), pa(4), 7
         frm060104_5.Caption = "外專發文-變更事項"
         m_blnClkChgEvnBtn = True
         m_bolChanged = True
   End Select
End Sub

Private Function FormSave() As Boolean
Dim i As Integer, varTmp As Variant
Dim strTmp(0 To 5) As String, strTmp1(0 To 5) As String
Dim strMemo605 As String   '2011/10/12 add by sonia
Dim stCP118 As String, stCP152 As String 'Added by Lydia 2018/09/11
Dim strAgreeOnDate As String 'Add By Sindy 2021/8/17

   '911105 nick  transation
   FormSave = True
 On Error GoTo CheckingErr
 cnnConnection.BeginTrans
 
   Select Case pa(1)
      Case "FCP"
         Dim strFLD As String
         Dim nMaxNo As String
         Dim nPos As Integer
         Dim aryCurr As Variant
         Dim aryAll As Variant
         Dim aryDate As Variant
         Dim nPosBegin As Integer
         Dim nPosEnd As Integer
         Dim nDot As Integer
         Dim strPA72 As String
         Dim strPA73 As String
         Dim strPA74 As String
      
         ' 計算逗號的總數(幾格)
         nDot = 0
         For nPos = 1 To Len(pa(72))
            If Mid(pa(72), nPos, 1) = "," Then nDot = nDot + 1
         Next nPos
        
         aryAll = Split(strCaseFee2, ",")
         aryCurr = Split(pa(72), ",")
         ' 找尋繳年費起始點位置
         nPosBegin = 0
         For nPos = 0 To UBound(aryAll)
            If aryAll(nPos) = Text7(0) Then
               nPosBegin = nPos
               Exit For
            End If
         Next nPos
         ' 找尋繳年費終止點位置
         nPosEnd = 0
         For nPos = 0 To UBound(aryAll)
            If aryAll(nPos) = Text7(1) Then
               nPosEnd = nPos
               Exit For
            End If
         Next nPos
         ' 組繳年費年度字串
         strFLD = Empty
         For nPos = 0 To nPosEnd
            If nPos > 0 Then: strFLD = strFLD & ","
            strFLD = strFLD & aryAll(nPos)
         Next nPos
         If nDot > nPosEnd Then
            strFLD = strFLD & String(nDot - nPosEnd, ",")
         End If
         strPA72 = strFLD

         ' 重新計算繳費年度共有幾欄
         nDot = 0
         For nPos = 1 To Len(strPA72)
            If Mid(strPA72, nPos, 1) = "," Then nDot = nDot + 1
         Next nPos
          
         ' 繳年費日期
         ReDim aryCurr(nDot)
         If InStr(pa(73), ",") > 0 Then
            aryDate = Split(pa(73), ",")
            ' 拷貝原資料
            For nPos = 0 To UBound(aryDate)
               If IsEmptyText(aryDate(nPos)) = False Then
                  If nDot > 0 Then
                     aryCurr(nPos) = aryDate(nPos)
                  End If
               End If
            Next nPos
         End If
         ' 填入繳年費日期新資料
         For nPos = nPosBegin To nPosEnd
            aryCurr(nPos) = DBDATE(Text9)
         Next nPos
         ' 讀取繳年費日期新資料
         strFLD = Empty
         For nPos = 0 To UBound(aryCurr)
            If nPos > 0 Then: strFLD = strFLD & ","
            strFLD = strFLD & aryCurr(nPos)
         Next nPos
         strPA73 = strFLD
          
         '費用是否雙倍
         ReDim aryCurr(nDot)
         If InStr(pa(74), ",") > 0 Then
            Dim aryFee As Variant
            aryFee = Split(pa(74), ",")
            ' 拷貝原資料
            For nPos = 0 To UBound(aryFee)
               If IsEmptyText(aryFee(nPos)) = False Then
                  If nDot > 0 Then
                     aryCurr(nPos) = aryFee(nPos)
                  End If
               End If
            Next nPos
         End If
         ' 填入新資料
         For nPos = nPosBegin To nPosEnd
            If Text6 = "Y" Then
                '只有起始年上雙倍記號
                If nPos = nPosBegin Then
                    aryCurr(nPos) = "Y"
                Else
                    aryCurr(nPos) = ""
                End If
            Else
               aryCurr(nPos) = Empty
            End If
         Next nPos
         ' 讀取費用是否雙倍新資料
         strFLD = Empty
         For nPos = 0 To UBound(aryCurr)
            If nPos > 0 Then: strFLD = strFLD & ","
            strFLD = strFLD & aryCurr(nPos)
         Next nPos
         strPA74 = strFLD
         'For i = 1 To 3
         '   strTmp(i) = ""
         'Next
         'strTmp(0) = TransDate(Text9, 2)
         'For i = 1 To Text7(1)
         '   strTmp(1) = strTmp(1) & i & ","
         '   strTmp(2) = strTmp(2) & strTmp(0) & ","
         '   strTmp(3) = strTmp(3) & Text6 & ","
         'Next
         'For i = 1 To 3
         '   If Right(strTmp(i), 1) = "," Then strTmp(i) = Left(strTmp(i), Len(strTmp(i)) - 1)
         'Next
         
         'strExc(1) = "UPDATE PATENT SET PA76=" & CNULL(Text11) & "," & _
         '   "PA72=" & cnull(chgsql(strTmp(1))) & ",PA73=" & cnull(chgsql(strTmp(2))) & "," & _
         '   "PA74=" & cnull(chgsql(strTmp(3))) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
         'strExc(1) = "UPDATE PATENT SET PA72=" & cnull(chgsql(strTmp(1))) & ",PA73=" & cnull(chgsql(strTmp(2))) & "," & _
         '   "PA74=" & cnull(chgsql(strTmp(3))) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
         strExc(1) = "UPDATE PATENT SET PA72=" & CNULL(strPA72) & ",PA73=" & CNULL(strPA73) & "," & _
            "PA74=" & CNULL(strPA74) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      Case "FG"
         
   End Select
   
   '911105 nick transation
   cnnConnection.Execute strExc(1)
   
   'strExc(2) = "UPDATE CASEPROGRESS SET cp27=" & CNULL(TransDate(Text9, 2)) & "," & _
   '   "cp14=" & CNULL(Text10) & ",cp64=" & CNULL(Text12) & " WHERE CP09='" & strReceiveNo & "'"
   ' 91.03.25 modify by louis (單引號)
   'Modify By Cheng 2002/07/04
'   strExc(2) = "UPDATE CASEPROGRESS SET cp27=" & CNULL(TransDate(Text9, 2)) & "," & _
'      "cp64=" & CNULL(ChgSQL(Text12)) & " WHERE CP09='" & strReceiveNo & "'"
'   strExc(2) = "UPDATE CASEPROGRESS SET cp27=" & CNULL(TransDate(Text9, 2)) & "," & _
'      "cp64=" & CNULL(ChgSQL(Text12)) & ",CP17=" & Val(GetCF08(pa(1), pa(9), m_CP10)) * IIf(Me.Text6.Text = "Y", 2, 1) & " WHERE CP09='" & strReceiveNo & "'"
    'Add By Cheng 2003/01/01
    '取得領證及繳年費相關費用
    GetPatentYearFee pa(9), pa(8), "Y00000000", m_CP10, Me.Text7(0).Text, Me.Text7(1).Text, IIf(Me.Text6.Text = "Y", True, False)
   
   'Add By Sindy 2015/9/23 +Instruction No. 儲存在進度備註
   If Text13.Enabled = True And Text13.Text <> "" Then
      Text12 = Text12 & ";" & Trim(Label9.Caption) & " " & Trim(Text13.Text) & ";"
   End If
   '2015/9/23 END
   
   'Added by Lydia 2018/09/11
   '電子送件有規費的一律設自動扣款(同內專) --敏莉
   stCP118 = txtCP118
   stCP152 = ""
   If txtCP118 = "Y" And Val(m_strOfficalFee) > 0 Then
      stCP118 = "A"
      stCP152 = Pub_FcpSetPayToday("2", Text9.Text, txtPayToday.Text)
   End If
   'end 2018/09/11
   
   'Modify by morgan 2005/8/4 加 cp110
   'Modify by Moragn 2008/1/25 +CP81
   'MODIFY BY SONIA 2015/9/21 add cp14
   'Modify By Sindy 2015/9/23 +,CP148=" & CNULL(TextCP148) & "
   'Modified by Lydia 2018/09/11 +CP118,CP152
   strExc(2) = "UPDATE CASEPROGRESS SET cp27=" & CNULL(TransDate(Text9, 2)) & ",cp14=" & CNULL(Text10) & "," & _
      "cp64=" & CNULL(ChgSQL(Text12)) & ",CP16=" & Val(m_strServiceFee) + Val(m_strOfficalFee) & _
      ",CP17=" & Val(m_strOfficalFee) & ",CP18=" & Val(m_strPoints) & ", cp84=" & Format(Val(m_strOfficalFee)) & _
      ",cp110=" & CNULL(m_CP110) & ",CP22=NULL,CP81=" & CNULL(m_CP81) & ",CP148=" & CNULL(textCP148) & _
      ",CP118='" & stCP118 & "',CP152=" & CNULL(stCP152, True) & " WHERE CP09='" & strReceiveNo & "'"
   
   '911105 nick transation
   cnnConnection.Execute strExc(2)
   
   '92.7.7 ADD BY SONIA
   If Text5(7) <> "" Then
   '92.7.7 END
      'Modified by Morgan 2014/11/20 外專改回舊規則
      ''Added by Morgan 2014/10/29
      'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
      '   strTmp(0) = PUB_GetOurDeadline(Text5(7).Text)
      'Else
      ''end 2014/10/29

      'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
      If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
         'Modify By Sindy 2021/8/17 + , , strAgreeOnDate
         strTmp(0) = PUB_GetFCPOurDeadline(Text5(7), 2, , strAgreeOnDate)
      Else
      'end 2019/7/11
         
         strTmp(0) = CompDate(2, -2, TransDate(Text5(7).Text, 2))
         
      End If 'Added by Morgan 2019/7/11
      
      'End If 'Added by Morgan 2014/10/29
      'end 2014/11/20
      
      'Modify By Cheng 2002/07/04
   '   strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
   '      "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
   '      "','" & pa(4) & "','" & strSales & "'," & 年費 & "," & strTmp(0) & "," & TransDate(Text5(7), 2) & "," & objPublicData.GetNextProgressNo & ")"
      'edit by nickc 2007/02/02 不用 dll 了
      'strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
         "VALUES ('" & strReceiveNo & "','" & pA(1) & "','" & pA(2) & "','" & pA(3) & _
         "','" & pA(4) & "','" & PUB_GetFCPSalesNo(pA(1), pA(2), pA(3), pA(4)) & "'," & 年費 & "," & strTmp(0) & "," & TransDate(Text5(7), 2) & "," & objPublicData.GetNextProgressNo & ")"
      '2008/10/13 modify by sonia
      'strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
         "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
         "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & strTmp(0) & "," & TransDate(Text5(7), 2) & "," & GetNextProgressNo & ")"
      '2008/11/26 MODIFY BY SONIA 改為關係企業
      'If ChangeCustomerL(pa(75)) = "Y33944010" Then
      '2009/2/5 modify by sonia 若年費代理人非Y33944的關係企業則不掛np15,FCP-011795
      '2009/8/4 MODIFY BY SONIA 加Y48840,Y48196,Y20624
      '2009/8/20 MODIFY BY SONIA 加Y21099
      'Modify by Morgan 2011/3/22 改先存變數才不用重複抓相同資料
      'If Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y33944" Or _
         Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y48840" Or _
         Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y48196" Or _
         Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y20624" Or _
         Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y21099" Then
      'Modified by Morgan 2013/5/2 函數要傳9碼
      'strExc(9) = Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6)
      strExc(9) = PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1")
      'end 2013/5/2
      
'2011/10/12 MODIFY BY SONIA 改用MODULE
'      If strExc(9) = "Y33944" Or strExc(9) = "Y48840" Or strExc(9) = "Y48196" Or _
'         strExc(9) = "Y20624" Or strExc(9) = "Y21099" Then
'
'      '2008/11/26 END
'         strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP15,NP22) " & _
'            "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
'            "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & strTmp(0) & "," & TransDate(Text5(7), 2) & ",'信函要傳真;'," & GetNextProgressNo & ")"
'      '2009/8/4 ADD BY SONIA Y49083年費備註
'      'Modify by Morgan 2011/3/22 改先存變數才不用重複抓相同資料
'      'ElseIf Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y49083" Then
'      ElseIf strExc(9) = "Y49083" Then
'         strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP15,NP22) " & _
'            "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
'            "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & strTmp(0) & "," & TransDate(Text5(7), 2) & ",'只需銀龍加蓋年費回傳章;'," & GetNextProgressNo & ")"
'      '2009/8/4 END
'
'      'Add by Morgan 2011/3/22 2011.03.19 指示信--Susan
'      ElseIf strExc(9) = "Y30011" Then
'         strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP15,NP22) " & _
'            "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
'            "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & strTmp(0) & "," & TransDate(Text5(7), 2) & ",'年費函需以EMail傳送,不寄紙本;'," & GetNextProgressNo & ")"
'
'      Else
'         strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
'            "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
'            "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & strTmp(0) & "," & TransDate(Text5(7), 2) & "," & GetNextProgressNo & ")"
'      End If
      'Modified by Morgan 2012/6/4 +pa26
      'Modified by Morgan 2013/9/11 改抓設定檔
      'strMemo605 = PUB_Get605Memo(strExc(9), ChangeCustomerL(pa(26)), pa(1) & pa(2) & pa(3) & pa(4))
      'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
      'strMemo605 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), ChangeCustomerL(pa(26)))
      strMemo605 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30))
      
      'Remove by Morgan 2011/10/19 改控制代理人 Y47735xxx 所有案件 --譚文容
      ''催年費函需另cc:Y47740  2011/9/7年費發文有加此段但此畫面沒寫 --譚文容
      'Select Case pa(1) & pa(2) & pa(3) & pa(4)
      '   Case "FCP032841000", "FCP031045000"
      '      strMemo605 = "催年費函需另cc:Y47740;" & strMemo605
      'End Select
      'Modify By Sindy 2021/8/17 + ,NP23
      strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP15,NP22,NP23) " & _
         "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
         "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & strTmp(0) & "," & TransDate(Text5(7), 2) & ",'" & ChgSQL(strMemo605) & "'," & GetNextProgressNo & "," & CNULL(strAgreeOnDate, True) & ")"
'2011/10/12 END
      '2008/10/13 end
         
      '911105 nick transation
      cnnConnection.Execute strExc(3)
   '92.7.7 ADD BY SONIA
   End If
   '92.7.7 END
      
   '911105 nick transation
   'FormSave = objLawDll.ExecSQL(3, strExc)
   
   'Add by Morgan 2004/6/24 延緩公告發文
   If m_str412CP09 <> "" Then
      'Added by Morgan 2022/12/27
      If strSrvDate(1) >= "20230101" Then
         strSql = "UPDATE CASEPROGRESS SET CP27=" & CNULL(TransDate(Text9, 2)) & ",CP64='" & m_str412AddCP64 & "'||CP64,cp110=" & CNULL(m_CP110) & ",CP118='" & stCP118 & "',CP130='" & m_CP130 & "' WHERE CP09='" & m_str412CP09 & "'"
      Else
      'end 2022/12/27
         strSql = "UPDATE CASEPROGRESS SET CP27=" & CNULL(TransDate(Text9, 2)) & " WHERE CP09='" & m_str412CP09 & "'"
      End If
      cnnConnection.Execute strSql, intI
   End If
   
   'Added by Lydia 2018/09/11 FCP電子送件若發文時若有規費，則自動產生行事曆。
   If txtCP118 = "Y" And Val(m_strOfficalFee) > 0 And stCP152 <> "" Then
       If Pub_AddReceiptCalendar1(pa(1), pa(2), pa(3), pa(4), m_CP10, stCP152) = True Then
       End If
   End If
   'end 2018/09/11
   
   'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：若又收文年費，「通知資訊變更」尚未發文時，提醒承辦工程師已收文年費繳費，不需通知資訊變更。年費發文時，自動取消收文「通知資訊變更」
   If pa(177) = "Y" Then
      Call PUB_ChkFCPlinkYearFee(pa)
   End If
   'end 2023/07/28
   
   PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130 'Add by Morgan 2009/3/20
   
   'cnnConnection.CommitTrans 'Mark by Lydia 2020/08/17 Transaction將發文和年證費請款函一併包入
'911105 nick
   Exit Function
CheckingErr:
   cnnConnection.RollbackTrans
   FormSave = False
   
End Function

Private Function FillValue(ByVal strValue As String) As String
 Dim varTemp As Variant
   If strValue = "" Then
      FillValue = String(20, ",")
   Else
      varTemp = Split(strValue, ",")
      FillValue = strValue & String(19 - UBound(varTemp), ",")
   End If
End Function

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label2(8) = pa(5)
      Case "英"
         Label2(8) = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label2(8) = pa(7)
   End Select
End Sub

'Add By Sindy 2018/11/8
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub Form_Activate()
    'Add By Cheng 2003/10/06
    '若有按下變更事項按鈕, 則重新讀取資料
    If m_blnClkChgEvnBtn = True Then
        ReadPatent
        Label2(0) = strReceiveNo
        m_blnClkChgEvnBtn = False
    End If
End Sub

Private Sub Form_Load()
Dim varTmp As Variant
   MoveFormToCenter Me
   intWhere = 國外_FC
   'Add By Sindy 2018/11/8
   If UCase(TypeName(m_PrevForm)) = UCase("frm06010310_1") Then
      Me.Text1 = m_PrevForm.Text1
      Me.Text2 = m_PrevForm.Text2
      Me.Text3 = m_PrevForm.Text3
      Me.Text4 = m_PrevForm.Text4
      strReceiveNo = m_PrevForm.strReceiveNo
   Else
   '2018/11/8 END
      With frm060104_1
         Me.Text1 = .Text1
         Me.Text2 = .Text2
         Me.Text3 = .Text3
         Me.Text4 = .Text4
         strReceiveNo = .Tag
      End With
   End If
   'Add by Morgan 2005/8/4
   ReDim pa(TF_PA)
   ReadPatent
   'Add by Morgan 2005/8/4
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1300
   lstNameAgent.Width = 1300
   
   Label2(0) = strReceiveNo
   Combo1.ListIndex = 0
   'Modify By Sindy 2019/12/19
   'Text7(1) = "1"
   If Val(Text7(1)) = 0 Then Text7(1) = "1"
   '2019/12/19 END
   'modify by sonia 90.10.10
   varTmp = Split(strCaseFee2, ",")
   If Text7(1) > UBound(varTmp) + 1 Then
      MsgBox "繳費年度大於應繳年度，請查明後再輸入 !", vbCritical
   ElseIf Text7(1) = UBound(varTmp) + 1 Then
      Text5(7).Text = ""
   Else
      'Add by Morgan 2004/6/24   '台灣無公告日的用發文日預估期限
      If m_bolNew Then
         'Modified by Morgan 2014/11/20 +系統別參數
         PUB_Get605NP Text1.Text, Text9.Text, Text7(1).Text, m_strDate, Val(txtCP71.Text)
         Text5(7).Text = Format(Val(m_strDate(1)) - 19110000)
      Else
         '92.11.25 modify by sonia
         'Text5(7).Text = TransDate(CompDate(0, Val(varTmp(Val(Text7(1).Text) - 1)), strCaseFee1), 1)
         Text5(7).Text = TransDate(CompDate(2, -1, CompDate(0, Val(varTmp(Val(Text7(1).Text) - 1)), strCaseFee1)), 1)
         '92.11.25 end
      End If
   End If
    m_blnClkChgEvnBtn = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2020/08/17
   Set frm060104_7 = Nothing
End Sub

Private Sub ReadPatent()
 Dim Lbl As Object, txt As Object, i As Integer, strTmp1(0 To 5) As String
'Add By Cheng 2002/10/22
Dim strEndDate As String '專用結束日
Dim strCP43 As String 'Added by Morgan 2014/2/6

   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   Select Case pa(1)
      Case "FCP"
         If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
            For i = 3 To 7
               If pa(i + 23) <> "" Then ChgType (i)
            Next
            Label2(8) = pa(5)
            'If pa(76) <> "" Then Text11 = pa(76): ChgType (11)
            
            strTmp1(0) = strReceiveNo
            For i = 1 To 4
               strTmp1(i) = pa(i)
            Next
            '92.7.7 ADD BY SONIA
            If GetMoneyDate(Val(pa(8)) + 10, pa(9), strTmp1, strCaseFee1, strCaseFee2, strEndDate) = True Then
               m_EndDate = strEndDate
            End If
            '92.7.7 END
            'Modify By Cheng 2002/10/22
'            If GetMoneyDate(Val(pa(8)), pa(9), strTmp1, strCaseFee1, strCaseFee2) = True Then
            If GetMoneyDate(Val(pa(8)), pa(9), strTmp1, strCaseFee1, strCaseFee2, strEndDate) = True Then
               'Add By Cheng 2002/10/22
               'strCaseFee1 = strEndDate
            End If
            
         End If
      Case "FG"
         If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
            
         End If
   End Select
   'Modify By Cheng 2002/07/04
'   strExc(0) = "select cp13,st02,cp06,cp27,cp14,cp64 from caseprogress" & _
'      ",staff where cp09='" & strReceiveNo & "' and cp13=st01(+)"
   'Modified by Lydia 2018/09/11 +CP118,CP82
   'Modify By Sindy 2018/11/9 + ,CP71,CP53,CP54
   'Modified by Lydia 2019/10/01 +CP60
   'Modified by Lydia 2020/06/16 +CP148
   'Modified by Lydia 2020/08/17 +GetEmailFlag(CP01||CP02||CP03||CP04) as eFlag
   strExc(0) = "select cp13,st02,cp06,cp27,cp14,cp64,cp07,cp10,CP110,cp43,cp142,cp14, " & _
                    "CP118,CP82,CP71,CP53,CP54,CP60,CP148,GetEmailFlag(CP01||CP02||CP03||CP04) as eFlag,CP164 " & _
                    "from caseprogress,staff where cp09='" & strReceiveNo & "' and cp13=st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      strCP43 = "" & .Fields("cp43") 'Added by Morgan 2014/2/6
      m_CP142 = "" & .Fields("cp142") 'Add By Sindy 2015/12/17
      m_CP164 = "" & .Fields("CP164") 'Add By Sindy 2021/4/20
      m_CP14 = "" & .Fields("CP14") 'Add by Sindy 2016/11/16
      m_CP110 = "" & .Fields("CP110")
      If Not IsNull(.Fields(0)) Then strSales = .Fields(0)
      If Not IsNull(.Fields(1)) Then Label2(1) = .Fields(1)
      If Not IsNull(.Fields(2)) Then Label2(2) = TransDate(.Fields(2), 1)
      If Not IsNull(.Fields(3)) Then
         Text9 = TransDate(.Fields(3), 1)
      Else
         Text9 = strSrvDate(2)
      End If
      Text9.Tag = Text9.Text 'Added by Lydia 2020/08/26 預設發文日
      
      'modify by sonia 2015/9/21 畫面雖不顯示,但承辦人為外專程序時,改為操作人員
      If Not IsNull(.Fields(4)) Then Text10 = .Fields(4): ChgType (10)
      'end 2015/9/21
      If Not IsNull(.Fields(5)) Then
         Text12 = .Fields(5)
         'Add By Sindy 2015/9/23
         If InStr(Text12, Trim(Label9.Caption)) > 0 Then
            Text13.Enabled = False
         End If
         '2015/9/23 END
      End If
      'Add By Cheng 2002/07/04
      m_CP07 = "" & .Fields(6).Value
      m_CP10 = "" & .Fields(7).Value
      'Added by Lydia 2018/09/11
      m_CP118 = "" & .Fields("cp118")  '電子送件
      If m_CP118 <> "" Then txtCP118.Text = "Y"
      
      m_CP82 = "" & .Fields("cp82") '發文時間
      'end 2018/09/11
      'Add By Siny 2018/11/9
      If Not IsNull(.Fields("cp71")) Then txtCP71 = .Fields("cp71")
      If Not IsNull(.Fields("cp53")) Then Text7(0) = .Fields("cp53")
      If Not IsNull(.Fields("cp54")) Then Text7(1) = .Fields("cp54")
      '2018/11/9 END
      m_CP60 = "" & .Fields("cp60") 'Added by Lydia 2019/10/01
      textCP148 = "" & .Fields("cp148") 'Added by Lydia 2020/06/23 預設為進度檔的設定
      'Added by Lydia 2020/08/17  Email維護
      m_eFlag = "" & .Fields("eFlag")
      lblEmail.Visible = False: txtEmail.Visible = False
      If m_CP60 = "" And Val(m_CP82) = 0 Then
          lblEmail.Visible = True: txtEmail.Visible = True
      End If
      'end 2020/08/17
   End If
   End With
   
   'Add by Morgan 2004/6/24'檢查是否有延緩公告未發文
   chk412.Visible = False
   lblCP71.Visible = False
   txtCP71.Visible = False
   txtCP71.Text = ""
   m_bolNew = False: m_bol412 = False
   If pa(9) = 台灣國家代號 Then
      If Val(pa(14)) = 0 Or Val(pa(14)) >= 930701 Then
         m_bolNew = True
         m_bol412 = PUB_Get412Data(pa, m_str412CP09, m_str412CP71)
         If m_bol412 = True Then
            chk412.Enabled = False
            chk412.Visible = True
            lblCP71.Visible = True
            txtCP71.Visible = True
            'Modified by Morgan 2016/3/11 105/3/9日起延緩公告最長改6個月(原3個月)
            'txtCP71.Text = "3"
            txtCP71.Text = "6"
            'end 2016/3/11
         End If
      End If
   End If
   
   'Add by Morgan 2008/1/25 設定案件是否可減免
   m_CP81 = ""
'Modify by Morgan 2010/6/28
'   strExc(1) = "'" & ChangeCustomerL(pa(26)) & "'"
'   For intI = 2 To 5
'      If pa(25 + intI) <> "" Then
'         strExc(1) = strExc(1) & ",'" & ChangeCustomerL(pa(25 + intI)) & "'"
'      Else
'         Exit For
'      End If
'   Next
'   strExc(0) = "select cu15,cu01||cu02 CNo,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),nvl(cu04,cu06)) CName from customer where (cu01||cu02) in (" & strExc(1) & ")"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      m_CP81 = "Y"
'      With RsTemp
'      Do While Not .EOF
'         If IsNull(.Fields("cu15")) Then
'            m_CP81 = "N"
'            MsgBox "申請人 [" & .Fields("CNo") & " " & .Fields("CName") & "] 尚未設定身分，本案將設為不可減免！"
'            Exit Do
'         ElseIf "" & .Fields("cu15") = "1" Then
'            m_CP81 = "N"
'            Exit Do
'         End If
'         .MoveNext
'      Loop
'      End With
'   'Add by Morgan 2009/2/13 預設不可減免---靜芳
'   Else
'      MsgBox "本案尚未輸入申請人將設為不可減免！"
'      m_CP81 = "N"
'   End If
'   'End 2008/1/25
   If PUB_GetFCPCaseDiscState(pa(1) & pa(2) & pa(3) & pa(4), m_DiscType) Then
      m_CP81 = "Y"
   Else
      m_CP81 = "N"
   End If
'end 2010/6/28
   
   'Added by Morgan 2014/11/4
   If strCP43 = "" Then
      MsgBox "請先輸入相關收文號且為下一程序領證期限的相關收文號！", vbExclamation
   Else
   'end if
      'Added by Morgan 2013/4/1
      'Modified by Morgan 2014/2/6 領證期限可能有兩筆 Ex.FCP-033029
      'strExc(0) = "SELECT NP09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07=601"
      strExc(0) = "SELECT NP09 FROM NEXTPROGRESS WHERE NP01='" & strCP43 & "' AND NP07=601 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      With RsTemp
      If intI = 1 Then
           m_strNP09 = "" & .Fields(0).Value
           '若法定期限為假日時, 抓大於法定期限最近的工作天
           m_strNP09_1 = ""
           If m_strNP09 <> "" Then
               m_strNP09_1 = DBDATE(PUB_GetLawDay(DBDATE(m_strNP09)))
            End If
      'Modified by Morgan 2014/11/4
      Else
         MsgBox "請修改相關收文號為下一程序領證期限的相關收文號！", vbExclamation
      'end 2014/11/4
      End If
      End With
      'end 2013/4/1
   End If 'Added by Morgan 2014/11/4
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 0 '發文日
         'Modify By Cheng 2002/07/04
'         If ChkDate(Text9) Or Val(Text9) > Val(strSrvDate(2)) Then
'            ChgType = True
'         End If
         If Not ChkDate(Text9) Then
         ElseIf Val(Text9.Text) > PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1))) Then
            MsgBox "發文日大於系統日下一個工作日, 請重新輸入!!!", vbExclamation + vbOKOnly
         Else
            ChgType = True
         End If

      Case 3, 4, 5, 6, 7
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(pa(i + 23), strTempName) Then
         If ClsLawLawGetName(pa(i + 23), strTempName) Then
            Label2(i) = strTempName
            ChgType = True
         End If
      Case 10
         'ADD BY SONIA 2015/9/21 承辦人為外專程序時,改為操作人員
         Text10 = GetFCPUser(Text10)
         'END 2015/9/21
         If ClsPDGetStaff(Text10, strTempName) Then
            Label2(9) = strTempName
            ChgType = True
         End If
      Case 11
         'If objLawDll.LawGetName(Text11, strTempName) Then
         '   Label2(10) = strTempName
         '   ChgType = True
         'End If
   End Select
End Function

Private Sub Text10_GotFocus()
  TextInverse Text10
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then
      If Not ChgType(10) Then Cancel = True
   Else
      MsgBox "承辦人不可空白 !", vbCritical
      Cancel = True
   End If
End Sub

Private Sub Text11_GotFocus()
  TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   'If Not ChgType(11) Then Cancel = True
   'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
   If Cancel = False Then
      If PUB_CheckStatus(Text11.Text) = False Then Cancel = True
   End If
End Sub

Private Sub Text12_GotFocus()
  TextInverse Text12
End Sub

Private Sub Text6_GotFocus()
  TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
'Modified by Morgan 2013/4/1
'   strExc(0) = "SELECT NP09 FROM NEXTPROGRESS WHERE NP01='" & strReceiveNo & "' AND NP07=601"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'   With RsTemp
'   If intI = 1 Then
'        m_strNP09 = "" & .Fields(0).Value
'        'Add By Cheng 2003/09/01
'        '若法定期限為假日時, 抓大於法定期限最近的工作天
'        m_strNP09_1 = ""
'        If m_strNP09 <> "" Then
'            m_strNP09_1 = DBDATE(PUB_GetLawDay(DBDATE(m_strNP09)))
'        End If
'        'Modify By Cheng 2003/09/01
''        If Val(Text9) > Val(.Fields(0)) Then
'        If DBDATE(Text9) > IIf(DBDATE(m_strNP09) >= DBDATE(m_strNP09_1), DBDATE(m_strNP09), DBDATE(m_strNP09_1)) Then
'           MsgBox "發文日大於下一程序中領證之法定期限時必須為 Y !", vbCritical
'           Text6 = "Y"
'        End If
'   End If
'   End With
ChkDouble
'end 2013/4/1
End Sub

Private Sub Text7_GotFocus(Index As Integer)
  TextInverse Text7(Index)
End Sub

Private Sub Text7_Validate(Index As Integer, Cancel As Boolean)
 Dim i As Integer, bolChk As Boolean, varTmp As Variant
   m_Pay = "Y"    '92.7.7 ADD BY SONIA
   If Text7(Index) <> "" Then
      If Index = 1 Then
         If ChkRange(Text7(0), Text7(1), "繳費年度") = True Then
            For i = Text7(0) To Text7(1)
               If InStr(pa(72), Format(i)) > 0 Then
                  bolChk = True
                  Exit For
               End If
            Next
            If bolChk = True Then
               MsgBox "繳費年度錯誤，請查明後再輸入 !", vbCritical
               Cancel = True
            Else
               varTmp = Split(strCaseFee2, ",")
               If Text7(1) > UBound(varTmp) + 1 Then
                  MsgBox "繳費年度大於應繳年度，請查明後再輸入 !", vbCritical
                  Cancel = True
               ElseIf Text7(1) = UBound(varTmp) + 1 Then
                  Text5(7).Text = ""
               Else
                  'Add by Morgan 2011/7/1
                  If m_CP81 = "Y" And pa(8) = "3" And Val(Text7(1)) < 3 And Val(Text7(1)) <> UBound(varTmp) + 1 Then
                     If UBound(varTmp) + 1 < 3 Then
                        strExc(1) = UBound(varTmp) + 1
                     Else
                        strExc(1) = 3
                     End If
                     'Modified by Morgan 2022/7/12 Ex:FCP-066150--何淑華
                     'MsgBox "繳費年度請輸入 " & strExc(1) & " 以上(可減免客戶1~3年免繳年費)!!"
                     If MsgBox("申請人為個人或中小企業1~3年可免繳年費，確定只繳" & Val(Text7(1)) & "年？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        Cancel = True
                     End If
                     'end 2022/7/12
                  End If
                  
                  'Add by Morgan 2004/6/24   '新法無公告日用發文日預估期限
                  If m_bolNew = True Then
                     'Modified by Morgan 2014/11/20 +系統別參數
                     PUB_Get605NP Text1.Text, Text9.Text, Text7(1).Text, m_strDate, Val(txtCP71.Text)
                     Text5(7).Text = Format(Val(m_strDate(1)) - 19110000)
                  Else
                     '92.11.25 modify by sonia
                     'Text5(7).Text = TransDate(CompDate(0, Val(varTmp(Val(Text7(1).Text) - 1)), strCaseFee1), 1)
                     Text5(7).Text = TransDate(CompDate(2, -1, CompDate(0, Val(varTmp(Val(Text7(1).Text) - 1)), strCaseFee1)), 1)
                     '92.11.25 end
                     '92.7.7 ADD BY SONIA
                     If Val(Text5(7).Text) > Val(TransDate(m_EndDate, 1)) Then
                        Text5(7).Text = ""
                        m_Pay = "N"
                     End If
                     '92.7.7 END
                  End If
               End If
            End If
         Else
            Cancel = True
         End If
      End If
   Else
      MsgBox "年度不可空白 !", vbCritical
      TextInverse Text7(Index)
   End If
End Sub

Private Sub Text8_GotFocus()
  TextInverse Text8
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

Private Sub Text9_LostFocus()
'Modified by Morgan 2013/4/1
''Add By Cheng 2002/07/04
'If m_CP07 <> "" Then
'   '判斷發文日是否大於法定期限(若法定期限不為工作天, 則為其下一個工作天)
'   If Val(Me.Text9.Text) > PUB_GetLawDay(Val(m_CP07)) Then
'      '費用雙倍
'      Me.Text6.Text = "Y"
'   Else
'      '取消費用雙倍
'      Me.Text6.Text = ""
'   End If
'End If
ChkDouble
'end 2013/4/1
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> "" Then
      If Not ChgType(0) Then
            Cancel = True
      'Added by Lydia 2018/09/11 當發文日有改時
      Else
            If Text9.Tag <> Text9 Then
                  Text9.Tag = Text9
                  txtPayToday.Text = Pub_FcpSetPayToday("1", Text9.Text, txtCP118.Text)
            End If
      'end 2018/09/11
      End If
   Else
      MsgBox "發文日不可空白 !", vbCritical
      Cancel = True
   End If
End Sub

'Add By Cheng 2002/05/21
Private Function CheckDataValid() As Boolean
Dim i As Integer, bolChk As Boolean, varTmp As Variant
Dim Cancel As Boolean

CheckDataValid = False
'檢查繳納年費年數
If Text7(0) = "" Then
   MsgBox "年度不可空白 !", vbCritical
   Me.Text7(0).SetFocus
   Text7_GotFocus 0
   Exit Function
End If

Text7_Validate 0, Cancel
If Cancel = True Then Exit Function

If Text7(1) = "" Then
   MsgBox "年度不可空白 !", vbCritical
   Me.Text7(1).SetFocus
   Text7_GotFocus 1
   Exit Function
End If

Text7_Validate 1, Cancel
If Cancel = True Then Exit Function

   m_Pay = "Y"    '92.7.7 ADD BY SONIA
   If ChkRange(Text7(0), Text7(1), "繳費年度") = True Then
   For i = Text7(0) To Text7(1)
      If InStr(pa(72), Format(i)) > 0 Then
         bolChk = True
         Exit For
      End If
   Next
   If bolChk = True Then
      MsgBox "繳費年度錯誤，請查明後再輸入 !", vbCritical
      Me.Text7(1).SetFocus
      Text7_GotFocus 1
      Exit Function
   Else
      varTmp = Split(strCaseFee2, ",")
      If Text7(1) > UBound(varTmp) + 1 Then
         MsgBox "繳費年度大於應繳年度，請查明後再輸入 !", vbCritical
         Me.Text7(1).SetFocus
         Text7_GotFocus 1
         Exit Function
      ElseIf Text7(1) = UBound(varTmp) + 1 Then
         Text5(7).Text = ""
      Else
         'Add by Morgan 2004/6/24   '台灣無公告日的用發文日預估期限
         If m_bolNew = True Then
            'Modified by Morgan 2014/11/20 +系統別參數
            PUB_Get605NP Text1.Text, Text9.Text, Text7(1).Text, m_strDate, Val(txtCP71.Text)
            Text5(7).Text = Format(Val(m_strDate(1)) - 19110000)
         Else
            '92.11.25 modify by sonia
            'Text5(7).Text = TransDate(CompDate(0, Val(varTmp(Val(Text7(1).Text) - 1)), strCaseFee1), 1)
            Text5(7).Text = TransDate(CompDate(2, -1, CompDate(0, Val(varTmp(Val(Text7(1).Text) - 1)), strCaseFee1)), 1)
            '92.11.25 end
            '92.7.7 ADD BY SONIA
            If Val(Text5(7).Text) > Val(TransDate(m_EndDate, 1)) Then
               Text5(7).Text = ""
               m_Pay = "N"
            End If
            '92.7.7 END
         End If
      End If
   End If
Else
   Me.Text7(1).SetFocus
   Text7_GotFocus 1
   Exit Function
End If
'檢查發文日
If Text9 <> "" Then
   If Not ChgType(0) Then
      Me.Text9.SetFocus
      Text9_GotFocus
      Exit Function
   End If
Else
   MsgBox "發文日不可空白 !", vbCritical
   Me.Text9.SetFocus
   Text9_GotFocus
   Exit Function
End If

   'Add by Morgan 2005/8/4
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
'Removed by Morgan 2014/1/7 與下面檢查重複(要抓最近一個工作日判斷)
'   'Added by Morgan 2013/2/27
'   If m_CP07 < DBDATE(Text9) And Text6 = "" Then
'      MsgBox "已過法定期限,費用是否雙倍應上Y!!"
'      Text6.SetFocus
'      Exit Function
'   End If
   
   'Added by Morgan 2013/4/1
   If ChkDouble(True) = False Then
      Text6.SetFocus
      Exit Function
   End If
   
   'Add By Sindy 2015/12/17 檢查是否有指定送件日期,若有不可小於指定日期送件
   If m_CP142 <> "" Then
      'Modified by Morgan 2017/3/14 可能會前一天作業,改判斷畫面上的發文日
      'Modify By Sindy 2021/11/11 淑華說之後可以含當天發文
      If m_CP142 > strSrvDate(1) Then
      'If m_CP142 >= DBDATE(Text9) Then
         'Add By Sindy 2021/4/20
         'Modify By Sindy 2021/10/20 + 3.之後
         If ((m_CP164 = "1" Or m_CP164 = "") And m_CP142 > DBDATE(Text9)) Or _
            m_CP164 = "3" Then '1.當天 3.之後
         '2021/4/20 END
            MsgBox "有指定送件日期（" & ChangeWStringToTDateString(m_CP142) & "），不可提前送件!!!"
            Exit Function
         End If
      End If
   End If
   '2015/12/17 END
   
   'Added by Lydia 2018/09/11
   If txtCP118 = "Y" Then
      If txtPayToday = "" Then
         MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
         txtPayToday.SetFocus
         Exit Function
      End If
   End If
   'end 2018/09/11
   
   'Added by Morgan 2025/8/11 --敏莉
   If textCP148 = "Y" Then
      If PUB_Get601605SpecDN(m_CP10, pa(1), pa(2), pa(3), pa(4), pa(8), pa(72), pa(26), pa(75), , , , , , , strExc(1)) = True Then
         MsgBox strExc(1) & "，此處不可再設特殊請款！", vbExclamation + vbOKOnly
         textCP148 = ""
         textCP148.SetFocus
         Exit Function
      End If
   End If
   'end 2025/8/11
   
CheckDataValid = True
End Function

'Add By Cheng 2002/12/31
'計算相關費用
Public Sub GetPatentYearFee( _
    strYF01 As String, strYF02 As String, strYF03 As String, _
    strYF04 As String, strYF05From As String, strYF05To As String, blnDouble As Boolean)
'strYF01  申請國家
'strYF02  專利種類
'strYF03  代理人
'strYF04  案件性質
'strYF05From  起始年度
'strYF05To  終止年度
'blnDouble  規費是否雙倍

Dim rsA As ADODB.Recordset
Dim StrSQLa As String
'Add by Morgan 2005/2/1 新型申請日
Dim stUtiAppDate As String
Dim ii As Integer
Dim iYear As Integer '繳費年度
Dim lngDDate As Long
Dim bol802 As Boolean
Dim bolNoServiceFee As Boolean '免收服務費

   m_strOfficalFee = 0
   m_strServiceFee = 0
   m_strPoints = 0
   m_lngDisc = 0
   m_lngFee1 = 0
   m_lngFee2 = 0
   m_lngDisc1Year = 0 '第一年減免金額 Add By Sindy 2020/3/31
   
    bol802 = PUB_ChkCPExist(pa, "802")
    
    '若案件性質為領證及繳年費, 則先取得領證相關費用
    If strYF04 = 領證及繳年費 Then
        StrSQLa = "Select * From PatentYearFee Where YF01='" & strYF01 & "' AND YF02='" & strYF02 & "' AND YF03='" & strYF03 & "' AND YF04='" & strYF04 & "' "
        intI = 1
        Set rsA = ClsLawReadRstMsg(intI, StrSQLa)
        If intI = 1 Then
            m_lngOfficalFee1 = Val(m_lngOfficalFee1) + Val("" & rsA.Fields("YF07").Value)
            'Modify by Morgan 2004/9/6 領證費不雙倍,年費才要
            '領證規費是否雙倍
            'If blnDouble = True Then m_strOfficalFee = Val(m_strOfficalFee) * 2
            m_strServiceFee = Val(m_strServiceFee) + Val("" & rsA.Fields("YF06").Value)
        End If
        
    End If
   
   m_strOfficalFee = m_lngOfficalFee1
   
   ii = 1
   '取得案件性質為年費的相關費用
   StrSQLa = "Select * From PatentYearFee Where YF01='" & strYF01 & "' AND YF02='" & strYF02 & "' AND YF03='" & strYF03 & "' AND YF04='" & 年費 & "' AND YF05>=" & Val(strYF05From) & " AND YF05<=" & Val(strYF05To) & " Order By YF05 "
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, StrSQLa)
   If intI = 1 Then
      rsA.MoveFirst
       While Not rsA.EOF
         bolNoServiceFee = False 'Add by Morgan 2011/7/27
         iYear = Val("" & rsA.Fields("YF05").Value)
         
         'Add by Morgan 2010/3/25
         If iYear = 1 Then
            m_lngOfficalFee1Year = Val("" & rsA.Fields("YF07").Value) '第一年年費 Add By Sindy 2018/11/9
            m_lngFee1 = Val(m_lngOfficalFee1) + Val("" & rsA.Fields("YF07").Value)
         Else
            m_lngFee2 = m_lngFee2 + Val("" & rsA.Fields("YF07").Value)
         End If
      
            m_strOfficalFee = Val(m_strOfficalFee) + Val("" & rsA.Fields("YF07").Value)
            'Add by Morgan 2008/1/25 考慮個人可減免
            If m_CP81 = "Y" Then
               'Modify by Morgan 2008/2/26 要考慮被異議 FCP-027155
               lngDDate = Val(pa(14))
               If lngDDate > 0 Then
                 '若被異議用該年費有效起日判斷
                 If bol802 = True Then
                    lngDDate = Val(pa(14)) + (iYear - 1) * 10000#
                 '未被異議用該年費有效迄日判斷
                 Else
                    lngDDate = Val(pa(14)) + (iYear) * 10000# - 1
                 End If
               End If
          
               '沒有公告日或公告日+繳費年度>930701的才能減免
               If (lngDDate = 0 Or lngDDate > 930701) Then
                  If iYear >= 1 And iYear <= 3 Then
                     m_strOfficalFee = m_strOfficalFee - 800
                     'Add By Sindy 2020/3/31
                     If iYear = 1 Then
                        m_lngDisc1Year = 800 '第一年減免金額
                     End If
                     '2020/3/31 END
                     m_lngDisc = m_lngDisc + 800
                     'Add by Morgan 2011/7/27 '設計可減免客戶2,3年免服務費
                     If iYear > 1 And Val("" & rsA.Fields("YF07").Value) = 800 Then
                        bolNoServiceFee = True
                     End If
                  ElseIf iYear >= 4 And iYear <= 6 Then
                     m_strOfficalFee = m_strOfficalFee - 1200
                     m_lngDisc = m_lngDisc + 1200
                  End If
               End If
            End If
               
           'Modify by Morgan 2004/9/6 領證費不雙倍,年費才要
           'm_strOfficalFee = Val(m_strOfficalFee) + Val("" & rsA.Fields("YF07").Value)
           '領證規費是否雙倍
           If blnDouble = True And ii = 1 Then
               'Modified by Morgan 2013/2/27
               'm_strOfficalFee = m_strOfficalFee * 2#
               m_strOfficalFee = m_strOfficalFee + Val("" & rsA.Fields("YF07").Value) - m_lngDisc
               'end 2013/2/27
               m_lngDisc = m_lngDisc * 2# 'Add by Morgan 2008/1/25
           End If
           'End
            
            If bolNoServiceFee = False Then '設計可減免客戶2,3年免服務費 Add by Morgan 2011/7/27
               m_strServiceFee = Val(m_strServiceFee) + Val("" & rsA.Fields("YF06").Value)
            End If
           rsA.MoveNext
           ii = ii + 1
       Wend
   End If
   
   m_strPoints = Val(m_strServiceFee) / 1000
    
'Removed by Morgan 2013/1/3 檢查統已無符合條件案件可取消
'
'    'Add by Morgan 2004/7/8
'    '台灣新型93.7.1以前申請,93.7.1(含)以後核准的規費可減免1500
''Modify by Morgan 2005/2/1 需判斷是否有改請或改請延期
''    If (pa(8) = "2" And pa(9) = "000" And Val(pa(10)) < 930701 And Val(pa(20)) >= 930701) Then
''      m_strOfficalFee = m_strOfficalFee - 1500
''    End If
'   m_bolSub = False
'   m_lngSub = 0
'   'Modify by Morgan 2011/5/18 日期全部都要轉西元年比較才不用考慮來源格式問題
'   If (pa(8) = "2" And pa(9) = "000" And DBDATE(pa(20)) >= DBDATE(930701)) Then
'      stUtiAppDate = PUB_UtiAppDate(pa, pa(10))
'      If DBDATE(stUtiAppDate) < DBDATE(930701) Then
'         m_strOfficalFee = m_strOfficalFee - 1500
'         m_bolSub = True
'         m_lngSub = 1500
'      End If
'   End If
''2005/2/1 end
'
'end 2013/1/3

   Set rsA = Nothing
End Sub

'Add by Morgan 2005/8/4
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   m_CP110 = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      MsgBox "出名代理人不可空白！", vbExclamation
   Else
      If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub

Private Sub StartLetter1(ByVal ET01 As String, ByVal ET03 As String)

   Dim strTxt(1 To 30) As String, strTmp As String, iItemNo As String
   Dim ii As Integer
   
    ii = 0
    EndLetter ET01, strReceiveNo, ET03, strUserNum
    
    If m_lngFee1 > 0 Then
          
'Removed by Morgan 2013/1/3 檢查統已無符合條件案件可取消
'
'      If m_bolSub Then
'        ii = ii + 1
'        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'          "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'          "','抵減金額','" & m_lngSub & "')"
'        ii = ii + 1
'        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'          "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'          "','抵減不印','♀')"
'      End If
'
'end 2013/1/3
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
        "','費用1','" & Format(m_lngFee1 - m_lngSub) & "')"
      
      If Val(Text7(1)) > 1 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選1','■ ')"
      
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
           "','年費迄年','" & PUB_ChgNumber2Chinese(Text7(1)) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
           "','費用2','" & Format(m_lngFee2) & "')"
        
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選1','□ ')"
      End If
      
      If m_lngDisc > 0 Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選2','■ ')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
           "','減免迄年','" & PUB_ChgNumber2Chinese(IIf(Val(Text7(1).Text) > 6, "6", Text7(1).Text)) & "')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
           "','減免金額','" & Format(m_lngDisc) & "')"
         
         '自然人=1
         ii = ii + 1
         'Modify by Morgan 2010/6/28 +其他減免身分
         'strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選5','■ ')"
         
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選5','" & IIf(InStr(m_DiscType, "1") > 0, "■ ", "□ ") & "')"
         
         
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選6','" & IIf(InStr(m_DiscType, "2") > 0, "■ ", "□ ") & "')"
         
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選7','" & IIf(InStr(m_DiscType, "3") > 0, "■ ", "□ ") & "')"
         'end 2010/6/28
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選2','□ ')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選5','□ ')"
      End If
   End If

   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
       "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
       "','規費','" & m_strOfficalFee & "')"
   
'cancel by sonia 2019/2/20 敏莉:取消加註
'   'Add By Sindy 2019/1/22
'   If m_strNA81Appl <> "" Then
'      ii = ii + 1
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'          "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'          "','外商國名申請人','" & m_strNA81Appl & "')"
'   End If
'   '2019/1/22 END
'end 2019/2/20
   
   'Added by Morgan 2013/4/1
   If Text6 = "Y" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','列印備註','逾期補繳')"
         
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','申請回復領證','♀')"
   End If
   'end 2013/4/1

   If m_bolNew Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
        "','機關文書','" & IIf(pa(8) = "2", "處分書", "審定書") & "')"
        
      iItemNo = 1
      If m_bol412 = True Then
         iItemNo = iItemNo + 1
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
           "','延緩公告申請','（暨申請延緩公告）" & "')"
         'Modified by Morgan 2012/10/17
         'strTmp = PUB_ChgNumber2Chinese(iItemNo) & "、申請延緩公告" & vbCrLf & "　　　1.延緩公告期限：" & txtCP71 & "個月" & vbCrLf & "　　　2.延緩公告理由：向國外申請專利需要"
         strTmp = PUB_ChgNumber2Chinese(iItemNo) & "、申請延緩公告" & vbCrLf & "　　　1.延緩公告期限：" & txtCP71 & "個月" & vbCrLf & "　　　2.延緩公告理由：因申請人商業需要"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
              "','延緩公告','" & strTmp & "')"
      End If

      '變更事項
      If m_bolChanged = True Then
         iItemNo = iItemNo + 1
         strTmp = GetContent(strReceiveNo, PUB_ChgNumber2Chinese(iItemNo))
         If strTmp <> "" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
              "','變更事項','" & strTmp & "')"
         End If
      End If
   End If

   If Not ClsLawExecSQL(ii, strTxt) Then
       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub


Private Function GetContent(ByVal stRecNo As String, Optional p_stNum As String = "三") As String

   Dim iNo As Integer '項目
   Dim i As Integer
   Dim stItem As String
   Dim strTemp As String
   
   GetContent = "": iNo = 0
   
On Error GoTo ErrHnd

   strSql = "select * from CHANGEEVENT where CE01='" & stRecNo & "'"
      
   CheckOC
   With adoRecordset
   
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         strTemp = ""
         '地址
         For i = 23 To 37
            stItem = "" & .Fields("CE" & Format(i))
            If stItem <> "" Then
               strTemp = strTemp & Chr(13) & "　　　" & "■ 地址（新址：" & Left(stItem, 20)
               If Len(stItem) > 20 Then
                  strTemp = strTemp & Chr(13) & "　　　　" & Mid(stItem, 21)
               End If
               strTemp = strTemp & "）"
               Exit For
            End If
         Next
         '公司代表人
         For i = 10 To 15
            stItem = "" & .Fields("CE" & Format(i))
            If stItem <> "" Then
               strTemp = strTemp & Chr(13) & "　　　" & "■ 公司代表人（附公司執照影本一份）"
               Exit For
            End If
         Next
         If i = 16 Then
            For i = 68 To 91
               stItem = "" & .Fields("CE" & Format(i))
               If stItem <> "" Then
                  strTemp = strTemp & Chr(13) & "　　　" & "■ 公司代表人（附公司執照影本一份）"
                  Exit For
               End If
            Next
         End If
         
         If strTemp <> "" Then
            iNo = iNo + 1
            GetContent = Chr(13) & "　　" & Format(iNo) & ".免繳規費事項：" & strTemp
            strTemp = ""
         End If
         
         '姓名或公司名稱
         For i = 4 To 8
            stItem = "" & .Fields("CE" & Format(i, "00"))
            If stItem <> "" Then
               strTemp = strTemp & Chr(13) & "　　　" & "■ 姓名或公司名稱（附證明文件一份）"
               Exit For
            End If
         Next
         '印章(申請人或代表人)
         stItem = "" & .Fields("CE51") & .Fields("CE53")
         If stItem <> "" Then
            strTemp = strTemp & Chr(13) & "　　　" & "■ 印章（附切結書及身分證或公司執照影本一份）"
         End If
         '代理人
         stItem = "" & .Fields("CE55")
         If stItem <> "" Then
            strTemp = strTemp & Chr(13) & "　　　" & "■ 代理人（附委任書一份）"
         End If
         
         If strTemp <> "" Then
            iNo = iNo + 1
            GetContent = GetContent & Chr(13) & "　　" & Format(iNo) & ".應繳規費三ＯＯ元事項：" & strTemp
            strTemp = ""
         End If
         
         If GetContent <> "" Then
            GetContent = p_stNum & "、變更事項" & GetContent
         End If
         
      End If
   
ErrHnd:

      If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
      
   End With
   
   CheckOC
End Function
'Added by Morgan 2013/4/1
'Modified by Morgan 2022/10/7
Public Function ChkDouble(Optional pbolSaveCheck As Boolean) As Boolean
   If Text9 <> "" And m_strNP09_1 <> "" Then
      If DBDATE(Text9) > m_strNP09_1 And Text6 <> "Y" Then
         MsgBox "發文日大於下一程序中領證之法定期限時費用是否要雙倍必須為 Y !", vbCritical
         If pbolSaveCheck = False Then
            Text6 = "Y"
         End If
      Else
         ChkDouble = True
      End If
   'Added by Morgan 2014/11/4
   ElseIf m_strNP09_1 = "" Then
      MsgBox "無法讀取原法定期限，請檢查相關總收文號是否為下一程序領證期限的相關收文號！", vbExclamation
   'end 2014/11/4
   End If
End Function
'end 2013/4/1

'Add By Sindy 2015/9/23
Private Sub TextCP148_GotFocus()
  TextInverse textCP148
  CloseIme
End Sub
Private Sub TextCP148_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub Text13_GotFocus()
  TextInverse Text13
  CloseIme
End Sub
Private Sub Text13_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub
'2015/9/23 END

'Added by Lydia 2018/09/11
Private Sub txtCP118_GotFocus()
   TextInverse txtCP118
   CloseIme
End Sub

Private Sub txtCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP118_Change()
    txtPayToday.Text = Pub_FcpSetPayToday("1", Text9.Text, txtCP118.Text)
End Sub

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
'end 2018/09/11

'Added by Lydia 2020/08/17
Private Sub txtPAID_GotFocus()
   TextInverse txtPAID
End Sub

Private Sub txtPAID_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
End Sub

Private Sub txtRecDate_GotFocus()
    TextInverse txtRecDate
End Sub

Private Sub txtRecDate_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtEmail_GotFocus()
    TextInverse txtEmail
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtRecDate_Validate(Cancel As Boolean)
   If txtRecDate.Tag <> txtRecDate.Text Then
       If txtRecDate = "Y" And textCP148 = "Y" Then
            txtEmail = "Y"
       End If
   End If
   txtRecDate.Tag = txtRecDate.Text
End Sub

Private Sub TextCP148_Validate(Cancel As Boolean)
   If textCP148.Tag <> textCP148.Text Then
       If txtRecDate = "Y" And textCP148 = "Y" Then
            txtEmail = "Y"
       End If
   End If
   textCP148.Tag = textCP148.Text
End Sub
