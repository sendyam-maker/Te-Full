VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010503_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "核駁函輸入"
   ClientHeight    =   5520
   ClientLeft      =   -2688
   ClientTop       =   2880
   ClientWidth     =   8460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   8460
   Begin VB.TextBox txtIDSPt 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   2
      Left            =   7770
      MaxLength       =   3
      TabIndex        =   29
      Top             =   5220
      Width           =   375
   End
   Begin VB.TextBox txtIDSFee 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   2
      Left            =   6840
      MaxLength       =   6
      TabIndex        =   28
      Top             =   5220
      Width           =   765
   End
   Begin VB.TextBox txtIDSPt 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   1
      Left            =   7770
      MaxLength       =   3
      TabIndex        =   27
      Top             =   4920
      Width           =   375
   End
   Begin VB.TextBox txtIDSFee 
      Alignment       =   2  '置中對齊
      Height          =   270
      Index           =   1
      Left            =   6840
      MaxLength       =   6
      TabIndex        =   26
      Top             =   4920
      Width           =   765
   End
   Begin VB.TextBox txtSNP23 
      Height          =   270
      Left            =   4980
      MaxLength       =   7
      TabIndex        =   73
      Top             =   4275
      Width           =   1215
   End
   Begin VB.TextBox Text22 
      Height          =   270
      Left            =   1245
      MaxLength       =   5
      TabIndex        =   7
      Top             =   2850
      Width           =   735
   End
   Begin VB.TextBox txtFee 
      Height          =   270
      Left            =   4680
      TabIndex        =   17
      Top             =   3630
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtPt 
      Height          =   270
      Left            =   6300
      TabIndex        =   18
      Top             =   3630
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.TextBox Text21 
      Height          =   270
      Left            =   6615
      MaxLength       =   15
      TabIndex        =   6
      Top             =   2550
      Width           =   1725
   End
   Begin VB.TextBox txtDispDate 
      Height          =   270
      Left            =   5175
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1950
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text20 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1455
      MaxLength       =   20
      TabIndex        =   25
      Top             =   5190
      Width           =   4215
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1245
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2235
      Width           =   7095
   End
   Begin VB.TextBox Text15 
      Height          =   270
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   24
      Top             =   4890
      Width           =   255
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   4125
      TabIndex        =   39
      Top             =   3090
      Width           =   4215
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   15
         Top             =   150
         Width           =   975
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   760
         MaxLength       =   2
         TabIndex        =   11
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   13
         Top             =   150
         Width           =   375
      End
      Begin VB.OptionButton Option4 
         Caption         =   "文到          天"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Value           =   -1  'True
         Width           =   1305
      End
      Begin VB.OptionButton Option4 
         Caption         =   "         月"
         Height          =   180
         Index           =   1
         Left            =   1530
         TabIndex        =   12
         Top             =   180
         Width           =   915
      End
      Begin VB.OptionButton Option4 
         Caption         =   "                      日"
         Height          =   225
         Index           =   2
         Left            =   2550
         TabIndex        =   14
         Top             =   180
         Width           =   1575
      End
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   1245
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1650
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   5385
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1650
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   1605
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1950
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   1245
      TabIndex        =   40
      Top             =   3090
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "文到當日"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "文到次日"
         Height          =   180
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.TextBox Text13 
      Height          =   270
      Left            =   1245
      MaxLength       =   4
      TabIndex        =   16
      Top             =   3630
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   270
      Index           =   0
      Left            =   1245
      MaxLength       =   7
      TabIndex        =   19
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   270
      Left            =   1245
      MaxLength       =   6
      TabIndex        =   21
      Top             =   4275
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   270
      Left            =   1245
      MaxLength       =   7
      TabIndex        =   22
      Top             =   4590
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Height          =   270
      Left            =   5385
      MaxLength       =   1
      TabIndex        =   23
      Top             =   4560
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7515
      TabIndex        =   32
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5475
      TabIndex        =   30
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6300
      TabIndex        =   31
      Top             =   45
      Width           =   1200
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frm04010503_3.frx":0000
      Left            =   1245
      List            =   "frm04010503_3.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   38
      Top             =   780
      Width           =   615
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2805
      MaxLength       =   2
      TabIndex        =   37
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2565
      MaxLength       =   1
      TabIndex        =   36
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1725
      MaxLength       =   6
      TabIndex        =   35
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1245
      MaxLength       =   3
      TabIndex        =   34
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   5175
      TabIndex        =   33
      Top             =   480
      Width           =   1632
   End
   Begin VB.TextBox Text14 
      Height          =   270
      Index           =   1
      Left            =   4980
      MaxLength       =   7
      TabIndex        =   20
      Top             =   3960
      Width           =   1215
   End
   Begin MSForms.TextBox Text19 
      Height          =   300
      Left            =   1245
      TabIndex        =   5
      Top             =   2550
      Width           =   4155
      VariousPropertyBits=   671107099
      MaxLength       =   32
      Size            =   "7329;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "2. 第二階段                    (           P)"
      Height          =   180
      Left            =   5865
      TabIndex        =   76
      Top             =   5265
      Width           =   2505
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "ＩＤＳ報價:  1. 第一階段                    (           P)"
      Height          =   180
      Left            =   4830
      TabIndex        =   75
      Top             =   4965
      Width           =   3540
   End
   Begin VB.Label lblsNP23 
      Caption         =   "約定期限:"
      Height          =   255
      Left            =   4200
      TabIndex        =   74
      Top             =   4283
      Width           =   975
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "國際分類:"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   72
      Top             =   2880
      Width           =   765
   End
   Begin VB.Label lblFee 
      AutoSize        =   -1  'True
      Caption         =   "費用:"
      Height          =   180
      Left            =   4185
      TabIndex        =   71
      Top             =   3675
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblPt 
      AutoSize        =   -1  'True
      Caption         =   "點數:"
      Height          =   180
      Left            =   5895
      TabIndex        =   70
      Top             =   3675
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "審查委員編號:"
      Height          =   180
      Index           =   2
      Left            =   5490
      TabIndex        =   69
      Top             =   2580
      Width           =   1125
   End
   Begin VB.Label lblDispDate 
      AutoSize        =   -1  'True
      Caption         =   "機關發文日:"
      Height          =   180
      Left            =   4185
      TabIndex        =   68
      Top             =   1980
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "主管機關文書:"
      Height          =   180
      Left            =   120
      TabIndex        =   67
      Top             =   5190
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   8340
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   120
      X2              =   8340
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "是否列印通知函        (N:不印)"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   66
      Top             =   4950
      Width           =   4005
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "核駁函日期:"
      Height          =   180
      Left            =   120
      TabIndex        =   65
      Top             =   1710
      Width           =   945
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "案件目前准駁:         (1:准 , 2:駁)"
      Height          =   180
      Left            =   4185
      TabIndex        =   64
      Top             =   1710
      Width           =   2580
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "專利權是否存在:             (Y/N)"
      Height          =   180
      Left            =   120
      TabIndex        =   63
      Top             =   1980
      Width           =   2595
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   120
      TabIndex        =   62
      Top             =   2235
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "來函期限:"
      Height          =   180
      Left            =   120
      TabIndex        =   61
      Top             =   3225
      Width           =   765
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "下一程序:"
      Height          =   180
      Left            =   120
      TabIndex        =   60
      Top             =   3675
      Width           =   765
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   120
      TabIndex        =   59
      Top             =   3990
      Width           =   765
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Left            =   4185
      TabIndex        =   58
      Top             =   3990
      Width           =   765
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   120
      TabIndex        =   57
      Top             =   4320
      Width           =   585
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "承辦期限:"
      Height          =   180
      Left            =   120
      TabIndex        =   56
      Top             =   4650
      Width           =   765
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "是否算案件數:            (N:不算)"
      Height          =   180
      Index           =   0
      Left            =   4185
      TabIndex        =   55
      Top             =   4650
      Width           =   2550
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "審查委員名稱:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   54
      Top             =   2550
      Width           =   1125
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   3
      Left            =   1245
      TabIndex        =   53
      Top             =   1350
      Width           =   1860
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3281;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   2
      Left            =   4830
      TabIndex        =   52
      Top             =   1110
      Width           =   1650
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2910;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   1
      Left            =   1245
      TabIndex        =   51
      Top             =   1110
      Width           =   1860
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3281;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   210
      Index           =   0
      Left            =   1890
      TabIndex        =   50
      Top             =   840
      Width           =   6270
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "11060;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   49
      Top             =   1350
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Left            =   4185
      TabIndex        =   48
      Top             =   1110
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   120
      TabIndex        =   47
      Top             =   1110
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   46
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   4185
      TabIndex        =   45
      Top             =   510
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   44
      Top             =   840
      Width           =   765
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   4
      Left            =   2565
      TabIndex        =   43
      Top             =   4275
      Width           =   1260
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2222;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   5
      Left            =   2565
      TabIndex        =   42
      Top             =   3675
      Width           =   1290
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2275;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   6
      Left            =   4185
      TabIndex        =   41
      Top             =   1350
      Width           =   2310
      ForeColor       =   255
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "4075;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm04010503_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (Text19,Label3)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim strReceiveNo As String, strTemp As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim intWhere As Integer
Dim m_strNP14 As String '記錄上個畫面所點選案件的相關人
Dim m_strCP10 As String
Dim m_strCP43 As String 'Added by Lydia 2023/06/15 被核駁收文之相關收文號
'Add by Morgan 2006/6/26
Dim m_901CP09 As String '901內部收文之總收文號
Dim m_901CP12 As String '901內部收文之業務區
Dim m_901CP13 As String '901內部收文之智權人員
Dim stCP12 As String, stCP13 As String '最新A類收文智權人員及業務區
Dim bolCancelClose As Boolean 'Add by Morgan 2007/5/4 是否取消閉卷
Dim m_bolFMP As Boolean, strCReceiveNo As String, stNP23 As String, stCP48Desc As String 'Add by Morgan 2009/12/1
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/15 是否為寰華
Dim m_CustX07166 As Boolean   '2012/11/26 add by sonia 是否順德(含關係企業)專利案件
Dim m_strCP14 As String       '2012/11/26 add by sonia 記錄上個畫面所點選收文號的承辦人(P非台灣案若原承辦人非工程師則改抓國內案承辦人)
Public str941CP14 As String   '內部收文941收文號及承辦人
'Added by Morgan 2014/1/14
Public m_DocNo As String
Public m_AppNo As String
'end 2014/1/14
'Added by Morgan 2014/4/17
Public m_DocWord As String
Public m_DeadLine As String
'end 2014/4/17
Dim RC_cp10 As String 'Add by Lydia 2014/11/18 台灣案主管機關來函輸入 (案件性質)
'Added by Morgan 2016/3/16
Dim m_CP27 As String '發文日
Dim m_bolEngCase As Boolean '是否工程師承辦
'---
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END
Dim m_USCaseNo As String 'Added by Morgan 2019/5/27 相關美國案本所案號(提IDS用)
Dim m_bolNoCP27 As Boolean '不上發文 Added by Morgan 2020/1/17
Dim m_bolReKeyInOK As Boolean 'Added by Morgan 2020/8/12 是否與2次確認期限一致
Dim m_bolBPFCase As Boolean '是否寶齡富錦
Dim m_bolW2001XCase As Boolean 'Added by Morgan 2021/9/22 是否顧服組W2001的4家客戶案件
Dim m_CustX69365 As Boolean 'Added by Morgan 2021/10/6 是否長庚醫院案件
Dim m_str1998CP09 As String 'Added by Morgan 2021/10/6 轉公文收文號
Dim m_bolFMPNoPrint As Boolean 'Added by Morgan 2023/4/10 FMP案是否列印中文定稿
Dim bolChk414for106 As Boolean, strFirstPriDate As String 'Added by Lydia 2023/06/15 寰華案:是否為「414恢復權利-主張優先權106」、最早優先權日
Dim strChoseBase As String 'Added by Lydia 2023/06/15 選擇的優先權基礎案
Dim m_bolBusy As Boolean 'Added by Morgan 2024/12/18
Dim m_bolAddB908 As Boolean 'Added by Morgan 2025/3/7 是否內部收文代辦退費
Dim bolChgRlt As Boolean 'Added by Morgan 2025/3/7 是否為申請案核駁(基本檔上駁)

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 12) As String, lTmp As Long
Dim ii As Integer
    
    ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
    ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','下一程序名稱','" & Label3(5).Caption & "')"
    ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','本所期限','" & Text14(0).Text & "')"
    ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','法定期限','" & Text14(1).Text & "')"
      
   If ET03 = "01" Then
      'MODIFY BY SONIA 90.10.21
      'Modify By Cheng 2002/01/08
      'Modify By Cheng 2002/01/09
   '   strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='1' AND YF03='Y99999000' AND YF04='" & Text13 & "' AND YF05=1"
   '   strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='Y99999000' AND YF04='" & Text13 & "' AND YF05=1"
       'Modify By Cheng 2003/01/14
       '內專代理人抓Y00000001
   '   strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='Y00000000' AND YF04='" & Text13 & "' AND YF05=1"
      'Modify by Morgan 2009/9/4 大陸再審由畫面輸入
      If txtFee.Visible = True Then
         If txtFee <> "" Then 'Added by Morgan 2024/4/30
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','費用','" & txtFee & "')"
               
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','點數','" & txtPt & "')"
               
            'Added by Morgan 2016/6/13 非台灣信函進度要存報價
            PUB_UpdateLP2930 strCReceiveNo, txtFee, txtPt
            'end 2016/6/13
            
         'Added by Morgan 2024/4/30
         ElseIf m_bolEngCase Then
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','費用','|#(紅字)【請程序報價】#|')"
         'end 2024/4/30
         End If
      Else
         strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='Y00000001' AND YF04='" & Text13 & "' AND YF05=1"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then lTmp = Val(RsTemp.Fields(0))
          ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','費用','" & lTmp & "')"
      End If
   '2009/8/28 ADD BY SONIA 大->台再審報價,新型核駁直接訴願故無再審報價
   ElseIf ET03 = "08" Then
      Select Case pa(8)
         Case "1"
            lTmp = 23000   '2012/3/27 MODIFY BY SONIA 20000->23000
         Case "3"
            lTmp = 14000   '2012/3/27 MODIFY BY SONIA 11000->14000
      End Select
       ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','費用','" & lTmp & "')"
   End If
   '2009/8/28 END
   
   'Add By Cheng 2002/06/21
   If Text20 <> "" Then
        ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','列印備註','" & Text20 & "')"
   End If
   'Add by Morgan 2005/11/16
   If Text13 <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','下一程序','" & Text13 & "')"
   End If
   'Add by Morgan 2007/9/11
   '台灣之發明及設計案件,若有同時辦美國案(主案,未閉卷,未核准),則於通知核准或核駁定稿中加入一段美國提IDS之提醒字眼
   'Moidfy by Morgan 2011/7/18 +控制美國需為發明案(設計不用)且和核准一樣改判斷美國案未領證未閉卷--郭
   'Modified by Morgan 2019/5/27 需輸入IDS報價，改存檔前檢查
   'If PA(9) = "000" And (PA(8) = "1" Or PA(8) = "3") Then
   '   strExc(1) = PUB_GetUSCaseNo(PA(1), PA(2), PA(3), PA(4))
   '   If strExc(1) <> "" Then
      If m_USCaseNo <> "" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','美國案本所案號','" & m_USCaseNo & "')"
            
         'Modified by Morgan 2019/6/3 第１階段報價金額大於０才寫，定稿要控制不出該報價文字
         If Val(txtIDSFee(1)) > 0 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','IDS報價1','" & txtIDSFee(1) & "')"
         End If
         
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','IDS報價2','" & txtIDSFee(2) & "')"
      End If
   'End If
   
   'Added by Morgan 2012/10/1
   '台灣發明初審駁且有最後通知
   If pa(9) = "000" And InStr("101,102,103", m_strCP10) > 0 Then
      If PUB_ChkCPExist(pa(), "1227") Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','初審駁且有最後通知','♀')"
      End If
   End If
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(ii, strTxt) Then MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   If Not ClsLawExecSQL(ii, strTxt) Then MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim strTmp As String
   'Dim bolEngLetter As Boolean 'Added by Morgan 2023/5/10 是否產生工程師用定稿 'Removed by Morgan 2024/4/30 改用 m_bolEngCase 控制
   
   'Added by Morgan 2024/12/18 有發生重複存檔故增加此判斷
   If m_bolBusy Then Exit Sub
   m_bolBusy = True
   'end 2024/12/18
   
   Select Case Index
      Case 0
         
         'Add by Morgan 2007/6/12
         If txtDispDate.Visible = True Then
            If txtDispDate = "" Then
               MsgBox "機關發文日不可空白！", vbCritical
               txtDispDate.SetFocus
               GoTo EXITSUB
            ElseIf ChkDate(txtDispDate) = False Then
               txtDispDate.SetFocus
               GoTo EXITSUB
            End If
         End If
         'end 2007/6/12
         
        'Modify By Cheng 2002/11/29
        '若來函性質為行政再審, 可不輸入期限
        'Modify By Cheng 2003/09/02
        '若來函性質為行政訴訟上訴, 可不輸入期限
'        If m_strCP10 <> 行政再審 Then
        'Modify by Morgan 2011/3/24 +行政上訴答辯 508
        'Modified by Morgan 2020/2/24 +PPH 431--Winfrey,玲玲
        If m_strCP10 <> 行政再審 And m_strCP10 <> "507" And m_strCP10 <> "508" And Not (pa(9) = "020" And m_strCP10 = "431") Then
            If Text14(0) = "" Or Text14(1) = "" Then
               MsgBox "本所期限、法定期限不可空白 !", vbCritical
               GoTo EXITSUB
            End If
        End If
         'Add By Cheng 2002/03/11
         '檢查本所期限
         With Me.Text14(0)
            If .Text <> "" Then
               If Len(.Text) = 8 Then
                  If .Text < strSrvDate(1) Then
                     MsgBox "本所期限不可小於系統日期!!!", vbExclamation
                     .SetFocus
                     TextInverse Text14(0)
                     GoTo EXITSUB
                  End If
               Else
                  If Val(.Text) + 19110000 < strSrvDate(1) Then
                     MsgBox "本所期限不可小於系統日期!!!", vbExclamation
                     .SetFocus
                     TextInverse Text14(0)
                     GoTo EXITSUB
                  End If
               End If
            End If
         End With
         
         'Add By Cheng 2002/05/06
         '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
         If Len(Me.Text14(0).Text) > 0 And Len(Me.Text17.Text) > 0 Then
            If Val(Me.Text14(0).Text) < Val(Me.Text17.Text) Then
               MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
               GoTo EXITSUB
            End If
         End If
         
        'Add By Cheng 2003/03/26
        '檢查機關文號
        If pa(9) = 台灣國家代號 Then
            If Me.Text9.Tag = Me.Text9.Text Then
                MsgBox "請輸入機關文號!", vbExclamation + vbOKOnly
                Me.Text9.SetFocus
                Text9_GotFocus
                GoTo EXITSUB
            End If
        End If
        
         'Add by Morgan 2008/5/20
         If Text19 = "" Then
            If pa(9) = 台灣國家代號 And InStr("101,103,104,107,301,303,304,306,307", m_strCP10) > 0 Then
               MsgBox "案件性質為【" & Label3(1) & "】時，審查委員不可空白！"
               Text19.SetFocus
               GoTo EXITSUB
            End If
         End If
         'end 2008/5/20
         
         'Add By Sindy 2012/3/7 內外專都只做申請案號第四碼<>'3'之新申請案件性質
         If pa(9) = 台灣國家代號 And Mid(Trim(Text1), 4, 1) <> "3" And InStr(NewCasePtyList, m_strCP10) > 0 Then
            If Text22.Text = "" Then
               MsgBox "國際分類不可空白！"
               Text22.SetFocus
               GoTo EXITSUB
            End If
         End If
         '2012/3/7 End
         
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then GoTo EXITSUB
         'Add by Morgan 2005/5/20
         '非台灣 宣告無效 詢問是否計算結餘
         If m_strCP10 = "803" Then
            'Modified by Lydia 2015/03/03 +pa01,pa02,pa03,pa04
            Pub_EndModCashMsg pa(9), pa(1), pa(2), pa(3), pa(4)
         End If
         
         'Add by Morgan 2007/5/4 若來函有期限但已閉卷
         bolCancelClose = False
         If pa(57) = "Y" And Text14(1) <> "" Then
            If MsgBox("本案目前為閉卷狀態，為管制期限將於存檔時取消閉卷，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
               GoTo EXITSUB
            End If
            bolCancelClose = True
         End If
         'end 2007/5/4
         
         'Added by Lydia 2015/12/17 對於已經閉卷的案件,後續若有官方來函是無期限的,全部都詢問user是否要取消閉卷,由user來判斷
         If pa(57) = "Y" And Text14(1) = "" And bolCancelClose = False Then
            If MsgBox("本案目前為閉卷狀態，您輸入的是無期限的來函，是否要取消閉卷？", vbYesNo + vbDefaultButton1) = vbYes Then
               bolCancelClose = True
            End If
         End If
         'end 2015/12/17
         
         'Added by Morgan 2012/5/24
         If txtFee.Visible = True Then
            If txtFee.Text = "" Then
               'Added by Morgan 2024/4/30 工程師承辦時可不輸費用
               If m_bolEngCase Then
                  txtPt.Text = ""
               Else
               'end 2024/4/30
               
                  MsgBox "請輸入費用！", vbExclamation
                  txtFee.SetFocus
                  GoTo EXITSUB
                  
               End If
            ElseIf txtPt.Text = "" Then
               MsgBox "請輸入點數！", vbExclamation
               txtPt.SetFocus
               GoTo EXITSUB
            'Added by Morgan 2016/12/22 P105805 核駁報價費用少輸1個0
            ElseIf 1000 * Val(txtPt) > Val(txtFee) Then
               MsgBox "費用輸入錯誤(不可少於點數)！", vbExclamation
               txtFee.SetFocus
               GoTo EXITSUB
            'end 2016/12/22
            End If
         End If
         'end 2012/5/24
         
         '2012/11/26 add by sonia
         If m_CustX07166 = False Then m_CustX07166 = PUB_CheckX07166Remind(pa(1), strReceiveNo, "1002", str941CP14)
         '2012/11/26 END
         
         

         'Added by Lydia Lydia 2023/06/15 寰華案:是否為「414恢復權利-主張優先權106」若被核駁比照一般來函輸入frm04010504_3＝＞後案官方來函性質「視為未主張」，若有兩個以上的優先權，則 show出該兩個優先權讓user勾選哪一個被視為未主張，若只有一個優先權，就直接自優先權資料處移至案件備註(刪除PriDate,寫案件備註)，該案若有以優先權計算期限的，則請重新計算期限。
         strChoseBase = ""
         If bolChk414for106 And strFirstPriDate <> "" Then
            Set RsTemp = PUB_ReadPDStateNew(pa, m_strCP10)
            If RsTemp.RecordCount = 1 Then
               strChoseBase = RsTemp.Fields("優先權號") & "|" & RsTemp.Fields("優先權日") & "|" & RsTemp.Fields("PD07")
            ElseIf RsTemp.RecordCount > 1 Then
                Set frm880012.grdDataList.Recordset = RsTemp
                Set frm880012.fmParent = Me
                frm880012.iTyp = "4"
                frm880012.Show vbModal
                If Me.Tag = "" Then
                   MsgBox "請選擇一個優先權資料!"
                   GoTo EXITSUB
                Else
                   strChoseBase = Me.Tag
                   Me.Tag = ""
                End If
            End If
         End If
         'end 2023/06/15
         
         'Modified by Morgan 2022/3/28 取消轉公文,改同其他3家直接報告,但本所期限改為 +14天-3個工作天 --黃教威
         'If m_CustX69365 = True Then Text15 = "" 'Added by Morgan 2021/10/6 長庚醫院案件要收[轉公文]先簡單報告
         'end 2022/3/28
         
         'Add By Sindy 2020/7/20
         If m_strIR01 <> "" Then
            '下載信件檔
            If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", , , True) = False Then
               Screen.MousePointer = vbDefault
               GoTo EXITSUB
            End If
            'Add By Sindy 2022/7/21
            If Left(Pub_StrUserSt03, 2) = "F2" Then
               If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
                  Screen.MousePointer = vbDefault
                  GoTo EXITSUB
               End If
            End If
            '2022/7/21 END
         End If
         '2020/7/20 END
            
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: GoTo EXITSUB
         
'Remove by Morgan 2007/4/18 P 的不用--郭
'         'Add by Morgan 2007/1/18 申請人為"福興"時彈訊息
'         If InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X43179") > 0 Then
'            MsgBox "請影印一份OA交付智權人員【" & GetStaffName(stCP13) & "】！"
'         End If
'         'end 2007/1/18
'end 2007/4/18

         'Added by Morgan 2023/5/10 工程師承辦的1201 通知修正,1002  核駁,1202 審查意見通知函 也要產生定稿以便撰寫信函使用
         'Modified by Morgan 2023/6/27 寶齡富錦且工程師承辦的來函除外
         'Removed by Morgan 2024/4/30 改用 m_bolEngCase 控制
         'If Text15.Text = "N" And Not m_bolFMP And Text16 <> strUserNum And Not m_bolBPFCase Then
         '   bolEngLetter = True
         'End If
         'end 2024/4/30
         'end 2023/5/10
         
         'Modified by Morgan 2023/5/10 +bolEngLetter
         'Modified by Morgan 2024/4/30
         'If Text15.Text <> "N" Or bolEngLetter Then   '通知函
         If Text15.Text <> "N" Or m_bolEngCase Then   '通知函
         'end 2024/4/30
            Select Case m_strCP10
               'modify by sonia 2019/11/8 +307分案(P-109152)
               Case 發明申請, 新型申請, 設計申請, 307 '0
                  If pa(9) = 台灣國家代號 Then '台灣 0
                     '2009/08/26 ADD BY SONIA 大->台定稿
                     If PUB_CheckCuNation(pa(26), Text2, Text3, Text4, Text5) = "1" Then     '大-->台 定稿
                        strTmp = "08"
                     Else
                     '2009/8/26 END
                        strTmp = "00"
                        'Add by Morgan 2004/7/5
                        '若發明, 新型申請有被主張國內優先權則註明15個月將被撤回
                        'Mark by Lydid 2025/07/01 整理所有國內對客戶的通知函定稿：協理確定要刪除的定稿
                        'If m_strCP10 <> 設計申請 Then
                        '   If PUB_ChkPriDate(pa(11), "", False) Then
                        '      strTmp = "07"
                        '   End If
                        'End If
                        'end 2025/07/01
                     End If
                  Else '非台灣 1
                     strTmp = "01"
                  End If
               Case 訴願 '4
                  strTmp = "04"
               Case 行政訴訟 '5
                  strTmp = "05"
                'Modify By Cheng 2002/12/29
'               Case 異議_專, 異議答辯, 舉發, 舉發答辯, 申請優先權證明
               Case 異議_專, 異議答辯, 舉發, 申請優先權證明
                  If pa(9) = 台灣國家代號 Then '台灣 2
                     strTmp = "02"
                  Else '非台灣 3
                     strTmp = "03"
                  End If
               'Add By Cheng 2002/12/29
               '舉發答辯(宣告無效答辯)
               Case 舉發答辯
                  If pa(9) = 台灣國家代號 Then '台灣 2
                     strTmp = "02"
                  Else '非台灣 6
                     strTmp = "06"
                  End If
               Case Else
                  strTmp = "00"
                  
                  'Added by Morgan 2013/2/22
                  '大陸復審核駁
                  If pa(9) = "020" And m_strCP10 = "107" Then
                     strTmp = "01"
                  End If
                  'end 2013/2/22
            End Select
            
            'Add by Morgan 2009/11/30
            If m_bolFMP Then
               'Modified by Morgan 2017/4/27 FMP已改出一般P案定稿以識別閉卷後來函--潘韻丞
               If pa(57) = "Y" Then
                  StartLetter "06", strTmp
                  NowPrint strReceiveNo, "06", strTmp, False, strUserNum, 0, , , , 1, , , , , , , , strCReceiveNo
               Else
               'end 2017/4/28
               
                  'Modify by Morgan 2010/6/10 改用通函
                  'StartLetter "06", strTmp
                  'NowPrint strReceiveNo, "06", strTmp, False, strUserNum, 0, , , , 1
                  'Modified by Morgan 2023/4/10 FMP案有EMail通知的就不在列印紙本
                  NowPrint strCReceiveNo, "07", "99", False, strUserNum, 0, , , , 1, , , , , , , , strCReceiveNo, , , , , m_bolFMPNoPrint
                  'end 2010/6/10
                  
               End If 'Added by Morgan 2017/4/28
            Else
            'end 2009/11/30
               StartLetter "06", strTmp
               'Modified by Morgan 2021/10/6 若有轉公文則優先
               'NowPrint strReceiveNo, "06", strTmp, False, strUserNum, 0, , , , , , , , , , , , strCReceiveNo
               'Modified by Morgan 2023/5/10 +bolEngLetter
               'Modified by Morgan 2024/4/30
               'NowPrint strReceiveNo, "06", strTmp, False, strUserNum, 0, , , , , , , , , , , , IIf(m_str1998CP09 <> "", m_str1998CP09, strCReceiveNo), , , , , bolEngLetter
               NowPrint strReceiveNo, "06", strTmp, False, strUserNum, 0, , , , , , , , , , , , IIf(m_str1998CP09 <> "", m_str1998CP09, strCReceiveNo), , , , , m_bolEngCase
               'end 2024/4/30
            End If
         End If
         'Add by Lydia 2014/11/18 台灣案主管機關來函輸入，若此案有工程師未發文的程序，發E-MAIL通知工程師收到來函的內容
         'Modified by Lydia 2022/08/15 開放P大陸案
         'If pa(9) = "000" And pa(1) = "P" Then
         'Modified by Lydia 2022/10/11 經查此設定並不適用於外專及日專，故請協助排除FMP案
         'If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" Then
         If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" And m_bolFMP = False Then
            'Modified by Morgan 2016/3/17 工程師承辦的都通知
            'PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), RC_cp10, strCReceiveNo
            'Modified by Lydia 2017/03/29 模組已改成，已收文未發文的承辦人全部發mail通知
            'PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), RC_cp10, IIf(m_bolEngCase, "", strCReceiveNo)
            'Modified by Lydia 2022/08/16 +申請國家
            'PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), RC_cp10, strCReceiveNo
            'Modified by Morgan 2023/6/27 寶齡富錦案件已有通知,要排除本次來函
            PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), RC_cp10, pa(9), strCReceiveNo, , , m_bolBPFCase
            'end 2016/3/17
         End If
         
         'Added by Lydia 2023/06/15 寰華案:「414恢復權利-主張優先權106」更新實審期限
         If bolChk414for106 = True Then
            '請彈跳提醒視窗：已更新實體審查期限為:
            strSql = "select '1' as ord1, cp06 from caseprogress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and cp10='416' and cp158=0 " & _
                     "union select '2' as ord1, np08 from nextprogress WHERE np02='" & pa(1) & "' AND np03='" & pa(2) & "' AND np04='" & pa(3) & "' AND np05='" & pa(4) & "' and np07='416' and np06 is null " & _
                     "order by ord1"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               MsgBox IIf(Left(strFirstPriDate, 1) = "Y", "已更新", "") & "實體審查期限為: " & ChangeWStringToTDateString("" & RsTemp.Fields("cp06")) & "，請通知承辦報告客戶。 ", vbExclamation + vbOKOnly
            End If
         End If
         'end 2023/06/15
            
         'Add By Sindy 2016/10/5
         If Me.m_strIR01 <> "" Then
            Unload frm04010503_1
            Unload frm04010503_2
            Unload Me
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         ElseIf Me.m_DocNo <> "" Then
         'Added by Morgan 2014/1/14
         'If Me.m_DocNo <> "" Then
         '2016/10/5 END
            Unload frm04010503_1
            Unload frm04010503_2
            Unload Me
            frm04010516.GoNext
         Else
         'end 2014/1/14
            frm04010503_1.Show
            frm04010503_1.Clear
            Unload frm04010503_2
            Unload Me
         End If 'Added by Morgan 2014/1/14
            
      Case 1
         frm04010503_2.Show
         Unload Me
      Case 2
         Unload frm04010503_1
         Unload frm04010503_2
         Unload Me
   End Select
   
'Added by Morgan 2024/12/18
   Exit Sub
   
EXITSUB:
   m_bolBusy = False
'end 2024/12/18
End Sub

Private Function FormSave() As Boolean
Dim intStep As Integer, strTxt(1 To 20) As String, strTmp As String, bolChk As Boolean
Dim i As Integer, strCe(99) As String
Dim strSql As String
' 90.07.17 modify by louis (暫存列印接洽結案單的資料)
Dim strProgressNo As String
Dim strCP115 As String 'Add by Morgan 2007/6/12
'Dim strCP48 As String '承辦期限 'Remove by Lydia 2021/11/05
Dim strCP26 As String, strCP27 As String, strCP133 As String, strCP134 As String 'Add by Morgan 2009/11/30
Dim st307Msg As String '分割案提醒訊息 Added by Morgan 2011/12/6
Dim str941ReceiveNo As String  '2012/11/26 ADD BY SONIA 內部收文941收文號
Dim str941CP06 As String 'Add By Sindy 2013/4/2
Dim strNP01 As String, strNP07 As String, strNP09 As String 'Add By Sindy 2013/4/29
Dim mRCno As String, mCCno As String, oSubject As String, oContext As String   'Add by Lydia 2014/10/16 FMP案會列印 C類接洽單, 請同時E-MAIL給畫面上之承辦人, 副本發給該員之工程師組別主管.
Dim strCMemo As String 'Added by Morgan 2020/9/15
'Dim bolAdd941 As Boolean '是否內部收文分析 Added by Morgan 2014/12/8 'Removed by Morgan 2016/3/18
Dim strCP64 As String 'Added by Morgan 2019/5/28
Dim strCP20 As String 'Added by Morgan 2019/8/8
Dim bolReKeyInCase As Boolean 'Added by Morgan 2023/4/10
Dim arrData As Variant 'Added by Lydia 2023/06/15
Dim bolCN431 As Boolean 'Added by Morgan 2025/3/12

On Error GoTo ErrorHandler
   
   FormSave = False
cnnConnection.BeginTrans
   
   bolReKeyInCase = False 'Added by Morgan 2023/4/10
   m_bolFMPNoPrint = False 'Added by Morgan 2023/4/10
   
   strProgressNo = Empty
   intStep = 1
   
   '2
   'Modify By Cheng 2002/07/23
'   'MODIFY BY SONIA 90.10.21
'   'If Text8 = "" Then MsgBox "專利權是否存在不可空白，請重新輸入 !", vbCritical: Exit Function
'   'strExc(0) = "PA17='" & Text8 & "',"
'   If Text8 <> "" Then
      strExc(0) = "PA17='" & text8 & "',"
'   Else
'      strExc(0) = "PA17=NULL,"
'   End If
'   If Text7 = "Y" Then strExc(0) = strExc(0) & "PA16='2',PA20=" & CNULL(TransDate(Text6, 2)) & ","
    'Modify By Cheng 2003/01/09
'   If (m_strCP10 >= "101" And m_strCP10 <= "105") Or m_strCP10 = "107" Or (m_strCP10 >= "301" And m_strCP10 <= "307") Or m_strCP10 = "802" Or m_strCP10 = "804" Then
   'Modified by Morgan 2020/12/18 改寫函數判斷以便共用及修改
   'If m_strCP10 <> "802" And m_strCP10 <> "804" Then 'Added by Morgan 2012/3/7 排除 802,804
   '   '2013/10/24 MODIFY BY SONIA 再加入卷宗性質判斷pa(23) = "1",P-083407的503不可更新,否則後續改變原處分也不會更新
   '   If pa(23) = "1" And ((Val(m_strCP10) >= 101 And Val(m_strCP10) <= 105) Or Val(m_strCP10) = 107 Or Val(m_strCP10) = 503 Or Val(m_strCP10) = 504 Or _
   '        (Val(m_strCP10) >= 301 And Val(m_strCP10) <= 307) Or (Val(m_strCP10) >= 801 And Val(m_strCP10) <= 805)) Then
   'Modified by Morgan 2025/3/7
   'If pa(23) = "1" And PUB_ChkIsRltPty(pa(1), m_strCP10, pa(9)) = True Then
   If bolChgRlt Then
   'end 2025/3/7
   'end 2020/12/18
         strExc(0) = strExc(0) & "PA16='" & Me.Text7.Text & "',"
         'Modify by Morgan 2004/11/30 爭議程序不更新基本檔准駁日
         'If IsEmptyText(Text6.Text) = False Then
         If IsEmptyText(Text6.Text) = False And Not (Val(m_strCP10) >= 801 And Val(m_strCP10) <= 805) Then
            strExc(0) = strExc(0) & "PA20=" & CNULL(TransDate(Text6.Text, 2)) & ","
         End If
         
      'End If 'Removed by Morgan 2020/12/18
   End If
   'Add By Sindy 2012/3/7 +國際分類更新
   strExc(0) = strExc(0) & "PA160=" & CNULL(Text22.Text) & ","
   '2012/3/7 End
   If Right(strExc(0), 1) = "," Then strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
   strTxt(intStep) = "UPDATE PATENT SET " & strExc(0) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
    'Add By Cheng 2002/11/08
    cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1
   
   '1
   'Modified by Morgan 2022/8/8/24 +3部分准駁
   If (frm04010503_2.Text6 = "1" Or frm04010503_2.Text6 = "3") Then
      If frm04010503_2.Text6 = "3" Then
         i = 1009
      Else
         i = 核駁
      End If
      
      If Left(m_strCP10, 1) = "1" Or Left(m_strCP10, 1) = "3" Then
         '92.6.8 MODIFY BY SONIA 同時更新機關文號且原無准駁者才更新
         'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='2',CP25=" & TransDate(Text6, 2) & _
         '   " WHERE CP09='" & strReceiveNo & "'"
         '2005/10/19 MODIFY BY SONIA 不判斷 CP25
         'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='2',CP25=" & TransDate(Text6, 2) & ",CP08=" & CNULL(Text9) & _
         '   " WHERE CP09='" & strReceiveNo & "' AND CP24 IS NULL AND CP25 IS NULL"
         'Modify  by Morgan 2008/5/21 +CP35,CP117
         strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='2',CP25=" & TransDate(Text6, 2) & ",CP08=" & CNULL(Text9) & _
            ",CP35='" & ChgSQL(Me.Text19.Text) & "',CP117='" & ChgSQL(Me.Text21.Text) & "' WHERE CP09='" & strReceiveNo & "' AND CP24 IS NULL"
         '2005/10/19 END
         '92.6.8 END
        'Add By Cheng 2002/11/08
        cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
      Else
         '92.6.8 MODIFY BY SONIA 同時更新機關文號且原無准駁者才更新
         'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='2',CP25=" & TransDate(Label3(3).Caption, 2) & _
         '   " WHERE CP09='" & strReceiveNo & "'"
         '2005/10/19 MODIFY BY SONIA 不判斷 CP25
         'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='2',CP25=" & TransDate(Label3(3).Caption, 2) & ",CP08=" & CNULL(Text9) & _
         '   " WHERE CP09='" & strReceiveNo & "' AND CP24 IS NULL AND CP25 IS NULL"
         'Modify  by Morgan 2008/5/21 +CP35,CP117
         strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='2',CP25=" & TransDate(Label3(3).Caption, 2) & ",CP08=" & CNULL(Text9) & _
            ",CP35='" & ChgSQL(Me.Text19.Text) & "',CP117='" & ChgSQL(Me.Text21.Text) & "' WHERE CP09='" & strReceiveNo & "' AND CP24 IS NULL"
         '2005/10/19 END
        '92.6.8 END
        'Add By Cheng 2002/11/08
        cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
      End If
   Else
      i = 改變原處分
   End If
    
    'Remove by Morgan 2008/5/21 改更新准駁時一併更新
    'Add By Cheng 2003/02/18
    '更新審查委員
     '+CP117
    'strTxt(intStep) = "Update CaseProgress Set CP35='" & ChgSQL(Me.Text19.Text) & "',CP117='" & ChgSQL(Me.Text21.Text) & "' Where CP09='" & strReceiveNo & "' "
    'cnnConnection.Execute strTxt(intStep)
    'intStep = intStep + 1
    'end 2008/5/21
    
   '3
   strCReceiveNo = AutoNo("C", 6)
   'Modify by Morgan 2007/6/13 加CP115
   If txtDispDate.Visible = True Then
      strCP115 = DBDATE(txtDispDate)
   Else
      strCP115 = "NULL"
   End If
   
   'Modify by Morgan 2009/12/1 +CP133,CP134
   'Modified by Morgan 2012/4/25 +不必限制大陸(台灣案延期要用)
   'If pa(9) <> "000" Then
      strCP133 = DBDATE(Text6)
      strCP134 = Val(Text11)
   'Else
   '   strCP133 = "NULL"
   '   strCP134 = "NULL"
   'End If
   'end 2012/4/25
   
   
   'Added by Morgan 2014/12/8
   'Modified by Morgan 2015/6/25 +501訴願,505參加訴願
   'Removed by Morgan 2016/3/16 改直接由工程師承辦
   'If pa(9) = 台灣國家代號 And (m_strCP10 = "803" Or m_strCP10 = "804" Or m_strCP10 = "501" Or m_strCP10 = "505") Then
   '   bolAdd941 = True
   'End If
   'end 2016/3/16
   'end 2014/12/8
   
   'Modify by Morgan 2009/11/30 FMP案不要上發文日
   If m_bolFMP Then
      strCP27 = "NULL"
   'Modified by Morgan 2014/12/8
   'Else
   '   strCP27 = strSrvDate(1)
   'Modified by Morgan 2016/3/16 改判斷是否工程師承辦
   'ElseIf (pa(9) = 台灣國家代號 And bolAdd941 = True) Then
   'Modified by Morgan 2021/9/22
   'ElseIf m_bolEngCase Then
   ElseIf m_bolEngCase Or m_bolBPFCase Or m_bolW2001XCase Then
   'end 2016/3/16
      strCP27 = "NULL"
   'Added by Morgan 2020/1/17
   ElseIf m_bolNoCP27 = True Then
      strCP27 = "NULL"
   'end 2020/1/17
   Else
      strCP27 = strSrvDate(1)
   'end 2014/12/8
   End If
   
   'Added by Morgan 2019/5/28 備註＋IDS報價
   strCP64 = ""
   If m_USCaseNo <> "" Then
      'Modified by Morgan 2019/6/3 第１階段報價金額大於０才寫
      'Modified by Morgan 2019/9/9 調整報價欄位名及定稿內容--郭
      If Val(txtIDSFee(1)) > 0 Then
         strCP64 = "IDS報價:1.第一階段 " & txtIDSFee(1) & "(" & txtIDSPt(1) & "P), 2.第二階段 " & txtIDSFee(2) & "(" & txtIDSPt(2) & "P);"
      Else
         strCP64 = "IDS報價:第二階段 " & txtIDSFee(2) & "(" & txtIDSPt(2) & "P);"
      End If
   End If
   'end 2019/5/27
   
   'Modified by Morgan 2019/8/8 FMP案的CP20要抓設定
   strCP20 = "N"
   If m_bolFMP Then
      strCP20 = PUB_GetCP20(pa(1), str(i), , pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
   End If
   
   'Added by Morgan 2020/2/24
   'FMP寰華案的PHH核駁 承辦期限=收文日+5個工作天
   If m_bolFMP Then
      If Text17 = "" And Left(Pub_StrUserSt03, 1) = "F" Then
         Text17 = TransDate(CompWorkDay(5, TransDate(Label3(3).Caption, 2)), 1)
      End If
   End If
   'end 2020/2/24
   
   'Modify by Morgan 2008/5/21 +CP35,CP117
   'Modified by Morgan 2012/4/30 +cp119=櫃檯收文日
   '2015/1/5 MODIFY BY SONIA FMP案可請款(P-098893)
   'strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP08," & _
      "CP09,CP10,CP12,CP13,CP14,CP48,CP20,CP32,CP26,CP27,CP43,CP115,CP35,CP117,CP133,CP134,cp119) VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & _
      "','" & TransDate(Label3(3).Caption, 2) & "'," & CNULL(TransDate(Text14(0), 2)) & "," & CNULL(TransDate(Text14(1), 2)) & _
      "," & CNULL(Text9) & ",'" & strCReceiveNo & "','" & i & "','" & stCP12 & "','" & stCP13 & _
      "','" & Text16 & "','" & TransDate(Text17, 2) & _
      "','N','N'," & CNULL(Text18) & "," & strCP27 & ",'" & strReceiveNo & "'," & strCP115 & _
      ",'" & ChgSQL(Me.Text19.Text) & "','" & ChgSQL(Me.Text21.Text) & "'," & strCP133 & "," & strCP134 & "," & DBDATE(Label3(3)) & ")"
   'Modified by Morgan 2019/5/28 +CP64
   strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP08," & _
      "CP09,CP10,CP12,CP13,CP14,CP48,CP20,CP32,CP26,CP27,CP43,CP64,CP115,CP35,CP117,CP133,CP134,cp119) VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & _
      "','" & TransDate(Label3(3).Caption, 2) & "'," & CNULL(TransDate(Text14(0), 2)) & "," & CNULL(TransDate(Text14(1), 2)) & _
      "," & CNULL(Text9) & ",'" & strCReceiveNo & "','" & i & "','" & stCP12 & "','" & stCP13 & _
      "','" & Text16 & "','" & TransDate(Text17, 2) & _
      "','" & strCP20 & "','" & IIf(m_bolFMP, "", "N") & "'," & CNULL(Text18) & "," & strCP27 & ",'" & strReceiveNo & "','" & ChgSQL(strCP64) & "'," & strCP115 & _
      ",'" & ChgSQL(Me.Text19.Text) & "','" & ChgSQL(Me.Text21.Text) & "'," & strCP133 & "," & strCP134 & "," & DBDATE(Label3(3)) & ")"
   'END 2007/6/13
    'Modify end 2004/2/9
    
   '91.11.12 END
    'Add By Cheng 2002/11/08
    cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1
   
    RC_cp10 = i 'Add by Lydia 2014/11/18 台灣案主管機關來函輸入 (案件性質)
       
   'Added by Lydia 2025/08/19 輸入C類來函時，去檢查上一道承辦人掛工程師，是否為未請款，若是，則發Mail通知工程師；
                  '核駁比照核准，指定特定案件性質
   If i = 核駁 And pa(1) = "P" And m_bolFMP = True And InStr("101,102,103,307", m_strCP10) > 0 Then
      If PUB_ChkFCPtoCP14CP60(pa(1), pa(2), pa(3), pa(4), CheckStr(i), strCReceiveNo, Text16) = True Then
      End If
   End If
   'end 2025/08/19
   
   'Added by Morgan 2014/1/14
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCReceiveNo, pa(1), pa(2), pa(3), pa(4), RC_cp10
   End If
   'end 2014/1/14
   
   'Add by Sindy 2016/10/5
   If m_strIR01 <> "" Then
      'Add By Sindy 2020/4/20 核駁函輸入後請將整封郵件存入系統
      'Modify By Sindy 2022/11/9 + IIf(pa(9) <> 台灣國家代號, "PAT", "RX")
      If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, strCReceiveNo, IIf(Pub_StrUserSt03 = "F22", "ALTR", IIf(pa(9) <> 台灣國家代號, "PAT", "RX"))) = False Then 'PAT.陸代郵件
         GoTo ErrorHandler
      End If
      '2020/4/20 END
      'Modified by Morgan 2020/8/12 +傳 strCReceiveNo, m_bolReKeyInOK
      'Modify By Sindy 2022/6/16 F2外專不做2次確認
      'Modified by Morgan 2022/12/21 有期限才要2次確認 +  Or Text14(1) = ""
      If Left(Pub_StrUserSt03, 2) = "F2" Or Text14(1) = "" Then
         'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", strCReceiveNo, "")
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010503_1", IIf(Pub_StrUserSt03 = "F22", strCReceiveNo, "")
         bolReKeyInCase = True 'Added by Morgan 2023/4/10
      Else
      '2022/6/16 END
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010503_1", strCReceiveNo, m_bolReKeyInOK
      End If
   End If
   '2016/10/5 END
   
   'Added by Morgan 2014/4/11 電子化-新增信函進度檔
   If pa(9) = "000" Then
      'Modified by Morgan 2014/12/8 改都要抓判發人(舉發及舉發答辯的分析由工程師撰寫但發文後來函改通知客戶且待判發)
      'strExc(1) = ""
      'If Text15 <> "N" Then
      'Added by Morgan 2016/3/22
      '工程師承辦的來函不必判發(在歷程判發)
      'Modified by Morgan 2021/9/22
      'If m_bolEngCase Then
      If m_bolEngCase Or m_bolBPFCase Or m_bolW2001XCase Then
         strExc(1) = ""
      'end 2016/3/22
      Else
         'Modified by Morgan 2018/8/1
         'strExc(1) = PUB_GetLetterJudge(pa(1), RC_cp10, m_strCP10, , pa(1), pa(2), pa(3), pa(4))
         strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), RC_cp10, , m_strCP10)
      End If 'Added by Morgan 2016/3/22
      'End If
      'end 2014/12/8
      'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
      'Modified by Morgan 2021/10/6 +長庚醫院案件會自動收[轉公文]先通知客戶，來函分析信由工程師撰寫
      'Modified by Morgan 2022/3/28 長庚醫院案件取消轉公文 --黃教威
      'PUB_AddLetterProgress strCReceiveNo, 1, IIf(Text15 <> "N" And m_CustX69365 = False, True, False), strExc(1), IIf(Text14(1) <> "", True, False), pa(26), RC_cp10, pa(75)
      PUB_AddLetterProgress strCReceiveNo, 1, IIf(Text15 <> "N", True, False), strExc(1), IIf(Text14(1) <> "", True, False), pa(26), RC_cp10, pa(75)
      'end 2022/3/28
      
   'Added by Morgan 2016/6/8 非臺灣案電子化
   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      'Added by Morgan 2021/9/22
      If m_bolEngCase Or m_bolBPFCase Or m_bolW2001XCase Then
         strExc(1) = ""
      Else
      'end 2021/9/22
         'Modified by Morgan 2018/8/1
         'strExc(1) = PUB_GetLetterJudge(pa(1), RC_cp10, m_strCP10, pa(9), pa(1), pa(2), pa(3), pa(4))
         strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), RC_cp10, pa(9), m_strCP10, , m_bolFMP)
      End If 'Added by Morgan 2021/9/22
      'Modified by Morgan 2021/10/6 +長庚醫院案件會自動收[轉公文]先通知客戶，來函分析信由工程師撰寫
      'Modified by Morgan 2022/3/28 長庚醫院案件取消轉公文 --黃教威
      'PUB_AddLetterProgress strCReceiveNo, 2, IIf(Text15 <> "N" And m_CustX69365 = False, True, False), strExc(1), IIf(Text14(1) <> "", True, False), pa(26), RC_cp10, pa(75)
      PUB_AddLetterProgress strCReceiveNo, 2, IIf(Text15 <> "N", True, False), strExc(1), IIf(Text14(1) <> "", True, False), pa(26), RC_cp10, pa(75)
      'end 2022/3/28
   'end 2016/6/8
   
   End If
   'end 2014/4/11
   
   '92.11.18 ADD BY SONIA
   If i = 改變原處分 Then
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='2' WHERE CP09='" & strCReceiveNo & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   '92.11.18 END
 
   '4
   strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07='" & 催審 & "'"
    'Add By Cheng 2002/11/08
    cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1
   
   '5
   If frm04010503_2.Text6 = "2" Then
      strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07='" & 改變原處分 & "'"
        'Add By Cheng 2002/11/08
        cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   
   '6
   'Add By Sindy 2013/4/29
   strNP01 = ""
   strNP07 = ""
   strNP09 = ""
   '2013/4/29 End
   If Text13 <> "" Then
      strProgressNo = GetNextProgressNo
      'Added by Lydia 2025/11/05 屬於2025/10/29更新所限,約定期限
      If m_bolFMP = False Then
         strExc(1) = PUB_GetPOurDeadline(DBDATE(Text14(1)), pa(9), stNP23, pa(1), Text13)
      End If
      'end 2025/11/05
      '智權人員存最近收文A類接洽記錄單的智權人員
      strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
         "NP09,NP10,NP13,NP14,NP22,NP23) VALUES ('" & strCReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & _
         "'," & Text13 & "," & CNULL(TransDate(Text14(0), 2)) & "," & CNULL(TransDate(Text14(1), 2)) & _
         "," & CNULL(stCP13) & "," & CNULL(Text9) & "," & CNULL(ChgSQL(m_strNP14)) & _
         "," & strProgressNo & "," & CNULL(stNP23, True) & ")"
        cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      'Add By Sindy 2013/4/29
      strNP01 = strCReceiveNo
      strNP07 = Text13
      strNP09 = CNULL(TransDate(Text14(1), 2))
      '2013/4/29 End
'Remove by Morgan 2009/12/1 改來函不自動上發文日控制
'
'      'Add by Morgan 2006/6/26
'      '國外部收文若有期限則自動內部收文901告知代理人,承辦人固定為78063黃得峻並列印內部收文接洽單
'      If m_bolFMP Then
'         m_901CP09 = AutoNo("B", 6)
'         '2008/12/2 modify by sonia 改FMP控管方式
'         'm_901CP13 = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
'         m_901CP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
'         '2008/12/2 END
'         m_901CP12 = GetSalesArea(m_901CP13)
'         strExc(1) = GetWorkDays(pa(1), pa(9), "901")
'         If strExc(1) = Empty Then strExc(1) = 7
'         'Add by Morgan 2008/5/26 若來函期限超過(含)3個月則告代的承辦期限為14天--阮威立
'         If Val(strExc(1)) < 14 Then
'            If DBDATE(Text14(1)) >= CompDate(1, 3, strSrvDate(1)) Then
'               strExc(1) = 14
'            End If
'         End If
'         'end 2008/5/26
'
'         'Modify by Morgan 2006/8/4 不必抓工作天--郭
'         'strCP48 = CompWorkDay(Val(strExc(1)), strSrvDate(1), 0)
'         strCP48 = CompDate(2, Val(strExc(1)), strSrvDate(1))
'         '2008/12/3 MODIFY BY SONIA 依FC代理人國籍抓預設承辦人
'         'strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
'            "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
'            "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & strCP48 & "," & strCP48 & _
'            ",'" & m_901CP09 & "','" & 告知代理人 & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
'            ",'85030','N','N','N','" & strCReceiveNo & "'," & strCP48 & ") "   '2008/2/5 MODIFY BY SONIA 78063離職改85030阮威立--郭
'         strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
'            "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
'            "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & strCP48 & "," & strCP48 & _
'            ",'" & m_901CP09 & "','" & 告知代理人 & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
'            "," & CNULL(PUB_GetFMCASECP14(pa(1), pa(2), pa(3), pa(4))) & ",'N','N','N','" & strCReceiveNo & "'," & strCP48 & ") "   '2008/2/5 MODIFY BY SONIA 78063離職改85030阮威立--郭
'         '2008/12/3 END
'         cnnConnection.Execute strSQL
'      End If
'
'end 2009/12/1

   End If
   
   '7
   strExc(0) = "SELECT * FROM CHANGEEVENT WHERE CE01='" & strReceiveNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         For i = 1 To 99
            If IsNull(.Fields(i - 1)) Then
               strCe(i) = ""
            Else
               strCe(i) = .Fields(i - 1)
            End If
         Next
      End With
      strExc(1) = ""
      
      '申請日
      If strCe(2) <> "" Then strExc(1) = strExc(1) & "CE03='2',"
      
      '申請人
      For i = 4 To 8
         If strCe(i) <> "" Then
            strExc(1) = strExc(1) & "CE09='1',"
            Exit For
         End If
      Next
      
      '代表人
      bolChk = False
      For i = 10 To 15
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If Not bolChk Then
         For i = 68 To 91
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
      End If
      If bolChk Then strExc(1) = strExc(1) & "CE16='1',"
      
      
      '申請地址
      For i = 23 To 37
         If strCe(i) <> "" Then
            strExc(1) = strExc(1) & "CE38='1',"
            Exit For
         End If
      Next
      
      '專利商標種類代號
      If strCe(39) <> "" Then strExc(1) = strExc(1) & "CE40='1',"
      
      '案件名稱
      For i = 41 To 43
         If strCe(i) <> "" Then
            strExc(1) = strExc(1) & "CE44='1',"
            Exit For
         End If
      Next
      
      '代表人中譯文
      bolChk = False
      For i = 63 To 64
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If Not bolChk Then
         For i = 92 To 99
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
      End If
      
      If bolChk Then strExc(1) = strExc(1) & "CE65='1',"
      
      If strExc(1) <> "" Then
         If Right(strExc(1), 1) = "," Then strExc(1) = Left(strExc(1), Len(strExc(1)) - 1)
         strTxt(intStep) = "UPDATE CHANGEEVENT SET " & strExc(1) & " WHERE CE01='" & strReceiveNo & "'"
        'Add By Cheng 2002/11/08
        cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
      End If
   End If
   
'   FormSave = objLawDll.ExecSQL(intStep - 1, strTxt)
   
   ' 90.07.05 modify by louis
   If frm04010503_2.Text6 = "2" Then
      strSql = "UPDATE PATENT SET PA16 = '2' " & _
               "WHERE PA01 = '" & pa(1) & "' AND " & _
                     "PA02 = '" & pa(2) & "' AND " & _
                     "PA03 = '" & pa(3) & "' AND " & _
                     "PA04 = '" & pa(4) & "' "
      cnnConnection.Execute strSql
   End If
   'Add by Morgan 2005/5/20
   '非台灣 宣告無效答辯 更新結餘
   If m_strCP10 = "803" Then
      Pub_UpdateEndModCash pa(1), pa(2), pa(3), pa(4)
   End If
   
   'Add by Morgan 2007/5/4
   If bolCancelClose = True Then
      strSql = "UPDATE PATENT SET PA57=NULL,PA58=NULL,PA59=NULL" & _
         " WHERE PA01 = '" & pa(1) & "' AND PA02 = '" & pa(2) & "'" & _
         " AND PA03 = '" & pa(3) & "' AND PA04 = '" & pa(4) & "' "
      cnnConnection.Execute strSql
   End If
   'end 2007/5/4
   
   'Added by Morgan 2011/12/6 台灣案的申復或再審期限要更新到分割案
   If pa(9) = "000" And Text13 = "107" Then
      strSql = "select cp09 from divisioncase,caseprogress" & _
         " where dc05='" & pa(1) & "' and dc06='" & pa(2) & "'" & _
         " and dc07='" & pa(3) & "' and dc08='" & pa(4) & "'" & _
         " and cp01(+)=dc01 and cp02(+)=dc02 and cp03(+)=dc03 and cp04(+)=dc04 and cp10='307' and cp27||cp57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         st307Msg = ""
         '可以有多個分割案
         Do While Not RsTemp.EOF
            strExc(1) = PUB_Update307RefTw(RsTemp(0))
            If strExc(1) <> "" Then
               st307Msg = st307Msg & strExc(1) & vbCrLf
            End If
            RsTemp.MoveNext
         Loop
      End If
   End If

   '2012/10/19 ADD BY SONIA Y53309審查意見通知1202或核駁要內部收文901,承辦期限為系統日起7天(日曆天)--吳若芬(因FCP案要加,此代理人雖無FMP案但仍先加入)
   '2013/1/24 modify by sonia 加Y51542
   'Modified by Morgan 2013/8/28 + Y34210 & X51446 --邱子瑜
   'Modified by Morgan 2013/8/30 + Y47453 & X55778 --羅惠蓮
   'Modified by Morgan 2013/9/6 + Y20065 --邱子瑜
'Modified by Morgan 2013/9/18 改呼叫共用函數
'   If m_bolFMP And (Left(pa(75), 6) = "Y53309" Or Left(pa(75), 6) = "Y51542" Or Left(pa(75), 6) = "Y20065" Or _
'      (Left(pa(75), 6) = "Y34210" And Left(pa(26), 6) = "X51446") Or _
'      (Left(pa(75), 6) = "Y47453" And Left(pa(26), 6) = "X55778")) Then
'
'      strExc(1) = 7
'      strExc(2) = 告知代理人
'      'Added by Morgan 2013/8/28
'      'Y34210 + X51446 14天 --邱子瑜
'      If Left(pa(75), 6) = "Y34210" And Left(pa(26), 6) = "X51446" Then
'         strExc(1) = 14
'      'Added by Morgan 2013/9/6
'      'Y20065 15天 --邱子瑜
'      ElseIf Left(pa(75), 6) = "Y20065" Then
'         strExc(1) = 15
'      End If
'
'      'Y51542 改收其他翻譯 --吳彩菱
'      If Left(pa(75), 6) = "Y51542" Then
'         strExc(2) = "927"
'      End If
'      'end 2013/8/28
'      strCP48 = CompDate(2, Val(strExc(1)), strSrvDate(1))
   ' If m_bolFMP And PUB_ChkAutoRec(pa(1), pa(75), pa(26), DBDATE(Text6), strExc(2), strCP48, , , pa(27), pa(28), pa(29), pa(30)) = True Then
   
   If m_bolFMP And pa(57) = "" Then 'Added by Morgan 2017/4/27 FMP未閉卷才交工程師報告客戶,已閉卷直接交FCP程序--潘韻丞(David 確認)
   
   'Add by Lydia 2014/12/3 核駁及審查意見通知函備註
       Dim sMemo As String
       Dim stBCP16 As String 'Added by Lydia 2022/01/05
        'Remove by Lydia 2021/11/05
        'strExc(2) = ""
        'strExc(7) = "": strExc(3) = "": strExc(4) = "": strExc(5) = ""
        'If Not IsNull(pa(27)) Then strExc(7) = ChangeCustomerL(pa(27))
        'If Not IsNull(pa(28)) Then strExc(3) = ChangeCustomerL(pa(28))
        'If Not IsNull(pa(29)) Then strExc(4) = ChangeCustomerL(pa(29))
        'If Not IsNull(pa(30)) Then strExc(5) = ChangeCustomerL(pa(30))
        'end 2021/11/05
        'Modified by Lydia 2021/11/05 分別傳回B類收文(承辦期限、所限)和C類來函(承辦期限和指定送件日期)
        'sMemo = PUB_GetIncomMemoNew(pa(1) & pa(2) & pa(3) & pa(4), pa(1), strExc(2), ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)), , strCP48, DBDATE(Text6) _
                    , strExc(7), strExc(3), strExc(4), strExc(5))
        Dim stBCP10 As String, stBCP48   As String, stBCP06 As String, stCCP48 As String, stCCP142 As String
        sMemo = PUB_GetIncomMemoNew(pa(1) & pa(2) & pa(3) & pa(4), pa(1), ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)) & "," & ChangeCustomerL(pa(27)) & "," & ChangeCustomerL(pa(28)) & "," & ChangeCustomerL(pa(29)) & "," & ChangeCustomerL(pa(30)), _
                       "", DBDATE(Text6), RC_cp10, stCCP48, stCCP142, stBCP10, stBCP48, stBCP06)

        'Added by Lydia 2021/11/05 更新C類來函的承辦期限和指定送件日期，一併更新指定送件日期之前CP164=2
        If stCCP48 <> "" Then
            'Modified by Lydia 2021/11/16 加註cp64
            strSql = "Update CaseProgress set cp48=" & stCCP48 & ", cp141='3', cp142=" & stCCP142 & ", cp164='2' " & _
                        ", cp64='客戶指定" & ChangeWStringToTDateString(stCCP142) & "之前送件;'||cp64 where cp09='" & strCReceiveNo & "' "
            cnnConnection.Execute strSql, intI
        End If
        'end 2021/11/05
        
        'Added by Lydia 2025/02/05 輸入中間程序來函時自動產生行事曆
        If PUB_AddSCforIncomMemo(pa(1), pa(2), pa(3), pa(4), strCReceiveNo, RC_cp10, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)) & "," & ChangeCustomerL(pa(27)) & "," & ChangeCustomerL(pa(28)) & "," & ChangeCustomerL(pa(29)) & "," & ChangeCustomerL(pa(30))) = False Then
            GoTo ErrorHandler
        End If
        'end 2025/02/05
        
      'Modified by Lydia 2021/11/05 PUB_GetIncomMemoNew已有另外抓B類收文設定
      'If m_bolFMP And Len(sMemo) > 0 Then
      '     If strExc(2) = "" Then strExc(2) = "901"
      If Len(stBCP10) > 0 Then
      'end 'Add by Lydia 2014/12/3
   'end 2013/9/18
         
         strCP20 = PUB_GetCP20(pa(1), stBCP10, stBCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4)) 'Added by Lydia 2022/01/05 之前的strCP20為C類來函設定
         'Modified by Lydia 2021/11/02 FMP案的CP20要抓設定(前面已有記錄) N=>strCP20
         'Modified by Lydia 2021/11/05 改變數
         'strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
            "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
            "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & _
            ",'" & AutoNo("B", 6) & "','" & strExc(2) & "','90'," & CNULL(stCP12) & "," & CNULL(stCP13) & _
            "," & CNULL(Text16) & ",'" & strCP20 & "','N','N','" & strReceiveNo & "'," & strCP48 & ") "
         'Modified by Morgan 2021/11/10 CP43應為核駁函收文號(原放點選的收文號)
         'Modified by Lydia 2022/01/05 +CP16
         strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
            "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48,CP06,CP16) VALUES " & _
            "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & _
            ",'" & AutoNo("B", 6) & "','" & stBCP10 & "','90'," & CNULL(stCP12) & "," & CNULL(stCP13) & _
            "," & CNULL(Text16) & ",'" & strCP20 & "','N','N','" & strCReceiveNo & "'," & stBCP48 & "," & stBCP06 & "," & CNULL(stBCP16, True) & ") "
         cnnConnection.Execute strSql, intI
      End If   '2012/10/19 END
      
   End If 'Added by Morgan 2017/4/27
   
   'Added by Lydia 2023/06/15 寰華案:「414恢復權利-主張優先權106」更新實審期限
   If bolChk414for106 = True Then
      '自優先權資料處移至案件備註
      If strChoseBase <> "" Then
         arrData = Split(strChoseBase, ";")
         strExc(4) = ""
         For intI = 0 To UBound(arrData)
            If Trim(arrData(intI)) <> "" Then
               Call PUB_GetPD060507(Trim(arrData(intI)), strExc(1), strExc(2), strExc(3)) '區分優先權資料
               strSql = "DELETE FROM PRIDATE WHERE PD01='" & pa(1) & "' AND PD02='" & pa(2) & "' AND PD03='" & pa(3) & "' AND PD04 ='" & pa(4) & "' "
               If strExc(1) <> "" Then strSql = strSql & "AND PD06='" & strExc(1) & "' "
               If strExc(2) <> "" Then strSql = strSql & "AND PD05=" & TransDate(strExc(2), 2) & " "
               If strExc(3) <> "" Then strSql = strSql & "AND PD07='" & strExc(3) & "' "
               cnnConnection.Execute strSql
               '備註的部份請詳列視為未主張的優先權國家、日期及優先權號
               strExc(4) = strExc(4) & IIf(Len(strExc(4)) > 0, "、", "") & IIf(strExc(3) <> "", PUB_GetNationName(strExc(3)) & ", ", "") & IIf(strExc(2) <> "", strExc(2) & ", ", "") & IIf(strExc(1) <> "", strExc(1) & ", ", "")
               strExc(4) = IIf(Right(strExc(4), 2) = ", ", Mid(strExc(4), 1, Len(strExc(4)) - 2), strExc(4))
            End If
         Next
         strSql = "UPDATE PATENT SET PA91='" & ChangeTStringToTDateString(strSrvDate(2)) & " 核駁-" & Label3(1) & "的優先權資料:" & strExc(4) & " ;'||PA91 WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "'  "
         cnnConnection.Execute strSql
      End If
             
      '更新公開和實審期限
      strExc(5) = PUB_GetFirstPriDate(pa())
      strExc(9) = ""
      If strExc(5) <> "" And strExc(5) <> strFirstPriDate Then
         '參考一般來函輸入frm04010504_3:「視為未主張」1918=>公開或實審期限的相關總收文號用申請程序的收文號
         strSql = "select cp09 from caseprogress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and instr('" & NewCasePtyList & "',cp10)>0 and cp159=0 order by cp05 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strExc(9) = "" & RsTemp(0)
         End If
         '模組沒有寫備註
         strSql = "Update CaseProgress Set CP64=sqldatet(to_char(sysdate,'yyyymmdd'))||'更新期限：原所限'||sqldatet(CP06)||'，原法限'||sqldatet(CP07)||';'||CP64 where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and cp10='416' and cp158=0 "
         cnnConnection.Execute strSql
         strSql = "Update NextProgress Set NP15=sqldatet(to_char(sysdate,'yyyymmdd'))||'更新期限：原所限'||sqldatet(NP08)||'，原法限'||sqldatet(NP09)||';'||NP15 where np02='" & pa(1) & "' AND np03='" & pa(2) & "' AND np04='" & pa(3) & "' AND np05='" & pa(4) & "' and np07='416' and np06 is null "
         cnnConnection.Execute strSql

         PUB_UpdCfpDate2 pa(1), pa(2), pa(3), pa(4), strExc(5), strExc(9)
         '請彈跳提醒視窗：已更新實體審查期限為: XXX/XX/XX=>移到存檔完成
         strFirstPriDate = "Y" & strFirstPriDate
      End If
   End If
   'end 2023/06/15
   
   'Add By Sindy 2013/4/2 台灣舉發及答辯自動產生一道分析(順德案件亦同),原工程師離職掛王副總
   'Modified by Morgan 2014/12/8 判斷移到前面以便控制來函發文日
   'If pa(9) = 台灣國家代號 And (m_strCP10 = "803" Or m_strCP10 = "804") Then
'Modified by Morgan 2016/3/16
'   If bolAdd941 = True Then
'   'end 2014/12/8
'
'      str941CP14 = m_strCP14
'      strExc(0) = "SELECT ST04,DECODE(ST04,'1',ST06,'1') ST06 FROM STAFF WHERE ST01='" & str941CP14 & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If "" & RsTemp(0).Value <> "1" Then str941CP14 = "71011" '原工程師離職掛王副總
'      End If
'
'      str941ReceiveNo = AutoNo("B", 6)
'      '本所期限為系統日+3個工作天
'      str941CP06 = CompWorkDay(3, strSrvDate(1), 0)
'      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06," & _
'         "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
'         "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & str941CP06 & _
'         ",'" & str941ReceiveNo & "','941','90'," & CNULL(GetSalesArea(PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4)))) & "," & CNULL(PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))) & _
'         ",'" & str941CP14 & "','N','N','N','" & strCReceiveNo & "'," & str941CP06 & ") "
'      cnnConnection.Execute strSql
'
''      If "" & RsTemp(1).Value <> "1" Then '分所
''         strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & CompWorkDay(2, strSrvDate(1), 0) & " WHERE EP02='" & str941ReceiveNo & "'"
''      Else
'         strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & str941ReceiveNo & "'"
''      End If
'      cnnConnection.Execute strSql
'
'      '更新承辦期限=本所期限,因ENGINEERPROGRESS_BEFORE5及CASEPROGRESS_AFTER6會造成承辦期限=本所期限-1天
'      strSql = "UPDATE CASEPROGRESS SET CP48=CP06 WHERE CP09='" & str941ReceiveNo & "'"
'      cnnConnection.Execute strSql
   If m_bolEngCase Then
      '不會稿,判發人73022
      'Modified by Morgan 2025/2/19 73022->pub_PMan
      pub_PMan = Pub_GetSpecMan("專利處特定編號")
      strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & ",EP34='N',EP40='" & Left(pub_PMan, 5) & "' WHERE EP02='" & strCReceiveNo & "'"
      cnnConnection.Execute strSql
      
      '承辦期限=系統日+3個工作天(沿用原規則)
      str941CP06 = CompWorkDay(3, strSrvDate(1), 0)
      strSql = "UPDATE CASEPROGRESS SET CP48=" & str941CP06 & " WHERE CP09='" & strCReceiveNo & "'"
      cnnConnection.Execute strSql
'end 2016/3/16

   '2012/11/26 ADD BY SONIA  順德及其關係企業加內部收文941分析,原工程師離職掛王副總,自動上齊備日(分所上下一工作日)以計算承辦期限,本所期限=承辦期限,但因ENGINEERPROGRESS_BEFORE5及CASEPROGRESS_AFTER6會造成承辦期限=本所期限-1天
   'If m_CustX07166 = True Then
   ElseIf m_CustX07166 = True Then
   '2013/4/2 End
      'str941CP14 = m_strCP14  '2013/2/8 cancel by sonia 移至basQuery的PUB_CheckX07166Remind
      
      strExc(0) = "SELECT ST04,DECODE(ST04,'1',ST06,'1') ST06 FROM STAFF WHERE ST01='" & str941CP14 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Memo by Amy 2024/07/16 抓str941CP14人員不會抓到離職,PUB_CheckX07166Remind有過濾
         If "" & RsTemp(0).Value <> "1" Then str941CP14 = "71011" '原工程師離職掛王副總
      End If
      
      str941ReceiveNo = AutoNo("B", 6)
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
         "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43) VALUES " & _
         "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & _
         ",'" & str941ReceiveNo & "','941','90'," & CNULL(stCP12) & "," & CNULL(stCP13) & _
         ",'" & str941CP14 & "','N','N','N','" & strCReceiveNo & "') "
      cnnConnection.Execute strSql
      If "" & RsTemp(1).Value <> "1" Then '分所
         strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & CompWorkDay(2, strSrvDate(1), 0) & " WHERE EP02='" & str941ReceiveNo & "'"
      Else
         strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & str941ReceiveNo & "'"
      End If
      cnnConnection.Execute strSql
      '更新本所期限=承辦期限,但因ENGINEERPROGRESS_BEFORE5及CASEPROGRESS_AFTER6會造成承辦期限=本所期限-1天
      strSql = "UPDATE CASEPROGRESS SET CP06=CP48 WHERE CP09='" & str941ReceiveNo & "' AND CP06 IS NULL"
      cnnConnection.Execute strSql
   End If
   '2012/11/26 END
   
   
   'Added by Morgan 2021/10/6 長庚醫院案件要收[轉公文]先簡單報告
   If m_CustX69365 = True Then
   
      'Removed by Morgan 2022/3/28 取消轉公文,改同其他3家直接報告,但本所期限改為 +14天-3個工作天 --黃教威
      'm_str1998CP09 = AutoNo("D", 6)
      'strSql = "INSERT INTO CASEPROGRESS(cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp27" & _
         ",cp32,cp43) SELECT cp01,cp02,cp03,cp04,cp05,cp06,cp07,'" & m_str1998CP09 & "','1998',cp12,cp13,'" & strUserNum & "'" & _
         ",'N','N'," & strSrvDate(1) & ",'N',cp09 FROM CASEPROGRESS WHERE CP09='" & strCReceiveNo & "'"
      'cnnConnection.Execute strSql, intI
      ''P案轉公文用系統的來函定稿，判發人同來函
      'strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), RC_cp10, pa(9), m_strCP10, , m_bolFMP)
      'PUB_AddLetterProgress m_str1998CP09, 0, True, strExc(1), IIf(Val(Text14(1)) > 0, True, False), pa(26), RC_cp10, pa(75)
      'PUB_SetX69365Case1998CP06 m_str1998CP09 '設定長庚醫院案件轉公文管制日(所限)
      'end 2022/3/28
      
      PUB_SetX69365CaseOACP06 strCReceiveNo '設定長庚醫院案件OA發文管制日(所限)
      
   'Added by Morgan 2022/4/15
   ElseIf m_bolW2001XCase Then
      'Added by Morgan 2022/11/4 原工程師離職只需通知王副總分案，程序分案會再通知新工程師--有跟郭確認過
      'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
      If Text16 = "99050" Then
         Call PUB_SendMail(strUserNum, "99050", strCReceiveNo, "分案通知")
      Else
      'end 2022/11/4
      
         strExc(0) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
         strExc(1) = strExc(0) & "案已收到「核駁」，請於「承辦期限」前完成分析通知函，謝謝。"
         strExc(2) = PUB_GetW2001InCC(pa(26), pa(158))
         'Modified by Morgan 2022/8/9 應該要發給承辦人CC給窗口及智權人員
         'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " select  '" & strUserNum & "','" & strExc(2) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",replace('" & ChgSQL(strExc(1)) & "','承辦期限',sqldatet(cp48)),'如旨' from caseprogress where cp09='" & strCReceiveNo & "'"
         'Modified by Morgan 2023/5/10 W2001-->stCP13
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " select  '" & strUserNum & "',cp14,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",replace('" & ChgSQL(strExc(1)) & "','承辦期限',sqldatet(cp48)),'如旨','" & stCP13 & ";" & strExc(2) & "'" & _
            " from caseprogress where cp09='" & strCReceiveNo & "'"
         cnnConnection.Execute strSql, intI
         
      End If 'Added by Morgan 2022/11/4
   'end 2022/4/15
      
   End If
   'end 2021/10/6
   
   If m_USCaseNo <> "" Then PUB_SetUsIDS pa(1), pa(2), pa(3), pa(4), strCReceiveNo, Text6.Text, , , , True     'Added by Morgan 2020/12/18 美國IDS期限管制
   
   'Added by Morgan 2020/4/10
   'FMP有期限之案件EMAIL通知(寰華案不必--敏莉)
   If m_bolFMP = True And Left(Pub_StrUserSt03, 1) <> "F" Then
      'Modified by Morgan 2020/9/15 未閉卷的併入工程師通知信
      If pa(57) = "Y" Then
         'Modified by Morgan 2023/4/10 +bolReKeyInCase
         'Modified by Morgan 2023/5/25 FMP電子化所有來函應該都要EMail通知
         'PUB_FMPCaseInform strCReceiveNo, , , , bolReKeyInCase
         PUB_FMPCaseInform strCReceiveNo, False, True, Left(Pub_StrUserSt03, 1) = "F", bolReKeyInCase
         'end 2023/5/25
      End If
      'end 2020/9/15
   End If
   'end 2020/4/10
   
   'Added by Morgan 2022/7/20 --Anny
   '收到CNIPA來函 "復審107之核駁1002",申請人編號為舊名字(X___001)時，則系統自動發email通知相關人員警示本案申請人已有新名字
   If m_strCP10 = "107" Then
      For intI = 1 To 5
         If (Len(pa(25 + intI)) = 9 And Right(pa(25 + intI), 1) <> "0") Then
            'Modified by Morgan 2023/4/10 +bolReKeyInCase
            PUB_POAInform pa(1), pa(2), pa(3), pa(4), strCReceiveNo, bolReKeyInCase
            Exit For
         End If
      Next
   End If
   'end 2022/7/20
   
   'Added by Morgan 2025/3/12
   '大陸案(431)PPH輸入核駁時，下一程序為(431)PPH, 法定期限為收文日加1個月(本所管制用，非官方期限，直接新增不2次確認)--玲玲/敏莉
   If RC_cp10 = 核駁 And pa(9) = "020" And m_strCP10 = "431" Then
      bolCN431 = True
      'Added by Morgan 2025/3/13
      '若已提過2次PPH則不必再新增NP--敏莉/玲玲
      strExc(0) = "select cp09 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='431' and cp09<>'" & strReceiveNo & "' and cp27>0 and cp159=0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
      'end 2025/3/13
         strExc(3) = GetNextProgressNo
         strExc(1) = CompDate(1, 1, strSrvDate(1))
         'Modified by Lydia 2025/10/29
         'strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         'strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22)" & _
         '   " VALUES ('" & strCReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & _
         '   "','431'," & strExc(2) & "," & strExc(1) & "," & CNULL(stCP13) & "," & strExc(3) & ")"
         'stNP23 = "NULL" 'Mark by Lydia 2025/11/06
         If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
            strExc(2) = PUB_GetPOurDeadline(strExc(1), pa(9), stNP23, pa(1), "431")
         Else
            strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         End If
         'end 2025/10/29
         strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22,NP23)" & _
            " VALUES ('" & strCReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & _
            "','431'," & strExc(2) & "," & strExc(1) & "," & CNULL(stCP13) & "," & strExc(3) & "," & stNP23 & ")"
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2025/3/12
   
   'Added by Morgan 2023/4/11 從下面移上來
   'Modified by Morgan 2025/3/12 +bolCN431
   If IsEmptyText(strProgressNo) = False Or bolCN431 = True Then
      If m_bolFMP And pa(57) = "" Then
      
         'Added by Morgan 2025/3/19 寰華案改正本給工程師,副本給主管--敏莉/Wilson
         If Left(Pub_StrUserSt03, 1) = "F" Then
            mRCno = Trim(Text16.Text)
            mCCno = PUB_GetFCPEngSup(mRCno)
         Else
         'end 2025/3/19
            mCCno = Trim(Text16.Text)
            mRCno = PUB_GetFCPEngSup(mCCno)
         End If
         
         If mCCno = mRCno Then mCCno = ""
            
         strExc(0) = "SELECT NVL(PA05,NVL(PA06,PA07)) pa05,nvl(FA05||' '||FA63,'') as faname1, nvl(FA04,'') as faname2, nvl(FA06,'') as faname3,CP48,NP23 " & _
                     "FROM PATENT,FAGENT,caseprogress,nextprogress WHERE substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) and CP01(+)=PA01 And CP02(+)=PA02 And CP03(+)=PA03 And CP04(+)=PA04 " & _
                     "And CP09 = '" & strCReceiveNo & "' and np01(+)=cp09 and np06(+) is null and PA01='" & pa(1) & "' and PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(0) = "" & RsTemp.Fields("PA05")
            strExc(1) = "" & RsTemp.Fields("faname1")
            strExc(2) = "" & RsTemp.Fields("faname2")
            strExc(3) = "" & RsTemp.Fields("faname3")
            strExc(4) = "" & RsTemp.Fields("CP48") '承辦期限
            strExc(5) = "" & RsTemp.Fields("NP23") '約定期限
         End If
         If Len(strExc(1)) > 0 Then '代理人名稱(英->中->日)
            strExc(1) = "代理人　：" & strExc(1)
         ElseIf Len(strExc(2)) > 0 Then
            strExc(1) = "代理人　：" & strExc(2)
         Else
            strExc(1) = "代理人　：" & strExc(3)
         End If
            
         oSubject = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
         
         '發E-Mail通知承辦人
         oContext = "※若改承辦人,請工程師主管轉寄給新的承辦人,及c.c.原承辦人" & vbCrLf & vbCrLf 'Added by Morgan 2020/9/17 從最後面移到最前面--敏莉
         'Modified by Morgan 2025/3/12 PPH核駁加案件性質中文
         oContext = oContext & _
                    "本所案號：" & oSubject & "　　" & vbTab & vbTab & "來函收文日：" & ChangeTStringToTDateString(Trim(Label3(3))) & vbCrLf & _
                    "專利名稱：" & strExc(0) & vbCrLf & _
                        strExc(1) & vbCrLf & _
                    "承辦人　：" & Trim(Label3(4)) & vbCrLf & _
                    "本所期限：" & ChangeTStringToTDateString(Trim(Text14(0))) & IIf(Trim(Text14(0)) = "", "　　　　", "") & "　　　　" & vbTab & vbTab & "法定期限：" & ChangeTStringToTDateString(Trim(Text14(1))) & vbCrLf & _
                    "承辦期限：" & IIf(Len(strExc(4)) > 0, ChangeWStringToTDateString(strExc(4)), "　　　　") & "　　　　" & vbTab & vbTab & "來函性質：核駁" & IIf(m_strCP10 = "431", "-" & Label3(1), "") & vbCrLf & _
                    "約定期限：" & ChangeWStringToTDateString(strExc(5)) & vbCrLf
             
         'Added by Morgan 2021/11/10 有收文告代時，主旨及內文都要加
         If stBCP10 = "901" Then
            oSubject = oSubject & "，同時內部收文【告代】"
            oContext = oContext & vbCrLf & "案件性質：告代" & vbCrLf
            oContext = oContext & "承辦期限：" & ChangeWStringToTDateString(stBCP48) & vbCrLf
            oContext = oContext & "本所期限：" & ChangeWStringToTDateString(stBCP06) & vbCrLf
         End If
         'end 2021/11/10
            
          'Modified by Morgan 2020/9/17 寰華案不用cc給Phoebe--敏莉
          If Left(Pub_StrUserSt03, 1) = "F" Then '寰華
            'Modified by Lydia 2022/05/10 寰華案與FMP案主旨一致
            'Modified by Lydia 2024/04/26 機械組要加註
            'Modified by Morgan 2025/3/12 PPH核駁加案件性質中文
            'Modified by Morgan 2025/3/19 --敏莉/Wilson
            'oSubject = IIf(pa(150) = "4", "【機械設計組】", "") & "FMP(寰華)案核駁" & IIf(m_strCP10 = "431", "-" & Label3(1), "") & "通知:" & oSubject & "，主管請分案，工程師請處理後續流程，謝謝！"
            oSubject = IIf(pa(150) = "4", "【機械設計組】", "") & "FMP(寰華案)核駁" & IIf(m_strCP10 = "431", "-" & Label3(1), "") & "通知:" & oSubject & "，工程師請處理後續流程，謝謝！"
            'end 2025/3/19
            'Added by Lydia 2022/04/22 核駁及一般來函皆CC給程序
            strExc(0) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
            If InStr(mCCno & ";", strExc(0)) = 0 And strExc(0) <> "" Then
                mCCno = mCCno & ";" & strExc(0)
            End If
            'end 2022/04/22
            'Added by Lydia 2023/01/09 FCP和寰華案 key C類來函，若key來函人員沒有在系統自動發Outlook的收件者中，副本請加上key來函人員;
            If InStr(mRCno & ";" & mCCno, strUserNum) = 0 Then
                mCCno = mCCno & IIf(mCCno <> "", ";", "") & strUserNum
            End If
            'end 2023/01/09
            
            If m_bolReKeyInOK Then oSubject = "(重發，請以此封為準)" & oSubject 'Added by Morgan 2023/4/17
          Else
            'Removed by Morgan 2023/5/25 不必再CC給FMP案外專程序窗口--敏莉
            'mCCno = mCCno & ";" & Pub_GetSpecMan("FMP案外專程序窗口")
            'end 2023/5/25
            'Modified by Morgan 2025/3/12 PPH核駁加案件性質中文
            oSubject = "FMP案核駁" & IIf(m_strCP10 = "431", "-" & Label3(1), "") & "通知:" & oSubject & "，主管請分案，工程師請處理後續流程，謝謝！"
          End If
          
          strCMemo = PUB_FCPCFormMemo(strCReceiveNo)  'Added by Morgan 2023/6/21
          If strCMemo <> "" Then
            oContext = oContext & vbCrLf & "備註:" & vbCrLf & strCMemo
          End If
         
         '需2次確認的來函改等職代輸入並確認後才通知
         'Modified by Morgan 2023/5/26 寰華除外
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc12,mc13)" & _
            " values('" & strUserNum & "','" & mRCno & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & ChgSQL(oSubject) & "','" & ChgSQL(oContext) & "','" & mCCno & "'," & IIf(bolReKeyInCase And Not Left(Pub_StrUserSt03, 1) = "F", "99999999", "0") & ",'" & strCReceiveNo & "')"
         cnnConnection.Execute strSql, intI
         m_bolFMPNoPrint = True
      End If
   End If
   'end 2023/4/11
   
   'Added by Morgan 2023/6/27
   If m_bolBPFCase Then Pub_COrderInform strCReceiveNo, , IIf(Text16 = "A0029", "", "A0029")
   'end 2023/6/27

   'Added by Morgan 2025/3/7 面詢未辦理，向官方辦理退費控管--玲玲
   If m_bolAddB908 Then
      strExc(9) = AutoNo("B", 6)
      strExc(1) = PUB_GetPHandler(pa(1) & pa(2) & pa(3) & pa(4))
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
         "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP64) VALUES " & _
         "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & _
         ",'" & strExc(9) & "','908','90','" & stCP12 & "','" & stCP13 & "'" & _
         ",'" & strExc(1) & "','N','N','N','" & strCReceiveNo & "','退請求面詢規費') "
      cnnConnection.Execute strSql, intI
   End If
   'end 2025/3/7

   If RC_cp10 = 核駁 And pa(9) = "000" Then PUB_ChkTW413 strReceiveNo, True 'Added by Morgan 2025/3/13
   
   'Add By Cheng 2002/11/08
   cnnConnection.CommitTrans
   FormSave = True
         
   'Add by Morgan 2011/12/6
   If st307Msg <> "" Then MsgBox st307Msg

   '2012/11/26 add by sonia 順德及其關係企業案件,若承辦人是王協理且未發文則要發EMail通知
   'Modify By Sindy 2013/4/2
   'If m_CustX07166 = True And str941CP14 = "71011" Then
   'Modified by Morgan2022/10/17
   'If m_CustX07166 = True Or (pa(9) = 台灣國家代號 And (m_strCP10 = "803" Or m_strCP10 = "804")) Then
   If str941ReceiveNo <> "" Then
   'end 2022/10/17
      'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
      If str941CP14 = "99050" Then
   '2013/4/2 End
         Call PUB_SendMail(strUserNum, "99050", str941ReceiveNo, "分案通知")
      'end 2024/07/16
      'Added by Morgan 2018/11/13 分析B類接洽單改承辦人是王副總要印紙本分案,其他存卷宗區並EMail通知承辦工程師
         g_PrtForm001.PrintCForm str941ReceiveNo
      Else
         g_PrtForm001.PrintCForm str941ReceiveNo, , , True
         Pub_COrderInform str941ReceiveNo, True
      'end 2018/11/13
      
      End If
   End If
   '2012/11/26 end
   
'   '2013/3/18 ADD BY SONIA 順德及其關係企業案件之分析B類收文加印接洽單
'   'Modify By Sindy 2013/4/2 台灣舉發及答辯同時跑一張B類接洽記錄單
'   'If m_CustX07166 = True Then
'   If m_CustX07166 = True Or (pa(9) = 台灣國家代號 And (m_strCP10 = "803" Or m_strCP10 = "804")) Then
'   '2013/4/2 End
'      g_PrtForm001.PrintCForm str941ReceiveNo  'B類接洽記錄單
'   End If
'   '2013/3/18 END
'end 2015/11/17
   
   'Add By Sindy 2013/4/29 列印案件回覆單
   'Modified by Morgan 2016/4/18 工程師承辦要印
   'If pa(9) = 台灣國家代號 And (m_strCP10 = "803" Or m_strCP10 = "804") And strNP07 <> "" Then
   'Modified by Morgan 2021/9/22
   'If m_bolEngCase Then
   'Modified by Morgan 2023/5/10 若工程師承辦時會預設程序的定稿，已有回覆單可不必再印
   'If m_bolEngCase Or m_bolBPFCase Or m_bolW2001XCase Then
   If m_bolBPFCase Then
   'end 2016/4/18
      Call g_PrtForm001.PrintReturnSheet(strNP01, strNP07, DBDATE(strNP09), , , , , pa(1) & pa(2) & pa(3) & pa(4))
   End If
   '2013/4/29 End

   'Added by Morgan 2024/10/4
   If pa(1) = "P" And pa(9) <> "000" And m_bolFMP = False Then
      If Pub_B911NotPay(pa(1), pa(2), pa(3), pa(4)) = True Then
          MsgBox "此案有未收款！", vbExclamation
      End If
   End If
   'end 2024/10/4
         
   ' 90.07.17 modify by louis (列印接洽結案單)->下一程序決定是否要辦(C類的接洽結案單   )
   If FormSave = True Then
      If IsEmptyText(strProgressNo) = False Then
         g_PrtForm001.PrintForm strProgressNo, pa(1), pa(2), pa(3), pa(4)

         'Remove by Morgan 2009/12/16
         'If m_bolFMP Then
         '   bol901 = True
         '   g_PrtForm001.PrintForm strProgressNo, pa(1), pa(2), pa(3), pa(4), m_901CP09
         '   bol901 = False
         'End If
         
         'Add by Lydia 2014/10/16 FMP案會列印 C類接洽單, 請同時E-MAIL給畫面上之承辦人, 副本發給該員之工程師組別主管.
         'Modified by Morgan 2017/4/27 FMP未閉卷才交工程師報告客戶,已閉卷直接交FCP程序--潘韻丞
          'If m_bolFMP = True Then
          If m_bolFMP And pa(57) = "" Then
          'end 2017/4/27
                            
            'Modified by Morgan 2020/9/15 自列印定稿後移來並+strCMemo
            'Modified by Lydia 2020/04/06 因應防疫在家上班作業，請將FMP案key來函產生的C類接洽記錄單回存到卷宗區
            '                                         比照FCP案C類接洽單同時列印並且上傳到卷宗區frm06010603_3: 原本就不傳入特殊備註,等到列印時再抓特殊備註
            'g_PrtForm001.PrintCForm strCReceiveNo, , stCP48Desc
            g_PrtForm001.PrintCFormNew strCReceiveNo, , stCP48Desc, True, strCMemo
                  
'Removed by Morgan 2023/4/11 EMail要2次確認完才要發，改移到上面先寫暫存
'            'Modified by Lydia 2020/08/24 改用模組
'            'strExc(0) = "SELECT ST01,ST04,decode(ST16,'1','T','2','R','3','S','4','T1','') mst16 FROM STAFF WHERE ST01='" & Trim(Text16.Text) & "' "
'            'intI = 1
'            'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            'If intI = 1 Then
'            '  strExc(0) = "" & RsTemp.Fields("mst16")
'            '  mCCno = Pub_GetSpecMan(strExc(0))
'            '  mRCno = RsTemp.Fields("ST01")
'            '  If mCCno = mRCno Then mCCno = "" '承辦人已是主管則不必再發副本
'            'End If
'            'Modified by Morgan 2020/9/15 改寄承辦工程師主管,cc承辦工程師; Phoebe(FMP案外專程序窗口)
'            'mRCno = Trim(Text16.Text)
'            'mCCno = PUB_GetFCPEngSup(mRCno)
'            'If mCCno = mRCno Then mCCno = ""
'            mCCno = Trim(Text16.Text)
'            mRCno = PUB_GetFCPEngSup(mCCno)
'            If mCCno = mRCno Then mCCno = ""
'            'end 2020/9/15
'            'end 2020/08/24
'
'            strExc(0) = "SELECT NVL(PA05,NVL(PA06,PA07)) pa05,nvl(FA05||' '||FA63,'') as faname1, nvl(FA04,'') as faname2, nvl(FA06,'') as faname3,CP48,NP23 " & _
'                        "FROM PATENT,FAGENT,caseprogress,nextprogress WHERE substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) and CP01(+)=PA01 And CP02(+)=PA02 And CP03(+)=PA03 And CP04(+)=PA04 " & _
'                        "And CP09 = '" & strCReceiveNo & "' and np01(+)=cp09 and np06(+) is null and PA01='" & pa(1) & "' and PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' "
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strExc(0) = "" & RsTemp.Fields("PA05")
'               strExc(1) = "" & RsTemp.Fields("faname1")
'               strExc(2) = "" & RsTemp.Fields("faname2")
'               strExc(3) = "" & RsTemp.Fields("faname3")
'               strExc(4) = "" & RsTemp.Fields("CP48") '承辦期限
'               strExc(5) = "" & RsTemp.Fields("NP23") '約定期限
'            End If
'            If Len(strExc(1)) > 0 Then '代理人名稱(英->中->日)
'               strExc(1) = "代理人　：" & strExc(1)
'            ElseIf Len(strExc(2)) > 0 Then
'               strExc(1) = "代理人　：" & strExc(2)
'            Else
'               strExc(1) = "代理人　：" & strExc(3)
'            End If
'
'            '發E-Mail通知承辦人
'            oContext = "※若改承辦人,請工程師主管轉寄給新的承辦人,及c.c.原承辦人" & vbCrLf & vbCrLf 'Added by Morgan 2020/9/17 從最後面移到最前面--敏莉
'            oContext = oContext & _
'                       "本所案號：" & oSubject & "　　" & vbTab & vbTab & "來函收文日：" & ChangeTStringToTDateString(Trim(Label3(3))) & vbCrLf & _
'                       "專利名稱：" & strExc(0) & vbCrLf & _
'                           strExc(1) & vbCrLf & _
'                       "承辦人　：" & Trim(Label3(4)) & vbCrLf & _
'                       "本所期限：" & ChangeTStringToTDateString(Trim(Text14(0))) & "　　　　" & vbTab & vbTab & "法定期限：" & ChangeTStringToTDateString(Trim(Text14(1))) & vbCrLf & _
'                       "承辦期限：" & IIf(Len(strExc(4)) > 0, ChangeWStringToTDateString(strExc(4)), "　　　　") & "　　　　" & vbTab & vbTab & "來函性質：核駁" & vbCrLf & _
'                       "約定期限：" & ChangeWStringToTDateString(strExc(5)) & vbCrLf
'
'             'Modified by Morgan 2020/9/15
'             'oSubject = pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
'             'oSubject = oSubject & "　收文-核駁，請自行去調卷處取卷，謝謝！"
'             oSubject = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
'
'            'Added by Morgan 2021/11/10 有收文告代時，主旨及內文都要加
'            If stBCP10 = "901" Then
'               oSubject = oSubject & "，同時內部收文【告代】"
'               oContext = oContext & vbCrLf & "案件性質：告代" & vbCrLf
'               oContext = oContext & "承辦期限：" & ChangeWStringToTDateString(stBCP48) & vbCrLf
'               oContext = oContext & "本所期限：" & ChangeWStringToTDateString(stBCP06) & vbCrLf
'            End If
'            'end 2021/11/10
'
'             'Modified by Morgan 2020/9/17 寰華案不用cc給Phoebe--敏莉
'             If Left(Pub_StrUserSt03, 1) = "F" Then '寰華
'               'Modified by Lydia 2022/05/10 寰華案與FMP案主旨一致
'               'oSubject = "FMP(寰華)案核駁通知:" & oSubject & "，主管請分案，卷隨後附上，謝謝！"
'               oSubject = "FMP(寰華)案核駁通知:" & oSubject & "，主管請分案，工程師請處理後續流程，謝謝！"
'               'Added by Lydia 2022/04/22 核駁及一般來函皆CC給程序
'               strExc(0) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
'               If InStr(mCCno & ";", strExc(0)) = 0 And strExc(0) <> "" Then
'                   mCCno = mCCno & ";" & strExc(0)
'               End If
'               'end 2022/04/22
'               'Added by Lydia 2023/01/09 FCP和寰華案 key C類來函，若key來函人員沒有在系統自動發Outlook的收件者中，副本請加上key來函人員;
'               If InStr(mRCno & ";" & mCCno, strUserNum) = 0 Then
'                   mCCno = mCCno & IIf(mCCno <> "", ";", "") & strUserNum
'               End If
'               'end 2023/01/09
'             Else
'               mCCno = mCCno & ";" & Pub_GetSpecMan("FMP案外專程序窗口")
'               'Modified by Morgan 2021/11/10 --淑華
'               'oSubject = "FMP案核駁通知:" & oSubject & "，主管請分案，工程師請自行去調卷處取卷，紙本公文後補，謝謝！"
'               oSubject = "FMP案核駁通知:" & oSubject & "，主管請分案，工程師請處理後續流程，謝謝！"
'               'end 2021/11/10
'             End If
'             If strCMemo <> "" Then
'               oContext = oContext & vbCrLf & "備註:" & vbCrLf & strCMemo
'             End If
'             'end 2020/9/15
'
'            PUB_SendMail strUserNum, mRCno, "", oSubject, oContext, "", "", , , , mCCno, "", "", ""

          End If  'end Lydia 2014/10/16
          
      End If
   End If
Exit Function

ErrorHandler:
    If FormSave = False Then cnnConnection.RollbackTrans
'Resume
'    FormSave = False
End Function

Private Sub Form_Initialize()
   'add by nickc 2007/02/02
   ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   Dim bPaper As Boolean
   
   MoveFormToCenter Me
   intWhere = 國內
   With frm04010503_2
      pa(1) = .Text2
      pa(2) = .Text3
      pa(3) = .Text4
      pa(4) = .Text5
      strReceiveNo = .Tag
      m_strNP14 = strExc(5)
      ReadPatent
   End With
   Combo2.ListIndex = 0
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010503_2.m_strIR01
   m_strIR02 = frm04010503_2.m_strIR02
   m_strIR03 = frm04010503_2.m_strIR03
   m_strIR04 = frm04010503_2.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
   'Add By Cheng 2002/06/21
   '若申請國家為台灣, 則帶出主管機關文號
   If pa(9) = 台灣國家代號 Then
      strExc(0) = "SELECT CF24 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & m_strCP10 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Me.Text20.Enabled = True
         Me.Text20.Text = "" & RsTemp(0).Value
      End If
   End If
    
    '若申請國家非台灣時, 鎖住來函期限欄位
    'Remove by Morgan 2008/5/13 不必鎖了
'    If pa(9) <> 台灣國家代號 Then
'        Me.Frame1.Enabled = False
'        Me.Frame2.Enabled = False
'    End If
    
    
    '游標預設在機關文號欄
   'Modify by Morgan 2008/5/21 非台灣停在核駁函日期 且預設期限選項為文到次日,月
   'Text6 = Label3(3).Caption
   'SendKeys "{Tab}"
   
   If pa(9) = 台灣國家代號 Then
      Text6 = Label3(3).Caption
      SendKeys "{Tab}"
      'Modified by Morgan 2015/6/24 +501,505
      'Modified by Morgan 2016/3/17 +判斷發文非111111(中間來所的舉發案由程序改定稿)
      If (m_strCP10 = "803" Or m_strCP10 = "804" Or m_strCP10 = "501" Or m_strCP10 = "505") And m_CP27 <> "19221111" Then Text15 = "N" 'Add By Sindy 2013/4/2 舉發及舉發答辯不出通知函
      
      '2018/12/26
      'If m_bolEngCase = True Then Text15 = "N": Text15.Enabled = False 'Added by Morgan 2016/3/17 工程師承辦的來函不出系統定稿
      If m_bolEngCase = True Then
         Text15 = "N"
      End If
   Else
      Option1(1).Value = True
      Option4(1).Value = True
      Text6.MaxLength = 8
      Text12.MaxLength = 8
   End If
   
   'Added by Morgan 2021/9/22
   If m_bolBPFCase Or m_bolW2001XCase Then
      Text15 = "N"
   End If
   'end 2021/9/22
   
   'Added by Morgan 2024/4/30 非FMP的核駁改可輸入，有可能會要工程師承辦--郭
   If Not m_bolFMP And pa(1) = "P" Then
      Text16.Enabled = True
   Else
   'end 2024/4/30
   
      Text16.Enabled = False 'Add by Morgan 2006/3/28 承辦人欄位不可修改
      
   End If
   m_CustX07166 = False '2012/11/26 add by sonia
   
   'Added by Morgan 2015/6/23
   '1001,1002,1202,1209,1802,1807,1809,1810 E化提醒
   If PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4), , , bPaper) = True And bPaper = False Then
      MsgBox "E化案件，不印前案!!", vbExclamation
   End If
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, i As Integer, rsTemp1 As New ADODB.Recordset
Dim strTmp As String
Dim m_Have808 As Boolean   'add by sonia 2018/12/27
    
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   Label3(3).Caption = frm04010503_1.Text5.Text
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label3(0) = pa(5)
      Label3(2) = pa(10)
      Text1 = pa(11)
      If pa(16) = "1" Then
         Label3(6) = "基本檔目前准駁 : 准"
      ElseIf pa(16) = "2" Then
         Label3(6) = "基本檔目前准駁 : 駁"
      Else
         Label3(6) = "基本檔目前准駁 : 無"
      End If
      text8 = pa(17)
   End If
   
   If pa(9) = 台灣國家代號 Then
      strTmp = "CPM03"
   Else
      strTmp = "CPM04"
   End If
   
   'Modified by Morgan 2016/3/16 +CP27
   'Modified b Lydia 2023/06/15 +CP43
   strExc(0) = "SELECT CP10," & strTmp & ",CP12,CP13,CP14,CP35,CP117,CP27,CP43 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE " & _
      "CP09='" & strReceiveNo & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP27 = "" & .Fields("CP27") 'Added by Morgan 2016/3/16
      m_strCP10 = "" & .Fields("cp10")
      m_strCP43 = "" & .Fields("cp43") 'Added by Lydia 2023/06/15
      Label3(1) = "" & .Fields(1)

'      If Left(m_strCP10, 1) = "1" Then
'         strExc(0) = "SELECT CP14 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10=" & 翻譯
'         intI = 1
'         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            'MODIFY BY SONIA 90.11.27不預設承辦人
'            'If Not IsNull(rsTemp1.Fields(0)) Then
'            '   Text16 = rsTemp1.Fields(0): ChgType 16
'            'End If
'         End If
'      Else
'         'MODIFY BY SONIA 90.11.27不預設承辦人
'         'If Not IsNull(.Fields(4)) Then
'         '   Text16 = .Fields(4): ChgType 16
'         'End If
'      End If

      'Added by Morgan 2021/1/28 從 Formsave 移來以便共用
      stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
      stCP12 = GetSalesArea(stCP13)
      'end 2021/1/28

      If Not IsNull(.Fields(5)) Then Text19 = .Fields(5)
      Text21 = "" & .Fields("CP117") 'Add by Morgan 2008/5/13
      '92.5.8 ADD BY SONIA 承辦人預設輸入人員及不算案件數
      'Modify by Morgan 2009/11/27 FMP 案承辦人另外預設
      'Text16 = strUserNum: ChgType 16
      'Text18 = "N"
      'Modified by Morgan 2021/1/28
      'If Left(.Fields("cp12"), 1) = "F" Then
      If Left(stCP12, 1) = "F" Then
      'end 2021/1/28
         m_bolFMP = True
         'Added by Lydia 2015/06/29 外專寰華的案件，在輸入各式審查機關來函的畫面，能帶出約定期限欄位
         txtSNP23.Locked = True
         'Modified by Morgan 2017/10/11 FMP預設承辦人比照FCP
         'Text16 = PUB_GetFmpCP14(pa)
         Text16 = PUB_GetFCPPromoterNo(strReceiveNo, 核駁, "" & .Fields("cp14"))
         'end 2017/10/11
         ChgType 16
         Text18 = ""
      Else
         m_bolFMP = False
         'Added by Lydia 2015/06/29 外專寰華的案件，在輸入各式審查機關來函的畫面，能帶出約定期限欄位
         lblsNP23.Visible = False: txtSNP23.Visible = False: txtSNP23.Locked = True
         
         'Modified by Morgan 2016/3/16
         '臺灣案舉發或舉發答辯的審定來函預設原工程師承辦,若離職則改游經理(73022),但假發文的則由程序承辦
         'Text16 = strUserNum: ChgType 16
         'modify by sonia 2018/11/9 +行政訴訟503,參加訴訟506
         If pa(9) = 台灣國家代號 And (m_strCP10 = "803" Or m_strCP10 = "804" Or m_strCP10 = "501" Or m_strCP10 = "505" Or m_strCP10 = "503" Or m_strCP10 = "506") And m_CP27 <> "19221111" Then
            m_bolEngCase = True
            If GetStaffName(.Fields("CP14")) <> "" Then
               strExc(1) = .Fields("CP14")
               'add by sonia 2024/7/15 A7010柯昱安調離也要改為游經理73022
               If GetStaffDepartment(strExc(1)) >= "P10" And GetStaffDepartment(strExc(1)) <= "P11" Then
               Else
                  'Modified by Morgan 2025/2/21 73022->left(pub_PMan,5)
                  pub_PMan = Pub_GetSpecMan("專利處特定編號")
                  strExc(1) = Left(pub_PMan, 5)
                  'end 2025/2/21
               End If
               'end 2024/7/15
            Else
               'Modified by Morgan 2025/2/19 73022-> left(pub_PMan,5)
               pub_PMan = Pub_GetSpecMan("專利處特定編號")
               strExc(1) = Left(pub_PMan, 5)
               'end 2025/2/21
            End If
            Text16 = strExc(1): ChgType 16
         Else
            Text16 = strUserNum: ChgType 16
         End If
         'end 2016/3/16
         
         'Added by Morgan 2021/3/12
         '寶齡富錦 Y55435 案件下列來函承辦人預設韻如
         '1202審查意見來函、1002核駁、1006最終核駁、1201通知修正、1209檢索報告、1205通知提供前案、1206通知要求選取、1203通知補充說明
         If pa(75) = "Y55435" Then
            'Modified by Morgan 2021/9/22
            'm_bolEngCase = True
            m_bolBPFCase = True
            'end 2021/9/22
            'Modified by Morgan 2023/6/27 預設最後收文的工程師--郭
            'Text16 = "A0029"
            If PUB_GetLastEng(pa(1), pa(2), pa(3), pa(4), strExc(1)) = True Then
               Text16 = strExc(1)
            Else
               Text16 = "A0029"
            End If
            'end 2023/6/27
            ChgType 16
         
         'Added by Morgan 2021/9/22
         '檢查案件是否為顧服組W2001的4家客戶X69365、X82504、X82708、X83239及其關係企業也掛在顧服組W2001者(只考慮第一申請人即可)
         '1. 來函承辦人預設該案最後之工程師(不限收文種類)；
         '2. 來函不上發文日
         ElseIf PUB_ChkIsRltPty(pa(1), m_strCP10, pa(9)) = True Then
            'Modified by Morgan 2021/10/6 長庚醫院案件要收[轉公文]先簡單報告 (111/3/28取消轉公文)
            If PUB_ChkIsW2001XCase(pa(1), pa(2), pa(3), pa(4), strExc(1), m_CustX69365) = True Then
               m_bolW2001XCase = True
               Text16 = strExc(1)
               ChgType 16
            End If
         'end 2021/9/22
         End If
         'end 2021/3/12

         Text18 = "N"
         m_strCP14 = "" & .Fields("CP14")   '2012/11/26 add by sonia
         '2012/12/5 ADD BY SONIA 記錄上個畫面所點選收文號的承辦人(P非台灣案若原承辦人非工程師則改抓國內案承辦人) P-093775
         If pa(9) <> 台灣國家代號 Then
            strExc(0) = "SELECT ST03 FROM STAFF WHERE ST01='" & m_strCP14 & "' AND ST03>='P1' AND ST03<='P11'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               m_strCP14 = PUB_GetInCaseCP14(pa(1), pa(2), pa(3), pa(4))
            End If
         End If
         '2012/12/5 END
      End If
      'end 2009/11/27
      
      '92.5.8 END
   End If
   End With
   
   'Added by Lydia 2023/06/15
   m_bolFMP2 = False
   If m_bolFMP = True Then  '判斷寰華案
      m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4))
   End If
   
   '寰華案:是否為「414恢復權利-主張優先權106」
   bolChk414for106 = False
   strFirstPriDate = ""
   If m_bolFMP2 = True And frm04010503_2.Text6 = "1" And m_strCP10 = "414" And m_strCP43 <> "" Then
      strExc(0) = "select c1.cp09 as cp09_1,c1.cp10 as cp10_1,c2.cp09 as cp09_2,c2.cp10 as cp10_2 from caseprogress c1, caseprogress c2 where c1.cp09='" & m_strCP43 & "' and c1.cp43=c2.cp09(+)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If "" & RsTemp.Fields("cp10_1") = "106" Or "" & RsTemp.Fields("cp10_2") = "106" Then
            bolChk414for106 = True
            strFirstPriDate = PUB_GetFirstPriDate(pa())
         End If
      End If
   End If
   'end 2023/06/15
   
   
   'Add by Morgan 2008/5/13 加判斷台灣案才抓
   If pa(9) = "000" Then
   '6
   strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & pa(1) & "' AND CPM02='" & m_strCP10 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
      If intI = 1 Then
         If Not IsNull(.Fields(1)) Then
            Option4(0).Value = True
            Text10 = .Fields(1)
            Text14(1) = TransDate(CompDate(2, .Fields(1), TransDate(Label3(3).Caption, 2)), 1)
         ElseIf Not IsNull(.Fields(2)) Then
            Option4(1).Value = True
            Text11 = .Fields(2)
            Text14(1) = TransDate(CompDate(1, .Fields(2), TransDate(Label3(3).Caption, 2)), 1)
         Else
            Text10 = ""
            Text11 = ""
            Option4(0).Value = True
         End If
         If Text14(1) <> "" And Not IsNull(.Fields(0)) Then
            If .Fields(0) = "1" Then
               Option1(0).Value = True
               Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
            Else
               Option1(1).Value = True
            End If
         End If
         If Not IsNull(.Fields(1)) Then
            If .Fields(1) = 60 Or .Fields(1) = 90 Then
               i = -4
            Else
               i = -2
            End If
         ElseIf Not IsNull(.Fields(2)) Then
            If .Fields(2) = 2 Then
               i = -4
            Else
               i = -2
            End If
         End If
         'add by sonia 2018/12/27 舉發,舉發答辯經過聽證程序後核發的審定書下一程序請設為"行政訴訟"法定期限為二個月,重為審定的除外P-119106
         m_Have808 = False
         strExc(0) = "SELECT CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='808' AND CP158>0 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modified by Morgan 2022/8/8/24 +3部分准駁
            If (frm04010503_2.Text6 = "1" Or frm04010503_2.Text6 = "3") And (m_strCP10 = "803" Or m_strCP10 = "804") Then
               m_Have808 = True
               Option4(1).Value = False: Option4(1).Value = True
               Text11 = 2: Text10 = ""
               Text14(1) = TransDate(CompDate(1, 2, TransDate(Label3(3).Caption, 2)), 1)
               i = -4
            End If
         End If
         'end 2018/12/27
         If Text14(1) <> "" Then
            'Added by Morgan 2014/10/9
            If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
               Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
            Else
            'end 2014/10/9
               Text14(0) = TransDate(CompDate(2, i, TransDate(Text14(1), 2)), 1)
            End If 'Added by Morgan 2014/10/9
         End If
        'Add By Cheng 2003/12/08
        '本所期限若非工作天則抓最近工作天
        Me.Text14(0).Text = TransDate(PUB_GetWorkDay1(Me.Text14(0).Text, True), 1)
      End If
   End With
   End If
   
   ' 90.10.5 modify by sonia (機關文號設預設內容)
   'Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   If pa(9) = 台灣國家代號 Then
      Select Case m_strCP10
         'Modified by Morgan 2013/7/31--玲玲
         'Case 發明申請, 新型申請, 設計申請, 追加申請, 聯合申請
         '   Text9.Text = "（" & strTmp & "）智專一（五）字第號"
         Case 發明申請
            Text9.Text = "（" & strTmp & "）智專一（五）字第號"
         Case 新型申請
            Text9.Text = "（" & strTmp & "）智專一（四）字第號"
         Case 設計申請, 追加申請, 聯合申請
            Text9.Text = "（" & strTmp & "）智專一（三）字第號"
         'end 2013/7/31
         
         'Modify By Cheng 2002/01/11
'         Case 申請優先權證明, 變更, 讓與
         Case 申請優先權證明, 變更, 讓與, 專利權讓與
            Text9.Text = "（" & strTmp & "）智專一（一）字第號"
         Case 訴願
            Text9.Text = "經訴字第號"
         Case Else
            Text9.Text = "（" & strTmp & "）智專一（二）字第號"
      End Select
        'Add By Cheng 2003/03/26
        '記錄機關文號的預設值
        Me.Text9.Tag = Me.Text9.Text
   End If
   
   'Added by Morgan 2014/1/14
   'Modified by Morgan 2014/4/17 +發文字,期限
   If m_DocWord <> "" Then
      Text9 = m_DocWord & "字第" & m_DocNo & "號"
   ElseIf m_DocNo <> "" Then
      Text9 = Replace(Text9, "第號", "第" & m_DocNo & "號")
   End If
   '期限
   'modify by sonia 2018/12/27 舉發,舉發答辯經過聽證程序後核發的審定書下一程序請設為"行政訴訟"法定期限為二個月,但公文的處理期限仍為30日,P-119106
   'If m_DeadLine <> "" Then
   If m_DeadLine <> "" And m_Have808 = False Then
      Option1(1).Value = True
      If Len(m_DeadLine) >= 7 Then
         Option4(2).Value = True
         Text12 = m_DeadLine
      'Modified by Morgan 2014/8/18 有日的期限
      ElseIf Right(m_DeadLine, 1) = "日" Then
         Option4(0).Value = True
         Text10 = Val(m_DeadLine)
      ElseIf Right(m_DeadLine, 1) = "月" Then
         Option4(1).Value = True
         Text11 = Val(m_DeadLine)
      'end 2014/8/18
      End If
   End If
   'end 2014/1/14

   
   'Add By Cheng 2002/07/23
   EnableTextBox text8, False
   '顯示專用權是否存在
   Me.text8.Text = "" & pa(17)
   
   ' 90.11.4 modify by SONIA (Disable是否更新基本檔准駁)
   EnableTextBox Text7, False
   'Add By Cheng 2002/07/23
   '顯示目前准駁
   Me.Text7.Text = "" & pa(16)
   Select Case m_strCP10
      Case 發明申請, 新型申請, 設計申請, 追加申請, 聯合申請, 答辯
         'Modify By Cheng 2002/07/23
'         Text7.Text = "Y"
      Case 改請發明, 改請新型, 改請設計, 改請追加, 改請聯合, 改請獨立, 分割
'         Text7.Text = "Y"
      Case 異議_專, 舉發
'         Text7.Text = "Y"
   End Select
   
   'Add By Sindy 2012/3/7 +國際分類
   Me.Text22.Text = "" & pa(160)
   '2012/3/7 End
   
   'Added by Morgan 2015/1/20
   '電子公文帶入審查委員及國際分類
   If m_DocNo <> "" Then
      If PUB_GetEDocData(m_DocNo, strExc(1), strExc(2)) Then
         Text19 = strExc(1)
         If Text22 = "" Then Text22 = Left(strExc(2), 4)
      End If
   End If
   'end 2015/1/20
   
   'Added by Morgan 2025/3/7
   bolChgRlt = False
   If pa(23) = "1" And PUB_ChkIsRltPty(pa(1), m_strCP10, pa(9)) = True Then
      bolChgRlt = True
   End If
   'end 2025/3/7
   
    'Modify By Cheng 2003/01/09
'   'Add By Cheng 2002/07/23
'   If (m_strCP10 >= "101" And m_strCP10 <= "105") Or m_strCP10 = "107" Or (m_strCP10 >= "301" And m_strCP10 <= "307") Or m_strCP10 = "802" Or m_strCP10 = "804" Then
   If m_strCP10 <> "802" And m_strCP10 <> "804" Then 'Added by Morgan 2012/3/7 排除 802,804
      'Modified by Morgan 2025/5/12 改抓變數
      'If (Val(m_strCP10) >= 101 And Val(m_strCP10) <= 105) Or Val(m_strCP10) = 107 Or Val(m_strCP10) = 503 Or Val(m_strCP10) = 504 Or _
      '     (Val(m_strCP10) >= 301 And Val(m_strCP10) <= 307) Or (Val(m_strCP10) >= 801 And Val(m_strCP10) <= 805) Then
      If bolChgRlt Then
           Me.Text7.Text = "2"
      End If
   End If
   'Remove by Morgan 2007/7/20 不再預設"N" --郭
   'If m_strCP10 = "804" Then
   '   Me.Text8.Text = "N"
   'End If
   'end 2007/7/20
   
   '承辦期限
   If frm04010503_2.Text6 = "1" Then
      strTmp = 核駁
   'Added by Morgan 2022/8/8/24 +3部分准駁
   ElseIf frm04010503_2.Text6 = "3" Then
      strTmp = 1009
   'end 2022/8/24
   Else
      strTmp = 改變原處分
   End If
   'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
   'Modify by Morgan 2009/12/1 FMP案改在掛法限時計算
   If m_bolFMP Then
      Text17 = ""
   Else
      'Modify by Morgan 2010/10/1
      'Text17 = TransDate(Pub_GetHandleDay(pa(1), pa(9), strTmp, TransDate(Label3(3).Caption, 2)), 1)
      If PUB_IfSetCP48 Then
         Text17 = TransDate(Pub_GetHandleDay(pa(1), pa(9), strTmp, TransDate(Label3(3).Caption, 2)), 1)
      Else
         Text17.Enabled = False
      End If
      'end 2010/10/1
   End If
   'end 2007/10/11
   
   'Add by Amy 2014/09/17 承辦人期限隱藏
   Label27.Visible = False
   Text17.Enabled = False
   Text17.Visible = False
   'end 2014/09/17
   
   '下一程序
   strExc(0) = "SELECT CF15 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & m_strCP10 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         If Not IsNull(.Fields(0)) Then Text13 = .Fields(0): ChgType 13
      End With
   End If
   
   'add by sonia 2018/12/27 舉發,舉發答辯經過聽證程序後核發的審定書下一程序請設為"行政訴訟"法定期限為二個月,重為審定的除外P-119106
   If m_Have808 = True Then
      Text13 = "503": ChgType 13
   End If
   'end 2018/12/27
   
   'Add by Morgan 2007/6/13 檢查65002是否為最後的代理人
   lblDispDate.Visible = False
   txtDispDate.Visible = False
   txtDispDate = ""
   If pa(9) = "000" Then
      If PUB_IsLatestAgent(pa(1), pa(2), pa(3), pa(4)) = True Then
         lblDispDate.Visible = True
         txtDispDate.Visible = True
         txtDispDate.MaxLength = 7
      End If
   End If
   'end 2007/6/13
   
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 13
         If pa(9) < "010" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetCaseProperty(pA(1), Text13.Text, strTempName, False) Then
            If ClsPDGetCaseProperty(pa(1), Text13.Text, strTempName, False) Then
               Label3(5) = strTempName
               ChgType = True
            Else
               Label3(5) = ""
            End If
         Else
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetCaseProperty(pA(1), Text13.Text, strTempName, True) Then
            If ClsPDGetCaseProperty(pa(1), Text13.Text, strTempName, True) Then
               Label3(5) = strTempName
               ChgType = True
            Else
               Label3(5) = ""
            End If
         End If
      Case 16
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(Text16.Text, strTempName) Then
         If ClsPDGetStaff(Text16.Text, strTempName) Then
            Label3(4) = strTempName
            ChgType = True
         Else
            Label3(4) = ""
         End If
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2018/11/13
   'Set frm04010503_3 = Nothing 'Removed by Morgan 2021/12/20 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub Combo2_Click()
   Select Case Combo2
      Case "中"
         Label3(0) = pa(5)
      Case "英"
         Label3(0) = pa(6)
      Case "日"
         Label3(0) = pa(7)
   End Select
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
End Sub

Private Sub Text10_LostFocus()
   'Add By Cheng 2003/04/01
   If Me.Text10.Text <> "" Then GetTime
   'Add by Morgan 2008/5/23 非台灣"天"跳離時到"下一程序"欄位
   If pa(9) <> 台灣國家代號 Then
      If Text13.Enabled = True Then Text13.SetFocus
   End If
End Sub

Private Sub Text11_GotFocus()
  TextInverse Text11
End Sub

Private Sub Text11_LostFocus()
   'Add By Cheng 2003/04/01
   If Me.Text11.Text <> "" Then GetTime
   'Add by Morgan 2008/5/23 非台灣"月"跳離時到"下一程序"欄位
   If pa(9) <> 台灣國家代號 Then
      If Text13.Enabled = True Then Text13.SetFocus
   End If
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_LostFocus()
   'Add by Morgan 2008/5/23 非台灣"日"跳離時到"下一程序"欄位
   If pa(9) <> 台灣國家代號 Then
      If Text13.Enabled = True Then Text13.SetFocus
   End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
      MsgBox "來函期限不可空白 !", vbCritical
      Cancel = True
   Else
      If ChkDate(Text12) Then
         'Add by Morgan 2008/5/23
         If pa(9) <> 台灣國家代號 Then
            If Val(Text12) < Val(strSrvDate(1)) Then
               MsgBox "來函期限不可小於系統日 !", vbCritical
               Cancel = True
            Else
               '轉民國年
               Text14(1) = TransDate(Text12, 1)
               'Added by Lydia 2025/10/29
               If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                  Text14(0) = TransDate(PUB_GetPOurDeadline(TransDate(Text14(1), 2), pa(9)), 1)
               Else
               'end 2025/10/29
                  'Add by Morgan 2009/12/23 FMP案所限=法限-7天
                  If m_bolFMP Then
                     Text14(0) = TransDate(CompDate(2, -7, TransDate(Text14(1), 2)), 1)
                  Else
                  'end 2009/12/23
                     '大陸案的所限=法限-10天
                     Text14(0) = TransDate(CompDate(2, -10, TransDate(Text14(1), 2)), 1)
                  End If
               End If 'Added by Lydia 2025/10/29
               Text14(0) = TransDate(PUB_GetWorkDay1(Text14(0), True), 1)
            End If
         Else
         'end 2008/5/23
            If Val(Text12) < Val(strSrvDate(2)) Then
               MsgBox "來函期限不可小於系統日 !", vbCritical
               Cancel = True
            Else
               Text14(1) = Text12
               'Added by Morgan 2014/10/9
               If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                  Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
               Else
               'end 2014/10/9
                  Text14(0) = TransDate(CompDate(2, -2, TransDate(Text14(1), 2)), 1)
                  'Add By Cheng 2003/12/08
                  '本所期限若非工作天則抓最近工作天
                  Me.Text14(0).Text = TransDate(PUB_GetWorkDay1(Me.Text14(0).Text, True), 1)
               End If 'Added by Morgan 2014/10/9
            End If
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub Text13_Change()
   'Modified by Morgan 2012/7/3 排除FMP案--敏惠
   If pa(9) = "020" And Text13 = "107" And m_bolFMP = False Then
      txtFee.Visible = True
      lblFee.Visible = True
      txtPt.Visible = True
      lblPt.Visible = True
   Else
      txtFee.Text = ""
      txtFee.Visible = False
      lblFee.Visible = False
      txtPt.Visible = False
      lblPt.Visible = False
   End If
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text13
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
    'Modify By Cheng 2002/11/29
    '若來函性質為行政再審則可不輸入下一程序
    'Modify By Cheng 2003/09/02
    '若來函性質為行政訴訟上訴也可不輸入下一程序
'    If m_strCP10 <> 行政再審 Then
    'Modify by Morgan 2011/3/24 +行政上訴答辯 508
    'Modified by Morgan 2020/2/24 +PPH 431--Winfrey,玲玲
    If m_strCP10 <> 行政再審 And m_strCP10 <> "507" And m_strCP10 <> "508" And Not (pa(9) = "020" And m_strCP10 = "431") Then
        If Text13 = "" Then
           MsgBox "下一程序不可空白 !", vbCritical
           Cancel = True
        Else
           'Add By Cheng 2002/01/04
           If Len(Me.Text13.Text) <> 3 Then
              MsgBox "下一程序欄位值必須為三碼 !", vbCritical
              Text13_GotFocus
              Cancel = True
              Exit Sub
           End If
           If ChgType(13) = False Then Cancel = True
        End If
    End If
End Sub

Private Sub Text14_GotFocus(Index As Integer)
  TextInverse Text14(Index)
End Sub

Private Sub Text14_LostFocus(Index As Integer)
   'Add By Cheng 2002/12/19
   '若為大陸案核駁, 則法定期限 = 本所期限
   If Index = 0 Then
      'Remove by Morgan 2008/5/23 不必在預設相同 --玲玲
      'If pa(9) <> 台灣國家代號 And frm04010503_2.Text6.Text = "1" Then Me.Text14(1).Text = Me.Text14(0).Text
   End If
End Sub

Private Sub Text14_Validate(Index As Integer, Cancel As Boolean)
 Static iTime As Integer
 Static strStatic(0 To 1) As String
   If Text14(Index) <> "" Then
      If Not ChkDate(Text14(Index)) Then
         Cancel = True
      Else
         If Index = 1 Then
            If Not ChkRange(Text14(0), Text14(1), "本所期限、法定期限") Then
               Cancel = True
            Else
               iTime = iTime + 1
               If iTime = 1 Then
                  strStatic(0) = Text14(0)
                  strStatic(1) = Text14(1)
               End If
               
               'Modified by Morgan 2014/5/2 改與其他程式一致，先檢查櫃台有沒有收文,再檢查期限是否相同
               'If Text14(0) <> strStatic(0) Or Text14(1) <> strStatic(1) Then
               
                  ' 90.07.10 midify by louis (申請國家非台灣, 不需檢查來函記錄檔)
                  If pa(9) < "010" Then
                     'edit by nickc 2007/02/05 不用 dll 了
                     'If objLawDll.ChkMRec(TransDate(Label3(3).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
                     If ClsLawChkMRec(TransDate(Label3(3).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
                        'If Text14(0) <> strStatic(0) Or Text14(1) <> strStatic(1) Then 'Removed by Morgan 2014/8/18
                           If Text14(0) <> TransDate(strExc(1), 1) Then
                              'Modified by Morgan 2014/8/18 改和一般來函一樣控制
                              'If MsgBox("與櫃台之來函收文記錄本所期限 ( " & TransDate(strExc(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then
                              
                              '   frm04010503_1.Show
                              '   Unload frm04010503_2
                              '   Unload Me
                              'Else
                              '   Text14(0) = ""
                              '   Text14(1) = ""
                              If MsgBox("與櫃台之來函收文記錄本所期限 ( " & TransDate(strExc(1), 1) & ") 不符，請確認 !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                                 Cancel = True
                                 Exit Sub
                              'end 2014/8/18
                              End If
                           ElseIf Text14(1) <> TransDate(strExc(2), 1) Then
                              'Modified by Morgan 2014/8/18 改和一般來函一樣控制
                              'If MsgBox("與櫃台之來函收文記錄法定期限 ( " & TransDate(strExc(2), 1) & ") 不符，請確認 !", vbCritical + vbYesNo) = vbNo Then
                              '   frm04010503_1.Show
                              '   Unload frm04010503_2
                              '   Unload Me
                              'Else
                              '   Text14(0) = ""
                              '   Text14(1) = ""
                              If MsgBox("與櫃台之來函收文記錄法定期限 ( " & TransDate(strExc(2), 1) & ") 不符，請確認 !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                                 Cancel = True
                                 Exit Sub
                              'end 2014/8/18
                              End If
                           End If
                        'End If 'Removed by Morgan 2014/8/18
                     'Modified by Morgan 2014/5/5 排除無期限電子公文
                     'Else
                     'modify by sonia 2018/2/8 電子公文都不檢查來函記錄檔
                     'ElseIf m_DocNo = "" Or Text14(1) <> "" Then
                     ElseIf m_DocNo = "" Then
                     'end 2014/5/5
                        If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbDefaultButton2 + vbYesNo) = vbNo Then
                           Cancel = True
                           Exit Sub
                        End If
                     
                     End If
                  End If
               'End If
               'end 2014/5/2
            End If
         End If
      End If
      
      If Index = 0 Then
         Text14(Index).Text = TransDate(PUB_GetWorkDay1(Me.Text14(Index).Text, True), 1)
         '2008/11/3 modify by sonia 因P-83487與郭確認再改回來
         'Modify by Morgan 2008/5/23
         'Modify by Morgan 2010/11/9
         'If Me.Text14(Index).Text < strSrvDate(1) Then
         If Val(Text14(Index).Text) < Val(strSrvDate(2)) Then
            MsgBox "本所期限不可小於系統日!!!", vbExclamation
         'If Val(TransDate(Text14(Index), 2)) <= Val(strSrvDate(1)) Then
         '   MsgBox "本所期限必須大於系統日!!!", vbExclamation
            Cancel = True
         End If
      End If
   End If
   If Cancel = True Then TextInverse Text14(Index)
End Sub

Private Sub Text16_GotFocus()
  TextInverse Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii) 'Added by Morgan 2024/4/30
End Sub

Private Sub Text17_GotFocus()
  TextInverse Text17
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
   If Text17 <> "" Then
      If ChkWorkDay(TransDate(Text17, 2)) Then
         ' 90.07.10 modify by louis
         If IsEmptyText(Text14(0)) = False Then
            'Modify by Morgan 2010/8/11 百年蟲
            'If Text17 > Text14(0) Then
            If Val(Text17) > Val(Text14(0)) Then
               MsgBox "承辦期限不可大於本所期限，請重新輸入 !", vbCritical
               Cancel = True
            End If
         End If
      Else
         MsgBox "承辦期限不正確，請重新輸入 !", vbCritical
         Cancel = True
      End If
   Else
      If Text13 <> "" Then
         'MODIFY BY SONIA 901031
         'MsgBox "有下一程序且有定義工作天數時不可空白 !", vbCritical
         'Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text17
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   If Text16 <> "" Then
      If ChgType(16) = False Then
         Cancel = True
         TextInverse Text16
      End If
   Else
      Label3(4) = ""
   End If
End Sub

Private Sub Text19_GotFocus()
  TextInverse Text19
End Sub

Private Sub Text19_Validate(Cancel As Boolean)
   If Text19 = "" And (m_strCP10 = 發明申請 Or m_strCP10 = 新型申請 Or m_strCP10 = 設計申請 Or m_strCP10 = 追加申請 Or m_strCP10 = 聯合申請 Or m_strCP10 = 答辯 Or m_strCP10 = 異議_專 Or m_strCP10 = 異議答辯 Or m_strCP10 = 舉發 Or m_strCP10 = 舉發答辯) Then
        'Modify By Cheng 2002/12/18
        '若為台灣案審查委員不可空白
        If pa(9) = 台灣國家代號 Then
            'Add by Morgan 2004/6/29
            '93.7.1以後新型核駁無審查委員
            If m_strCP10 = 新型申請 And Val(Text6) >= 930701 Then Exit Sub
            
            MsgBox "審查委員不可空白 !", vbCritical
            Cancel = True
        End If
   Else
      If CheckLengthIsOK(Text19, 32) = False Then
         Cancel = True
      End If
   End If
End Sub

Private Sub Text21_GotFocus()
   TextInverse Text21
End Sub

Private Sub Text21_LostFocus()
   'Add by Morgan 2008/5/23 非台灣"審查委員編號"跳離時到"月"欄位
   If pa(9) <> 台灣國家代號 Then
      If Option4(0).Value = True Then
         If Text10.Enabled = True Then Text10.SetFocus
      ElseIf Option4(1).Value = True Then
         If Text11.Enabled = True Then Text11.SetFocus
      ElseIf Option4(2).Value = True Then
         If Text12.Enabled = True Then Text12.SetFocus
      End If
   End If
End Sub

Private Sub Text21_Validate(Cancel As Boolean)
   If Not CheckLengthIsOK(Text21, Text21.MaxLength) Then
      Cancel = True
   End If
End Sub

'Add By Sindy 2012/3/7
Private Sub Text22_GotFocus()
   InverseTextBox Text22
End Sub

'Add By Sindy 2012/3/7
Private Sub Text22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = "" Then
      MsgBox "核駁函日期不可空白 !", vbCritical
      Cancel = True
   Else
      'Add by Morgan 2008/5/21
      If pa(9) <> 台灣國家代號 Then
         If Len(Text6) <> 8 Then
            MsgBox "非台灣案時核駁函日期請輸西元格式！"
            Cancel = True
         ElseIf ChkDate(Text6) Then
            If Val(Text6) > Val(strSrvDate(1)) Then
               MsgBox "核駁函日期不可大於系統日 !", vbCritical
               Cancel = True
            End If
         Else
            Cancel = True
         End If
      'end 2008/5/21
      ElseIf ChkDate(Text6) Then
         If Val(Text6) > Val(strSrvDate(2)) Then
            MsgBox "核駁函日期不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text6
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   'Modify By Cheng 2002/07/23
'   If KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then
'      KeyAscii = 0
'      Beep
'   End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   'Modify By Cheng 2002/07/23
'   If KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then
'      KeyAscii = 0
'      Beep
'   End If
End Sub

Private Sub Text9_GotFocus()
'  TextInverse Text9
Dim intPos As Integer
'Modify By Cheng 2002/04/22
'將游標設定在機關文號欄的"專"的後面
With Me.Text9
   If Len("" & .Text) > 0 Then
      intPos = InStr("" & .Text, "專")
      If intPos > 0 Then
         .SelStart = intPos
         .SelLength = 0
      End If
   End If
End With
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 = "" Then
        'Modify By Cheng 2002/12/18
        '若為台灣案機關文號不可空白
        If pa(9) = 台灣國家代號 Then
            MsgBox "機關文號不可空白 !", vbCritical
            Cancel = True
        End If
   Else
      'Modify by Morgan 2011/1/3 機關文號欄位改長度(百年問題)改抓MaxLength屬性控制
      If CheckLengthIsOK(Text9, Text9.MaxLength) = False Then
         Cancel = True
      End If
   End If
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim arrCaseNo() As String 'Added by Morgan 2021/2/25

TxtValidate = False
'Remove by Morgan 2008/5/23 存檔時不必再檢查否則期限會重算
'If Me.Text12.Enabled = True Then
'   Cancel = False
'   Text12_Validate Cancel
'   If Cancel = True Then
'      Me.Text12.SetFocus
'      Text12_GotFocus
'      Exit Function
'   End If
'End If

   'Added by Morgan 2021/12/20 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/20

If Me.Text13.Enabled = True Then
   Cancel = False
   Text13_Validate Cancel
   If Cancel = True Then
      Me.Text13.SetFocus
      Text13_GotFocus
      Exit Function
   End If
End If

For Each objTxt In Text14
   If objTxt.Enabled = True Then
      Cancel = False
      Text14_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.Text14(objTxt.Index).SetFocus
         Text14_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next

If Me.Text16.Enabled = True Then
   Cancel = False
   Text16_Validate Cancel
   If Cancel = True Then
      Me.Text16.SetFocus
      Text16_GotFocus
      Exit Function
   End If
End If

If Me.Text17.Enabled = True Then
   Cancel = False
   Text17_Validate Cancel
   If Cancel = True Then
      Me.Text17.SetFocus
      Text17_GotFocus
      Exit Function
   End If
End If

If Me.Text6.Enabled = True Then
   Cancel = False
   Text6_Validate Cancel
   If Cancel = True Then
      Me.Text6.SetFocus
      Text6_GotFocus
      Exit Function
   End If
End If

If Me.Text9.Enabled = True Then
   Cancel = False
   Text9_Validate Cancel
   If Cancel = True Then
      Me.Text9.SetFocus
      Text9_GotFocus
      Exit Function
   End If
End If

If Me.Text19.Enabled = True Then
   Cancel = False
   Text19_Validate Cancel
   If Cancel = True Then
      Me.Text19.SetFocus
      Text19_GotFocus
      Exit Function
   End If
End If

   'Added by Morgan 2019/5/24 從寫定稿例外欄位移來此處先作檢查
   'IDS報價檢查
   m_USCaseNo = ""
   'Modified by Morgan 2023/12/26 +803舉發、804舉發答辯--郭
   If (pa(9) = "000" And (pa(8) = "1" Or pa(8) = "3") And InStr("101,103,301,303,107,307,803,804", m_strCP10) > 0) Then
      m_USCaseNo = PUB_GetUSCaseNo(pa(1), pa(2), pa(3), pa(4))
      'Added by Morgan 2020/3/5
      If m_USCaseNo <> "" Then
         strExc(0) = "1.請確認引證前案是否與前一次審查意見通知書相同。" & vbCrLf & _
            "2.請確認美國案 " & m_USCaseNo & " 是否已提出相同引證前案的IDS。" & vbCrLf & _
            "若二者均相同可不必通知IDS報價"
         strExc(0) = strExc(0) & vbCrLf & vbCrLf & "【是】:要通知    【否】:不通知    【取消】:回畫面" 'Added by Morgan 2020/12/18
         intI = MsgBox(strExc(0), vbYesNoCancel + vbInformation + vbDefaultButton3, "是否通知IDS報價？")
         If intI = vbCancel Then
            Exit Function
         ElseIf intI = vbNo Then
            m_USCaseNo = ""
         End If
      End If
      'end 2020/3/5
      If m_USCaseNo <> "" Then
         If txtIDSFee(1) = "" Or txtIDSFee(2) = "" Or txtIDSPt(1) = "" Or txtIDSPt(2) = "" Then
            If MsgBox("尚未輸入ＩＤＳ報價，是否 EMail 通知 CFP 程序人員報價？", vbYesNo + vbDefaultButton2 + vbExclamation, "ＩＤＳ報價") = vbYes Then
               strExc(0) = "核駁函"
               'Added by Morgan 2023/5/24
               strExc(5) = ""
               Do
                  strExc(5) = InputBox("請輸入引證前案檔案數量：")
                  If Val(strExc(5)) > 0 Then
                     Exit Do
                  ElseIf strExc(5) = "" Then
                     MsgBox "未輸入引證前案檔案數量，取消 EMail 通知！", vbExclamation
                     Exit Function
                  Else
                     MsgBox "引證前案檔案數量必須大於 0，請重新輸入！", vbExclamation
                  End If
               Loop
               'end 2023/5/24
                  
               strExc(2) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
               'Modified by Morgan 2021/2/25 考慮會有多個美國案
               'strExc(1) = PUB_GetCFPHandler(m_USCaseNo)
               'strExc(4) = strExc(2) & " 案已收到" & strExc(0) & "，請提供相關美國案( " & m_USCaseNo & " )的IDS報價！"
               arrCaseNo = Split(m_USCaseNo, "、")
               For ii = LBound(arrCaseNo) To UBound(arrCaseNo)
                  strExc(4) = strExc(2) & " 案已收到" & strExc(0) & "，請提供相關美國案( " & arrCaseNo(ii) & " )的IDS報價！"
                  strExc(1) = PUB_GetCFPHandler(arrCaseNo(ii))
               'end 2021/2/25
                  If strExc(1) <> "" Then
                     'Modified by Morgan 2019/9/9 調整報價欄位名及定稿內容--郭
                     'Modified by Morgan 2023/5/24 +引證前案檔案數量
                     strExc(3) = "引證前案共: " & strExc(5) & " 件" & vbCrLf & _
                                 "IDS報價:" & vbCrLf & _
                                 "　1.第一階段　　　(　P)" & vbCrLf & _
                                 "　2.第二階段　　　(　P)" & vbCrLf & vbCrLf & _
                                 "**　若該案已是第二階段，則第一階段請輸　0　**"
                     PUB_SendMail strUserNum, strExc(1), "", strExc(4), strExc(3)
                  End If
               Next 'Added by Morgan 2021/2/25
               
            ElseIf txtIDSFee(1) = "" Then
               txtIDSFee(1).SetFocus
            ElseIf txtIDSPt(1) = "" Then
               txtIDSPt(1).SetFocus
            ElseIf txtIDSFee(2) = "" Then
               txtIDSFee(2).SetFocus
            ElseIf txtIDSPt(2) = "" Then
               txtIDSPt(2).SetFocus
            End If
            Exit Function
         End If
      End If
   End If
   'end 2019/5/24

'Added by Morgan 2014/5/15 電子化-檢查pdf檔
If pa(9) = "000" Then
   If PUB_CheckPDF(pa(1), pa(2), pa(3), pa(4), 1, m_DocNo) = False Then
      Exit Function
   End If
End If
'end 2014/5/15

   'Added by Morgan 2020/1/17
   '大陸案,有通知函,程序承辦,非掛號(無期限)
   m_bolNoCP27 = False
   'Removed by Morgan 2024/1/30 取消--郭
   'If pa(9) = "020" And Text15 <> "N" And PUB_GetST03(Text16) = "P12" And Text14(1) = "" Then
   '   If PUB_GetCustomerValue(pa(26), "CU182") = "Y" Then
   '      If MsgBox("請確認是否已收到公文正本？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
   '         m_bolNoCP27 = True
   '      End If
   '   End If
   'End If
   'end 2020/1/17
   
   'Added by Morgan 2020/8/12
   '若為來函期限2次確認退回時需檢查法限是否一致
   If m_strIR01 <> "" Then
      If PUB_ChkReKeyInOk(m_strIR01, m_strIR02, m_strIR03, m_strIR04, Text14(1).Text, m_bolReKeyInOK) = False Then
         Text14(1).SetFocus
         Exit Function
      End If
   End If
   'end 2020/8/12
   
   'Added by Morgan 2024/4/30
   m_bolEngCase = False
   'Modified by Morgan 2024/5/7
   'If GetStaffDepartment(Text16) <> "P12" Then
   If Not m_bolFMP And Text16.Enabled And GetStaffDepartment(Text16) <> "P12" Then
      m_bolEngCase = True
      Text15.Text = "N"
   End If
   'end 2024/4/30
   
   'Added by Morgan 2025/3/7 面詢未辦理，向官方辦理退費控管--玲玲
   m_bolAddB908 = False
   If bolChgRlt And pa(9) = "000" Then
      strExc(0) = "select * from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='407' and cp27>19221111" & _
         " and not exists(select * from caseprogress b where cp01=a.cp01 and cp02=a.cp02 and cp03=a.cp03 and cp04=a.cp04 and cp10='408' and cp27>a.cp27)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         intI = MsgBox("本案曾申請【請求面詢】但未辦理，請問是否自動收文【代辦退費】？", vbQuestion + vbYesNoCancel + vbDefaultButton3)
         If intI = vbYes Then
            m_bolAddB908 = True
         ElseIf intI = vbCancel Then
            Exit Function
         End If
      End If
   End If
   'end 2025/3/7
   
   TxtValidate = True
End Function

Private Sub GetTime()
   Dim i As Integer
   'Add by Morgan 2007/6/13
   Dim strFromDate As String '期限起算日
   If txtDispDate.Visible = True Then
      strFromDate = DBDATE(txtDispDate)
   Else
      'Modify by Morgan 2008/5/22 大陸案用核駁日期計算
      If pa(9) <> 台灣國家代號 Then
         strFromDate = DBDATE(Text6)
      Else
         strFromDate = DBDATE(Label3(3))
      End If
   End If
   
    '文到天數
    If Option4(0).Value = True Then
      If Text10 <> "" Then    '2007/11/21 ADD BY SONIA
        Text14(1) = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
        If Option1(0).Value = True Then Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
        If Text10 = "60" Or Text10 = "90" Then
            i = -4
        Else
            i = -2
        End If
      End If
    '文到月數
    ElseIf Option4(1).Value = True Then
      If Text11 <> "" Then    '2007/11/21 ADD BY SONIA
        Text14(1) = TAIWANDATE(AddMonth(strFromDate, Val(Text11)))
        If Option1(0).Value = True Then Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
        If Text11 = "2" Then
            i = -4
        Else
            i = -2
        End If
      End If
    End If
    
   'Modify by Morgan 2008/5/12 大陸案的所限=法限-10天
   'Modify by Morgan 2009/12/1 FMP案所限=法限-7天
   If m_bolFMP Then
      i = -7
   ElseIf pa(9) <> "000" Then
      i = -10
   End If
   'end 2008/5/12
   
   If Text14(1) <> "" Then
      'Added by Lydia 2025/10/29
      If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
         Text14(0) = TransDate(PUB_GetPOurDeadline(Text14(1), pa(9)), 1)
      Else
      'end 2025/10/29
         'Added by Morgan 2014/10/9
         If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
            Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
         Else
         'end 2014/10/9
            Text14(0) = TransDate(CompDate(2, i, TransDate(Text14(1), 2)), 1)
         End If 'Added by Morgan 2014/10/9
      End If 'Added by Lydia 2025/10/29
   End If
    'Add By Cheng 2003/12/08
    '本所期限若非工作天則抓最近工作天
    Me.Text14(0).Text = TransDate(PUB_GetWorkDay1(Me.Text14(0).Text, True), 1)
    
       'Add by Morgan 2009/12/1
      '承辦期限,承辦期限狀況,約定期限
      If m_bolFMP And Val(Text11) > 0 Then
         strExc(1) = PUB_GetFmpCP48(DBDATE(Label3(3)), DBDATE(Text14(0)), DBDATE(Text14(1)), strFromDate, Val(Text11), stNP23, stCP48Desc)
         Text17 = TransDate(strExc(1), 1)
         'Added by Lydia 2015/06/29 外專寰華的案件，在輸入各式審查機關來函的畫面，能帶出約定期限欄位
          txtSNP23.Text = TransDate(stNP23, 1)
          
      End If
      
'Modified by Morgan 2016/9/21 改呼叫共用函數
'      'Added by Morgan 2015/4/20
'      '先正達OA承辦期限設7個工作天,若下一程序為 804,501-509時設2個工作天(24Hr)
'      If InStr("Y4830900,Y4830901,Y4830902,Y4830903,Y4830904,Y4830905,Y4830908,Y5132600", Left(pa(75) & "000", 8)) > 0 Then
'         If Text13 = "804" Or Text13 >= "501" And Text13 <= "509" Then
'            Text17 = TransDate(CompWorkDay(2, TransDate(Label3(3).Caption, 2), 0), 1)
'         Else
'            Text17 = TransDate(CompWorkDay(7, TransDate(Label3(3).Caption, 2), 0), 1)
'         End If
'
'      'Added by Morgan 2015/7/3 --吳彩菱
'      'Y51753+X45149010 承辦天數:14 起算日期:系統日
'      ElseIf Left(pa(75) & "000", 8) = "Y5175300" And Left(pa(26) & "000", 8) = "X4514901" Then
'         Text17 = TransDate(CompDate(2, 14, strSrvDate(1)), 1)
'
'      End If
'      'end 2015/4/20
   If m_bolFMP Then
      'Modified by Morgan 2021/3/23 +DBDATE(Text6)
      Call Pub_SetExceptCP48(pa(75), pa(26), IIf(frm04010503_2.Text6 = "1", 核駁, 改變原處分), TransDate(Label3(3).Caption, 2), Text17, Text13.Text, , , DBDATE(Text6))
   End If
'end 2016/9/21

End Sub

Private Sub txtDispDate_GotFocus()
   TextInverse txtDispDate
End Sub

Private Sub txtIDSFee_GotFocus(Index As Integer)
   TextInverse txtIDSFee(Index)
End Sub

Private Sub txtIDSFee_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtIDSPt_GotFocus(Index As Integer)
   TextInverse txtIDSPt(Index)
End Sub

Private Sub txtIDSPt_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub
