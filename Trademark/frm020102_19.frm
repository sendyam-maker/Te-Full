VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_19 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(條碼)"
   ClientHeight    =   5508
   ClientLeft      =   5280
   ClientTop       =   648
   ClientWidth     =   9132
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5508
   ScaleWidth      =   9132
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   4530
      MaxLength       =   4
      TabIndex        =   9
      Top             =   4080
      Width           =   540
   End
   Begin VB.TextBox textPeriod 
      Height          =   264
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   4
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox textCP84 
      Alignment       =   1  '靠右對齊
      Height          =   264
      Left            =   3330
      TabIndex        =   1
      Top             =   2880
      Width           =   1425
   End
   Begin VB.TextBox textTM81 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   2250
      Width           =   2532
   End
   Begin VB.TextBox textTM80 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   5430
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   1950
      Width           =   3645
   End
   Begin VB.TextBox textTM79 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   1950
      Width           =   2532
   End
   Begin VB.TextBox textTM78 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   5430
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   1650
      Width           =   3645
   End
   Begin VB.TextBox textWord 
      Height          =   264
      Left            =   1800
      MaxLength       =   1
      TabIndex        =   8
      Top             =   4080
      Width           =   372
   End
   Begin VB.TextBox textFee 
      Height          =   264
      Left            =   5760
      TabIndex        =   5
      Top             =   3480
      Width           =   1092
   End
   Begin VB.TextBox textSP22 
      Height          =   264
      Left            =   1200
      MaxLength       =   2000
      TabIndex        =   3
      Top             =   3180
      Width           =   7752
   End
   Begin VB.TextBox textCF09 
      Height          =   264
      Left            =   5400
      MaxLength       =   12
      TabIndex        =   7
      Top             =   3780
      Width           =   828
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   4968
      TabIndex        =   12
      Top             =   36
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8244
      TabIndex        =   15
      Top             =   36
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6192
      TabIndex        =   13
      Top             =   36
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7020
      TabIndex        =   14
      Top             =   36
      Width           =   1200
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1200
      MaxLength       =   1
      TabIndex        =   6
      Top             =   3780
      Width           =   372
   End
   Begin VB.TextBox textTM21 
      BorderStyle     =   0  '沒有框線
      Height          =   240
      Left            =   5460
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2262
      Width           =   1092
   End
   Begin VB.TextBox textTM22 
      BorderStyle     =   0  '沒有框線
      Height          =   240
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2262
      Width           =   1092
   End
   Begin VB.TextBox textCP18 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   7710
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2910
      Width           =   1245
   End
   Begin VB.TextBox textCP27 
      Height          =   264
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2880
      Width           =   1092
   End
   Begin VB.TextBox textCP22 
      Height          =   264
      Left            =   5760
      MaxLength       =   1
      TabIndex        =   2
      Top             =   2880
      Width           =   372
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1050
      Width           =   2532
   End
   Begin VB.TextBox textCP13 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5430
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1050
      Width           =   3645
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5430
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   750
      Width           =   3645
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1230
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   450
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   750
      Width           =   2532
   End
   Begin VB.TextBox textTM44 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5430
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   450
      Width           =   3645
   End
   Begin VB.TextBox textCP14 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5430
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1350
      Width           =   3645
   End
   Begin VB.TextBox textTM23 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1650
      Width           =   2532
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1200
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   2550
      Width           =   7812
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13779;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP18 
      Height          =   492
      Left            =   1260
      TabIndex        =   11
      Top             =   4920
      Width           =   7692
      VariousPropertyBits=   -1467989989
      MaxLength       =   60
      ScrollBars      =   2
      Size            =   "13568;868"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   492
      Left            =   1260
      TabIndex        =   10
      Top             =   4380
      Width           =   7692
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13568;868"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   18
      Left            =   3720
      TabIndex        =   64
      Top             =   4080
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "期登記期"
      Height          =   180
      Left            =   1920
      TabIndex        =   63
      Top             =   3525
      Width           =   720
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "發文規費："
      Height          =   180
      Left            =   2400
      TabIndex        =   62
      Top             =   2940
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2 :"
      Height          =   180
      Index           =   17
      Left            =   4440
      TabIndex        =   57
      Top             =   1692
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3 :"
      Height          =   180
      Index           =   16
      Left            =   150
      TabIndex        =   56
      Top             =   1992
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4 :"
      Height          =   180
      Index           =   15
      Left            =   4440
      TabIndex        =   55
      Top             =   1992
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5 :"
      Height          =   180
      Index           =   14
      Left            =   150
      TabIndex        =   54
      Top             =   2292
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "(Y:修改)"
      Height          =   180
      Left            =   2280
      TabIndex        =   53
      Top             =   4125
      Width           =   645
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "是否修改定稿內容 :"
      Height          =   180
      Left            =   150
      TabIndex        =   52
      Top             =   4125
      Width           =   1530
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "下次繳年費期限 :"
      Height          =   180
      Left            =   4440
      TabIndex        =   51
      Top             =   3525
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "正片號碼 :"
      Height          =   180
      Left            =   150
      TabIndex        =   50
      Top             =   3222
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "大約                                 可接獲回音"
      Height          =   180
      Index           =   12
      Left            =   4440
      TabIndex        =   49
      Top             =   3822
      Width           =   2850
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "案件備註 :"
      Height          =   180
      Left            =   150
      TabIndex        =   48
      Top             =   4920
      Width           =   810
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "進度備註 :"
      Height          =   180
      Left            =   150
      TabIndex        =   47
      Top             =   4380
      Width           =   810
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
      Height          =   180
      Left            =   1620
      TabIndex        =   46
      Top             =   3822
      Width           =   2745
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "列印定稿 :"
      Height          =   180
      Left            =   150
      TabIndex        =   45
      Top             =   3822
      Width           =   810
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "原使用期間 :"
      Height          =   180
      Left            =   4440
      TabIndex        =   44
      Top             =   2292
      Width           =   990
   End
   Begin VB.Line Line2 
      X1              =   6600
      X2              =   6720
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "第"
      Height          =   180
      Left            =   675
      TabIndex        =   41
      Top             =   3525
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點數 :"
      Height          =   180
      Index           =   10
      Left            =   7230
      TabIndex        =   40
      Top             =   2910
      Width           =   450
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "發文日 :"
      Height          =   180
      Left            =   150
      TabIndex        =   39
      Top             =   2922
      Width           =   630
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "(N:不出名)"
      Height          =   180
      Left            =   6180
      TabIndex        =   38
      Top             =   2925
      Width           =   825
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "是否出名 :"
      Height          =   180
      Left            =   4920
      TabIndex        =   37
      Top             =   2925
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱 :"
      Height          =   180
      Left            =   150
      TabIndex        =   35
      Top             =   2610
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "審定號數/申請案號 :"
      Height          =   180
      Left            =   150
      TabIndex        =   34
      Top             =   1092
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員 :"
      Height          =   180
      Index           =   11
      Left            =   4440
      TabIndex        =   33
      Top             =   1092
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號 :"
      Height          =   180
      Index           =   9
      Left            =   4440
      TabIndex        =   32
      Top             =   792
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質 :"
      Height          =   180
      Index           =   6
      Left            =   150
      TabIndex        =   31
      Top             =   1392
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號 :"
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   30
      Top             =   492
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   29
      Top             =   792
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人 :"
      Height          =   180
      Index           =   2
      Left            =   4440
      TabIndex        =   28
      Top             =   495
      Width           =   840
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "承辦人 :"
      Height          =   180
      Left            =   4440
      TabIndex        =   27
      Top             =   1392
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人1 :"
      Height          =   180
      Left            =   150
      TabIndex        =   26
      Top             =   1692
      Width           =   720
   End
End
Attribute VB_Name = "frm020102_19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/27 Form2.0已修改 textTm44/textCP13/textCP14/textCP44_2/textTM23(申請人名).../cmbTM05/textCP64/textSP18
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
' 智權人員
Dim m_CP13 As String
' 承辦人
Dim m_CP14 As String
'
' 91.09.02 modify by louis
' 加廠商號碼比對
Dim m_SP19 As String
Dim m_SP24 As String
Dim m_SP25 As String

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
Dim m_CP44 As String
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

Private Sub cmdCCCCode_Click()
   frmCCCCode.SetData 0, m_TM01, True
   frmCCCCode.SetData 1, m_TM02, False
   frmCCCCode.SetData 2, m_TM03, False
   frmCCCCode.SetData 3, m_TM04, False
   frmCCCCode.SetData 4, m_SP24, False
   frmCCCCode.SetData 5, m_SP25, False
   frmCCCCode.Show vbModal, Me
   If frmCCCCode.GetData(0) = "1" Then
      m_SP24 = frmCCCCode.GetData(1)
      m_SP25 = frmCCCCode.GetData(2)
   End If
   Unload frmCCCCode
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
            'GetCu103ByCustomer Me, m_TM23
            Call Pub_GetDataFrm020102(m_TM23, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
            
            'edit by nickc 2006/01/20
            'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Then
            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
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
            End If
            'add by nickc 2007/08/10 多申請人也要
            If m_TM78 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, m_TM78
            Call Pub_GetDataFrm020102(m_TM78, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
            
            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
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
            End If
            End If
            If m_TM79 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, m_TM79
            Call Pub_GetDataFrm020102(m_TM79, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
            
            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
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
            End If
            End If
            If m_TM80 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, m_TM80
            Call Pub_GetDataFrm020102(m_TM80, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
            
            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
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
            End If
            End If
            If m_TM81 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, m_TM81
            Call Pub_GetDataFrm020102(m_TM81, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
            
            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
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
      ' 列印定稿
      If textPrint <> "N" Then
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
   textTM15.BackColor = &H8000000F
   textTM21.BackColor = &H8000000F
   textTM22.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   'add by nickc
   textTM78.BackColor = &H8000000F
   textTM79.BackColor = &H8000000F
   textTM80.BackColor = &H8000000F
   textTM81.BackColor = &H8000000F
   
   textTM45.BackColor = &H8000000F
      
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textTM44.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   
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
      ' 審定號數
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      'add by nickc 2007/02/01
      ElseIf IsNull(rsTmp.Fields("TM12")) = False Then
        textTM15 = rsTmp.Fields("TM12")
      End If
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 使用期間(起)
      If IsNull(rsTmp.Fields("TM21")) = False Then
         'Modified by Lydia 2017/08/25 改顯示為西元年
         textTM21 = ChangeWStringToWDateString(rsTmp.Fields("TM21"))
      End If
      ' 使用期間(迄)
      If IsNull(rsTmp.Fields("TM22")) = False Then
         'Modified by Lydia 2017/08/25 改顯示為西元年
         textTM22 = ChangeWStringToWDateString(rsTmp.Fields("TM22"))
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      'add by nickc 2007/02/01
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = rsTmp.Fields("TM78")
         textTM78 = GetCustomerName(rsTmp.Fields("TM78"), 0)
      End If
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = rsTmp.Fields("TM79")
         textTM79 = GetCustomerName(rsTmp.Fields("TM79"), 0)
      End If
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = rsTmp.Fields("TM80")
         textTM80 = GetCustomerName(rsTmp.Fields("TM80"), 0)
      End If
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = rsTmp.Fields("TM81")
         textTM81 = GetCustomerName(rsTmp.Fields("TM81"), 0)
      End If
      
      
      ' 案件備註
      textSP18 = Empty
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
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP05")
      End If
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP06")
      End If
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP07")
      End If
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      'add by nickc 2007/02/01
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("SP58")) = False Then
         m_TM78 = rsTmp.Fields("SP58")
         textTM78 = GetCustomerName(rsTmp.Fields("SP58"), 0)
      End If
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("SP59")) = False Then
         m_TM79 = rsTmp.Fields("SP59")
         textTM79 = GetCustomerName(rsTmp.Fields("SP59"), 0)
      End If
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("SP65")) = False Then
         m_TM80 = rsTmp.Fields("SP65")
         textTM80 = GetCustomerName(rsTmp.Fields("SP65"), 0)
      End If
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("SP66")) = False Then
         m_TM81 = rsTmp.Fields("SP66")
         textTM81 = GetCustomerName(rsTmp.Fields("SP66"), 0)
      End If
      
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         'Add By Sindy 2009/06/29
         textTM15 = rsTmp.Fields("SP11")
      End If
      
      ' 案件備註
      textSP18 = Empty
      If IsNull(rsTmp.Fields("SP18")) = False Then
         textSP18 = rsTmp.Fields("SP18")
      End If
      SetTMSPFieldOldData "SP18", textSP18, 0
      ' 91.09.02 modify by louis
      ' 廠商號碼
      m_SP19 = Empty
      If IsNull(rsTmp.Fields("SP19")) = False Then
         m_SP19 = rsTmp.Fields("SP19")
      End If
      ' 使用期間(起)
      If IsNull(rsTmp.Fields("SP20")) = False Then
         'Modified by Lydia 2017/08/25 改顯示為西元年
         textTM21 = ChangeWStringToWDateString(rsTmp.Fields("SP20"))
      End If
      ' 使用期間(迄)
      If IsNull(rsTmp.Fields("SP21")) = False Then
         'Modified by Lydia 2017/08/25 改顯示為西元年
         textTM22 = ChangeWStringToWDateString(rsTmp.Fields("SP21"))
      End If
      'Add By Cheng 2002/07/17
      m_SP24 = Empty
      If IsNull(rsTmp.Fields("SP24")) = False Then
         m_SP24 = rsTmp.Fields("SP24")
      End If
      'Add By Cheng 2002/07/17
      m_SP25 = Empty
      If IsNull(rsTmp.Fields("SP25")) = False Then
         m_SP25 = rsTmp.Fields("SP25")
      End If
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("SP72"))
      m_textPrint = textPrint
      SetTMSPFieldOldData "SP72", textPrint, 0
      ' 正片號碼
      textCP22 = Empty
      Select Case m_CP10
         Case "803", "804":
            EnableTextBox textSP22, True
            '911017 nick
            'If IsNull(rsTmp.Fields("SP22")) = False Then
            '   textCP22 = rsTmp.Fields("SP22")
            'End If
            '因為改成table
            textSP22 = Empty
            Dim nick911017rs As New ADODB.Recordset
            Dim nickstrsql  As String
            '911018 nick 只能抓該收文號的
            nickstrsql = "select BC02 from barcode where bc01 ='" & m_CP09 & "' "
            nick911017rs.CursorLocation = adUseClient
            nick911017rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
            If nick911017rs.RecordCount <> 0 Then
                nick911017rs.MoveFirst
                Do While nick911017rs.EOF = False
                    textSP22 = textSP22 & CheckStr(nick911017rs.Fields(0).Value)
                    nick911017rs.MoveNext
                    If nick911017rs.EOF = False Then
                        textSP22 = textSP22 & ","
                    End If
                Loop
            End If
            SetTMSPFieldOldData "SP22", textSP22, 0
         Case Else
            EnableTextBox textSP22, False
      End Select
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgress()
Dim strTemp As String
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
      m_CP44 = CheckStr(rsTmp.Fields("CP44"))
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
      ' 智權人員
      'Add By Cheng 2002/07/17
      m_CP13 = Empty
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      ' 承辦人員
      m_CP14 = Empty
      If IsNull(rsTmp.Fields("CP14")) = False Then
         m_CP14 = rsTmp.Fields("CP14")
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      
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
      ' 彼所案號
      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         strCP45 = rsTmp.Fields("CP45")
      End If
      SetCPFieldOldData "CP45", strCP45, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      'Remove by Lydia 2017/08/21
      ' 授權期間(起)
      'textCP53 = Empty
      'strTemp = Empty
      'If IsNull(rsTmp.Fields("CP53")) = False Then
      '   textCP53 = rsTmp.Fields("CP53")
      'End If
      'SetCPFieldOldData "CP53", strTemp, 1
      ' 授權期間(迄)
      'textCP54 = Empty
      'strTemp = Empty
      'If IsNull(rsTmp.Fields("CP54")) = False Then
      '   textCP54 = rsTmp.Fields("CP54")
      'End If
      'SetCPFieldOldData "CP54", strTemp, 1
      'end 2017/08/21
      'Added by Lydia 2017/08/21 移除TextBox
      SetCPFieldOldData "CP53", "" & rsTmp.Fields("CP53"), 1
      SetCPFieldOldData "CP54", "" & rsTmp.Fields("CP54"), 1
      'end 2017/08/21
      
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
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
      
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
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
      
   strSql = "SELECT CF09 FROM CASEFEE WHERE CF01='" & m_TM01 & "' AND CF02='" & m_TM10 & "' AND CF03='" & m_CP10 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CF09")) = False Then textCF09 = rsTmp.Fields("CF09")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
    'Add By Cheng 2003/07/15
    '若為條碼製作正片, 預設開啟Word
    If m_TM01 = "TB" And m_CP10 = "803" Then
        Me.textWord.Text = "Y"
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
   
   Call PUB_TCaseEFeeRemind(m_CP09) 'Add By Sindy 2016/5/9 內商電子收文請款提醒訊息
   
  'Added by Lydia 2017/08/21 TB案繳年費的第?期登記期,預設為前一次發文繳年費之CP53的值+1
   If m_TM01 = "TB" And m_CP10 = "708" Then
      strSql = "SELECT nvl(cp53,1) a1 FROM caseprogress WHERE cp01='" & m_TM01 & "' AND cp02='" & m_TM02 & "' AND cp03='" & m_TM03 & "' AND cp04='" & m_TM04 & "' " & _
               "and cp10='708' and cp158>0 and cp09 <> '" & m_CP09 & "' order by cp158 desc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount = 0 Then
         textPeriod = "1"
      Else
         textPeriod = Val(rsTmp.Fields("a1")) + 1
      End If
      rsTmp.Close
      Set rsTmp = Nothing
      textPeriod.Enabled = True
   Else
      textPeriod = ""
      textPeriod.Enabled = False
   End If
   'end 2017/08/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm020102_19 = Nothing
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

' 使用期間起
'Remove by Lydia 2017/08/21
'Private Sub textCP53_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'
'   If IsEmptyText(textCP53) = False Then
'      If CheckIsDate(textCP53, False) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "請輸入正確的使用期間起日"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP53_GotFocus
'      End If
'   End If
'End Sub

' 使用期間迄
'Remove by Lydia 2017/08/21
'Private Sub textCP54_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'
'   If IsEmptyText(textCP54) = False Then
'      If CheckIsDate(textCP54, False) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "請輸入正確的使用期間迄日"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP54_GotFocus
'      End If
'   End If
'   ' 新使用期間
'   If IsEmptyText(textCP53) = False And IsEmptyText(textCP54) = False Then
'      If Val(textCP53) > Val(textCP54) Then
'         Cancel = True
'         strTit = "檢核資料"
'         strMsg = "使用期間範圍不正確"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP53_GotFocus
'      End If
'   End If
'End Sub

Private Sub textFee_Validate(Cancel As Boolean)
'Add By Cheng 2002/06/18
If Len(Me.textFee.Text) > 0 Then
   If Len(Me.textFee.Text) = 8 Then
      If CheckIsDate(DBDATE(Me.textFee.Text)) = False Then
         Cancel = True
         Me.textFee.SetFocus
         textFee_GotFocus
      End If
   Else
      If CheckIsTaiwanDate(Me.textFee.Text) = False Then
         Cancel = True
         Me.textFee.SetFocus
         textFee_GotFocus
      End If
   End If
End If
End Sub

' 案件備註
Private Sub textSP18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP18, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP18_GotFocus
   End If
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

' 列印定稿
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

' 正片號碼
Private Sub textSP22_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textSP22) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textSP22)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textSP22, nIndex)
      If m_CP10 = "803" Then
         ' 91.09.02 modify by louis 13碼
         If Len(strTemp) <> 13 Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "正片號碼<" & strTemp & ">不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP22_GotFocus
            GoTo EXITSUB
         Else
            ' 正片號碼前幾碼必須與廠商號碼相同
            '911017 nick 檢查條碼公式已經改為前三碼為 471 ，廠商號碼緊接在後，再來不管
            'If m_SP19 <> Left(strTemp, Len(m_SP19)) Then
            If (Mid(strTemp, 1, 3) <> "471" Or Mid(strTemp, 4, Len(m_SP19)) <> m_SP19) Then
               '911017 nick
               Cancel = True
               
               strTit = "檢核資料"
               strMsg = "正片號碼<" & strTemp & ">不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSP22_GotFocus
               GoTo EXITSUB
            End If
         End If
      End If
      If m_CP10 = "804" Then
         If Len(strTemp) <> 13 Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "正片號碼<" & strTemp & ">不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textSP22_GotFocus
            GoTo EXITSUB
         Else
            ' 91.09.02 modify by louis
            ' 不檢查條碼公式
            'If IsBarcodeCorrect(strTemp) = False Then
            '   strTit = "檢核資料"
            '   strMsg = "正片號碼<" & strTemp & ">不正確"
            '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            '   textSP22_GotFocus
            '   GoTo EXITSUB
            'End If
            ' 正片號碼前幾碼必須與廠商號碼相同
            '911017 nick 檢查條碼公式已經改為前三碼為 471 ，廠商號碼緊接在後，再來不管
            'If m_SP19 <> Left(strTemp, Len(m_SP19)) Then
            If (Mid(strTemp, 1, 3) <> "471" Or Mid(strTemp, 4, Len(m_SP19)) <> m_SP19) Then
               '911017 nick
               Cancel = True
               
               strTit = "檢核資料"
               strMsg = "正片號碼<" & strTemp & ">不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSP22_GotFocus
               GoTo EXITSUB
            End If
         End If
      End If
   Next nIndex
   
   ' 檢查是否重覆
   For nIndex = 1 To nCount
      strTemp = GetSubString(textSP22, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textSP22, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "正片號碼<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textSP22_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   
EXITSUB:
End Sub

Private Sub textWord_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否修改定稿內容
Private Sub textWord_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textWord) = False Then
      Select Case textWord
         Case "", " ", "Y":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textWord_GotFocus
      End Select
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
   ' 是否出名
   SetCPFieldNewData "CP22", textCP22
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' 專用期間(起)
   'SetCPFieldNewData "CP53", DBDATE(textCP53) 'Remove by Lydia 2017/08/21
   ' 專用期間(迄)
   'SetCPFieldNewData "CP54", DBDATE(textCP54) 'Remove by Lydia 2017/08/21
   ' 進度備註
   SetCPFieldNewData "CP64", textCP64
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
   Select Case m_TM01
      Case "T", "TF", "FCT":
         ' 案件備註
         SetTMSPFieldNewData "TM58", textSP18
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
         ' 正片號碼
         Select Case m_CP10
            Case "803", "804":
               SetTMSPFieldNewData "SP22", textSP22
            Case Else
         End Select
         ' 案件備註
         SetTMSPFieldNewData "SP18", textSP18
         'add by nickc 2006/11/17
         If textPrint <> "N" Then
             SetTMSPFieldNewData "SP72", textPrint
         Else
             SetTMSPFieldNewData "SP72", m_textPrint
         End If
         
         'Added by Lydia 2017/08/21 條碼案繳年費第?期登記期
         If m_TM01 = "TB" And m_CP10 = "708" Then
            SetCPFieldNewData "CP53", textPeriod
            SetCPFieldNewData "CP54", textPeriod
         End If
         'end 2017/08/21
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
               ' 91.03.25 modify by louis (單引號)
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
'Modify By Cheng 2002/11/07
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
               ' 91.03.25 modify by louis (單引號)
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
   '911017 nick 因為 改成另一個table
   '所以要獨立存檔
    'Modify By Cheng 2002/11/07
'   SaveBarCode
   If SaveBarCode = False Then GoTo ErrorHandler
'Add By Cheng 2002/11/07
Exit Function
ErrorHandler:
    OnUpdateServicePractice = False
End Function

'911017 nick 儲存 barcode 檔 ok
'Modify By Cheng 2002/11/07
'Function SaveBarCode()
Function SaveBarCode() As Boolean
'911017 nick 先刪除
Dim nickstrsql As String
Dim nickIndex As Integer
Dim strBC01 As String
Dim strBC02 As String
Dim ArrBarcode As Variant
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
SaveBarCode = True

'911018 nick 只能刪除該收文號的
nickstrsql = "delete barcode where bc01='" & m_CP09 & "' "
cnnConnection.Execute nickstrsql
'911017 nick 再新增
ArrBarcode = Split(textSP22, ",")
For nickIndex = 0 To UBound(ArrBarcode)
    If Len(ArrBarcode(nickIndex)) <> 0 Then
        strBC02 = ArrBarcode(nickIndex)
        strBC01 = m_CP09
        nickstrsql = "insert into barcode (bc01,bc02,bc03) values ('" & strBC01 & "','" & strBC02 & "',null) "
        cnnConnection.Execute nickstrsql
    End If
Next nickIndex
'Add By Cheng 2002/11/07
Exit Function
ErrorHandler:
    SaveBarCode = False
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
'Add By Cheng 2002/11/07
Exit Function
ErrorHandler:
    OnUpdateCaseProgress = False
End Function

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strSql As String
   Dim strDate As String
   Dim strNP22 As String
   Dim bolSysDt As Boolean 'Add By Sindy 2010/12/28
   Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2012/9/10
   Dim strNP07 As String, strNP08 As String 'Add By Sindy 2012/9/10
   
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
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新案件進度檔
    'Modify By Cheng 2002/11/07
'   OnUpdateCaseProgress
   If OnUpdateCaseProgress = False Then GoTo ErrorHandler
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有輸入下次繳年費期間限, 新增繳年費的記錄到下一程序檔
   If IsEmptyText(textFee) = False Then
      strNP22 = GetNextProgressNo()
    'Modify By Cheng 2003/09/01
'      strDate = DBDATE(DateSerial(Val(DBYEAR(textFee)), Val(DBMONTH(textFee)), Val(DBDAY(textFee)) - 2))
      'Modify By Sindy 2014/10/6 台灣案之本所期限設定
      If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
         strDate = PUB_GetOurDeadline(DBDATE(textFee))
      Else
      '2014/10/6 END
         strDate = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(textFee))))
      End If
        'Modify By Cheng 2003/11/24
        '重抓智權人員
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & "708" & "," & _
'                          DBDATE(strDate) & "," & DBDATE(textFee) & ",'" & m_CP13 & "'," & strNP22 & ")"
      'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & "708" & "," & _
                          DBDATE(strDate) & "," & DBDATE(textFee) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & "708" & "," & _
                          PUB_GetWorkDay1(strDate, True) & "," & DBDATE(textFee) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 繳年費不印
      ' 列印國內案件接洽及結案記錄單
      ' g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
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
   
   'Add By Sindy 2012/9/10
   ' 若有審查天數, 新增一筆催審期限的記錄到下一程序檔
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
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
   End If
   rsTmp.Close
   '2012/9/10 End
   
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
   
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans

     'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
    PUB_CheckEMail m_CP44, m_CP116
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
   'Add by Amy 2021/12/27檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        GoTo EXITSUB
   End If

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
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' 新使用期間
   'Remove by Lydia 2017/08/21
'   If IsEmptyText(textCP53) = False And IsEmptyText(textCP54) = False Then
'      If Val(textCP53) > Val(textCP54) Then
'         strTit = "檢核資料"
'         strMsg = "使用期間範圍不正確"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP53.SetFocus
'         GoTo EXITSUB
'      End If
'   End If
    'end 2017/08/21
    
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
   
   'Added by Lydia 2017/08/21 TB欲發文之案件性質為'繳年費'708時，'第?期登記期'及'下次繳年費期限'欄皆不可空白；
   If m_TM01 = "TB" And m_CP10 = "708" Then
      If textPeriod = "" Then
         MsgBox "第?期登記期不可空白!!!", vbExclamation + vbOKOnly
         textPeriod.SetFocus
         GoTo EXITSUB
      End If
      If textFee = "" Then
         MsgBox "下次繳年費期限不可空白!!!", vbExclamation + vbOKOnly
         textFee.SetFocus
         GoTo EXITSUB
      End If
   End If
   'end 2017/08/21
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textWord_GotFocus()
   InverseTextBox textWord
End Sub

Private Sub textFee_GotFocus()
   InverseTextBox textFee
End Sub

Private Sub textSP18_GotFocus()
   InverseTextBox textSP18
End Sub

Private Sub textSP22_GotFocus()
   InverseTextBox textSP22
End Sub

Private Sub textCP22_GotFocus()
   InverseTextBox textCP22
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

'Remove by Lydia 2017/08/21
'Private Sub textCP53_GotFocus()
'   InverseTextBox textCP53
'End Sub
'
'Private Sub textCP54_GotFocus()
'   InverseTextBox textCP54
'End Sub
'end 2017/08/21

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
Dim strTM23Nation As String
Dim strSql As String
Dim arrSP22
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   'add by nickc 2006/06/29
   If textPrint = "1" Then
        Select Case m_CP10
           ' 移轉
           Case "501":
              ' 清除定稿例外欄位檔原有資料
              EndLetter "01", m_CP09, "29", strUserNum
           ' 繳年費
           Case "708":
              ' 清除定稿例外欄位檔原有資料
              EndLetter "01", m_CP09, "32", strUserNum
           ' 條碼申請
           Case "802":
              ' 清除定稿例外欄位檔原有資料
              EndLetter "01", m_CP09, "28", strUserNum
              ' 回音
              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       "VALUES ('" & "01" & "','" & m_CP09 & "','" & "28" & "','" & strUserNum & "'," & _
                       "'" & "回音" & "','" & textCF09 & "')"
              cnnConnection.Execute strSql
             'Add By Cheng 2003/07/15
             '製作正片
           Case "803":
              ' 清除定稿例外欄位檔原有資料
              EndLetter "01", m_CP09, "28", strUserNum
              ' 片數
             If Me.textSP22.Text <> "" Then
                 arrSP22 = Split(Me.textSP22.Text, ",")
                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          "VALUES ('" & "01" & "','" & m_CP09 & "','" & "28" & "','" & strUserNum & "'," & _
                          "'" & "片數" & "','" & UBound(arrSP22) + 1 & "')"
                 cnnConnection.Execute strSql
             End If
              ' 正片號碼
              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       "VALUES ('" & "01" & "','" & m_CP09 & "','" & "28" & "','" & strUserNum & "'," & _
                       "'" & "正片號碼" & "','" & Replace(Me.textSP22.Text, ",", "、") & "')"
              cnnConnection.Execute strSql
        End Select
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
   
   'add by nickc 2006/06/29
   If textPrint = "1" Then
        Select Case m_CP10
           ' 移轉
           Case "501":
              ' 列印定稿
'              NowPrint m_CP09, "01", "29", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
               ET03 = "29" 'Modify By Sindy 2012/1/12
           ' 繳年費
           Case "708":
              ' 列印定稿
'              NowPrint m_CP09, "01", "32", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
               ET03 = "32" 'Modify By Sindy 2012/1/12
           ' 條碼申請
           Case "802":
              ' 列印定稿
'              NowPrint m_CP09, "01", "28", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
               ET03 = "28" 'Modify By Sindy 2012/1/12
           ' 製作正片
           Case "803":
              ' 列印定稿
'              NowPrint m_CP09, "01", "28", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
               ET03 = "28" 'Modify By Sindy 2012/1/12
        End Select
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
   
   'Remove by Lydia 2017/08/21
'   If Me.textCP53.Enabled = True Then
'      Cancel = False
'      textCP53_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
'
'   If Me.textCP54.Enabled = True Then
'      Cancel = False
'      textCP54_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
   'end 2017/08/21
   
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
   
   If Me.textSP18.Enabled = True Then
      Cancel = False
      textSP18_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textFee.Enabled = True Then
      Cancel = False
      textFee_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textWord.Enabled = True Then
      Cancel = False
      textWord_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   '920113 nick 新增
   If Me.textSP22.Enabled = True Then
       Cancel = True
       textSP22_Validate Cancel
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

'Added by Lydia 2017/08/21
Private Sub textPeriod_GotFocus()
   InverseTextBox textPeriod
End Sub

Private Sub textPeriod_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPeriod) = False Then
      If IsNumeric(textPeriod) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的數值"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPeriod_GotFocus
      End If
   End If
End Sub
'end 2017/08/21

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
