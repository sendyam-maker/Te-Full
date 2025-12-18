VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_a 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文-年費"
   ClientHeight    =   5630
   ClientLeft      =   -130
   ClientTop       =   1030
   ClientWidth     =   8710
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5630
   ScaleWidth      =   8710
   Begin VB.TextBox txtEmail 
      Height          =   270
      Left            =   7380
      MaxLength       =   1
      TabIndex        =   13
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox txtRecDate 
      Height          =   270
      Left            =   5010
      MaxLength       =   1
      TabIndex        =   12
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox txtPAID 
      Height          =   270
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   11
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox txtPayToday 
      Height          =   270
      Left            =   5010
      MaxLength       =   1
      TabIndex        =   10
      Top             =   4980
      Width           =   375
   End
   Begin VB.TextBox Text13 
      Height          =   270
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TextCP148 
      Height          =   270
      Left            =   7380
      MaxLength       =   1
      TabIndex        =   5
      Top             =   3150
      Width           =   375
   End
   Begin VB.TextBox txtCP118 
      Height          =   270
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   9
      Top             =   4980
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   34
      Top             =   900
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   33
      Top             =   900
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   32
      Top             =   900
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   31
      Top             =   900
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060104_a.frx":0000
      Left            =   1080
      List            =   "frm060104_a.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   21
      Top             =   2415
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   3
      Left            =   2895
      TabIndex        =   16
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6180
      TabIndex        =   19
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5340
      TabIndex        =   18
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "繳費記錄(&Y)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7395
      TabIndex        =   20
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "變更事項(&R)"
      Height          =   400
      Index           =   4
      Left            =   4125
      TabIndex        =   17
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.TextBox Text5 
      Height          =   855
      Index           =   8
      Left            =   1230
      TabIndex        =   8
      Top             =   4080
      Width           =   5865
      VariousPropertyBits=   671105051
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "10345;1508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   7
      Left            =   1230
      TabIndex        =   14
      Top             =   3750
      Width           =   1095
      VariousPropertyBits=   679493661
      BackColor       =   -2147483648
      MaxLength       =   8
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   6
      Left            =   4080
      TabIndex        =   7
      Top             =   3450
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   5
      Left            =   1230
      TabIndex        =   6
      Top             =   3450
      Width           =   795
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1402;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   4
      Top             =   3150
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1931;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   3
      Left            =   1650
      TabIndex        =   3
      Top             =   3150
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "661;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   2
      Left            =   5640
      TabIndex        =   2
      Top             =   2850
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "661;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   1
      Left            =   1860
      TabIndex        =   1
      Top             =   2850
      Width           =   495
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "873;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text5 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   2850
      Width           =   495
      VariousPropertyBits=   671105051
      MaxLength       =   2
      Size            =   "873;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7170
      TabIndex        =   15
      Top             =   3750
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
      TabIndex        =   66
      Top             =   5325
      Width           =   1860
   End
   Begin VB.Label lblRecDate 
      AutoSize        =   -1  'True
      Caption         =   "當天請款:             (Y:是)"
      Height          =   210
      Left            =   4080
      TabIndex        =   65
      Top             =   5310
      Width           =   1815
   End
   Begin VB.Label lblPAID 
      AutoSize        =   -1  'True
      Caption         =   "已收款:           (1-不寄D/N, 2-寄D/N)"
      Height          =   180
      Left            =   750
      TabIndex        =   64
      Top             =   5325
      Width           =   2700
   End
   Begin VB.Label lblPayToday 
      AutoSize        =   -1  'True
      Caption         =   "電子送件是否當日扣款:           (Y/N)"
      Height          =   180
      Left            =   3120
      TabIndex        =   63
      Top             =   5025
      Width           =   2745
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Instruction No. "
      Height          =   180
      Left            =   390
      TabIndex        =   62
      Top             =   165
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否特殊請款:"
      Height          =   180
      Left            =   6210
      TabIndex        =   61
      Top             =   3210
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(Y:是)"
      Height          =   180
      Left            =   7800
      TabIndex        =   60
      Top             =   3210
      Width           =   465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:          (Y: 是)"
      Height          =   180
      Left            =   240
      TabIndex        =   59
      Top             =   5025
      Width           =   2085
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人:"
      Height          =   180
      Left            =   6180
      TabIndex        =   58
      Top             =   3810
      Width           =   945
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      X1              =   180
      X2              =   8690
      Y1              =   2805
      Y2              =   2805
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   11
      Left            =   5220
      TabIndex        =   57
      Top             =   3465
      Width           =   3270
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
      Index           =   10
      Left            =   2070
      TabIndex        =   56
      Top             =   3465
      Width           =   855
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
      Index           =   9
      Left            =   1080
      TabIndex        =   55
      Top             =   1200
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "5186;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "巳繳年費:"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   54
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   53
      Top             =   2415
      Width           =   765
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Left            =   240
      TabIndex        =   52
      Top             =   2100
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Left            =   4080
      TabIndex        =   51
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Left            =   240
      TabIndex        =   50
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Left            =   4080
      TabIndex        =   49
      Top             =   1500
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Left            =   240
      TabIndex        =   48
      Top             =   1500
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Left            =   4080
      TabIndex        =   47
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   46
      Top             =   900
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   4080
      TabIndex        =   45
      Top             =   570
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   240
      TabIndex        =   44
      Top             =   570
      Width           =   585
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   43
      Top             =   570
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2805;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   1
      Left            =   4890
      TabIndex        =   42
      Top             =   570
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2805;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   2
      Left            =   4890
      TabIndex        =   41
      Top             =   900
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2805;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   40
      Top             =   1500
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
      Left            =   4890
      TabIndex        =   39
      Top             =   1500
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
      Index           =   5
      Left            =   1080
      TabIndex        =   38
      Top             =   1800
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
      Left            =   4890
      TabIndex        =   37
      Top             =   1800
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
      Index           =   7
      Left            =   1080
      TabIndex        =   36
      Top             =   2100
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
      Index           =   8
      Left            =   1740
      TabIndex        =   35
      Top             =   2415
      Width           =   6930
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "12224;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "下次繳費日:"
      Height          =   180
      Left            =   240
      TabIndex        =   30
      Top             =   3810
      Width           =   945
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書:              (Y:Word)"
      Height          =   180
      Left            =   240
      TabIndex        =   29
      Top             =   3210
      Width           =   2625
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   3480
      TabIndex        =   28
      Top             =   3210
      Width           =   585
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   240
      TabIndex        =   27
      Top             =   3495
      Width           =   585
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "年費代理人:"
      Height          =   180
      Left            =   3120
      TabIndex        =   26
      Top             =   3495
      Width           =   945
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   240
      TabIndex        =   25
      Top             =   4080
      Width           =   765
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "繳納第:                至                年 年費"
      Height          =   180
      Left            =   240
      TabIndex        =   24
      Top             =   2895
      Width           =   2790
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "是否逾期補繳:              (Y:是)"
      Height          =   180
      Left            =   4410
      TabIndex        =   22
      Top             =   2895
      Width           =   2220
   End
End
Attribute VB_Name = "frm060104_a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/15 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/8/4 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String, m_CP110 As String, m_AgentName As String

Dim intWhere As Integer
Dim strTemp As String, strSales As String
Dim strCaseFee(1 To 2) As String 'strCaseFee(1) 國家檔中繳費年度，strCaseFee(2) 國家檔中起算日
'Add By Cheng 2002/07/04
Dim m_CP07 As String '法定期限
Dim m_CP10 As String '案件性質
Dim m_CP14 As String 'Add By Sindy 2016/11/16
'Add By Cheng 2003/01/01
Dim m_strOfficalFee As String '規費
Dim m_strServiceFee As String '服務費
Dim m_strPoints As String '點數
'92.9.15 add by sonia
Dim m_strNP09 As String
Dim m_strNP09_1 As String
'92.9.15 end
'Add By Cheng 2003/10/06
Dim m_blnClkChgEvnBtn As Boolean '是否按下變更事項按鈕
Dim m_CP81 As String '可否減免
Dim m_lngDisc As Long '減免金額
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/20 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_DiscType As String 'Add by Morgan 2010/6/28 減免身分
'Add by Morgan 2012/9/13
Public m_bolBeCalled As Boolean '整批呼叫
Public m_CP01 As String
Public m_CP02 As String
Public m_CP03 As String
Public m_CP04 As String
Public m_CP09 As String
Public m_CP84 As String
Dim m_bol2ndY As Boolean 'Added by Morgan 2012/10/3 次年是否為補繳
Public m_CP152 As String 'Added by Morgan 2013/5/15
Dim m_CP60 As String 'Added by Lydia 2015/02/26
Dim m_CP142 As String 'Add By Sindy 2015/12/17
Dim m_CP164 As String 'Add By Sindy 2021/4/20
'Added by Lydia 2018/09/11
Dim m_CP118 As String '是否電子送件
Dim m_CP82 As String '發文時間
Dim m_AddMcRecord As String 'Added by Lydia 2020/08/17 人工Email維護(語法)
Dim m_eFlag As String 'Added by Lydia 2020/08/17 是否e/E化


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 5) As String, strTmp As String
   
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   If Text5(0).Text = Text5(1).Text Then
      strTmp = "第 " & Text5(0) & " 年年費"
   Else
      strTmp = "第 " & Text5(0) & " 至 " & Text5(1) & " 年年費"
   End If
   
   strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','第幾年至幾年費','" & strTmp & "')"
   
   If Text5(2) = "Y" Then
      strTxt(2) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','列印備註','(逾期補繳)')"
      'ADD BY SONIA 2014/4/1 有收414才加印 FCP-029467
      strExc(0) = "select * from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
               " and cp10='414' and cp57 is null and cp27 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strTxt(5) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                    "','申請回復領證','♀')"
      Else
         strTxt(5) = ""
      End If
      '2014/4/1 END
   End If
      
   'Add by Morgan 2008/1/28
   If m_lngDisc > 0 Then
      strTxt(3) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','減免金額','" & m_lngDisc & "')"
         
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
      
      strTxt(4) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','減免說明','" & strExc(1) & "')"
      'end 2010/6/28
   End If
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(2, strTxt) Then
   If Not ClsLawExecSQL(5, strTxt) Then                     '20140324MODIFY By eric
'   If Not ClsLawExecSQL(4, strTxt) Then                    '20140324REMARK By eric
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim bolChk As Boolean, ET03 As String
   Dim nFrm As Form  'Added by Lydia 2019/10/01
   
   Select Case Index
      'Modify by Morgan 2009/3/26 將同時發文併入
      'Case 0 '確定
      Case 0, 3
      
         'Added by Lydia 2020/08/17 Email維護
         m_AddMcRecord = ""
         If txtEmail.Visible = True And txtEmail.Text = "Y" Then
            strExc(5) = Pub_FcpSetPayToday("2", Text5(4).Text, txtPayToday.Text) '扣款日
            '開啟Email畫面
            Call PUB_GetFCPEmpMail("2", strReceiveNo, m_eFlag, textCP148, txtPAID, txtRecDate, IIf(Text5(2).Text = "Y", "逾期補繳", ""), strExc(5), strExc(1), strExc(2), strExc(3), strExc(4))
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
               'Modified by Lydia 2010/09/11
               'If m_AddMcRecord = "Y" Then
               If m_AddMcRecord = "" Or m_AddMcRecord = "Y" Then
                   MsgBox "Email維護未確認，請重新確認Email !", vbCritical, "檢核資料"
                   Exit Sub
               End If
            End If
         End If
         'end 2020/08/17
         
         ' 設定滑鼠游標為等待狀態
         Screen.MousePointer = vbHourglass
         
         If Process() = False Then Screen.MousePointer = vbDefault: Exit Sub
            
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
         
         If txtCP118 = "" Then 'Added by Morgan 2012/9/13
         
            If Text5(3) = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            
            'Modified by Morgan 2012/7/9 中小企業要印條款改新定稿
            'ET03 = "00"
            ET03 = "02"
            'Add by Morgan 2008/3/18
            '年費申請人不出名
            If pa(143) = "N" Then
               'Modified by Morgan 2012/7/9 中小企業要印條款改新定稿
               'ET03 = "01"
               ET03 = "03"
            End If
            'end 2008/3/18
            
            StartLetter "01", ET03
            NowPrint strReceiveNo, "01", ET03, bolChk, strUserNum, 0
            
         End If 'Added by Morgan 2012/9/13
                 
         'Added by Lydia 2019/10/01 領證/年費發文直接產生:承辦單+請款定稿+帳單(請款單)
         'Modified by Lydia 2020/08/17 重新發文不產生請款單 (P.S 修改條件,請一併更新FormSave的.CommitTrans)
         'If m_bolBeCalled = False And m_CP60 = "" Then '排除年費整批發文
         If m_bolBeCalled = False And m_CP60 = "" And Val(m_CP82) = 0 Then
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
            Call frm060307.SetData(1, txtPAID.Text) '已收款
            Call frm060307.SetData(2, txtRecDate.Text)  '當天請款
            Call frm060307.SetData(3, IIf(Text5(2).Text = "Y", "逾期補繳", "")) '逾期補繳(來源表單的設定之描述)
            If m_AddMcRecord <> "" Then
                Call frm060307.SetData(4, m_AddMcRecord)
            End If
            'end 2020/08/17
            frm060307.Show
            Call frm060307.cmdOK_Click(0)
            'Modified by Lydia 2020/08/17 Transaction將發文和年證費請款函一併包入
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
             'cnnConnection.CommitTrans 'Remove by Lydia 2020/08/17 非外部呼叫改在FormSave執行
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
         
         If m_bolBeCalled = False Then 'Added by Morgan 2012/9/13
         
            'Add by Morgan 2008/2/20 檢查代理人Email
            PUB_CheckEMail pa(75), pa(144)
            If pa(145) <> "" Then
               PUB_CheckEMail pa(75), pa(145)
            End If
            'end 2008/2/20
            
            If Index = 0 Then
               'Add By Sindy 2023/11/9
               If frm060104_1.bolIsEMPFlow = True Then
                  frm090202_4.QueryData
               End If
               '2023/11/9 End
               'Add By Cheng 2002/04/30
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
            
         End If 'Added by Morgan 2012/9/13
         
         Unload Me
      Case 1
         If m_bolBeCalled = False Then 'Added by Morgan 2012/9/13
            frm060104_1.Show
         End If 'Added by Morgan 2012/9/13
         Unload Me
      Case 2
         Set frm060104_b.oParent = Me 'Add by Morgan 2011/10/5
         frm060104_b.LoadMe pa(1), pa(2), pa(3), pa(4), 1
         Me.Hide
'Remove by Morgan 2009/3/26
'      Case 3 '同時發文
'         'Add By Cheng 2002/05/21
'         If CheckDataValid = False Then Exit Sub
'         If TxtValidate = False Then Exit Sub
'         'Add by Morgan 2009/3/20 設定是否算發文室案件
'         If ModifyDispatch(strReceiveNo, m_CP09s, m_CP123s, Text5(4)) = False Then
'             Exit Sub
'         End If
'         'end 2009/3/20
'
'         ' 設定滑鼠游標為等待狀態
'         Screen.MousePointer = vbHourglass
'         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
'         ' 設定滑鼠游標為預設
'         Screen.MousePointer = vbDefault
'
'         'frm060104_1.Text1 = pa(1)
'         'frm060104_1.Text2 = pa(2)
'         'frm060104_1.Text3 = pa(3)
'         'frm060104_1.Text4 = pa(4)
'         'frm060104_1.Command1_Click
'         frm060104_1.Show
'         ' 90.08.06 modify by louis
'         frm060104_1.ReQuery
'         Unload Me
      Case 4
         Me.Hide
         frm060104_5.LoadMe strReceiveNo, pa(1), pa(2), pa(3), pa(4), 10
         frm060104_5.Caption = "外專發文-變更事項"
        m_blnClkChgEvnBtn = True
   End Select
End Sub

Private Function FormSave() As Boolean
   Dim i As Integer
   Dim strTmp(0 To 3) As String
   Dim strFLD As String
   Dim nMaxNo As String
   Dim nPos As Integer
   Dim aryCurr As Variant
   Dim aryAll As Variant
   Dim aryDate As Variant
   Dim nPosBegin As Integer
   Dim nPosEnd As Integer
   Dim nDot As Integer
   'Add By Cheng 2002/11/08
   Dim ii As Integer
   Dim stCP118 As String, stCP152 As String 'Added by Lydia 2018/09/11
   Dim strAgreeOnDate As String 'Add By Sindy 2012/8/17
   
 '911105 nick transation
 FormSave = True
  On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' 計算逗號的總數(幾格)
   nDot = 0
   For nPos = 1 To Len(pa(72))
      If Mid(pa(72), nPos, 1) = "," Then nDot = nDot + 1
   Next nPos
    'Add By Cheng 2002/11/08
    If nDot <> 0 Then
        If "" & pa(73) = "" Then
            '若繳年費日期逗號數與繳費年度逗號數不同時, 重新補上
            For ii = 1 To nDot
                pa(73) = pa(73) & ","
            Next ii
        End If
        If "" & pa(74) = "" Then
            '若繳年費是否雙倍逗號數與繳費年度逗號數不同時, 重新補上
            For ii = 1 To nDot
                pa(74) = pa(74) & ","
            Next ii
        End If
    End If
   
   aryAll = Split(strCaseFee(2), ",")
   aryCurr = Split(pa(72), ",")
   ' 找尋繳年費起始點位置
   nPosBegin = 0
   For nPos = 0 To UBound(aryAll)
      If aryAll(nPos) = Format(Val(Text5(0))) Then
         nPosBegin = nPos
         Exit For
      End If
   Next nPos
   ' 找尋繳年費終止點位置
   nPosEnd = 0
   For nPos = 0 To UBound(aryAll)
      If aryAll(nPos) = Format(Val(Text5(1))) Then
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
   pa(72) = strFLD
   
   ' 重新計算繳費年度共有幾欄
   nDot = 0
   For nPos = 1 To Len(pa(72))
      If Mid(pa(72), nPos, 1) = "," Then nDot = nDot + 1
   Next nPos
   
   ' 繳年費日期
   ReDim aryCurr(nDot)
   
   'Modify by Morgan 2007/3/8 只繳過一年的(沒有",")也要考慮否則第一次的資料會被清掉
   'If InStr(pa(73), ",") > 0 Then
      aryDate = Split(pa(73), ",")
      ' 拷貝原資料
      For nPos = 0 To UBound(aryDate)
         If IsEmptyText(aryDate(nPos)) = False Then
            If nDot > 0 Then
               aryCurr(nPos) = aryDate(nPos)
            End If
         End If
      Next nPos
   'End If
   
   ' 填入新資料
   For nPos = nPosBegin To nPosEnd
      aryCurr(nPos) = DBDATE(Text5(4))
   Next nPos
   ' 讀取新資料
   strFLD = Empty
   For nPos = 0 To UBound(aryCurr)
      If nPos > 0 Then: strFLD = strFLD & ","
      strFLD = strFLD & aryCurr(nPos)
   Next nPos
   pa(73) = strFLD
   
   '費用是否雙倍
   ReDim aryCurr(nDot)
   'Modify by Morgan 2007/3/8 只繳過一年的(沒有",")也要考慮否則第一次的資料會被清掉
   'If InStr(pa(74), ",") > 0 Then
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
   'End If
   
   ' 填入新資料
   For nPos = nPosBegin To nPosEnd
      If Text5(2) = "Y" Then
            '只有起始年要雙倍
            If nPos = nPosBegin Then
                aryCurr(nPos) = "Y"
            'Added by Morgan 2012/10/3 次年補繳也要紀錄
            ElseIf m_bol2ndY = True And nPos = nPosBegin + 1 Then
                aryCurr(nPos) = "Y"
            Else
                aryCurr(nPos) = ""
            End If
      Else
         aryCurr(nPos) = Empty
      End If
   Next nPos
   ' 讀取新資料
   strFLD = Empty
   For nPos = 0 To UBound(aryCurr)
      If nPos > 0 Then: strFLD = strFLD & ","
      strFLD = strFLD & aryCurr(nPos)
   Next nPos
   pa(74) = strFLD
 
   'For i = Text5(0) To Text5(1)
   '   pa(72) = pa(72) & "," & i
   'Next
   'For i = Text5(0) To Text5(1)
   '   pa(73) = pa(73) & "," & TransDate(Text5(4), 2)
   'Next
   'For i = Text5(0) To Text5(1)
   '   If i = Text5(0) Then
   '      pa(74) = pa(74) & ",Y"
   '   Else
   '      pa(74) = pa(74) & ","
   '   End If
   'Next
   'For i = 72 To 74
   '   If Left(pa(i), 1) = "," Then pa(i) = Mid(pa(i), 2)
   'Next
   '
   strTmp(0) = ""
    strExc(1) = "UPDATE PATENT SET PA76=" & CNULL(ChangeCustomerL(Text5(6))) & "," & _
      "PA72='" & pa(72) & "',PA73='" & pa(73) & "'," & _
      "PA74='" & pa(74) & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      
   '911105 nick transation
   cnnConnection.Execute strExc(1)
   
   'Add By Sindy 2015/9/23 +Instruction No. 儲存在進度備註
   If Text13.Enabled = True And Text13.Text <> "" Then
      Text5(8) = Text5(8) & ";" & Trim(Label11.Caption) & " " & Trim(Text13.Text) & ";"
   End If
   '2015/9/23 END
   
   'Added by Lydia 2018/09/11
   '電子送件有規費的一律設自動扣款(同內專) --敏莉
   stCP118 = txtCP118
   stCP152 = ""
   If m_bolBeCalled = False Then
        If txtCP118 = "Y" And Val(m_strOfficalFee) > 0 Then
           stCP118 = "A"
           stCP152 = Pub_FcpSetPayToday("2", Text5(4).Text, txtPayToday.Text)
        End If
   Else
        stCP152 = m_CP152
   End If
   'end 2018/09/11
   
   'Modify by morgan 2005/8/4 加 cp110
   'Modify by Moragn 2008/1/25 +CP81
   'Modified by Morgan 2012/9/13 +CP118
   'Modified by Morgan 2013/5/15 +CP152
   'Modify By Sindy 2015/9/23 +,CP148=" & CNULL(TextCP148) & "
   'Modified by Lydia 2018/09/11 +CP118,CP152
   ' strExc(2) = "UPDATE CASEPROGRESS SET CP27=" & CNULL(TransDate(Text5(4), 2)) & "," & _
      "CP14=" & CNULL(ChgSQL(Text5(5))) & ",cp64=" & CNULL(ChgSQL(Text5(8))) & _
      ",CP16=" & Val(m_strServiceFee) + Val(m_strOfficalFee) & _
      ",CP17=" & Val(m_strOfficalFee) & ",CP18=" & Val(m_strPoints) & _
      ", cp84=" & Format(Val(m_strOfficalFee)) & _
      ",cp110=" & CNULL(m_CP110) & ",CP22=decode('" & m_CP110 & "',NULL,'N',NULL),CP81=" & CNULL(m_CP81) & _
      ",CP118='" & txtCP118 & "',CP152=" & CNULL(m_CP152, True) & ",CP148=" & CNULL(TextCP148) & " WHERE CP09='" & strReceiveNo & "'"
    strExc(2) = "UPDATE CASEPROGRESS SET CP27=" & CNULL(TransDate(Text5(4), 2)) & "," & _
      "CP14=" & CNULL(ChgSQL(Text5(5))) & ",cp64=" & CNULL(ChgSQL(Text5(8))) & _
      ",CP16=" & Val(m_strServiceFee) + Val(m_strOfficalFee) & _
      ",CP17=" & Val(m_strOfficalFee) & ",CP18=" & Val(m_strPoints) & _
      ", cp84=" & Format(Val(m_strOfficalFee)) & _
      ",cp110=" & CNULL(m_CP110) & ",CP22=decode('" & m_CP110 & "',NULL,'N',NULL),CP81=" & CNULL(m_CP81) & _
      ",CP148=" & CNULL(textCP148) & ",CP118='" & stCP118 & "',CP152=" & CNULL(stCP152, True) & " " & _
      "WHERE CP09='" & strReceiveNo & "'"
      
   '911105 nick transation
   cnnConnection.Execute strExc(2)
    'Modify By Cheng 2004/01/05
    '若有下次繳費日, 才要新增下一程序
    If Me.Text5(7).Text <> "" Then
        'Modified by Morgan 2014/11/20 外專改回舊規則
        ''Added by Morgan 2014/10/29
        ' If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
        '    strTmp(0) = PUB_GetOurDeadline(Text5(7).Text)
        ' Else
        ' 'end 2014/10/29
        
         'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
         If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
            strTmp(0) = PUB_GetFCPOurDeadline(Text5(7), 2, , strAgreeOnDate)
         Else
         'end 2019/7/11
         
            strTmp(0) = CompDate(2, -2, TransDate(Text5(7).Text, 2))
            
         End If 'Added by Morgan 2019/7/11
            
        'End If 'Added by Morgan 2014/10/29
        'end 2014/11/20
         
       'edit by nickc 2007/02/02 不用 dll 了
       'strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
          "VALUES ('" & strReceiveNo & "','" & pA(1) & "','" & pA(2) & "','" & pA(3) & _
          "','" & pA(4) & "','" & PUB_GetFCPSalesNo(pA(1), pA(2), pA(3), pA(4)) & "'," & 年費 & "," & CNULL(strTmp(0)) & "," & _
          CNULL(TransDate(Text5(7), 2)) & "," & objPublicData.GetNextProgressNo & ")"
       '2008/10/13 modify by sonia
       'strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
          "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
          "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & CNULL(strTmp(0)) & "," & _
          CNULL(TransDate(Text5(7), 2)) & "," & GetNextProgressNo & ")"
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
      
'Modify by Morgan 2011/6/13 改抓共用函數
'      If strExc(9) = "Y33944" Or strExc(9) = "Y48840" Or strExc(9) = "Y48196" Or _
'         strExc(9) = "Y20624" Or strExc(9) = "Y21099" Then
'
'      '2008/11/26 END
'         strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,np15,NP22) " & _
'            "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
'            "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & CNULL(strTmp(0)) & "," & _
'            CNULL(TransDate(Text5(7), 2)) & ",'信函要傳真;'," & GetNextProgressNo & ")"
'       '2009/8/4 ADD BY SONIA Y49083年費備註
'       'Modify by Morgan 2011/3/22 改先存變數才不用重複抓相同資料
'       'ElseIf Mid(PUB_GetReceiver(pa(1), pa(2), pa(3), pa(4), "605", "1"), 1, 6) = "Y49083" Then
'       ElseIf strExc(9) = "Y49083" Then
'         strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,np15,NP22) " & _
'            "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
'            "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & CNULL(strTmp(0)) & "," & _
'            CNULL(TransDate(Text5(7), 2)) & ",'只需銀龍加蓋年費回傳章;'," & GetNextProgressNo & ")"
'       '2009/8/4 END
'
'      'Add by Morgan 2011/3/22 2011.03.19 指示信--Susan
'      ElseIf strExc(9) = "Y30011" Then
'         strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,np15,NP22) " & _
'            "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
'            "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & CNULL(strTmp(0)) & "," & _
'            CNULL(TransDate(Text5(7), 2)) & ",'年費函需以EMail傳送,不寄紙本;'," & GetNextProgressNo & ")"
'
'       Else
'         strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
'            "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
'            "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & CNULL(strTmp(0)) & "," & _
'            CNULL(TransDate(Text5(7), 2)) & "," & GetNextProgressNo & ")"
'       End If
'       '2008/10/13 end
      'Modified by Morgan 2012/6/4 +pa26
      'Modified by Morgan 2013/9/11 改抓設定檔
      'strExc(1) = PUB_Get605Memo(strExc(9), ChangeCustomerL(pa(26)), pa(1) & pa(2) & pa(3) & pa(4))
      'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
      'strExc(1) = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), ChangeCustomerL(pa(26)))
      strExc(1) = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "605", strExc(9), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30))
      
      'Remove by Morgan 2011/10/19 改控制代理人 Y47735xxx 所有案件 --譚文容
      ''Add by Morgan 2011/9/7 --譚文容
      ''催年費函需另cc:Y47740
      'Select Case pa(1) & pa(2) & pa(3) & pa(4)
      '   'Modify by Morgan 2011/10/6 +FCP031045000--譚文容
      '   Case "FCP032841000", "FCP031045000"
      '      strExc(1) = "催年費函需另cc:Y47740;" & strExc(1)
      'End Select
      ''end 2011/9/7
      'Modify By Sindy 2021/8/17 + ,NP23
      strExc(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,np15,NP22,NP23) " & _
         "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
         "','" & pa(4) & "','" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "'," & 年費 & "," & CNULL(strTmp(0)) & "," & _
         CNULL(TransDate(Text5(7), 2)) & ",'" & ChgSQL(strExc(1)) & "'," & GetNextProgressNo & "," & CNULL(strAgreeOnDate, True) & ")"
'end 2011/6/13
      
       '911105 nick transation
       cnnConnection.Execute strExc(3)
    End If
    
   PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130 'Add by Morgan 2009/3/20
   
    'Added by Lydia 2015/02/26 若已開請款單則換承辦人或核稿人時發Mail通知靜芳
   If m_CP60 > "X" Then
      'Modified by Lydia 2019/10/17 本所案號+"-"
      'PUB_PointReAssignInform Text1 & Text2 & Text3 & Text4, m_CP60, Text5(5).Tag, Text5(5).Text
      PUB_PointReAssignInform pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)), m_CP60, Text5(5).Tag, Text5(5).Text
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
   
   'Added by Lydia 2020/08/17 Transaction將發文和年證費請款函一併包入 (P.S 修改條件,請一併更新FormSave的.CommitTrans)
   If Not (m_bolBeCalled = False And m_CP60 = "" And Val(m_CP82) = 0) Then
      cnnConnection.CommitTrans
   End If
   'end 2020/08/17

'911105 nick
   Exit Function
CheckingErr:
   Resume
   cnnConnection.RollbackTrans
   FormSave = False
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
   MoveFormToCenter Me
   intWhere = 國外_FC
   
   'Added by Morgan 2012/9/13
   '考慮被批次發文呼叫情形
   If m_bolBeCalled Then
      Text1 = m_CP01
      Text2 = m_CP02
      Text3 = m_CP03
      Text4 = m_CP04
      strReceiveNo = m_CP09
   Else
   'end 2012/9/13
      With frm060104_1
         Text1 = .Text1
         Text2 = .Text2
         Text3 = .Text3
         Text4 = .Text4
         strReceiveNo = .Tag
      End With
   End If 'Added by Morgan 2012/9/13
   
   'Add by Morgan 2005/8/4
   ReDim pa(TF_PA)
   ReadPatent
   Label2(0) = strReceiveNo
   
   'Add by Morgan 2005/8/4
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   'Modify by Morgan 2008/3/17
   'PUB_SetOurAgent lstNameAgent, pa(), m_CP110
   If pa(143) = "N" Then
      lstNameAgent.Enabled = False
      m_CP110 = ""
   Else
      'Modified by Morgan 2020/3/20 +CP10
      PUB_SetOurAgent lstNameAgent, pa(), m_CP110, m_CP10, True
   End If
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1300
   lstNameAgent.Width = 1300

   Combo1.ListIndex = 0
    m_blnClkChgEvnBtn = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2020/08/17
   Set frm060104_a = Nothing
End Sub

Private Sub ReadPatent()
 Dim Lbl As Object, txt As Object, i As Integer
 Dim strTmp(0 To 5) As String, varTmp As Variant, strTmp1(0 To 5) As String
 '92.9.15 add by sonia
 Dim StrSQLa As String
 Dim rsA As New ADODB.Recordset
 '92.9.15 end
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
            Label2(9) = pa(72)
            If pa(76) <> "" Then Text5(6) = pa(76): ChgType 9
      
            strTmp1(0) = strReceiveNo
            
            For i = 1 To 4
               strTmp1(i) = pa(i)
            Next
            If GetMoneyDate(Val(pa(8)), pa(9), strTmp1, strCaseFee(1), strCaseFee(2)) = True Then
               'Remove by Morgan 2006/4/24 GetMoneyDate已改
'               'Modify by Morgan 2004/12/14 舊法新型專用期12年
'               If pA(9) = "000" And pA(8) = "2" And Val(pA(14)) > 0 And Val(pA(14)) < 930701 Then
'                  strCaseFee(2) = "1,2,3,4,5,6,7,8,9,10,11,12"
'               End If
            End If
            '2008/5/15 ADD BY SONIA 若尚未發證則依公式計算專用期止日
            If pa(25) = "" Then
               If GetMoneyDate(Val(pa(8)) + 10, pa(9), strTmp1, strCaseFee(1), strCaseFee(2), pa(25)) Then   '抓專用期起止日
                   pa(25) = ChangeWStringToTString(pa(25))
                   If GetMoneyDate(pa(8), pa(9), strTmp1, strCaseFee(1), strCaseFee(2)) = True Then
                      'Modify by Morgan 2004/12/14 舊法新型專用期12年
                      If pa(9) = "000" And pa(8) = "2" And Val(pa(14)) > 0 And Val(pa(14)) < 930701 Then
                         strCaseFee(2) = "1,2,3,4,5,6,7,8,9,10,11,12"
                      End If
                   End If
               End If
            End If
            '2008/5/15 END
         End If
      Case "FG"
      
   End Select
   'Modify By Cheng 2002/07/04
'   strExc(0) = "select cp13,st02,cp06,cp27,cp14,cp64,CP07 from caseprogress" & _
'      ",staff where cp09='" & strReceiveNo & "' and cp13=st01(+)"
'Modified by Lydia 2015/02/26 +cp60
   'Modified by Lydia 2018/09/11 +cp118,cp82
   'Modified by Sindy 2019/12/13 +,cp53,cp54
   'Modified by Lydia 2020/06/23 +CP148
   'Modified by Lydia 2020/08/17 +GetEmailFlag(CP01||CP02||CP03||CP04) as eFlag
   strExc(0) = "select cp13,st02,cp06,cp27,cp14,cp64,CP07,CP10,CP110,CP60, " & _
                    "CP142,CP118,CP82,cp53,cp54,CP148,GetEmailFlag(CP01||CP02||CP03||CP04) as eFlag,cp164 " & _
                    "from caseprogress,staff where cp09='" & strReceiveNo & "' and cp13=st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
   If intI = 1 Then
      m_CP110 = "" & .Fields("CP110")
      m_CP142 = "" & .Fields("CP142") 'Add By Sindy 2015/12/17
      m_CP164 = "" & .Fields("CP164") 'Add By Sindy 2021/4/20
      m_CP14 = "" & .Fields("CP14") 'Add by Sindy 2016/11/16
      
      'Add By Sindy 2019/12/13 當601領證及605年費key繳費年度而產生電子送件申請書時，將key的年度自動帶到發文作業的年度。
      Text5(0) = "" & .Fields("cp53")
      Text5(1) = "" & .Fields("cp54")
      '2019/12/13 END
      
      If Not IsNull(.Fields(0)) Then strSales = .Fields(0)
      If Not IsNull(.Fields(1)) Then Label2(1) = .Fields(1)
      If Not IsNull(.Fields(2)) Then Label2(2) = TransDate(.Fields(2), 1)
      If Not IsNull(.Fields(3)) Then
         Text5(4) = TransDate(.Fields(3), 1)
      Else
         Text5(4) = strSrvDate(2)
      End If
      Text5(4).Tag = Text5(4).Text 'Added by Lydia 2020/08/26 預設發文日
      
      If Not IsNull(.Fields(4)) Then Text5(5) = .Fields(4): ChgType (5 + 3)
      'Added by Lydia 2015/02/26
      Text5(5).Tag = Text5(5).Text
      If Not IsNull(.Fields("CP60")) Then
         m_CP60 = .Fields("CP60")
      Else
         m_CP60 = ""
      End If
      'end 2015/02/26
      
      If Not IsNull(.Fields(5)) Then
         Text5(8) = .Fields(5)
         'Add By Sindy 2015/9/23
         If InStr(Text5(8), Trim(Label11.Caption)) > 0 Then
            Text13.Enabled = False
         End If
         '2015/9/23 END
      End If
      
      If Not IsNull(.Fields(6)) Then strTemp = TransDate(.Fields(6), 1)
      'Add By Cheng 2002/07/04
      m_CP07 = "" & .Fields(6).Value
      m_CP10 = "" & .Fields(7).Value
      
      'Added by Lydia 2018/09/11
      m_CP118 = "" & .Fields("cp118")  '電子送件
      If m_CP118 <> "" Then txtCP118.Text = "Y"
      
      m_CP82 = "" & .Fields("cp82") '發文時間
      'end 2018/09/11
      textCP148 = "" & .Fields("cp148") 'Added by Lydia 2020/06/23 預設為進度檔的設定
      'Added by Lydia 2020/08/17  Email維護
      m_eFlag = "" & .Fields("eFlag")
      lblEmail.Visible = False: txtEmail.Visible = False
      If m_bolBeCalled = False And m_CP60 = "" And Val(m_CP82) = 0 Then
          lblEmail.Visible = True: txtEmail.Visible = True
      End If
      'end 2020/08/17
   End If
   End With
   
'Modify by Morgan 2008/10/8 因NP的年費期限可能會延期故需重算原期限
'   '92.9.15 add by sonia
'   '取得下一程序的法定期限
'   StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " And Np07='" & m_CP10 & "' Order By NP09 Desc "
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsA.RecordCount > 0 Then
'       m_strNP09 = "" & ChangeWStringToTString(rsA("NP09").Value)
'   Else
'       m_strNP09 = "" & m_CP07
'   End If
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
'   m_strNP09_1 = m_strNP09
   m_strNP09_1 = PUB_GetNextFeeDate(pa)
'end 2008/10/8
    
   '若法定期限為假日時, 抓大於法定期限最近的工作天
   If m_strNP09_1 <> "" Then
      m_strNP09_1 = DBDATE(PUB_GetLawDay(DBDATE(m_strNP09_1)))
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
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 0 '發文日
'         If ChkDate(Text5(4)) Or Val(Text5(4)) > Val(strSrvDate(2)) Then
'            ChgType = True
'         End If
         If Not ChkDate(Text5(4)) Then
         ElseIf Val(Text5(4).Text) > PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1))) Then
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
      Case 8
         'ADD BY SONIA 2015/9/21 單筆發文承辦人為外專程序時,改為操作人員,整批發文不改
         If m_bolBeCalled = False Then Text5(5) = GetFCPUser(Text5(5))
         'END 2015/9/21
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(Text5(5), strTempName) Then
         If ClsPDGetStaff(Text5(5), strTempName) Then
            Label2(10) = strTempName
            ChgType = True
         End If
      Case 9
         'edit by nickc 2007/02/05 不用 dll 了
         'If objLawDll.LawGetName(Text5(6), strTempName) Then
         If ClsLawLawGetName(Text5(6), strTempName) Then
            Label2(11) = strTempName
            ChgType = True
         End If
   End Select
End Function

Private Sub Text5_GotFocus(Index As Integer)
   InverseTextBox Text5(Index)
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 2, 3
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 5, 6
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Text5_LostFocus(Index As Integer)
Select Case Index
Case 4 '發文日
   'Add By Cheng 2002/07/04
   If m_CP07 <> "" Then
      '判斷發文日是否大於法定期限(若法定期限不為工作天, 則為其下一個工作天)
      '92.9.15 modify by sonia
      'If DBDATE(Me.Text5(4).Text) > DBDATE(PUB_GetLawDay(DBDATE(m_CP07))) Then
      'Modify by Morgan 2008/10/8 需用原年費期限判斷
      'If DBDATE(Text5(4).Text) > IIf(DBDATE(m_strNP09) >= DBDATE(m_strNP09_1), DBDATE(m_strNP09), DBDATE(m_strNP09_1)) Then
      If DBDATE(Text5(4).Text) > m_strNP09_1 Then
      'end 2008/10/8
      '92.9.15 end
         '費用雙倍
         Me.Text5(2).Text = "Y"
      Else
         '取消費用雙倍
         'Modify by Morgan 2005/3/1 已設雙倍不可清除
         'Me.Text5(2).Text = ""
      End If
   End If
End Select
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
 Dim i As Integer, bolChk As Boolean, varTmp As Variant, varTmpNICK As Variant, TMPnick060104 As Integer
'Add By Cheng 2002/10/29
Dim strNextFeeDate As String '下次繳費日

   Select Case Index
      Case 0, 1
         If Text5(Index) <> "" Then
            If Index = 0 Then
               If pa(72) = "" Then
                  If Text5(0) <> "1" Then
                     MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
                     Cancel = True
                  End If
               Else
                  varTmpNICK = Split(pa(72), ",")
                  For TMPnick060104 = UBound(varTmpNICK) To 0 Step -1
                     If Trim(varTmpNICK(TMPnick060104)) <> "" Then
                        Exit For
                     End If
                  Next TMPnick060104
                  'NICKC    900606
                  'If Text5(0) <> Right(pa(72), 1) + 1 Then
                  If Text5(0) <> Val(varTmpNICK(TMPnick060104)) + 1 Then
                     MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
                     Cancel = True
                  End If
               End If
            ElseIf Index = 1 Then
               If ChkRange(Text5(0), Text5(1), "繳費年度") = True Then
                  For i = Text5(0) To Text5(1)
                     If InStr(pa(72), Format(i)) > 0 Then
                        bolChk = True
                        Exit For
                     End If
                  Next
                  If bolChk = True Then
                     MsgBox "繳費年度重覆，請查明後再輸入 !", vbCritical
                     Cancel = True
                  Else
                     varTmp = Split(strCaseFee(2), ",")
                     'Modify by Morgan 2006/4/24 改判斷繳費迄年是否繳超過專用期
                     'If Text5(1) > UBound(varTmp) + 1 Then
                     strExc(0) = TransDate(CompDate(0, Text5(1) - 1, strCaseFee(1)), 1)
                     If Val(strExc(0)) > Val(pa(25)) Then
                        MsgBox "繳費年度大於應繳年度，請查明後再輸入 !", vbCritical
                        Cancel = True
                     ElseIf Text5(1) = UBound(varTmp) + 1 Then
                        Text5(7).Text = ""
                     Else
                        'Modify By Cheng 2002/10/29
                        '原算出的下次繳費日多一天
'                        Text5(7).Text = TransDate(CompDate(0, Val(varTmp(Val(Text5(1).Text) - 1)), strCaseFee(1)), 1)
                        strNextFeeDate = CompDate(0, Val(varTmp(Val(Text5(1).Text) - 1)), strCaseFee(1))
                        'Modify By Cheng 2002/11/22
                        '避免計算下次繳費日時出錯
                        If strNextFeeDate <> "" Then
                            'Modify By Cheng 2002/12/04
'                            Text5(7).Text = TransDate(Replace("" & DateSerial(Left(strNextFeeDate, 4), Mid(strNextFeeDate, 5, 2), Right(strNextFeeDate, 2) - 1), "/", ""), 1)
                            Text5(7).Text = ChangeWDateStringToTString(DateSerial(Left(strNextFeeDate, 4), Mid(strNextFeeDate, 5, 2), Right(strNextFeeDate, 2) - 1))
                        Else
                            Text5(7).Text = ""
                        End If
                        'Add By Cheng 2004/01/05
                        '若計算出的下次繳費年度>=專用期止日, 則清空下次繳費日(存檔時不產生下一程序)
                        If Me.Text5(7).Text <> "" Then
                            If DBDATE(Me.Text5(7).Text) >= DBDATE(pa(25)) Then
                                Me.Text5(7).Text = ""
                            End If
                        End If
                        'End
                     End If
                  End If
               Else
                  Cancel = True
               End If
            End If
         Else
            MsgBox "年度不可空白 !", vbCritical
            Cancel = True
         End If
      Case 2
        'Add By Cheng 2003/01/03
        '若案件進度檔有法定期限
        If strTemp <> "" Then
            'Modify By Cheng 2002/07/04
    '         If Text5(4) > strTemp Then
            '92.9.15 modify by sonia
            'If DBDATE(Text5(4)) > DBDATE(PUB_GetLawDay(DBDATE(strTemp))) Then
            'Modify by Morgan 2008/10/8 需用原年費期限判斷
            'If DBDATE(Text5(4)) > IIf(DBDATE(m_strNP09) >= DBDATE(m_strNP09_1), DBDATE(m_strNP09), DBDATE(m_strNP09_1)) Then
            If DBDATE(Text5(4)) > m_strNP09_1 Then
            'end 2008/10/8
            '92.9.15 end
               If Text5(Index) <> "Y" Then
                  MsgBox "發文日大於法定期限則此欄必須為 Y !", vbCritical
                  Text5(Index) = "Y"
               End If
            End If
        '若案件進度檔無法定期限
        Else
            MsgBox "案件進度檔無法定期限資料!!!", vbExclamation + vbOKOnly
            Cancel = True
        End If
      Case 4 '發文日
         If Text5(Index) <> "" Then
            If Not ChgType(0) Then
                  Cancel = True
            'Added by Lydia 2018/09/11 當發文日有改時
            Else
                  If Text5(Index).Tag <> Text5(Index) Then
                        Text5(Index).Tag = Text5(Index)
                        txtPayToday.Text = Pub_FcpSetPayToday("1", Text5(Index).Text, txtCP118.Text)
                  End If
            'end 2018/09/11
            End If
         Else
            MsgBox "發文日不可空白 !", vbCritical
            Cancel = True
         End If
      Case 5
         If Text5(Index) <> "" Then
            If Not ChgType(Index + 3) Then Cancel = True
         Else
            MsgBox "承辦人不可空白 !", vbCritical
            Cancel = True
         End If
      Case 6
         If Text5(Index) <> "" Then
            If Not ChgType(Index + 3) Then Cancel = True
         End If
         'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
         If Cancel = False Then
            If PUB_CheckStatus(Text5(Index).Text) = False Then Cancel = True
         End If
   End Select
   If Cancel = True Then TextInverse Text5(Index)
End Sub

'Add By Cheng 2002/05/21
Private Function CheckDataValid() As Boolean
Dim i As Integer, bolChk As Boolean, varTmp As Variant, varTmpNICK As Variant, TMPnick060104 As Integer

CheckDataValid = False
'檢查繳納年費年數
If Me.Text5(0).Text = "" Then
   MsgBox "年度不可空白 !", vbCritical
   Me.Text5(0).SetFocus
   Text5_GotFocus 0
   Exit Function
End If
If pa(72) = "" Then
   If Text5(0) <> "1" Then
      MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
      Me.Text5(0).SetFocus
      Text5_GotFocus 0
      Exit Function
   End If
Else
   varTmpNICK = Split(pa(72), ",")
   For TMPnick060104 = UBound(varTmpNICK) To 0 Step -1
      If Trim(varTmpNICK(TMPnick060104)) <> "" Then
         Exit For
      End If
   Next TMPnick060104
   'NICKC    900606
   'If Text5(0) <> Right(pa(72), 1) + 1 Then
   If Text5(0) <> Val(varTmpNICK(TMPnick060104)) + 1 Then
      MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
      Me.Text5(0).SetFocus
      Text5_GotFocus 0
      Exit Function
   End If
End If
If Me.Text5(1).Text = "" Then
   MsgBox "年度不可空白 !", vbCritical
   Me.Text5(1).SetFocus
   Text5_GotFocus 1
   Exit Function
End If
If ChkRange(Text5(0), Text5(1), "繳費年度") = True Then
   For i = Text5(0) To Text5(1)
      If InStr(pa(72), Format(i)) > 0 Then
         bolChk = True
         Exit For
      End If
   Next
   If bolChk = True Then
      MsgBox "繳費年度重覆，請查明後再輸入 !", vbCritical
      Me.Text5(1).SetFocus
      Text5_GotFocus 1
      Exit Function
   Else
      varTmp = Split(strCaseFee(2), ",")
      'Modify by Morgan 2006/4/24 改判斷繳費迄年是否繳超過專用期
      'If Text5(1) > UBound(varTmp) + 1 Then
      strExc(0) = TransDate(CompDate(0, Text5(1) - 1, strCaseFee(1)), 1)
      If Val(strExc(0)) > Val(pa(25)) Then
      '2006/4/24 end
         MsgBox "繳費年度大於應繳年度，請查明後再輸入 !", vbCritical
         Me.Text5(1).SetFocus
         Text5_GotFocus 1
         Exit Function
      ElseIf Text5(1) = UBound(varTmp) + 1 Then
         Text5(7).Text = ""
      Else
         Text5(7).Text = TransDate(CompDate(0, Val(varTmp(Val(Text5(1).Text) - 1)), strCaseFee(1)), 1)
        'Add By Cheng 2004/01/05
        '若計算出的下次繳費年度>=專用期止日, 則清空下次繳費日(存檔時不產生下一程序)
        If Me.Text5(7).Text <> "" Then
            If DBDATE(Me.Text5(7).Text) >= DBDATE(pa(25)) Then
                Me.Text5(7).Text = ""
            End If
        End If
        'End
      End If
   End If
Else
   Me.Text5(1).SetFocus
   Text5_GotFocus 1
   Exit Function
End If
'檢查費用是否要雙倍
If strTemp <> "" Then
    '92.9.15 MODIFY BY SONIA
    'If DBDATE(Text5(4)) > DBDATE(PUB_GetLawDay(DBDATE(strTemp))) Then
    'Modify by Morgan 2008/10/8 需用原年費期限判斷
    'If DBDATE(Text5(4)) > IIf(DBDATE(m_strNP09) >= DBDATE(m_strNP09_1), DBDATE(m_strNP09), DBDATE(m_strNP09_1)) Then
    If DBDATE(Text5(4)) > m_strNP09_1 Then
    'end 2008/10/8
    '92.9.15 END
       If Text5(2) <> "Y" Then
          MsgBox "發文日大於法定期限則此欄必須為 Y !", vbCritical
          Text5(2) = "Y"
       End If
    End If
Else
    MsgBox "案件進度檔無法定期限資料!!!", vbExclamation + vbOKOnly
    Exit Function
End If
'檢查發文日
If Text5(4) <> "" Then
   If Not ChgType(0) Then
      Me.Text5(4).SetFocus
      Text5_GotFocus 4
      Exit Function
   End If
Else
   MsgBox "發文日不可空白 !", vbCritical
   Me.Text5(4).SetFocus
   Text5_GotFocus 4
   Exit Function
End If
'檢查承辦人
If Text5(5) <> "" Then
   If Not ChgType(5 + 3) Then
      Me.Text5(5).SetFocus
      Text5_GotFocus 5
      Exit Function
   End If
Else
   MsgBox "承辦人不可空白 !", vbCritical
   Me.Text5(5).SetFocus
   Text5_GotFocus 5
   Exit Function
End If

CheckDataValid = True
End Function

'Modify By Cheng 2003/01/01
''Add By Cheng 2002/07/04
''取得案件收費表的規費
'Private Function GetCF08(strCF01 As String, strCF02 As String, strCF03 As String) As String
'Dim rsA As New ADODB.Recordset
'Dim strSQLA As String
'
'GetCF08 = "0"
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'strSQLA = "Select CF08 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "' AND CF08 IS NOT NULL"
'rsA.CursorLocation = adUseClient
'rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'   GetCF08 = rsA.Fields(0).Value
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'End Function

'Add By Cheng 2002/12/31
'計算相關費用
Private Sub GetPatentYearFee( _
    strYF01 As String, strYF02 As String, strYF03 As String, _
    strYF04 As String, strYF05From As String, strYF05To As String, blnDouble As Boolean)
'strYF01  申請國家
'strYF02  專利種類
'strYF03  代理人
'strYF04  案件性質
'strYF05From  起始年度
'strYF05To  終止年度
'blnDouble  規費是否雙倍

PUB_GetPatentYearFee strYF01, strYF02, strYF03, strYF04, strYF05From, strYF05To, blnDouble, m_CP81, pa(14), Text5(4), m_strOfficalFee, m_strServiceFee, m_lngDisc, m_bol2ndY

'Removed by Morgan 2012/10/12 改呼叫共用函數
'Dim rsA As New ADODB.Recordset
'Dim StrSQLa As String
'Dim ii As Integer
'Dim iYear As Integer '繳費年度
'Dim lngOfficalFee As Long
'
'    m_strOfficalFee = 0
'    m_strServiceFee = 0
'    m_strPoints = 0
'    m_lngDisc = 0
'
'    ii = 1
'    '取得案件性質為年費的相關費用
'    'Modify By Cheng 2003/01/02
''    strSQLA = "Select * From PatentYearFee Where YF01='" & strYF01 & "' AND YF02='" & strYF02 & "' AND YF03='" & strYF03 & "' AND YF04='" & 年費 & "' AND YF05>=" & Val(strYF05From) & " AND YF05<=" & Val(strYF05To)
'    StrSQLa = "Select * From PatentYearFee Where YF01='" & strYF01 & "' AND YF02='" & strYF02 & "' AND YF03='" & strYF03 & "' AND YF04='" & 年費 & "' AND YF05>=" & Val(strYF05From) & " AND YF05<=" & Val(strYF05To) & " Order By YF05 "
'    rsA.CursorLocation = adUseClient
'    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'    While Not rsA.EOF
'        lngOfficalFee = Val(rsA.Fields("YF07").Value)
'
'        'Add by Morgan 2008/1/25 考慮個人可減免
'         iYear = Val(rsA.Fields("YF05").Value)
'         If m_CP81 = "Y" Then
'            If iYear >= 1 And iYear <= 3 Then
'               lngOfficalFee = lngOfficalFee - 800
'               m_lngDisc = m_lngDisc + 800
'            ElseIf iYear >= 4 And iYear <= 6 Then
'               lngOfficalFee = lngOfficalFee - 1200
'               m_lngDisc = m_lngDisc + 1200
'            End If
'         End If
'
'         '起始那年年費是否雙倍
'         If blnDouble = True Then
'            If ii = 1 Then
'               lngOfficalFee = lngOfficalFee * 2
'               m_lngDisc = m_lngDisc * 2 'Add by Morgan 2008/1/25
'            End If
'         End If
'
'        m_strOfficalFee = Val(m_strOfficalFee) + lngOfficalFee
'        m_strServiceFee = Val(m_strServiceFee) + Val(rsA.Fields("YF06").Value)
'        rsA.MoveNext
'        'Add By Cheng 2003/01/02
'        ii = ii + 1
'    Wend
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
    
    m_strPoints = Val(m_strServiceFee) / 1000
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean

   Dim objTxt As Object
   Dim ii As Integer
   Dim Cancel As Boolean

   TxtValidate = False
   For Each objTxt In Me.Text5
      If objTxt.Enabled = True Then
         Cancel = False
         Text5_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next

   'Add by Morgan 2005/8/4
   If lstNameAgent.Enabled = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If
   
   '取得領證及繳年費相關費用
   GetPatentYearFee pa(9), pa(8), "Y00000000", m_CP10, Me.Text5(0).Text, Me.Text5(1).Text, IIf(Me.Text5(2).Text = "Y", True, False)
   If m_bolBeCalled Then
      If m_CP84 <> m_strOfficalFee Then
         MsgBox "CSV上傳規費[" & m_CP84 & "]與系統計算[" & m_strOfficalFee & "]不符，請確認！", vbExclamation
         Exit Function
      End If
   End If
   
   'Add By Sindy 2015/12/17 檢查是否有指定送件日期,若有不可小於指定日期送件
   If m_CP142 <> "" Then
      'Modify By Sindy 2021/11/11 淑華說之後可以含當天發文
      'If m_CP142 >= strSrvDate(1) Then
      If m_CP142 > strSrvDate(1) Then
         'Add By Sindy 2021/4/20
         'Modify By Sindy 2021/10/20 + 3.之後
         If ((m_CP164 = "1" Or m_CP164 = "") And m_CP142 > strSrvDate(1)) Or _
            m_CP164 = "3" Then '1.當天 3.之後
         '2021/4/20 END
            MsgBox "有指定送件日期（" & ChangeWStringToTDateString(m_CP142) & "），不可提前送件!!!"
            Exit Function
         End If
      End If
   End If
   '2015/12/17 END
   
   'Added by Lydia 2019/01/14
   If txtCP118 = "Y" Then
      If txtPayToday = "" Then
         MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
         txtPayToday.SetFocus
         Exit Function
      End If
   End If
   'end 2019/01/14
      
   TxtValidate = True
End Function

'Add by Morgan 2005/8/4
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   
   If lstNameAgent.Enabled = False Then Exit Sub 'Add by Morgan 2008/3/17
   
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
'Add by Morgan 2008/3/19 申請書的申請人檢查
'個案設年費申請人不出名
Private Function CheckStop1() As Boolean
   strExc(0) = "select '" & pa(26) & "' CuNo,pa01,pa02,pa03,pa04 from patent,caseprogress" & _
      " where pa01='FCP' and '" & ChangeCustomerL(pa(26)) & "' in (pa26,pa27,pa28,pa29,pa30)" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp05>20080310" & _
      " and cp10 in ('101','102','103','105') and rownum<2"
      
   If pa(27) <> "" Then
      strExc(0) = strExc(0) & " union select '" & pa(27) & "' CuNo,pa01,pa02,pa03,pa04" & _
         " from patent,caseprogress where pa01='FCP' and '" & ChangeCustomerL(pa(27)) & "' in (pa26,pa27,pa28,pa29,pa30)" & _
         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp05>20080310" & _
         " and cp10 in ('101','102','103','105') and rownum<2"
   End If
   If pa(28) <> "" Then
      strExc(0) = strExc(0) & " union select '" & pa(28) & "' CuNo,pa01,pa02,pa03,pa04" & _
         " from patent,caseprogress where pa01='FCP' and '" & ChangeCustomerL(pa(28)) & "' in (pa26,pa27,pa28,pa29,pa30)" & _
         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp05>20080310" & _
         " and cp10 in ('101','102','103','105') and rownum<2"
   End If
   If pa(29) <> "" Then
      strExc(0) = strExc(0) & " union select '" & pa(29) & "' CuNo,pa01,pa02,pa03,pa04" & _
         " from patent,caseprogress where pa01='FCP' and '" & ChangeCustomerL(pa(29)) & "' in (pa26,pa27,pa28,pa29,pa30)" & _
         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp05>20080310" & _
         " and cp10 in ('101','102','103','105') and rownum<2"
   End If
   If pa(30) <> "" Then
      strExc(0) = strExc(0) & " union select '" & pa(30) & "' CuNo,pa01,pa02,pa03,pa04" & _
         " from patent,caseprogress where pa01='FCP' and '" & ChangeCustomerL(pa(30)) & "' in (pa26,pa27,pa28,pa29,pa30)" & _
         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp05>20080310" & _
         " and cp10 in ('101','102','103','105') and rownum<2"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(1) = "本案有設定年費申請人不出名，但下列申請人 97/3/10 以後有收新案：" & vbCrLf
      With RsTemp
      Do While Not .EOF
         strExc(1) = strExc(1) & vbCrLf & "申請人 [ " & .Fields("CuNo") & " ] ，新案 [ " & .Fields("pa01") & "-" & .Fields("pa02") & "-" & .Fields("pa03") & "-" & .Fields("pa04") & " ]"
         .MoveNext
      Loop
      End With
      strExc(1) = strExc(1) & vbCrLf & vbCrLf & "是否要先清除年費申請人不出名設定？"
      If MsgBox(strExc(1), vbYesNo + vbDefaultButton1) = vbYes Then
         CheckStop1 = True
      End If
   End If

End Function
'多人申請
Private Function CheckStop2() As Boolean
   strExc(0) = "select cu01||cu02 CuNo,pa01,pa02,pa03,pa04" & _
      " from customer,patent,caseprogress where " & ChgCustomer(pa(26)) & " and cu123='N'" & _
      " and pa01='FCP' and cu01||cu02 in (pa26,pa27,pa28,pa29,pa30)" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp05>20080310" & _
      " and cp10 in ('101','102','103','105') and rownum<2"
      
   strExc(0) = strExc(0) & " union select cu01||cu02 CuNo,pa01,pa02,pa03,pa04" & _
      " from customer,patent,caseprogress where " & ChgCustomer(pa(27)) & " and cu123='N'" & _
      " and pa01='FCP' and cu01||cu02 in (pa26,pa27,pa28,pa29,pa30)" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp05>20080310" & _
      " and cp10 in ('101','102','103','105') and rownum<2"
      
   If pa(28) <> "" Then
      strExc(0) = strExc(0) & " union select cu01||cu02 CuNo,pa01,pa02,pa03,pa04" & _
      " from customer,patent,caseprogress where " & ChgCustomer(pa(28)) & " and cu123='N'" & _
      " and pa01='FCP' and cu01||cu02 in (pa26,pa27,pa28,pa29,pa30)" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp05>20080310" & _
      " and cp10 in ('101','102','103','105') and rownum<2"
   End If
   If pa(29) <> "" Then
      strExc(0) = strExc(0) & " union select cu01||cu02 CuNo,pa01,pa02,pa03,pa04" & _
      " from customer,patent,caseprogress where " & ChgCustomer(pa(29)) & " and cu123='N'" & _
      " and pa01='FCP' and cu01||cu02 in (pa26,pa27,pa28,pa29,pa30)" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp05>20080310" & _
      " and cp10 in ('101','102','103','105') and rownum<2"
   End If
   If pa(30) <> "" Then
      strExc(0) = strExc(0) & " union select cu01||cu02 CuNo,pa01,pa02,pa03,pa04" & _
      " from customer,patent,caseprogress where " & ChgCustomer(pa(30)) & " and cu123='N'" & _
      " and pa01='FCP' and cu01||cu02 in (pa26,pa27,pa28,pa29,pa30)" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp05>20080310" & _
      " and cp10 in ('101','102','103','105') and rownum<2"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(1) = "本案下列申請人有設定不辦理重新委任，但 97/3/10 以後有收新案：" & vbCrLf
      With RsTemp
      Do While Not .EOF
         strExc(1) = strExc(1) & vbCrLf & "申請人 [ " & .Fields("CuNo") & " ] ，新案 [ " & .Fields("pa01") & "-" & .Fields("pa02") & "-" & .Fields("pa03") & "-" & .Fields("pa04") & " ]"
         .MoveNext
      Loop
      End With
      strExc(1) = strExc(1) & vbCrLf & vbCrLf & "是否要先清除申請人不辦理重新委任註記？"
      If MsgBox(strExc(1), vbYesNo + vbDefaultButton1) = vbYes Then
         CheckStop2 = True
      End If
   End If
End Function
'end 2008/3/19

Public Function Process() As Boolean
 'Added by Lydia 2018/09/11
 Dim strFilePath As String '記錄智慧局收文文號
 Dim strNewCP64 As String '保留進度備註
 'end 2018/09/11
 
   'Add By Cheng 2002/05/21
   If CheckDataValid = False Then Exit Function
   If TxtValidate = False Then Exit Function
   'Add by Morgan 2008/3/19 第三人繳年費或部分申請人出名案件檢查若97.3.10以後有收文新案時提醒
   '個案有設定年費申請人不出名
   If pa(143) = "N" Then
      If CheckStop1 = True Then
         Exit Function
      End If
   '多人申請
   ElseIf pa(27) <> "" Then
      If CheckStop2 = True Then
         Exit Function
      End If
   End If
   'end 2008/3/19
   'Add by Morgan 2009/4/28
   'Modified by Morgan 2016/5/16 +是傳否電子送件參數
   If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text5(4), , IIf(txtCP118 <> "", True, False)) = False Then
      Exit Function
   End If
   
   'Add by Morgan 2012/9/13
   'Modified by Lydia 2018/09/11 電子送件
   'If txtCP118 <> "" Then
   '   m_CP123s = ""
   'Else
   ''end 2012/9/13
   strNewCP64 = Text5(8)
   If txtCP118 = "Y" Then
       m_CP123s = ""
       If m_bolBeCalled = False Then '排除整批
            strExc(0) = InputBox("請輸入智慧局收文文號!!")
            If strExc(0) = "" Then
               Exit Function
            Else
               strFilePath = strExc(0)  '記錄智慧局收文文號
               strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & Text5(8)
            End If
       End If
   Else
    'end 2018/09/11
        If m_CP123s = "Y" Then
        'end 2009/4/28
           'Add by Morgan 2009/3/20 設定是否算發文室案件
           'modify by sonia 2014/6/23 加傳發文規費,一定有規費先用1, P-108903
           If ModifyDispatch(strReceiveNo, m_CP09s, m_CP123s, 1, Text5(4)) = False Then
               Exit Function
           End If
           'end 2009/3/20
        End If
   End If 'Add by Morgan 2012/9/13
   
    'Added by Lydia 2018/09/11 依據輸入的智慧局收文號(受理號,ex: 1073066637-0)，將本機C:\E-SET\RdcDocDir\(收文號ex: 1073066637-0)的pdf檔自動搬移到卷宗區(by Phoebe);
    If txtCP118.Text = "Y" And strFilePath <> "" And m_bolBeCalled = False Then
        strExc(1) = m_CP82
        If Val(m_CP82) > 0 Then
            If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                 strExc(1) = ""
            End If
        End If
        If Val(strExc(1)) = 0 Then
           'Modified by Lydia 2019/03/22 +傳入發文日
           If Pub_AutoEsetToCpp(True, pa(1), pa(2), pa(3), pa(4), pa(8), Label2(0).Caption, m_CP10, strFilePath, Text5(4).Text) = False Then
                 Exit Function
           End If
        End If
    End If
    'end 2018/09/11

   'Added by Lydia 2018/09/11 檢查完畢，更新備註欄位
   Text5(8).Text = strNewCP64
   
   'Add by Sindy 2021/11/15 檢查畫面上的物件是否含有Unicode文字
   If PUB_ChkUniText(Me, True, True) = False Then
      Exit Function
   End If
    
   If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Function
   
   Process = True
   
End Function

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

Private Sub txtCP118_GotFocus()
   TextInverse txtCP118
   CloseIme
End Sub

Private Sub txtCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Lydia 2018/09/11
Private Sub txtCP118_Change()
    txtPayToday.Text = Pub_FcpSetPayToday("1", Text5(4).Text, txtCP118.Text)
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
