VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010402 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請案號輸入"
   ClientHeight    =   5304
   ClientLeft      =   240
   ClientTop       =   1440
   ClientWidth     =   8736
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5304
   ScaleWidth      =   8736
   Begin VB.TextBox txtFavDt 
      Height          =   270
      Left            =   7605
      MaxLength       =   7
      TabIndex        =   24
      Top             =   3450
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   90
      TabIndex        =   75
      Top             =   4770
      Width           =   8520
      Begin VB.TextBox Text27 
         Height          =   270
         Left            =   5580
         MaxLength       =   8
         TabIndex        =   32
         Top             =   150
         Width           =   975
      End
      Begin VB.TextBox Text26 
         Height          =   270
         Left            =   3870
         MaxLength       =   2
         TabIndex        =   31
         Text            =   "2"
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox Text15 
         Height          =   270
         Left            =   7425
         MaxLength       =   8
         TabIndex        =   33
         Top             =   150
         Width           =   975
      End
      Begin VB.TextBox Text25 
         Height          =   270
         Left            =   1170
         MaxLength       =   8
         TabIndex        =   30
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "法定期限:"
         Height          =   180
         Index           =   4
         Left            =   4770
         TabIndex        =   79
         Top             =   180
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "本所期限:"
         Height          =   180
         Index           =   3
         Left            =   6615
         TabIndex        =   78
         Top             =   195
         Width           =   765
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "補文件期限: 文到           月"
         Height          =   180
         Index           =   1
         Left            =   2475
         TabIndex        =   77
         Top             =   195
         Width           =   2025
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "官方發文日:"
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   76
         Top             =   195
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdDeadLine 
      Caption         =   "補件資料"
      Height          =   400
      Left            =   2790
      TabIndex        =   74
      Top             =   60
      Width           =   1300
   End
   Begin VB.TextBox Text24 
      Height          =   270
      Left            =   6735
      MaxLength       =   15
      TabIndex        =   21
      Top             =   3150
      Width           =   1770
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   3285
      TabIndex        =   18
      Top             =   2835
      Width           =   945
   End
   Begin VB.CommandButton cmdPriDate 
      Caption         =   "優先權資料(&P)"
      Height          =   400
      Left            =   4185
      TabIndex        =   70
      Top             =   60
      Width           =   1300
   End
   Begin VB.CheckBox cheEnc 
      Caption         =   "國際申請日通知書"
      Height          =   255
      Index           =   5
      Left            =   6660
      TabIndex        =   11
      Top             =   1920
      Width           =   1770
   End
   Begin VB.CheckBox cheEnc 
      Caption         =   "國際申請號"
      Height          =   255
      Index           =   4
      Left            =   5265
      TabIndex        =   10
      Top             =   1920
      Width           =   1275
   End
   Begin VB.CheckBox cheEnc 
      Caption         =   "受理通知書"
      Height          =   255
      Index           =   3
      Left            =   3870
      TabIndex        =   9
      Top             =   1920
      Width           =   1275
   End
   Begin VB.CheckBox cheEnc 
      Caption         =   "圖式"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   8
      Top             =   1920
      Width           =   870
   End
   Begin VB.CheckBox cheEnc 
      Caption         =   "專利說明書"
      Height          =   255
      Index           =   1
      Left            =   1395
      TabIndex        =   7
      Top             =   1920
      Width           =   1365
   End
   Begin VB.TextBox Text20 
      Height          =   270
      Left            =   7605
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1590
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1410
      MaxLength       =   7
      TabIndex        =   4
      Top             =   1590
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   3645
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1590
      Width           =   2445
   End
   Begin VB.TextBox txtSplitDate 
      Height          =   270
      Left            =   6735
      MaxLength       =   8
      TabIndex        =   19
      Top             =   2850
      Width           =   1095
   End
   Begin VB.TextBox Text19 
      Enabled         =   0   'False
      Height          =   270
      Left            =   5175
      MaxLength       =   50
      TabIndex        =   29
      Top             =   4440
      Visible         =   0   'False
      Width           =   3300
   End
   Begin VB.TextBox Text18 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4125
      MaxLength       =   1
      TabIndex        =   13
      Top             =   2235
      Width           =   375
   End
   Begin VB.TextBox Text21 
      Height          =   270
      Left            =   1410
      MaxLength       =   50
      TabIndex        =   15
      Top             =   2550
      Width           =   2475
   End
   Begin VB.TextBox Text22 
      Height          =   270
      Left            =   6735
      MaxLength       =   8
      TabIndex        =   16
      Top             =   2550
      Width           =   1095
   End
   Begin VB.TextBox Text23 
      Height          =   270
      Left            =   1410
      TabIndex        =   17
      Top             =   2850
      Width           =   1095
   End
   Begin VB.TextBox Text17 
      Height          =   270
      Left            =   6735
      MaxLength       =   8
      TabIndex        =   14
      Top             =   2235
      Width           =   1095
   End
   Begin VB.TextBox Text16 
      Height          =   270
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   12
      Top             =   2235
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   6480
      TabIndex        =   35
      Top             =   45
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1260
      TabIndex        =   50
      Text            =   "P"
      Top             =   540
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7704
      TabIndex        =   36
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5670
      TabIndex        =   34
      Top             =   45
      Width           =   800
   End
   Begin VB.TextBox Text14 
      Height          =   270
      Left            =   1260
      MaxLength       =   4
      TabIndex        =   28
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   4950
      MaxLength       =   1
      TabIndex        =   23
      Top             =   3450
      Width           =   495
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1620
      MaxLength       =   1
      TabIndex        =   22
      Top             =   3450
      Width           =   495
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1260
      MaxLength       =   7
      TabIndex        =   3
      Top             =   855
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1728
      MaxLength       =   6
      TabIndex        =   0
      Top             =   540
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   270
      Left            =   2580
      MaxLength       =   1
      TabIndex        =   1
      Top             =   540
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   270
      Left            =   2820
      MaxLength       =   2
      TabIndex        =   2
      Top             =   540
      Width           =   375
   End
   Begin MSForms.TextBox Text8 
      Height          =   300
      Left            =   1395
      TabIndex        =   20
      Top             =   3150
      Width           =   3795
      VariousPropertyBits=   671107099
      MaxLength       =   32
      Size            =   "6694;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text13 
      Height          =   300
      Left            =   5490
      TabIndex        =   27
      Top             =   4110
      Width           =   2985
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "5265;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text12 
      Height          =   300
      Left            =   1260
      TabIndex        =   26
      Top             =   4110
      Width           =   2985
      VariousPropertyBits=   671107099
      MaxLength       =   250
      Size            =   "5265;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text11 
      Height          =   300
      Left            =   1260
      TabIndex        =   25
      Top             =   3780
      Width           =   7215
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "12726;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblFavDt 
      AutoSize        =   -1  'True
      Caption         =   "優惠期日期:"
      Height          =   180
      Left            =   6615
      TabIndex        =   80
      Top             =   3495
      Width           =   945
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "審查委員編號:"
      Height          =   180
      Index           =   5
      Left            =   5490
      TabIndex        =   73
      Top             =   3195
      Width           =   1125
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "審查委員名稱:"
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   72
      Top             =   3195
      Width           =   1125
   End
   Begin VB.Label Label25 
      Caption         =   "幣別:"
      Height          =   180
      Index           =   0
      Left            =   2790
      TabIndex        =   71
      Top             =   2895
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "附件:"
      Height          =   180
      Left            =   180
      TabIndex        =   69
      Top             =   1950
      Width           =   405
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "是否收到副本:           (Y:是)"
      Height          =   180
      Left            =   6390
      TabIndex        =   68
      Top             =   1635
      Width           =   2085
   End
   Begin VB.Label lblAppDate 
      AutoSize        =   -1  'True
      Caption         =   "申請日期:"
      Height          =   180
      Left            =   180
      TabIndex        =   67
      Top             =   1635
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   2790
      TabIndex        =   66
      Top             =   1635
      Width           =   765
   End
   Begin VB.Label lblSplitCase 
      AutoSize        =   -1  'True
      Caption         =   "分割案提交日:"
      Height          =   180
      Left            =   5490
      TabIndex        =   65
      Top             =   2880
      Width           =   1125
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號:"
      Enabled         =   0   'False
      Height          =   180
      Left            =   4395
      TabIndex        =   64
      Top             =   4485
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "(Y:是)"
      Height          =   180
      Left            =   4605
      TabIndex        =   63
      Top             =   2280
      Width           =   465
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "是否PCT案:"
      Height          =   180
      Left            =   3135
      TabIndex        =   62
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "代理人D/N NO:"
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   61
      Top             =   2580
      Width           =   1185
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "帳單日期:"
      Height          =   180
      Index           =   2
      Left            =   5490
      TabIndex        =   60
      Top             =   2580
      Width           =   765
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "帳單金額:"
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   59
      Top             =   2895
      Width           =   795
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "PCT提交日:"
      Height          =   180
      Left            =   5490
      TabIndex        =   58
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "發明已提實審:"
      Height          =   180
      Left            =   180
      TabIndex        =   57
      Top             =   2280
      Width           =   1125
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "(N:未提)"
      Height          =   180
      Left            =   2010
      TabIndex        =   56
      Top             =   2280
      Width           =   645
   End
   Begin VB.Label lblPriDate 
      AutoSize        =   -1  'True
      Caption         =   "PriorityDate"
      Height          =   180
      Left            =   1260
      TabIndex        =   55
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "優先權日:"
      Height          =   180
      Left            =   180
      TabIndex        =   54
      Top             =   1200
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   180
      X2              =   8520
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   8520
      Y1              =   1470
      Y2              =   1470
   End
   Begin MSForms.Label Label4 
      Height          =   180
      Index           =   2
      Left            =   7365
      TabIndex        =   53
      Top             =   555
      Width           =   1140
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "2011;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   210
      Index           =   1
      Left            =   1860
      TabIndex        =   52
      Top             =   4485
      Width           =   2400
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "4233;370"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   180
      Index           =   0
      Left            =   4620
      TabIndex        =   51
      Top             =   540
      Width           =   1620
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "2857;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "專利種類:"
      Height          =   180
      Left            =   6555
      TabIndex        =   49
      Top             =   555
      Width           =   765
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   48
      Top             =   4485
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱(日):"
      Height          =   180
      Left            =   4395
      TabIndex        =   47
      Top             =   4155
      Width           =   1065
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱(英):"
      Height          =   180
      Left            =   180
      TabIndex        =   46
      Top             =   4155
      Width           =   1065
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱(中):"
      Height          =   180
      Left            =   180
      TabIndex        =   45
      Top             =   3825
      Width           =   1065
   End
   Begin VB.Label Label10 
      Caption         =   "(Y:Word)"
      Height          =   255
      Left            =   5505
      TabIndex        =   44
      Top             =   3465
      Width           =   735
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "是否修改通知函內容:"
      Height          =   180
      Left            =   3240
      TabIndex        =   43
      Top             =   3495
      Width           =   1665
   End
   Begin VB.Label Label8 
      Caption         =   "(N:不印)"
      Height          =   255
      Left            =   2220
      TabIndex        =   42
      Top             =   3465
      Width           =   735
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "列印客戶通知函:"
      Height          =   180
      Left            =   180
      TabIndex        =   41
      Top             =   3495
      Width           =   1305
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "附件:"
      Height          =   180
      Left            =   4095
      TabIndex        =   40
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   180
      TabIndex        =   39
      Top             =   855
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請國家 :"
      Height          =   180
      Left            =   3780
      TabIndex        =   38
      Top             =   540
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   37
      Top             =   540
      Width           =   765
   End
End
Attribute VB_Name = "frm04010402"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/16 改成Form2.0 (Text8,Text11,Text12,Text13,Label4)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim pA(T_PA) As String, cp(T_CP) As String
Dim pa() As String, cp() As String

Dim intWhere As Integer, CP09 As String
Dim m_strSitu As String '目前有三種情況分別為"A"=實審已收未發,"B"=實審已收已發,"C"=未收實審
Dim strFirstPriDate As String  '最早的優先權日期
'Add by Morgan 2004/6/7
Dim m_bol307Exist As Boolean '是否為分割案
Dim m_st416NP08 As String  '分割案實審期限
'Add by Morgan 2005/11/7
Dim strPriority(1 To 5) As String '優先權資料 'Modify by Amy 2014/04/11 +strPriority(5)
Dim m_bolRePriDate As Boolean '優先權資料需重新輸入
Dim m_has202CP09 As String     '2009/3/24 add by sonia 是否更新未發文之補文件期限
Dim m_si880017 As Single '補件期限按鈕回傳狀態
Dim m_strUnSaveData As String '待新增補文件期限
Dim m_bolFMP As Boolean '是否FMP案
Dim m_1003CP09 As String '通知補文件收文號
Dim m_430CP09 As String, m_430CP25 As String '保密審查收文號,核准日 Add by Morgan 2010/5/26
'Added by Morgan 2014/3/5
Dim m_strUnSaveData2 As String '待新增修正期限
Dim m_1201CP09 As String '通知修正收文號
Dim strOldcp09 As String 'Modify by Amy 2014/08/12 由FormSave搬過來
Dim str240Date As String 'Added by Lydia 2015/05/01 補優先權轉讓證明(所限)
Dim m_bolUpd430CP47 As Boolean 'Added by Morgan 2015/9/14 更新保密審查提申日
Dim m_strCP10 As String, m_bolAddLP As Boolean 'Add by Morgan 2016/5/30
'Add By Sindy 2016/9/21
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/9/21 END
Dim m_bolNoCP27 As Boolean '不上發文 Added by Morgan 2020/1/16
'Modified by Morgan 2021/1/21 改共用
Dim strCP13New As String '收文智權人員
Dim strCP12New As String '收文業務區
   
Private Sub StartLetter2(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt() As String, i As Integer
   
   EndLetter ET01, CP09, ET03, strUserNum
   
   i = 0
   If Text18 = "Y" Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','PCT案','♀')"
   Else
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','非PCT案','♀')"
   End If
   
   '未提實審
   If Text16 = "N" Then
      strExc(0) = "select np09 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "'" & _
         " and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06||np07='416'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
            "','實審未收文','♀')"
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
            "','實審法定期限','" & RsTemp(0) & "')"
      End If
   '已提實審(不考慮已收未發情形)
   Else
      i = i + 1
      ReDim Preserve strTxt(i)
      'Modified by Morgan 2018/11/22 實審已收文->實審已發文,實審未發文
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','" & IIf(m_strSitu = "A", "實審未發文", "實審已發文") & "','♀')"
   End If
   'Modified by Morgan 2013/7/18 +232補優先權證明
   strExc(0) = "select np15 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "'" & _
      " and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06||np07 in ('202','232')"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','有缺文件','♀')"
      intI = 0
      strExc(1) = ""
      With RsTemp
      Do While Not .EOF
         intI = intI + 1
         Select Case "" & .Fields(0)
            Case "委託書"
               strExc(1) = strExc(1) & vbCrLf & "(" & intI & ") An original Power of Attorney executed by a representative of the applicant"
            Case "轉讓證明"
               strExc(1) = strExc(1) & vbCrLf & "(" & intI & ") An original Certificate of Assignment of Right of Priority signed by the inventors" & _
                  vbCrLf & "   or a notarized copy of the Worldwide Assignment duly executed by the inventors"
            Case "優先權證明"
               strExc(1) = strExc(1) & vbCrLf & "(" & intI & ") A certified copy of the priority document"
            Case Else
               strExc(1) = strExc(1) & vbCrLf & "(" & intI & ") " & .Fields(0)
         End Select
         .MoveNext
      Loop
      End With
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','補件內容','" & strExc(1) & "')"
   End If
   
   'Add by Morgan 2010/3/4
   If cp(10) = "110" Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "select '" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & "','大陸案申請號'" & _
         " ,pa11 from casemap,patent where cm01='" & cp(1) & "' and cm02='" & cp(2) & "'" & _
         " and cm03='" & cp(3) & "' and cm04='" & cp(4) & "' and cm10='4'" & _
         " and pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08"
   End If
   
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)

   Dim strTxt(1 To 11) As String, i As Integer
   Dim strEnc As String '附件
   Dim iIdx As Integer
   
   EndLetter ET01, CP09, ET03, strUserNum
   i = 1
   '910628 Sieg
   If Text15.Visible And Text15 <> "" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','本所期限','" & TransDate(Text15.Text, 2) & "')"
      i = i + 1
   End If
   
   'Added by Morgan 2016/3/14 有保密審查通知書
   If m_430CP09 <> "" And m_430CP25 <> "" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','有保密審查通知書','♀')"
      i = i + 1
   End If
   'end 2016/3/14
   
   'Add by Morgan 2005/10/25 有副本時判斷附件
   If Text20.Text = "Y" Then
      strEnc = Empty
      For iIdx = 1 To 5
         If cheEnc(iIdx).Value = 1 Then
            strEnc = strEnc & IIf(strEnc <> Empty, "、", Empty) & cheEnc(iIdx).Caption
         End If
      Next
      If strEnc <> Empty Then
         
         strEnc = PUB_ReContent(strEnc)
         'Added by Morgan 2016/3/14 有保密審查通知書
         If m_430CP09 <> "" And m_430CP25 <> "" Then
            strEnc = strEnc & "，以及向外國申請專利保密審查之意見通知書影本"
         End If
         'end 2016/3/14
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
            "','附件','" & ChgSQL(strEnc) & "')"
         i = i + 1
      End If
      'Add by Morgan 2008/1/18
      If pa(9) = "013" Then
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
            "','副本','副本')"
         i = i + 1
      End If
      
   End If
   
   'Add by Morgan 2004/6/7 未提實審分割案的實審期限
   If ET03 = "23" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','實審期限','" & m_st416NP08 & "')"
      i = i + 1
    End If
    
   'Add by Morgan 2004/7/5
   '新型,改請新型副本定稿加註修正期限，申(改)請日起二個月
   If InStr("24, 35, 64, 75", ET03) > 0 Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','修正期限起算日','" & IIf(cp(10) = "102", "申請日", "改請日") & "')"
      i = i + 1
   End If
    
   'Add by Morgan 2005/10/14 分割案提交日
   If txtSplitDate <> "" Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','分割案提交日'," & TransDate(txtSplitDate, 2) & ")"
      i = i + 1
   End If
      
   'Add by Morgan 2009/7/23
   If pa(8) = "1" And pa(9) = "000" Then
      strExc(0) = "SELECT * FROM CASEPROGRESS WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & _
         "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10='417' AND CP57 IS NULL"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
            "','公開','公開')"
         i = i + 1
      End If
   End If
   'end 2009/7/23
   
   'Add by Morgan 2009/12/24 補件內容
   strExc(0) = "select np15 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "'" & _
      " and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06||np07='202'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','有缺文件','♀')"
      i = i + 1
      intI = 0
      strExc(1) = ""
      With RsTemp
      strExc(1) = "" & .Fields(0)
      .MoveNext
      Do While Not .EOF
         strExc(0) = "" & .Fields(0)
         .MoveNext
         If .EOF Then
            strExc(1) = strExc(1) & "及" & strExc(0)
         Else
            strExc(1) = strExc(1) & "、" & strExc(0)
         End If
      Loop
      End With
      
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','補件內容','" & strExc(1) & "')"
      i = i + 1
   End If
   
   'Add by Morgan 2010/3/1 分割或已收文提早公開則定稿不要帶預定公開段落
   If cp(10) = "307" Or PUB_ChkCPExist(cp, "417") Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','公開段不印','♀')"
      i = i + 1
   End If
   
   'Added by Lydia 2015/05/01 當通知副本時補優先權轉讓證明若未發文,將期限帶入定稿
   If Len(str240Date) > 0 Then
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & CP09 & "','" & ET03 & "','" & strUserNum & _
         "','補優先權轉讓證明期限','" & str240Date & "')"
      i = i + 1
   End If
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(i - 1, strTxt) Then
   If Not ClsLawExecSQL(i - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub cmdDeadLine_Click()
   If Text15 <> "" And Text27 <> "" Then
      ModifyAddDeadline1 cp(9), Text15, Text27, m_si880017, True, m_strUnSaveData, m_bolFMP, m_strUnSaveData2
   Else
      MsgBox "請輸入補文件期限！"
      Text25.SetFocus
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim strTmp As String, strTmp2 As String, bolChk As Boolean
   
   Select Case Index
      Case 2
         Unload frm04010401
         Unload Me
      Case 3
         Unload Me
         frm04010401.Show
      Case 0
         ' 90.07.18 modify by louis
         If Text5.Enabled And IsEmptyText(Text5) Then
            MsgBox "請輸入申請日期", vbOKOnly + vbCritical, "檢核資料"
            Text5.SetFocus
            Exit Sub
         End If
         
         If IsEmptyText(Text6) Then
            'Modify by Morgan 2005/10/26 加判斷台灣案或有副本才檢查
            If pa(9) = 台灣國家代號 Or Text20.Text = "Y" Then
               MsgBox "請輸入申請案號", vbOKOnly + vbCritical, "檢核資料"
               Text6.SetFocus
               Exit Sub
            End If
         End If

        'Modify By Cheng 2003/05/28
'         If pa(9) = 台灣國家代號 And Left("" & Me.Text6.Text, 2) <> Left("" & Me.Text5.Text, 2) Then
         'Remove by Morgan 2010/8/20 改由共用函數檢查(後面的程式有做)
         'If pa(9) = 台灣國家代號 And Me.Text3.Text = "0" And Left("" & Me.Text6.Text, 2) <> Left("" & Me.Text5.Text, 2) And Left(cp(10), 1) <> "3" Then
         '   MsgBox "申請案號的前二碼必須為申請年度 !", vbCritical
         '   Text6.SetFocus
         '   Exit Sub
         'End If
         If IsEmptyText(Text11) And IsEmptyText(Text12) And IsEmptyText(Text13) Then
            MsgBox "請輸入專利名稱", vbOKOnly + vbCritical, "檢核資料"
            Text11.SetFocus
            Exit Sub
         End If
         
        'Add By Cheng 2002/11/01
        If (Me.Text21.Text = "" Xor Me.Text22.Text = "") Or (Me.Text21.Text = "" Xor Me.Text23.Text = "") Or (Me.Text22.Text = "" Xor Me.Text23.Text = "") Then
            MsgBox "代理人D/N NO , 帳單日期 及 帳單金額 " & Chr(10) & Chr(13) & "三欄位必須同時輸入或不輸入資料!!!", vbExclamation + vbOKOnly
            If Me.Text21.Text = "" Then Me.Text21.SetFocus:     Text21_GotFocus:     Exit Sub
            If Me.Text22.Text = "" Then Me.Text22.SetFocus:     Text22_GotFocus:     Exit Sub
            If Me.Text23.Text = "" Then Me.Text23.SetFocus:     Text23_GotFocus:     Exit Sub
        End If
        
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         'Add by Morgan 2004/3/30
         '檢查分割案件關係檔紀錄
         If Text14 = "307" Then
            Dim stPA10 As String
            If CheckDivCase(stPA10) = False Then
               If MsgBox("無此案之原案紀錄，確定要繼續？", vbExclamation + vbYesNo) = vbNo Then
                  Exit Sub
               End If
            'Modify by Morgan 2005/3/22
            'ElseIf (stPA10 = "" Or stPA10 <> ChangeTStringToWString(Text5)) Then
            ElseIf (stPA10 = "" Or stPA10 <> TransDate(Text5, 2)) Then
               MsgBox "原案申請日與原案不符，請查明！", vbCritical
               Exit Sub
            End If
         End If
         
         'Add by Morgan 2005/11/8 非台灣案有優先權資料且有副本
         If Pub_StrUserSt03 <> "M51" And pa(9) <> "000" And strPriority(1) <> "" And Text20 = "Y" And m_bolRePriDate = False Then
            MsgBox "本案有優先權資料,請重新輸入以便與原資料檢核！", vbCritical
            Exit Sub
         End If
         
         'Add by Morgan 2009/10/7
         '大陸一案兩請申請日輸入時必須互相檢查發明及新型之申請日是否為同一天,若不是,則show訊息告知user。
         If pa(9) = "020" And InStr("101,102", cp(10)) > 0 And Text5 <> "" Then
            strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,ptm04 C2,pa10 C3" & _
               " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & cp(1) & "' and cm02='" & cp(2) & "' and cm03='" & cp(3) & "' and cm04='" & cp(4) & "'" & _
               " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & cp(1) & "' and cm06='" & cp(2) & "' and cm07='" & cp(3) & "' and cm08='" & cp(4) & "') X" & _
               ",patent,patenttrademarkmap where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 and pa10>0 and ptm01(+)='1' and ptm02(+)=pa08"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If (Val(DBDATE(Text5)) >= 20091001 Or Val(RsTemp("C3")) >= 20091001) And Val(DBDATE(Text5)) <> Val(RsTemp("C3")) Then
                  If MsgBox("本案與" & RsTemp("C2") & "案 " & RsTemp("C1") & " 為一案兩請，但申請日不同日！是否仍要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                     Exit Sub
                  End If
               End If
               'Add by Morgan 2009/12/23
               If Text20 = "Y" Then
                  MsgBox "請檢查是否有一案兩請之聲明請求書！", vbExclamation
               End If
            End If
         End If
         
         'Add By Sindy 2020/7/20
         If m_strIR01 <> "" Then
            '下載信件檔
            'Modify By Sindy 2022/11/10 + IIf(pa(9) <> 台灣國家代號, "PAT", "RX")
            If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", IIf(pa(9) <> 台灣國家代號, "PAT", "RX"), , True) = False Then
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            'Add By Sindy 2022/7/21
            If Left(Pub_StrUserSt03, 2) = "F2" Then
               If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            '2022/7/21 END
         End If
         '2020/7/20 END
         
         'Add By Cheng 2002/11/05
         'Modified by Lydia 2021/02/02 改成TextBox可複製
         'frm04010401.lblBillNo.Caption = ""
         frm04010401.txtBillno.Text = ""
         
         'Remove by Morgan 2009/11/23 改加按鈕控制且只抓收文號相關期限逐筆輸入
         ''2009/3/24 ADD BY SONIA 大陸案若有補文件期限,檢查是否已有未發文之補文件期限,詢問是否更新期限或另行新增B類補文件期限
         'm_has202CP09 = ""
         'If pa(9) = "020" Then
         '   Check202
         'End If
         ''2009/3/24 END
         'end 2009/11/23
         
         If FormSave = False Then
            MsgBox "存檔失敗，請洽系統管理人 !", vbCritical
            Exit Sub
         End If
         
        'Add by Lydia 2015/01/26 FMP案件之申請案號輸入(僅第一次), 請加入自動發e-mail給外專收文人員
        'Modified by Lydia 2025/08/03 香港案之申請案號輸入，全部都要通知;
         'If Len(Trim(pa(11))) = 0 And m_bolFMP Then
         If (Len(Trim(pa(11))) = 0 Or pa(9) = "013") And m_bolFMP Then
             strExc(0) = "": strExc(1) = "": strExc(2) = ""
             If ClsPDGetCustomerNameAndAddress(pa(26), strExc(0)) Then
             End If
             If Len(Text11) > 0 Then strExc(2) = Text11
             If Len(Text12) > 0 And Len(strExc(2)) = 0 Then strExc(2) = Text12
             If Len(Text13) > 0 And Len(strExc(2)) = 0 Then strExc(2) = Text13
             
             'Email內文
             strExc(2) = "本所案號：" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & vbCrLf & vbCrLf & _
                              "專利名稱：" & strExc(2) & vbCrLf & vbCrLf & _
                              "申請人　：" & strExc(0) & vbCrLf & vbCrLf & _
                              "申請案號：" & Trim(Text6)

             'Added by Lydia 2021/01/04 增加檢查進度檔是否尚有告知代理人901/主動補正203/補正204尚未發文
             strExc(3) = "select cp05,cp09,cp14,nvl(cpm04,cpm03) cpm0403 from caseprogress,casepropertymap " & _
                              "where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in ('901','203','204') and cp158=0 and cp159=0 " & _
                              "and cp01=cpm01(+) and cp10=cpm02(+) order by cp05 desc "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(3))
             strExc(3) = "": strExc(4) = "": strExc(5) = "": strExc(6) = ""
             If intI = 1 Then
                Do While Not RsTemp.EOF
                     If strExc(3) = "" Then   '收文號抓第一筆
                        strExc(3) = "" & RsTemp.Fields("cp09")
                        strExc(4) = "" & RsTemp.Fields("cp14")
                     End If
                     strExc(5) = strExc(5) & "," & RsTemp.Fields("cpm0403")
                     RsTemp.MoveNext
                Loop
             Else
                strExc(3) = "": strExc(4) = "": strExc(5) = "": strExc(6) = ""
             End If
             'Email收件人
             strExc(0) = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
             'Email主旨
             strExc(1) = "通知 " & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & " 的" & IIf(m_strCP10 = 通知申請日, "提申日", "申請案號") & "已輸入！"
             
             If strExc(3) <> "" Then
                  'Email主旨加註：※本案尚有告知代理人/主動補正/補正尚未發文
                  'strExc(1) = strExc(1) & " ※本案尚有" & Mid(strExc(5), 2) & "尚未發文" 'Mark by lydia 2022/08/24 重複
                  'Email收件人：承辦人員+該進度之工程師,若無則抓最後一道收文之工程師
                  If strExc(4) = "" Then
                       strSql = "select max(cp05||cp14) from caseprogress,staff where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp159=0 and cp14=st01(+) and st03='F21' and st04='1' and st01 not like 'F%' "
                       intI = 1
                       Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                       If intI = 1 Then
                           If "" & RsTemp(0) <> "" Then
                               strExc(4) = Mid("" & RsTemp(0), 9)
                           End If
                       End If
                  End If
                  If strExc(4) <> "" Then
                      strExc(0) = strExc(0) & ";" & strExc(4)
                      'CC
                      strExc(6) = PUB_GetFCPEngSup(strExc(4))
                  End If
                  'Email內文加註：※本案尚有告知代理人/主動補正/補正尚未發文，請承辦告申日後，卷退工程師主管分案進行後續流程。
                  'Modified by Lydia 2022/03/09 改備註
                  'strExc(2) = strExc(2) & vbCrLf & vbCrLf & "※本案尚有" & Mid(strExc(5), 2) & "尚未發文，請承辦告申日後，卷退工程師主管分案進行後續流程。"
                  strExc(2) = strExc(2) & vbCrLf & vbCrLf & "※本案尚有" & Mid(strExc(5), 2) & "尚未發文，主管請分案，工程師請處理後續流程，謝謝！"
                  'Added by Lydia 2022/03/09 主旨比照內文加註
                  strExc(1) = strExc(1) & "※本案尚有" & Mid(strExc(5), 2) & "尚未發文，主管請分案，工程師請處理後續流程，謝謝！"
                  'Added by Lydia 2022/08/25 通知FMP案管制人
                  strExc(7) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
                  If InStr(strExc(0) & ";" & strExc(6), strExc(7)) = 0 Then
                      strExc(6) = strExc(6) & IIf(strExc(6) <> "", ";", "") & strExc(7)
                  End If
                  'Added by Lydia 2022/08/24 加發給操作人員;
                  'Modified by Lydia 2022/08/25 限F22外專程序人員---當FMP案管制人休假
                  If InStr(strExc(0) & ";" & strExc(6), strUserNum) = 0 And Pub_StrUserSt03 = "F22" Then
                      strExc(6) = strExc(6) & IIf(strExc(6) <> "", ";", "") & strUserNum
                  End If
                  'end 2022/08/24
             End If
             'end 2021/01/04
             
             'Modified by Morgan 2016/8/17 --秀玲
             'PUB_SendMail strUserNum, PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4)), CP09, "通知 " & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & " 的申請案號已輸入！", strExc(2)
             'Modified by Lydia 2021/01/04 改用變數
             'PUB_SendMail strUserNum, PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4)), CP09, "通知 " & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & " 的" & IIf(m_strCP10 = 通知申請日, "提申日", "申請案號") & "已輸入！", strExc(2) '提申日已輸入 or 申請案號已輸入
             PUB_SendMail strUserNum, strExc(0), CP09, strExc(1), strExc(2), , , , , , strExc(6)
         End If
         If Text9 <> "N" Then '通知函
            If Text10 = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If

            Select Case pa(9)
               Case 台灣國家代號
                  strTmp = "01"
                  'Add By Cheng 2002/11/01
                  '若專利種類為發明
                  If pa(8) = "1" Then
                       
                        '未提實審
                        If Me.Text16.Text = "N" Then
                           '2009/2/24 ADD BY SONIA for 大對台定稿
                           If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
                              strTmp = "66"
                           Else
                           '2009/2/24 END
                              strTmp = "08"
                              'Add by Morgan 2004/6/7   分割案
                              If m_bol307Exist = True Then
                                strTmp = "23"
                              End If
                           End If
                        '已提實審
                        Else
                            'add by toni  20080904 for 大對台定稿
                           If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
                              strTmp = "65"
                           Else

                              strTmp = "09"
                              'Add by Morgan 2006/4/27 加已收文未發文定稿 10
                              If m_strSitu = "A" Then
                                 strTmp = "10"
                              End If
                          End If
                        End If
                  'Add by Morgan 2004/6/14
                  '新型副本定稿加註修正期限
                  'Modify by Morgan 2004/7/5
                  '加改請新型副本定稿加註修正期限
                  ElseIf (Val(Text14) = 102 Or Val(Text14) = 302) And Val(Text5) >= 930701 Then
                     strTmp = "24"
                     'add by toni  20080904 for 大對台定稿
                     If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
                        strTmp = "65"
                     End If
                  Else
                     strTmp = "01"
                     'add by toni  20080904 for 大對台定稿
                     If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
                        strTmp = "65"
                     End If
                  End If
                  
               'Modify by Morgan 2005/10/25 非台灣通知函精簡化
               Case "056"  'PCT
                  'Modify by Morgan 2005/10/31
                  '有副本才要出開窗的定稿
                  If Text20.Text = "Y" Then
                     strTmp = "04"
                  Else
                     strTmp = "07"
                  End If
                  
               Case Else
                  'Modify by Morgan 2005/10/31
                  '有副本才要出開窗的定稿
                  If Text20.Text = "Y" Then
                     '若為PCT案(只有發明才有PCT案)
                     If Me.Text18.Text = "Y" Then
                        strTmp = "03"
                     Else
                        strTmp = "02"
                     End If
                     'Add by Morgan 2009/11/23 FMP-受理通知(有副本)
                     If m_bolFMP Then
                        'Added by Morgan 2024/6/13 香港設計
                        If pa(9) = "013" And pa(8) = "3" Then
                           strTmp2 = "54"
                        Else
                        'end 2024/6/13
                        
                           strTmp2 = "52"
                        End If
                     End If
                  Else
                     If Me.Text18.Text = "Y" Then
                        strTmp = "06"
                     Else
                        strTmp = "05"
                     End If
                     'Add by Morgan 2009/11/23 FMP-受理通知(申請日)
                     If m_bolFMP Then
                        'Added by Morgan 2024/6/13 香港設計
                        If pa(9) = "013" And pa(8) = "3" Then
                           strTmp2 = "53"
                        Else
                        'end 2024/6/13
                           strTmp2 = "51"
                        End If
                     End If
                  End If
           End Select
           
           'Remove by Morgan 2006/8/18 改用定稿特殊符號控制
'           'Modify by Morgan 2005/10/26
'           '台灣若有優先權資料則出不同定稿
'            If lblPriDate <> "" And Val(strTmp) < 12 And pA(9) = "000" Then
'               strTmp = Format(Val(strTmp) + 11, "00")
'            End If
            'end 2006/8/18
            
            'Romove by Morgan 2006/8/18 改用系統例外欄位"<客戶案號或專利種類/P>"替代
'            'Add by Morgan 2004/6/28
'            '若有客戶案件案號則本所案號印客戶案件案號不印申請國家專利種類
'            If pA(9) = 台灣國家代號 And pA(48) <> "" Then
'               strTmp = Format(Val(strTmp) + 40, "00")
'            End If
            'end 2006/8/18
            
            StartLetter "03", strTmp
            '910628 Sieg
            If Text15.Visible And Text15 <> "" Then
               bolChk = True
            End If
            
            'Modify by Amy 2014/08/12 +傳C類收文號 for P台灣案電子化
            'Add by Morgan 2009/11/23
            If m_bolFMP Then
               'Modified by Morgan 2025/4/11 FMP不再印紙本--品薇
               NowPrint CP09, "03", strTmp, bolChk, strUserNum, 0, , , , 1, , , , , , , , strOldcp09, , , , , True
               If strTmp2 <> "" Then
                  strUserNum = strFMPNum
                  StartLetter2 "03", strTmp2
                  'Modified by Morgan 2016/5/30 不可傳LD18否則FCP承辦執行定維護時會開E化的畫面
                  'NowPrint CP09, "03", strTmp2, False, strUserNum, 0, , , , , , , , , , , , strOldcp09
                  NowPrint CP09, "03", strTmp2, False, strUserNum
                  strUserNum = strUser1Num
               End If
            Else
            'end 2009/11/23
               NowPrint CP09, "03", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strOldcp09
            End If
            
            'Added by Morgan 2016/6/22
            If Left(Pub_StrUserSt03, 1) <> "F" And m_bolAddLP = True Then
               If bolChk Then
                  frm1105_1.m_RecNo = strOldcp09
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & m_strCP10 & ".CUS.PDF"
                  frm1105_1.Show
               End If
            End If
            'end 2014/08/12
            
         End If
         ' 90.07.05 modify by louis (顯示訊息)
         'MsgBox "存檔成功 !", vbInformation
         ' modify by sonia 不顯示訊息
         
         'Add By Sindy 2016/9/21
         If Me.m_strIR01 <> "" Then
            Unload frm04010401
            Unload Me
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         Else
         '2016/9/21 END
            
            Unload Me
           'Modify By Cheng 2002/12/18
           '不要清除輸入條件
   '         'Modify By Cheng 2002/10/24
   '         '保留系統類別
   ''         frm04010401.Text1 = ""
   '         frm04010401.Text2 = ""
   '         frm04010401.Text3 = ""
   '         frm04010401.Text4 = ""
            frm04010401.Show
            'Modify BY Cheng 2002/10/24
   '         frm04010401.Text1.SetFocus
            frm04010401.Text2.SetFocus
           'Add By Cheng 2002/12/18
            TextInverse frm04010401.Text2
         End If
   End Select
End Sub

'Add by Morgan 2005/11/8
'優先權資料
Private Sub cmdPriDate_Click()
   '非台灣
   If pa(9) <> "000" Then
      m_bolRePriDate = True
      'Modify by Morgan 2007/3/5
      'ModifyPriority strPriority(1), strPriority(2), strPriority(3), , m_bolRePriDate, pa(1) & pa(2) & pa(3) & pa(4)
      'Modify by Amy 2014/04/11 +strPriority(5)
      ModifyPriority strPriority(1), strPriority(2), strPriority(3), , m_bolRePriDate, pa(1) & pa(2) & pa(3) & pa(4), pa(9), True, strPriority(4), strPriority(5)
      'end 2007/3/5
   End If
End Sub

Private Sub Form_Initialize()
    'add by nickc 2007/02/02
    ReDim pa(1 To TF_PA) As String
    ReDim cp(1 To TF_CP) As String
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   
   Text1.Text = frm04010401.Text1
   Text2.Text = frm04010401.Text2
   Text3.Text = frm04010401.Text3
   Text4.Text = frm04010401.Text4
   Text7.Text = frm04010401.Text5
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010401.m_strIR01
   m_strIR02 = frm04010401.m_strIR02
   m_strIR03 = frm04010401.m_strIR03
   m_strIR04 = frm04010401.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
   ReadPatent
   
   'Add by Morgan 2011/4/11
   '特殊備註預設不印客戶通知函
   If InStr(cp(64), "郵寄申請") > 0 Then
      Text9 = "N"
   End If
   
   'Modify by Morgan 2004/11/26 分割案也不可改
   'Modify by Morgan 2006/5/16 香港案也不可改
   'Modify by Morgan 2006/11/2 香港的發明才不可改--玲玲
   'If (Val(Text14) >= 301 And Val(Text14) <= 307) Or pA(9) = "013" Then
   'Modify by Morgan 2006/11/8 加PCT案
   'If (Val(Text14) >= 301 And Val(Text14) <= 307) Or (pA(9) = "013" And pA(8) = "1") Then
   If (Val(Text14) >= 301 And Val(Text14) <= 307) Or (pa(9) = "013" And pa(8) = "1") Or (pa(46) = "Y") Then
      Text5.Enabled = False
   Else
      Text5.Enabled = True
   End If
   
   'Add by Morgan 2005/10/26
   '非台灣
   If pa(9) <> "000" Then
      'Modify by Morgan 2006/10/26 PCT或分割案要先判斷(PCT的申請日已改在分案時便輸入)
      'PCT案的提交日=提申日
      If Text18 = "Y" Then
         Text17.Tag = DBDATE(cp(47))
         Text17 = ""
         If Text17.Tag <> Empty Then Text9 = "N"
      '一般案
      ElseIf Text5.Enabled = True Then
         Text5.Tag = DBDATE(pa(10))
         Text5 = ""
         '不是第一次輸申請日時預設不印通知函,副本上Y時再預設印
         If Text5.Tag <> Empty Then Text9 = "N"
      '分割案的提交日=提申日
      ElseIf txtSplitDate.Enabled = True Then
         txtSplitDate.Tag = DBDATE(cp(47))
         txtSplitDate = ""
         If txtSplitDate.Tag <> Empty Then Text9 = "N"
      End If
   'Add by Morgan 2005/11/8
      cmdPriDate.Visible = True
   Else
      cmdPriDate.Visible = False
   End If
   
   '預設附件
   If pa(9) <> "000" Then
      'Add by Morgan 2008/1/18 香港沒有附件--敏惠
      If pa(9) = "013" Then
         For intI = 1 To 5
            cheEnc(intI).Value = 0
            cheEnc(intI).Enabled = False
         Next
      Else
      'end 2008/1/18
         If pa(8) <> "3" Then
            cheEnc(1).Value = 1
         End If
         cheEnc(2).Value = 1
         If pa(9) = "056" Then
            cheEnc(4).Value = 1
            cheEnc(5).Value = 1
         Else
            cheEnc(3).Value = 1
            cheEnc(4).Enabled = False
            cheEnc(5).Enabled = False
         End If
      End If
   End If
   '2005/10/26 end
   
   'Add by Morgan 2008/5/13
   If cp(44) <> "" Then
      PUB_Add2Combo Combo1, cp(44)
   End If

   'add by sonia 2019/7/30 PCT進入國家階段提醒
   If InStr(NewCasePtyList, cp(10)) > 0 And pa(46) = "Y" Then
      MsgBox "本案為PCT國際申請案進入國家階段案件，請確認代理人提交專利局的PCT相關資料是否正確無誤。"
   End If
   'end 2019/7/30
End Sub

Public Sub SetData(ByVal strNo As String, ByVal CPFIELD As String)
   ChgCaseNo strNo, pa
   CP09 = CPFIELD
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2009/8/17
   'Set frm04010402 = Nothing 'Removed by Morgan 2021/12/16 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
   If Text11 = "" And Text12 = "" And Text13 = "" Then
      MsgBox "專利名稱不可同時空白 !", vbCritical
      Text11.SetFocus
   End If
End Sub

Private Function ChgType(ByVal iSitu As Integer) As Boolean
 Dim strTmp As String, bolTmp As Boolean, i As Integer
   ChgType = False
   Select Case iSitu
      Case 6
         If Text6 = "" Then
            'Modify by Morgan 2005/10/19 可以只輸申請日
            'MsgBox "申請案號不可空白 !", vbCritical
            ChgType = True
            '2005/10/19 end
         Else
            i = 2
            If pa(9) = 台灣國家代號 Then
               i = 0
               'Modified by Morgan 2021/6/24 衍生設計除外 Ex:P-127439--陳玲玲
               If Len("" & Me.Text5.Text) > 0 And cp(10) <> "125" Then
                    'Modify By Cheng 2003/05/28
                    '若為母案(即非追加聯合案), 則要檢查申請案號
                    If Me.Text3.Text = "0" Then
                        'Modify by Morgan 2010/6/22 修改100年問題
                        ''2005/5/20 MODIFY BY SONIA 改請案不檢查申請年度
                        ''If Left("" & Me.Text6.Text, 2) <> Left("" & Me.Text5.Text, 2) Then
                        'If Left("" & Me.Text6.Text, 2) <> Left("" & Me.Text5.Text, 2) And Left(cp(10), 1) <> "3" Then
                        ''2005/5/20 END
                        '   MsgBox "申請案號的前二碼必須為申請年度 !", vbCritical
                        If Val(Left(Text6, 1)) > "1" Then
                           strExc(1) = Val(Left(Text6, 2))
                        Else
                           strExc(1) = Val(Left(Text6, 3))
                        End If
                        strExc(2) = Trim(Val(Text5) \ 10000)
                        If strExc(1) <> strExc(2) And Left(cp(10), 1) <> "3" Then
                           MsgBox "申請案號的前二(三)碼必須為申請年度 !", vbCritical
                        'end 2010/6/22
                           Exit Function
                        End If
                        'Add by Morgan 2004/11/5 "聯合或追加案不可為母案(本所號第3欄為0) !"
                        'Modify by Morgan 2010/6/22 修改100年問題
                        'If Len(Me.Text6.Text) > 8 Then
                        If Len(Me.Text6.Text) > 9 Then
                           'Modified by Morgan 2012/12/19 +衍生設計、
                           'Removed by Morgan 2021/6/1 取消,母案可能非本所辦理 Ex:P-127284--玲玲
                           'MsgBox "衍生設計、聯合或追加案不可為母案(本所號第3欄為0) !", vbCritical
                           'Exit Function
                           'end 2021/6/1
                        End If
                    End If
               End If
            Else
                'Modify By Cheng 2002/11/29
'               If pa(9) = 大陸國家代號 And pa(46) <> "Y" Then
               If pa(9) = 大陸國家代號 And Me.Text18.Text <> "Y" Then
                  i = 1
               End If
            End If
            If i <> 2 Then
                'Modify By Cheng 2003/05/28
                '若為母案(即非追加聯合案), 則要檢查申請案號
                '93.4.28 modify by sonia
                'If Me.Text3.Text = "0" Then
                'Modify by Morgan 2010/8/20 追加聯合案也可檢查
                'If Me.Text3.Text = "0" And cp(10) <> "117" Then
                If cp(10) <> "117" Then
                'end 2010/8/20
                '93.4.28 end
                    'Modify by Morgan 2004/8/17 改以專利種類判斷
                    'ChgType = ChkAppNo(Text6.Text, Val(Mid(Text14, 3, 1)), i)
                    '2005/6/14 MODIFY BY SONIA
                    'ChgType =.(Text6.Text, Val(pa(8)), i)
                    ChgType = ChkAppNo(Text6.Text, Val(pa(8)), i, Val(pa(23)), , cp(10))
                    '2005/6/14 END
                Else
                    ChgType = True
                End If
            Else
               ChgType = True
            End If
            
            'Added by Morgan 2012/8/21 檢查申請案號是否重複
            If ChgType = True Then
               ChgType = PUB_ChkAppNo(Text6.Text, pa(1), pa(2), pa(9))
            End If
            'end 2012/8/21
         End If
      Case 14
         If pa(9) = 台灣國家代號 Then
            bolTmp = False
         Else
            bolTmp = True
         End If
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(pA(1), Text14, strTmp, BolTmp) Then
         If ClsPDGetCaseProperty(pa(1), Text14, strTmp, bolTmp) Then
            Label4(1) = strTmp
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetPatentTrademarkKind(專利, pA(8), strTmp, BolTmp, pA(9)) = 1 Then
            If ClsPDGetPatentTrademarkKind(專利, pa(8), strTmp, bolTmp, pa(9)) = 1 Then
               Label4(2) = strTmp
               ChgType = True
            Else
               Label4(2) = ""
            End If
         Else
            Label4(1) = ""
         End If
   End Select
End Function

Private Sub Text14_Validate(Cancel As Boolean)
   If Text14 = "" Then
      MsgBox "案件性質不可空白 !", vbCritical
      Cancel = True
   Else
      If cp(10) <> Text14.Text Then
         If pa(9) <> 台灣國家代號 Then
            If Text14 = "101" Or Text14 = "102" Or Text14 = "103" Then
               If ChgType(14) Then
                  Cancel = Not ChgType(6)
               Else
                  Cancel = True
               End If
            Else
               MsgBox "案件性質錯誤，請重新輸入 !", vbCritical
               Cancel = True
            End If
         Else
            MsgBox "非台灣案時才可修改 !", vbCritical
            Text14 = cp(10)
         End If
      End If
   End If
   If Cancel = True Then TextInverse Text14
    SplitCheck
End Sub
'分割案欄位控制
Private Sub SplitCheck()
    '若為分割時顯示分割案提交日否則隱藏
    If Text14 = "307" Then
        lblSplitCase.Visible = True
        txtSplitDate.Visible = True
        lblAppDate.Caption = "原案申請日期:"
    Else
        lblSplitCase.Visible = False
        txtSplitDate.Visible = False
        lblAppDate.Caption = "申請日期:"
    End If
    
End Sub

Private Sub Text16_GotFocus()
    TextInverse Me.Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 78 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text17_GotFocus()
    TextInverse Me.Text17
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
   
   If (Text17.Text = "") And (pa(9) = "020") And (Text18.Text = "Y") Then
         MsgBox "申請國家為大陸(020)的PCT案，PCT提交日不可為空白！", vbCritical
         Cancel = True
   ElseIf Me.Text17.Text <> "" Then
         'Modify by Morgan 2005/4/21 改輸西元
         If CheckIsDate(Text17) = False Then
            Cancel = True
         End If
   End If
   
   'Added by Morgan 2018/11/26
   '寰華案PCT提交日變更提醒--敏莉
   '不必限定寰華案--秀玲(郭確認OK)
   strExc(1) = ""
   If Text18 = "Y" And cp(47) <> "" Then
      If DBDATE(cp(47)) <> DBDATE(Text17) Then
         strExc(1) = "PCT提交日和前次不一致，請輸入正確提交日後向大陸代理人確認原因！"
      End If
   End If
   'end 2018/11/26

   'Add by Morgan 2005/11/2 修改時雙重檢查
   If Cancel = False Then
      If CheckReKey(Text17, , strExc(1)) = True Then
         Text17.Tag = Text17
      Else
         Cancel = True
      End If
   End If
   
   If Cancel = True Then
      Text17_GotFocus
   End If
End Sub

Private Sub Text18_GotFocus()
    TextInverse Me.Text18
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 89 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text20_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text20.IMEMode = 2
   CloseIme
   TextInverse Text20
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   ElseIf KeyAscii = Asc("Y") Then
      Text9 = ""
   End If
End Sub

Private Sub Text21_GotFocus()
    TextInverse Me.Text21
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text22_GotFocus()
    TextInverse Me.Text22
End Sub

Private Sub Text22_Validate(Cancel As Boolean)
    If Me.Text22.Text <> "" Then
      'Modify by Morgan 2005/4/21 改輸西元
      If CheckIsDate(Text22) = False Then
         Cancel = True
      'Add by Morgan 2006/4/25 檢查不可大於系統日
      ElseIf Val(Text22) > Val(strSrvDate(1)) Then
         MsgBox "帳單日期不可大於系統日！", vbExclamation
         Cancel = True
      End If
   End If
   If Cancel = True Then
      Text22_GotFocus
   End If
End Sub

Private Sub Text23_GotFocus()
    TextInverse Me.Text23
End Sub

Private Sub Text23_Validate(Cancel As Boolean)
    'Add By Cheng 2002/11/01
    If Me.Text23.Text <> "" Then
        If IsNumeric(Me.Text23.Text) = False Then
            MsgBox "帳單金額輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            Text23_GotFocus
        'Add by Morgan 2004/1/30
        ElseIf Val(Text23) <> 0 Then
            If cp(44) = "" Then
                MsgBox "該筆進度資料無代理人，不可輸入帳單!!!", vbExclamation + vbOKOnly
                Cancel = True
                TextInverse Text23
            End If
        'Add end ---------------
        End If
    End If
End Sub

Private Sub Text24_GotFocus()
   TextInverse Text24
End Sub

Private Sub Text24_Validate(Cancel As Boolean)
   If Not CheckLengthIsOK(Text24, Text24.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Function SetDate() As Boolean
   If Text25 <> "" And Text26 <> "" Then
      strExc(1) = CompDate(1, Val(Text26), Text25)
      Text27 = strExc(1)
      Text27.Tag = Text27
   End If
   SetDate = SetDate2
End Function

Private Function SetDate2() As Boolean
   If Text27 <> "" Then
      'Added by Lydia 2025/10/29
      If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
         Text15 = PUB_GetPOurDeadline(Text27, pa(9))
      Else
      'end 2025/10/29
         Text15 = PUB_GetWorkDay1(CompDate(2, -7, Text27), True)
      End If  'Added by Lydia 2025/10/29
      If Val(Text15) < Val(strSrvDate(1)) Then
         Text15 = strSrvDate(1)
      End If
   End If
   SetDate2 = True
End Function

Private Sub Text25_GotFocus()
   TextInverse Text25
End Sub

Private Sub Text25_Validate(Cancel As Boolean)
   If Text25 <> "" And Text25.Tag <> Text25 Then
      If Len(Text25) <> 8 Then
         MsgBox "必須輸入西元年！"
         Cancel = True
      ElseIf Not ChkDate(Text25) Then
         Cancel = True
      ElseIf Text26 <> "" Then
         Cancel = Not SetDate
      End If
   End If
   If Cancel Then
      TextInverse Text25
   Else
      Text25.Tag = Text25
   End If
End Sub

Private Sub Text26_GotFocus()
   TextInverse Text26
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text26_Validate(Cancel As Boolean)
   If Text26 <> "" And Text26.Tag <> Text26 Then
      Cancel = Not SetDate
   End If
   If Cancel = False Then
      Text26.Tag = Text26
   End If
End Sub

Private Sub Text27_GotFocus()
   TextInverse Text27
End Sub

Private Sub Text27_Validate(Cancel As Boolean)
   If Text27 <> "" And Text27.Tag <> Text27 Then
      If Len(Text27) <> 8 Then
         MsgBox "必須輸入西元年！"
         Cancel = True
      ElseIf Not ChkDate(Text27) Then
         Cancel = True
      Else
         Cancel = Not SetDate2
      End If
   End If
   If Cancel = False Then
      Text27.Tag = Text27
   End If
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If IsEmptyText(Text5) Then
      MsgBox "申請日不可空白，請重新輸入 !", vbCritical
      Cancel = True
   Else
      'Modify by Morgan 2005/3/22 大陸案之申請日入改以西元年輸入
      If pa(9) = 台灣國家代號 Then
         If CheckIsTaiwanDate(Text5) = False Then
            Cancel = True
         ElseIf Val(Text5) > Val(strSrvDate(2)) Then
            Cancel = True
            MsgBox "申請日不正確或大於系統日，請重新輸入 !", vbCritical
         End If
      Else
         If CheckIsDate(Text5) = False Then
            Cancel = True
         ElseIf Val(Text5) > Val(strSrvDate(1)) Then
            Cancel = True
            MsgBox "申請日不正確或大於系統日，請重新輸入 !", vbCritical
         End If
      End If
   End If
   
   'Add by Morgan 2005/10/19 修改時雙重檢查
   If Cancel = False Then
      If CheckReKey(Text5) = True Then
         Text5.Tag = Text5
      Else
         Cancel = True
      End If
   End If
   
   If Cancel = True Then
      Text5_GotFocus
   End If
End Sub

Private Sub Text6_Change()
   'Add by Morgan 2016/5/30
   If Text6 <> "" Then
      If Text20 <> "Y" Then
         'Modify By Sindy 2019/10/23
'         Text20 = "Y"
'         Text9 = ""
         '請取消非台灣案,申請案號輸入,輸入申請案時,是否收到副本Y,不預設Y,由USER自行輸入
         'Modified by Morgan 2019/11/13 外專程序操作時一律要預設--敏莉
         'If pa(9) = "000" Then
         If pa(9) = "000" Or Pub_StrUserSt03 = "F22" Then
         'end 2019/11/13
            Text20 = "Y"
            Text9 = ""
         End If
         '2019/10/23 END
      End If
   Else
      Text20 = cp(145)
   End If
   'end 2016/5/30
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modify by Morgan 2005/6/30 再加"(",")"
   'edit by nickc 2005/06/29 允許輸入斜線
   'Modify by Morgan 2010/11/5 +"-"
   If Not (KeyAscii = 8 Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46 Or KeyAscii = 47 Or KeyAscii = Asc("(") Or KeyAscii = Asc(")") Or KeyAscii = Asc("-")) Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   Cancel = Not ChgType(6)
   If Cancel = True Then TextInverse Text6
End Sub

Private Sub Text8_GotFocus()
   TextInverse text8
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Not CheckLengthIsOK(text8, text8.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> "" Then
      If Text9 <> "N" Then
         ShowMsg MsgText(9044)
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

Private Function FormSave() As Boolean
   Dim intStep As Integer, strTxt(1 To 20) As String, i As Integer
   Dim intMax As Long
   'Add By Cheng 2002/10/31
   Dim rsA As New ADODB.Recordset
   Dim StrSQLa As String
   Dim rsB As New ADODB.Recordset
   Dim StrSqlB As String
   'Add By Cheng 2002/11/01
   Dim strTemp As String
   Dim strTemp1 As String
   Dim strTemp2 As String
   Dim strDateS(0 To 5) As String
   Dim dobDateAdd As Double
   Dim strStartDate As String
   'Add By Cheng 2002/11/05
   Dim strBillNo As String '帳單編號
   'Add by Morgan 2005/11/4
   Dim stCP47 As String
   'Add by Morgan 2006/4/17
   Dim strCP06 As String, strCP07 As String
   Dim strNP06 As String
   '2008/4/28 ADD BY SONIA
   Dim strCaseProperty As String
   Dim varTemp As Variant
   Dim yearTemp As String
   '2008/4/28 END
   Dim strUpdCpSQL As String
   Dim Tmp As String                        '20080918 add by toni 訊息存備註
   Dim st307Msg As String
   Dim arrData() As String
   Dim stCP64 As String 'Added by Morgan 2012/6/13
   Dim stNP23 As String, stCP48Desc As String, st1201CP64 As String 'Added by Morgan 2014/3/5
   Dim strLDate As String, strLtitle As String 'Added by Lydia 2015/09/09 大陸案公開期限(日期,備註)
   'Modified by Lydia 2016/1/14 下一程序公開期限999的收文號掛相關總收文號CP09
   'Dim tmpCp06 As String, tmpCp09 As String 'Added by Lydia 2015/09/09 所限,新案收文號
   Dim tmpCp06 As String
   Dim strErrMsg As String 'Added by Morgan 2016/6/30
   Dim tmpBol As Boolean 'Added by Lydia 2018/07/09
   
   m_bolAddLP = False 'Added by Morgan 2016/6/22
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
   
   cnnConnection.BeginTrans

   Screen.MousePointer = vbHourglass
   
   'Add by Morgan 2005/11/8 有優先權資料且有重輸才要
   If strPriority(1) <> Empty And m_bolRePriDate = True Then
      'Modify by Morgan 2007/4/25 加strPriority(4)
      'Modify by Amy 2014/04/11 +strPriority(5)
      ClsPDSavePriority pa, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)
      strFirstPriDate = PUB_GetFirstPriDate(cp) 'Add by Morgan 2006/5/12 重讀最早優先權日
   End If
   
   CP09 = cp(9)
   intStep = 1
   strUpdCpSQL = ""
   '1
   'Modify by Morgan 2005/11/4 整理並加若未收達時一併更新
   '更新提申日
   If pa(9) <> "000" Then
      stCP47 = ""
      '香港案用提交日更新
      'Modify by Morgan 2006/11/2 只有發明--玲玲
      'If pA(9) = "013" And Val(Text17) > 0 Then
      If (pa(9) = "013" And pa(8) = "1") And Val(Text17) > 0 Then
         stCP47 = TransDate(Text17, 2)
      '若是PCT案用提交日更新,不再存基本檔備註
      ElseIf Text18 = "Y" And Val(Text17) > 0 Then
         stCP47 = TransDate(Text17, 2)
      '分割案用提交日更新,不再存基本檔備註
      ElseIf Val(txtSplitDate) > 0 Then
         stCP47 = TransDate(txtSplitDate, 2)
      '其他非台灣用申請日更新
      ElseIf Val(Text5) > 0 Then
         stCP47 = TransDate(Text5, 2)
      End If
      If stCP47 <> "" Then
         'Modify by Morgan 2008/5/13 移到下面一起做
         'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP47=" & stCP47 & ",CP46=NVL(CP46," & stCP47 & ") WHERE CP09='" & CP09 & "'"
         'cnnConnection.Execute strTxt(intStep)
         'intStep = intStep + 1
         strUpdCpSQL = strUpdCpSQL & ",CP47=" & stCP47 & ",CP46=NVL(CP46," & stCP47 & ")"
         
      End If
   End If
   
   'Added by Morgan 2012/2/10
   If cp(145) <> Text20 Then
      strUpdCpSQL = strUpdCpSQL & ",CP145=" & CNULL(Text20)
   End If
   'end 2012/2/10
   
   '2 3開頭不等於307 edit by Toni   20090918
   'modify by sonia 2023/3/21 308改請衍生設計在發文時已更新,此處不可再更新P-130147-1
   If Left(Text14, 1) = "3" And Text14 <> "307" And Text14 <> "308" Then
      'Modify by Morgan 2008/5/13 移到下面一起做
      'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP30=" & CNULL(pa(11)) & " WHERE CP09='" & CP09 & "'"
      'cnnConnection.Execute strTxt(intStep)
      'intStep = intStep + 1
      strUpdCpSQL = strUpdCpSQL & ",CP30=" & CNULL(pa(11))
   End If
   
   If strUpdCpSQL <> "" Then
      strTxt(intStep) = "UPDATE CASEPROGRESS SET " & Mid(strUpdCpSQL, 2) & " WHERE CP09='" & CP09 & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   
    '更新彼所案號(大陸案時處理)
   If Me.Text19.Visible And Me.Text19.Enabled Then
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP45=" & CNULL(ChgSQL(Me.Text19.Text)) & " WHERE CP09='" & CP09 & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      '2005/12/19 ADD BY SONIA 更新相同本所案號之相同代理人的彼所案號，若是彼所案號空的話
      'Modified by Morgan 2011/12/26 申請號來時不一定會有彼所案號,程序會先輸入空白
      'strTxt(intStep) = "update caseprogress set cp45=" & CNULL(ChgSQL(Me.Text19.Text)) & " where cp09 in (select cp09 from caseprogress where cp45 is null and " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp09<'C' AND cp44 in (select cp44 from caseprogress where cp09='" & CP09 & "' ))"
      'Modified by Morgan 2012/2/15 取消 cp09<'C' 條件(C類也會有發文作業,有代理人就要更新彼號,資料才會一致)
      strTxt(intStep) = "update caseprogress set cp45=" & CNULL(ChgSQL(Me.Text19.Text)) & " where cp09 in (select cp09 from caseprogress where rtrim(cp45) is null and " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND cp44 in (select cp44 from caseprogress where cp09='" & CP09 & "' ))"
      cnnConnection.Execute strTxt(intStep), intI
      intStep = intStep + 1
      '2005/12/19 END
   End If
    
   'Add by Toni  20080918  改請時案件備註加註改請
   Tmp = ""
   If Left(Text14, 1) = "3" And Text14 <> "307" Then
      If Text14 = "301" Or Text14 = "302" Or Text14 = "303" Then '2010/1/15 ADD BY SONIA
         '先抓原申請案件性質
         strExc(0) = "SELECT CPM03 FROM caseprogress,CASEPROPERTYMAP WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & "and instr('" & NewCasePtyList & "',cp10)>0 and cp01=CPM01(+) and cp10=CPM02(+)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI > 0 Then
            Tmp = RsTemp.Fields(0)
         End If
      End If   '2010/1/15 END
      '再抓改請案件性質
      strExc(0) = "SELECT CPM03 FROM CASEPROPERTYMAP WHERE CPM01=" & CNULL(cp(1)) & " and CPM02=" + CNULL(cp(10))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI > 0 Then
         Tmp = ";" + ChangeTStringToTDateString(cp(27)) + Tmp + RsTemp.Fields(0)
      End If
      
   End If
   'end by Toni 20080918
    
   'Modify by Morgan 2007/8/16 加PA140(是否收到副本)
    '更新是否為PCT案欄位
   Select Case Text14
      '93.4.28 add by sonia 積體電路不更新專利種類
      'Modify by Morgan 2004/11/26 聯合，追加不要更新專利種類
      'Modified by Morgan 2012/2/10 取消 pa140 改放 cp145
      'Modified by Morgan 2012/10/8 +衍生設計125
      'Modified by Morgan 2020/5/25 +改請衍生設計308 Ex: P123933-1
      Case "117", "104", "105", "125"
          strTxt(intStep) = "UPDATE PATENT SET PA10=" & CNULL(TransDate(Text5, 2), True) & ",PA11=" & CNULL(ChgSQL(Text6)) & "," & _
            "PA05=" & CNULL(ChgSQL(Text11)) & ",PA06=" & CNULL(ChgSQL(Text12)) & ",PA07=" & CNULL(ChgSQL(Text13)) & "," & _
            "PA46 ='" & ChgSQL(Trim(Me.Text18.Text)) & "'  WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      '93.4.28 end
      Case "110"
         strTxt(intStep) = "UPDATE PATENT SET PA10=" & CNULL(TransDate(Text5, 2), True) & ",PA11=" & CNULL(ChgSQL(Text6)) & "," & _
            "PA05=" & CNULL(ChgSQL(Text11)) & ",PA06=" & CNULL(ChgSQL(Text12)) & ",PA07=" & CNULL(ChgSQL(Text13)) & "," & _
            "PA08='1' , PA46 ='" & ChgSQL(Trim(Me.Text18.Text)) & "'  WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      
      Case "307"
         strTxt(intStep) = "UPDATE PATENT SET PA10=" & CNULL(TransDate(Text5, 2), True) & ",PA11=" & CNULL(ChgSQL(Text6)) & "," & _
             "PA05=" & CNULL(ChgSQL(Text11)) & ",PA06=" & CNULL(ChgSQL(Text12)) & ",PA07=" & CNULL(ChgSQL(Text13)) & "," & _
             "PA46 ='" & ChgSQL(Trim(Me.Text18.Text)) & "'  WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      
      Case "109"
         strTxt(intStep) = "UPDATE PATENT SET PA10=" & CNULL(TransDate(Text5, 2), True) & ",PA11=" & CNULL(ChgSQL(Text6)) & "," & _
            "PA05=" & CNULL(ChgSQL(Text11)) & ",PA06=" & CNULL(ChgSQL(Text12)) & ",PA07=" & CNULL(ChgSQL(Text13)) & "," & _
            "PA46 ='" & ChgSQL(Trim(Me.Text18.Text)) & "'  WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      
      'add by Morgan 2004/8/19 加 "304", "305", "306"
      Case "304", "305", "306"
         '2008/9/22 modify by sonia 案件備註加註改請
         'strTxt(intStep) = "UPDATE PATENT SET PA10=" & TransDate(Text5, 2) & ",PA11=" & CNULL(ChgSQL(Text6)) & "," & _
         '    "PA05=" & CNULL(ChgSQL(Text11)) & ",PA06=" & CNULL(ChgSQL(Text12)) & ",PA07=" & CNULL(ChgSQL(Text13)) & "," & _
         '    "PA46 ='" & ChgSQL(Trim(Me.Text18.Text)) & "'  WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
         strTxt(intStep) = "UPDATE PATENT SET PA10=" & CNULL(TransDate(Text5, 2), True) & ",PA11=" & CNULL(ChgSQL(Text6)) & "," & _
             "PA05=" & CNULL(ChgSQL(Text11)) & ",PA06=" & CNULL(ChgSQL(Text12)) & ",PA07=" & CNULL(ChgSQL(Text13)) & "," & _
             "PA46 ='" & ChgSQL(Trim(Me.Text18.Text)) & "',PA91=PA91||'" & Tmp & "'  WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      
      'add by Toni 20080918
      Case "301", "302", "303"
         strTxt(intStep) = "UPDATE PATENT SET PA10=" & CNULL(TransDate(Text5, 2), True) & ",PA11=" & CNULL(ChgSQL(Text6)) & "," & _
             "PA05=" & CNULL(ChgSQL(Text11)) & ",PA06=" & CNULL(ChgSQL(Text12)) & ",PA07=" & CNULL(ChgSQL(Text13)) & "," & _
             "PA08='" & Mid(Text14, 3, 1) & "' , PA46 ='" & ChgSQL(Trim(Me.Text18.Text)) & "',PA91=PA91||'" & Tmp & "'  WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      'end by Toni 20080918
      
      'add by sonia 2023/3/21 308,309不可更新PA08(P-130147-1)
      Case "308", "309"
         strTxt(intStep) = "UPDATE PATENT SET PA10=" & CNULL(TransDate(Text5, 2), True) & ",PA11=" & CNULL(ChgSQL(Text6)) & "," & _
             "PA05=" & CNULL(ChgSQL(Text11)) & ",PA06=" & CNULL(ChgSQL(Text12)) & ",PA07=" & CNULL(ChgSQL(Text13)) & "," & _
             "PA46 ='" & ChgSQL(Trim(Me.Text18.Text)) & "'  WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      'end 2023/3/21
      
      Case Else
         strTxt(intStep) = "UPDATE PATENT SET PA10=" & CNULL(TransDate(Text5, 2), True) & ",PA11=" & CNULL(ChgSQL(Text6)) & "," & _
             "PA05=" & CNULL(ChgSQL(Text11)) & ",PA06=" & CNULL(ChgSQL(Text12)) & ",PA07=" & CNULL(ChgSQL(Text13)) & "," & _
             "PA08='" & Mid(Text14, 3, 1) & "' , PA46 ='" & ChgSQL(Trim(Me.Text18.Text)) & "'  WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   End Select
   '92.10.23 END
    'Add By Cheng 2002/11/06
    cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1
   
   'Add by Morgan 2008/9/3 若有被主張優先權時需更新相關期限資料
   If Text5 <> "" And Text6 <> "" Then
      strExc(0) = "select pd01,pd02,pd03,pd04 from pridate where PD06='" & pa(1) & pa(2) & pa(3) & pa(4) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(6) = ChgSQL(Text6)
         'Add by Morgan 2010/6/21
         '台灣案申請號大於1開頭的前面補0再回寫到優先權號
         If Not bolNewAppNoFormat Then 'Add by Morgan 2010/8/23 改9碼格式後就不必再補0
            If pa(9) = "000" Then
               If Val(Left(Text6, 1)) > "1" Then
                  strExc(6) = "0" & strExc(6)
               End If
            End If
         End If
         'end 2010/6//21
         
         strSql = "update pridate set pd05=" & CNULL(DBDATE(Text5), True) & ",pd06='" & strExc(6) & "',pd07='" & pa(9) & "',pd08='" & pa(8) & "'" & _
            " where pd06='" & pa(1) & pa(2) & pa(3) & pa(4) & "'"
         cnnConnection.Execute strSql, intI
         With RsTemp
            .MoveFirst
            Do While Not .EOF
               PUB_UpdCfpDate1 .Fields("pd01"), .Fields("pd02"), .Fields("pd03"), .Fields("pd04"), True
               .MoveNext
            Loop
         End With
      End If
      
      'Added by Morgan 2012/1/11 若有收文申請優先權證明發Mail通知承辦人可送件
      'Modify by Amy 2014/07/24 增加436 優先權存取碼
      'strExc(0) = "select * from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='405' and cp27||cp57 is null"
      strExc(0) = "Select cp14," & IIf(pa(9) = "020", "cpm04", "cpm03") & " cpm From CaseProgress,CasePropertyMap " & _
                       "Where cp01='" & pa(1) & "' And cp02='" & pa(2) & "' And cp03='" & pa(3) & "' And cp04='" & pa(4) & "' " & _
                       "And cp10 In ('405','436') And cp27||cp57 is null And cp01=cpm01(+) And cp10=cpm02(+)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Do While RsTemp.EOF = False
            strExc(1) = "" & RsTemp("cp14")
            If strExc(1) = "" Then strExc(1) = strUserNum
            strExc(2) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
            strExc(2) = strExc(2) & " 案已輸入申請號，" & RsTemp.Fields("cpm") & "可送件！"
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                        " values( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd')" & _
                        ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','如旨')"
            cnnConnection.Execute strSql, intI
            RsTemp.MoveNext
         Loop
      End If
   End If
   
   '3
   If Text14 <> cp(10) Then
      cp(9) = "B" & CompAutoNumberYear(GetTaiwanThisYear) 'Add by Morgan 2011/2/24
      cp(10) = Text14.Text
      cp(5) = strSrvDate(1)
      cp(26) = "N"
      cp(31) = ""
      cp(32) = "N"
      cp(20) = "N"
      cp(16) = ""
      cp(17) = ""
      cp(18) = ""
      cp(19) = ""
      '91.11.12 ADD BY SONIA
      cp(27) = strSrvDate(1)
      '91.11.12 END
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.SaveNewCaseProgressDatabase("B", cp, intWhere) Then
      If Not ClsPDSaveNewCaseProgressDatabase("B", cp, intWhere) Then
        GoTo ErrorHandler
      End If
   End If

   '4
   'Modify by Morgan 2005/10/26
   '新增C收文:台灣案時判斷有輸申請案號，非台灣案時都要
   'If Text6 <> "" Then
   If (pa(9) = 台灣國家代號 And Text6 <> "") Or (pa(9) <> 台灣國家代號) Then
      'Add by Morgan 2016/5/30
      If Text6 <> "" Then
         'Modified by Morgan 2019/10/28 非台灣案副本未上Y,視為通知申請日--玲玲
         'm_strCP10 = 通知申請案號
         If pa(9) <> 台灣國家代號 And Text20 <> "Y" Then
            m_strCP10 = 通知申請日
         Else
            m_strCP10 = 通知申請案號
         End If
         'end 2019/10/28
      Else
         m_strCP10 = 通知申請日
      End If
      'end 2016/5/30
      
      strOldcp09 = AutoNo("C", 6)
      'Remove by Lydia 2016/1/14
      'tmpCp09 = strOldcp09 'Added by Lydia 2015/09/09
      'Modify by Morgan 2008/5/21 +CP35,CP117
      'Modified by Morgan 2012/5/25 +CP119
      'Modified by Morgan 2020/1/16 +m_bolNoCP27
      strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
         "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43,CP35,CP117,CP119,CP145) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
         pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & "NULL" & ",'" & _
         strOldcp09 & "','" & m_strCP10 & "'," & CNULL(strCP12New) & "," & _
         CNULL(strCP13New) & ",'" & strUserNum & "','N','N','N'," & IIf(m_bolNoCP27, "NULL", strSrvDate(1)) & ",'" & CP09 & "'" & _
         "," & CNULL(text8) & "," & CNULL(Text24) & "," & Val(DBDATE(Text7)) & ",'" & Text20 & "')"
   
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      
      'Add by Amy 2014/08/12 P台灣案電子化
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) And pa(9) = 台灣國家代號 Then
         If Text9 <> "N" Then
            'Modified by Morgan 2014/12/17 案件性質改"通知申請案號"
            'Modified by Morgan 2018/8/1
            'strExc(1) = PUB_GetLetterJudge(pa(1), m_strCP10, , , pa(1), pa(2), pa(3), pa(4))
            strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), m_strCP10)
            PUB_AddLetterProgress strOldcp09, 1, True, strExc(1), False, pa(26), m_strCP10, pa(75), True
            m_bolAddLP = True
         End If
         
      'Added by Morgan 2016/5/26
      ElseIf Left(Pub_StrUserSt03, 1) <> "F" Then
         '客戶通知函
         If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And pa(9) <> 台灣國家代號 Then
            'Modified by Morgan 2018/8/1
            'strExc(1) = PUB_GetLetterJudge(pa(1), m_strCP10, , pa(9), pa(1), pa(2), pa(3), pa(4))
            strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), m_strCP10, pa(9), , , m_bolFMP)
            
            'Modified by Morgan 2019/1/21 不必檢查ALTR(無附件)
            'intI = IIf(Text20 = "Y", 2, 1)
            If m_strCP10 = 通知申請日 Then
               strSql = "update caseprogress set cp121='Y' where cp09='" & strOldcp09 & "'"
               cnnConnection.Execute strSql, intI
               intI = 0
            Else
               intI = IIf(Text20 = "Y", 2, 1)
            End If
            'end 2019/1/21
            
            PUB_AddLetterProgress strOldcp09, intI, IIf(Text9 = "N", False, True), strExc(1), False, pa(26), m_strCP10, pa(75)
            m_bolAddLP = True
         End If
      'end 2016/5/26
      
      End If
      'end 2014/08/12
   
      If Text15.Visible And Text15 <> "" Then
         'Modify by Morgan 2006/11/12 補文件的承辦人要抓原申請案工程師
         'Add by Morgan 2006/9/11 大陸有補件期限時自動收C類的通知補文件及B類的補文件
         If pa(9) = "020" Then
            If m_si880017 = 1 Then 'Added by Morgan 2014/3/5
               strExc(1) = AutoNo("C", 6)
               m_1003CP09 = strOldcp09
               'Modify by Morgan 2009/11/19 +CP133,CP134
               'Modified by Morgan 2012/5/7 +CP06,CP07
               'Modified by Morgan 2012/5/25 +CP119
               strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10," & _
                  "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43,CP133,CP134,CP119) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
                  pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & DBDATE(Text15) & "," & DBDATE(Text27) & ",'" & _
                  strExc(1) & "','" & 通知補文件 & "'," & CNULL(strCP12New) & "," & _
                  CNULL(strCP13New) & ",'" & strUserNum & "','N','N','N'," & strSrvDate(1) & ",'" & CP09 & "'," & CNULL(DBDATE(Text25), True) & "," & CNULL(Text26, True) & "," & Val(DBDATE(Text7)) & ")"
               
               cnnConnection.Execute strTxt(intStep)
               intStep = intStep + 1
            End If
            
'Modify by Morgan 2009/11/19 改點補文件資料按鈕維護
'            '2009/3/24 modify by sonia 依使用者選擇是否更新未發文補文件期限
'            'strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10," & _
'               "CP12,CP13,CP14,CP20,CP26,CP32,CP43) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
'               pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & PUB_GetWorkDay1(TransDate(Text15, 2), True) & _
'               "," & TransDate(Text15, 2) & ",'" & AutoNo("B", 6) & "','" & 補文件 & "'," & CNULL(strCP12New) & "," & _
'               CNULL(strCP13New) & ",'" & cp(14) & "','N','N','N','" & strOldcp09 & "')"
'            If m_has202CP09 <> "" Then
'               strTxt(intStep) = "UPDATE CASEPROGRESS SET CP06=" & PUB_GetWorkDay1(TransDate(Text15, 2), True) & _
'                  ",CP07=" & TransDate(Text15, 2) & " WHERE CP09='" & m_has202CP09 & "'"
'            Else
'               strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10," & _
'                  "CP12,CP13,CP14,CP20,CP26,CP32,CP43) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
'                  pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & PUB_GetWorkDay1(TransDate(Text15, 2), True) & _
'                  "," & TransDate(Text15, 2) & ",'" & AutoNo("B", 6) & "','" & 補文件 & "'," & CNULL(strCP12New) & "," & _
'                  CNULL(strCP13New) & ",'" & cp(14) & "','N','N','N','" & strOldcp09 & "')"
'            End If
'            '2009/3/24 end
'
'            cnnConnection.Execute strTxt(intStep)
'            intStep = intStep + 1
'
'            strNP06 = "Y"
'         Else
'            strNP06 = ""
'         End If
'         'end 2006/9/11
'
'         '若本所期限非工作天則抓最近的工作天
'         strTxt(intStep) = "declare intMax number;begin   select max(np22)+1 into intMax from nextprogress;IF intMax IS NULL THEN intMax:=1; END IF;"
'         strTxt(intStep) = strTxt(intStep) & "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP06,NP07,NP08," & _
'            "NP09,NP10,NP22) VALUES ('" & strOldcp09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & _
'            pa(4) & "'," & CNULL(strNP06) & "," & 補文件 & "," & PUB_GetWorkDay1(TransDate(Text15, 2), True) & "," & TransDate(Text15, 2) & ",'" & strCP13New & "',intMax);"
'         strTxt(intStep) = strTxt(intStep) & " end;"
'
'         cnnConnection.Execute strTxt(intStep)
'         intStep = intStep + 1
            
            'Added by Lydia 2025/10/29
            stNP23 = ""
            If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
               strExc(1) = PUB_GetPOurDeadline(DBDATE(Text27), pa(9), stNP23, pa(1), "202")
            End If
            'end 2025/10/29
            strSql = "update caseprogress set cp06=" & DBDATE(Text15) & ",cp07=" & DBDATE(Text27) & " where cp43='" & cp(9) & "' and cp10='202' and cp57 is null and cp27 is null"
            cnnConnection.Execute strSql, intI
            'Modified by Lydia 2025/10/29 +NP23
            strSql = "update nextprogress set np08=" & DBDATE(Text15) & ",np09=" & DBDATE(Text27) & ",np23=" & IIf(stNP23 = "", "NP23", stNP23) & " where np01='" & cp(9) & "' and np06||np07='202'"
            cnnConnection.Execute strSql, intI
            stCP64 = ""
            If m_strUnSaveData <> "" Then
               arrData = Split(m_strUnSaveData, vbCrLf)
               For i = LBound(arrData) To UBound(arrData)
                  If arrData(i) <> "" Then
                     'Modified by Lydia 2025/10/29 +NP23
                     strSql = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22,NP23)" & _
                        " SELECT '" & m_1003CP09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','202'" & _
                        "," & CNULL(DBDATE(Text15), True) & "," & CNULL(DBDATE(Text27), True) & ",'" & strCP13New & "'" & _
                        "," & CNULL(ChgSQL(arrData(i))) & _
                        ",NP22," & CNULL(stNP23, True) & " FROM (SELECT MAX(NP22)+1 NP22 FROM NEXTPROGRESS) X"
                     cnnConnection.Execute strSql, intI
                     
                     stCP64 = arrData(i) & ";" & stCP64
                  End If
               Next
            End If
            
            'Added by Morgan 2012/6/13
            If stCP64 <> "" Then
               strSql = "update caseprogress set cp64='" & ChgSQL(stCP64) & "'||cp64 where cp09='" & m_1003CP09 & "'"
               cnnConnection.Execute strSql, intI
            End If
            'end 2012/6/13
            
            'Added by Morgan 2014/3/5
            st1201CP64 = ""
            If m_strUnSaveData2 <> "" Then
               m_1201CP09 = AutoNo("C", 6)
               strExc(1) = PUB_GetFmpCP48(strSrvDate(1), DBDATE(Text15), DBDATE(Text27), DBDATE(Text25), Text26, stNP23, stCP48Desc)
               'Modified by Morgan 2017/10/11 FMP預設承辦人比照FCP
               'strExc(2) = PUB_GetFmpCP14(pa)
               strExc(2) = PUB_GetFCPPromoterNo(cp(9), 通知修正, cp(14))
               'end 2017/10/11
               strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10," & _
                  "CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48,CP133,CP134,CP119) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
                  pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & DBDATE(Text15) & "," & DBDATE(Text27) & ",'" & _
                  m_1201CP09 & "','" & 通知修正 & "'," & CNULL(strCP12New) & "," & _
                  CNULL(strCP13New) & ",'" & strExc(2) & "','N','N','N','" & CP09 & "'," & CNULL(strExc(1), True) & "," & CNULL(DBDATE(Text25), True) & "," & CNULL(Text26, True) & "," & Val(DBDATE(Text7)) & ")"
               
               cnnConnection.Execute strTxt(intStep)
               intStep = intStep + 1
               'Added by Lydia 2025/10/29
               If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                  strExc(1) = PUB_GetPOurDeadline(DBDATE(Text27), pa(9), stNP23, pa(1), "204")
               End If
               'end 2025/10/29
               arrData = Split(m_strUnSaveData2, vbCrLf)
               For i = LBound(arrData) To UBound(arrData)
                  If arrData(i) <> "" Then
                     strSql = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22,NP23)" & _
                        " SELECT '" & m_1201CP09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','204'" & _
                        "," & CNULL(DBDATE(Text15), True) & "," & CNULL(DBDATE(Text27), True) & ",'" & strCP13New & "'" & _
                        "," & CNULL(ChgSQL(arrData(i))) & _
                        ",NP22," & CNULL(stNP23, True) & " FROM (SELECT MAX(NP22)+1 NP22 FROM NEXTPROGRESS) X"
                     cnnConnection.Execute strSql, intI
                     
                     st1201CP64 = arrData(i) & ";" & st1201CP64
                  End If
               Next
               
            End If
            If st1201CP64 <> "" Then
               strSql = "update caseprogress set cp64='" & ChgSQL(st1201CP64) & "'||cp64 where cp09='" & m_1201CP09 & "'"
               cnnConnection.Execute strSql, intI
            End If
            'end 2014/3/5
            
         End If
'end 2009/11/19

      End If
      
      'Add by Morgan 2010/5/11
      '若無通知補文件時不管是否PCT案都用提申日+2個月更新補文件的期限-->郭雅娟P-94435
      If m_1003CP09 = "" And stCP47 <> "" Then
         strExc(1) = PUB_GetWorkDay1(CompDate(1, 2, stCP47), True)
         strSql = "update caseprogress set cp06=" & strExc(1) & " where cp43='" & cp(9) & "' and cp10='202' and cp57 is null and cp27 is null and cp07 is null"
         cnnConnection.Execute strSql, intI
         strSql = "update nextprogress set np08=" & strExc(1) & " where np01='" & cp(9) & "' and np06||np07='202' and np09 is null"
         cnnConnection.Execute strSql, intI
      End If
      
      '5
      intStep = 1
      
   End If
   'Modify by Morgan 2009/6/8 +指定提申995,最終提申996
   'strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & CP09 & "' And NP07 in ( " & 收達 & "," & 提申 & ")"
   strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & CP09 & "' And NP07 in ( " & 收達 & "," & 提申 & ",995,996)"
   cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1
   
   'Modify by Morgan 2009/3/12 積體電路佈局117不要掛
   'If pa(8) = "1" And (m_strSitu = "A" Or m_strSitu = "C") Then
   'Modify by Morgan 2009/7/10 不必限制專利種類(澳門新型及設計也有實審制度)
   'If pa(8) = "1" And (m_strSitu = "A" Or m_strSitu = "C") And cp(10) <> "117" Then
   'Modified by Morgan 2022/1/17 分割案規則不同要排除(台灣案在分割發文,大陸案在後面有函數更新) Ex:P-128925--玲玲
   If (m_strSitu = "A" Or m_strSitu = "C") And cp(10) <> "117" And cp(10) <> "307" Then
      'Modify by Morgan 2009/7/10 改呼叫共用函數
      PUB_UpdExamDate pa(1), pa(2), pa(3), pa(4), cp(9), strFirstPriDate
      
'        strTemp = "" '法定期限
'        strTemp1 = "" '本所期限
'        '2006/1/27 MODIFY BY SONIA
'        'If pa(9) = 大陸國家代號 And strFirstPriDate <> "" Then
'        If (pa(9) = 大陸國家代號 Or pa(9) = "056") And strFirstPriDate <> "" Then
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'            StrSQLa = "Select NA27 From Nation Where NA01='" & pa(9) & "' "
'            rsA.CursorLocation = adUseClient
'            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount > 0 Then
'                strTemp = CompDate(1, rsA("NA27").Value, strFirstPriDate)
'                'strTemp = CompDate(2, -1, strTemp) 'Remove by Morgan 2006/11/9 大陸不必減一天
'                strDates(1) = cp(1)
'                strDates(2) = pa(9)
'                strDates(3) = TransDate(strTemp, 2)
'                GetCtrlDT strDates
'                strTemp1 = strDates(0)
'            End If
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'        'edit by nickc 2007/02/02 不用 dll 了
'        'ElseIf objPublicData.GetNationTaxEx(Val(pA(8)) + 3, pA(9), strTemp, strTemp1, , , False) = 0 Then
'        ElseIf ClsPDGetNationTaxEx(Val(pa(8)) + 3, pa(9), strTemp, strTemp1, , , False) = 0 Then
'             dobDateAdd = Val(strTemp1)
'             '2006/5/9 ADD BY SONIA 申請日要以畫面上的值計算
'             pa(10) = TransDate(Text5, 2)
'             '2006/5/9 END
'             strStartDate = GetStartDate(strTemp, cp(), pa())
'             If strStartDate = "" Then
'                strStartDate = TransDate(Text5, 2)
'             End If
'             If strStartDate <> "" Then
'                strStartDate = CompDate(1, dobDateAdd, strStartDate)
'                If pa(9) = "000" Then 'Add by Morgan 2006/11/9 加控制台灣才要減一天
'                  strStartDate = CompDate(2, -1, strStartDate)
'                End If
'                strTemp = strStartDate
'                strDates(1) = cp(1)
'                strDates(2) = pa(9)
'                strDates(3) = TransDate(strTemp, 2)
'                GetCtrlDT strDates
'                strTemp1 = strDates(0)
'            End If
'        End If
'        '若本所期限非工作天則抓最近的工作天
'        strTemp1 = PUB_GetWorkDay1(strTemp1, True)
'        '若取得本所及法定期限
'        If strTemp1 <> "" And strTemp <> "" Then
'            'Add by Morgan 2006/5/12 台灣改請發明時實審期限為"原實審期限"or"改請日+30"取較大者
'            If pa(9) = 台灣國家代號 And cp(10) = "301" Then
'               strExc(0) = CompDate(2, 30, cp(27))
'               If Val(strExc(0)) > Val(strTemp) Then
'                  strTemp = strExc(0)
'                  strDates(1) = pa(1)
'                  strDates(2) = pa(9)
'                  strDates(3) = TransDate(strTemp, 2)
'                  GetCtrlDT strDates
'                  strTemp1 = PUB_GetWorkDay1(strDates(0), True)
'               End If
'            End If
'
'            If m_strSitu = "A" Then
'                strTxt(intStep) = "Update CaseProgress Set CP06 = " & strTemp1 & ",CP07=" & strTemp & " Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10='" & 實體審查 & "' And CP27 IS NULL "
'                cnnConnection.Execute strTxt(intStep)
'                intStep = intStep + 1
'            ElseIf m_strSitu = "C" Then
'                strExc(0) = "Select NP01,NP07,NP22 From Nextprogress Where NP22= (SELECT MAX(NP22) FROM NEXTPROGRESS WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP07='416' AND NP06 IS NULL ) "
'                intI = 1
'                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'                If intI = 0 Then
'                  strTxt(intStep) = "declare intMax number;begin   select max(np22)+1 into intMax from nextprogress;IF intMax IS NULL THEN intMax:=1; END IF;"
'                  strTxt(intStep) = strTxt(intStep) & "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                                            " Values ('" & CP09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','416'," & strTemp1 & "," & strTemp & ",'" & strCP13New & "',intMax); "
'                  strTxt(intStep) = strTxt(intStep) & " end;"
'
'                  cnnConnection.Execute strTxt(intStep)
'                  intStep = intStep + 1
'                Else
'                  strTxt(intStep) = "update nextprogress set np08=" + strTemp1 + ",np09=" + strTemp + " WHERE NP01='" & RsTemp.Fields("NP01").Value & "' AND NP07='416' AND NP22=" & RsTemp.Fields("NP22").Value
'                  cnnConnection.Execute strTxt(intStep)
'                  intStep = intStep + 1
'                End If
'            End If
'        End If
   End If
   
   '若有輸入代理人D/N No, 帳單日期 及 帳單金額, 則新增國外帳單資料
   If Me.Text21.Text <> "" And Me.Text22.Text <> "" And Me.Text23.Text <> "" Then
       'Modify by Morgan 2008/5/13 +傳幣別
       'If PUB_AddNewFBillData(CP09, Me.Text21.Text, Me.Text22.Text, Me.Text23.Text, strBillNo) = False Then
       If PUB_AddNewFBillData(CP09, Me.Text21.Text, Me.Text22.Text, Me.Text23.Text, strBillNo, Combo1.Text) = False Then
           'Modified by Morgan 2016/6/30 錯誤訊息不可放在 Transaction 內
           'MsgBox "新增國外帳單資料作業失敗!!!", vbExclamation + vbOKOnly
           strErrMsg = "新增國外帳單資料作業失敗!!!"
           'end 2016/6/30
           Screen.MousePointer = vbDefault
           GoTo ErrorHandler
       Else
         'Added by Morgan 2016/6/30 非臺灣案電子化
         'Removed by Morgan 2025/8/13 帳單已全部都電子化
         'If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
         'end 2025/8/13
            '檢查帳單是否存在
            If PUB_CheckInvoicePDF(pa(1), pa(2), pa(3), pa(4), CP09, strErrMsg, , True, strBillNo) = False Then
               GoTo ErrorHandler
            End If
         'End If
         'end 2016/6/30
          'Modified by Lydia 2021/02/02 改成TextBox可複製
          'frm04010401.lblBillNo.Caption = "" & strBillNo
          frm04010401.txtBillno.Text = "" & strBillNo
       End If
   End If

            
   'Add by Morgan 2006/4/17 發明且有國外案時，用國內案預估公開日更新國外新案的期限
   'Modify by Morgan 2006/6/1 加判斷未收文主張國際優先權
   'Mofified by Morgan 2021/3/12 國內或國外案都要排除PCT案 Ex:P-126808 --郭
   If pa(8) = "1" And pa(46) = "" Then
      'Modified by Morgan 2015/1/19 澳門除外--郭 Ex.P-109280
      strSql = "SELECT CM01,CM02,CM03,CM04,CP06,CP07,CP09" & _
         " FROM CASEMAP,PATENT,CASEPROGRESS" & _
         " WHERE " & ChgCaseMap(pa(1) & pa(2) & pa(3) & pa(4), 0, 1) & " AND CM10='0'" & _
         " AND PA01(+)=CM01 AND PA02(+)=CM02 AND PA03(+)=CM03 AND PA04(+)=CM04 AND PA57 IS NULL AND PA09<>'044' AND PA46 IS NULL" & _
         " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP27 IS NULL AND CP31='Y' AND CP57 IS NULL" & _
         " AND NOT EXISTS(SELECT * FROM CASEPROGRESS WHERE CP01=CM01 AND CP02=CM02 AND CP03=CM03 AND CP04=CM04 AND CP10='106' AND CP57 IS NULL)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
      If intI = 1 Then
         '法定期限=預估公開日=申請日(最早優先權日)+18個月
         If strFirstPriDate <> "" Then
            strCP07 = CompDate(1, 18, TransDate(strFirstPriDate, 2))
         Else
            strCP07 = CompDate(1, 18, TransDate(Text5, 2))
         End If
         'Added by Lydia 2025/10/29
         If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
            strCP06 = PUB_GetPOurDeadline(strCP07, pa(9))
         Else
         'end 2025/10/29
            '本所期限=法定期限-10天
            strCP06 = PUB_GetWorkDay1(CompDate(2, -10, strCP07), True)
         End If 'Added by Lydia 2025/10/29
         If strCP06 < strSrvDate(1) Then
            strCP06 = strSrvDate(1)
         End If
         With RsTemp
            Do While Not .EOF
               strExc(1) = "" & .Fields("CM01")
               strExc(2) = "" & .Fields("CM02")
               strExc(3) = "" & .Fields("CM03")
               strExc(4) = "" & .Fields("CM04")
               'Added by Morgan 2013/7/1
               strExc(5) = "期限來源:" & Right("  " & pa(1), 3) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "(預定公開日);"
               If IsNull(.Fields("CP06")) Or strCP06 < "" & .Fields("CP06") Then
                  'Modified by Morgan 2013/7/1 +CP64
                  strSql = "Update caseprogress set CP06=" & strCP06 & ",CP07=" & strCP07 & _
                     ",CP64=SUBSTR(CP64,1,INSTR(CP64,'期限來源:')-1)||'" & strExc(5) & "'||SUBSTR(CP64,INSTR(CP64,';',instr(CP64,'期限來源:'))+1)" & _
                     " WHERE CP09='" & .Fields("CP09") & "'"
                  cnnConnection.Execute strSql
               End If
               .MoveNext
            Loop
         End With
      End If
   End If
   '2006/4/17 end
   
   'Add by Morgan 2006/5/12 若申請國家為PCT時掛119(進入國家階段)期限,法限=申請日(最早優先權日)+30月;所限=法限-2月
   If pa(9) = "056" And Text5 <> "" Then
      '起算日
      If strFirstPriDate <> "" Then
         strExc(0) = strFirstPriDate
      Else
         strExc(0) = TransDate(Text5, 2)
      End If
      '法限
      strExc(1) = CompDate(1, 30, strExc(0))
      '所限
      'Modified by Lydia 2025/10/31 屬於2025/10/29更新所限 ; 非台灣案(大陸、香港、澳門和PCT)都是相同算法---郭經理(電話)
      'strExc(2) = CompDate(1, -2, strExc(1))
      'strExc(2) = PUB_GetWorkDay1(strExc(2), True)
      'Added by Lydia 2025/11/06 檢查資料有申請國家為056的FMP(P-115159)，問Phoebe又沒有，就先補Code
      If m_bolFMP = True Then
         strExc(2) = CompDate(1, -2, strExc(1))
         strExc(2) = PUB_GetWorkDay1(strExc(2), True)
      Else
      'end 2025/11/06
         strExc(2) = PUB_GetPOurDeadline(strExc(1), pa(9), stNP23, pa(1), "119")
         'end 2025/10/31
      End If
      strSql = "Select NP01,NP07,NP22 From Nextprogress Where NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP07='119' AND NP06 IS NULL  ORDER BY NP22 DESC"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
      If intI = 1 Then
         'Modified by Lydia 2025/10/29 +NP23
         strSql = "update nextprogress set np08=" + strExc(2) + ",np09=" + strExc(1) + ",np23=" + IIf(stNP23 = "", "NP23", stNP23) + " WHERE NP01='" & RsTemp("NP01") & "' AND NP22=" & RsTemp("NP22")
      Else
         strSql = "declare intMax number;begin   select max(np22)+1 into intMax from nextprogress;IF intMax IS NULL THEN intMax:=1; END IF;"
         'Modified by Lydia 2025/10/29 +NP23
         strSql = strSql & "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22,np23) " & _
            " Values ('" & CP09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','119'," & strExc(2) & "," & strExc(1) & ",'" & strCP13New & "',intMax, " & CNULL(stNP23, True) & "); "
         strSql = strSql & " end;"
      End If
      cnnConnection.Execute strSql
   End If
   
   'Add by Morgan 2008/1/2
   'Modified by Morgan 2021/6/11 大陸發明新型110/6/1新法改最早優先權日起16個月,此處不可再更新
   'If pa(9) <> "000" And pa(9) <> "056" And Text20 = "Y" Then
   If pa(9) <> "000" And pa(9) <> "056" And Text20 = "Y" And Not (DBDATE(Text5) >= "20210601" And pa(9) = "020" And (pa(8) = "1" Or pa(8) = "2")) Then
      '法限=提申日+3個月
      strExc(1) = CompDate(1, 3, Text5)
      'Added by Lydia 2025/10/29
      If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
         strExc(2) = PUB_GetPOurDeadline(strExc(1), pa(9))
      Else
      'end 2025/10/29
      '所限=法限-10天
         strExc(2) = PUB_GetWorkDay1(CompDate(2, -10, strExc(1)), True)
      End If 'Added by Lydia 2025/10/29
      'Modified by Morgan 2021/6/11 +232
      strExc(0) = "select cp09 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in ('202','232') and cp27 is null and instr(cp64,'優先權證明')>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "update caseprogress set cp06=" & CNULL(strExc(2), True) & ",cp07=" & CNULL(strExc(1), True) & " where cp09='" & RsTemp.Fields(0) & "'"
         cnnConnection.Execute strSql, intI
      'Add by Morgan 2010/3/25 '下一程序也要更新
      Else
         strSql = "update nextprogress set np08=" & CNULL(strExc(2), True) & ",np09=" & CNULL(strExc(1), True) & " where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07 in ('202','232') and instr(np15,'優先權證明')>0 and np06 is null"
         cnnConnection.Execute strSql, intI
      End If
      
   End If
   
   'Added by Lydia 2015/05/01 P案國外指示信補件期限點選「轉讓證明」,自動產生一道Ｂ類收文(240補轉讓證明)承辦人掛PS2，期限設定為3個月,本所提前10天。
   str240Date = ""
   'Modified by Lydia 2021/04/01 FMP寰華國外指示信補件期限點選期限控管：選「優先權轉讓證明」，請自動產生期限於下一程序(240補優先權轉讓證明)
                                                 '輸入通知申請案號輸入時，若下一程序或進度檔有(240補優先權轉讓證明)未發文，請更新期限，法定期限設定為申請日加3個月，本所期限為法定期限前10天。
   'If pa(1) = "P" And pa(9) <> "000" And m_bolFMP = False And stCP47 <> "" Then
   If pa(1) = "P" And pa(9) <> "000" And stCP47 <> "" Then
      strExc(0) = "select cp09 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='240' and cp27 is null "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
        '若通知申請案號輸入時(240)補優先權轉讓證明仍未發文，請更新期限，期限設定為提申日加3個月,本所提前10天
         strExc(7) = CompDate(1, 3, stCP47) '法限不限工作天
         'Added by Lydia 2025/10/29
         If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
            strExc(6) = PUB_GetPOurDeadline(strExc(7), pa(9))
         Else
            strExc(6) = PUB_GetWorkDay1(CompDate(2, -10, strExc(7)), True)
         End If 'Added by Lydia 2025/10/29
         '通知副本時,將期限帶入定稿
         'Modified by Lydia 2021/04/01
         'If Text20.Text = "Y" Then
         If (m_bolFMP = False And Text20.Text = "Y") Or m_bolFMP = True Then
            'Modifed by Lydia 2021/04/01 改成西元年
            'str240Date = ChangeWStringToTString(strExc(6))
            'str240Date = Left(str240Date, 3) & "年" & Mid(str240Date, 4, 2) & "月" & Mid(str240Date, 6, 2) & "日"
            str240Date = Mid(strExc(6), 1, 4) & "年" & Mid(strExc(6), 5, 2) & "月" & Mid(strExc(6), 7, 2) & "日"
         End If
         
         strExc(0) = "UPDATE CASEPROGRESS set cp06=" & CNULL(strExc(6), True) & ",cp07=" & CNULL(strExc(7), True) & " where cp09=" & CNULL(RsTemp(0))
         cnnConnection.Execute strExc(0), intI
      End If
      'Added by Lydia 2021/04/01 FMP寰華國外指示信補件期限點選期限控管：若下一程序有(240補優先權轉讓證明)未收文，請將期限帶入定稿
      If str240Date = "" And m_bolFMP = True Then
          strExc(0) = "select np01,np22 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07='240' and np06 is null "
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
          If intI = 1 Then
            '請更新期限，期限設定為提申日加3個月,本所提前10天
             strExc(7) = CompDate(1, 3, stCP47) '法限不限工作天
             strExc(6) = PUB_GetWorkDay1(CompDate(2, -10, strExc(7)), True)
             '將期限帶入定稿
             str240Date = Mid(strExc(6), 1, 4) & "年" & Mid(strExc(6), 5, 2) & "月" & Mid(strExc(6), 7, 2) & "日"
             strExc(0) = "update nextprogress set np08=" & CNULL(strExc(6), True) & ", np09=" & CNULL(strExc(7), True) & " where np01=" & CNULL(RsTemp.Fields("np01")) & " and np22=" & CNULL(RsTemp.Fields("np22"))
             cnnConnection.Execute strExc(0), intI
          End If
      End If
      'end 2021/04/01
   End If
   'end 2015/05/01
   
   
   '2008/4/28 add by sonia 澳門發明掛第4年年費期限
   'Modify by Morgan 2011/8/29 +新型也要
   'If pa(9) = "044" And pa(8) = "1" Then
   If pa(9) = "044" And (pa(8) = "1" Or pa(8) = "2") Then
      strCaseProperty = ""
      If GetNP07(pa(9), pa(8), strTemp) Then
         strCaseProperty = strTemp
      End If
      If ClsPDGetNationTax(Val(pa(8)), pa(9), strTemp, strTemp1, strTemp2) Then
         '起算日為申請日
         strExc(0) = TransDate(Text5, 2)
         '法限
         yearTemp = GetMoneyYears(pa(72))
         varTemp = Split(strTemp1, ",")
         dobDateAdd = varTemp(yearTemp - 1)
         If strCaseProperty = "605" Then
            strExc(1) = CompDate(0, (dobDateAdd - 1), strExc(0))
         Else
            strExc(1) = CompDate(0, dobDateAdd, strExc(0))
         End If
         '所限
         strTemp = TransDate(strExc(1), 2)
         strDateS(1) = cp(1)
         strDateS(2) = pa(9)
         strDateS(3) = TransDate(strExc(1), 2)
         GetCtrlDT strDateS
         strTemp1 = strDateS(0)
         'Added by Lydia 2025/10/29
         stNP23 = ""
         If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
            strTemp1 = PUB_GetPOurDeadline(strTemp, pa(1), stNP23, pa(1), strCaseProperty)
         End If
         'end 2025/10/29
         strSql = "Select NP01,NP22 From Nextprogress Where NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP07='" & strCaseProperty & "' AND NP06 IS NULL  ORDER BY NP22 DESC"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            'Modified by Lydia 2025/10/29 +NP23
            strSql = "update nextprogress set np08=" + strTemp1 + ",np09=" + strTemp & ",NP23=" + IIf(stNP23 = "", "NP23", stNP23) + " WHERE NP01='" & RsTemp("NP01") & "' AND NP22=" & RsTemp("NP22")
         Else
            strSql = "declare intMax number;begin   select max(np22)+1 into intMax from nextprogress;IF intMax IS NULL THEN intMax:=1; END IF;"
            'Modified by Lydia 2025/10/29 +NP23
            strSql = strSql & "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22,NP23) " & _
               " Values ('" & CP09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & strCaseProperty & "'," & strTemp1 & "," & strTemp & ",'" & strCP13New & "',intMax," & CNULL(stNP23, True) & "); "
            strSql = strSql & " end;"
         End If
         cnnConnection.Execute strSql
      End If
   End If
   '2008/4/28 END
   
   'Add by Morgan 2009/9/4 大陸分割案期限控管
   If cp(10) = "307" And pa(9) = "020" And Val(txtSplitDate) > 0 Then
      st307Msg = PUB_Update307Ref(cp(9))
   End If
   
   'Add by Morgan 2009/11/19
   '大陸有主張國內優先權的案子,被主張的先申請案要自動上閉卷
   If pa(9) = "020" Then
      If PUB_ChkCPExist(cp, "121", 2) Then
         strSql = "update patent a set pa57='Y',pa58=" & strSrvDate(1) & ",pa59='88' where (pa01,pa02,pa03,pa04) in" & _
            "(select b.pa01,b.pa02,b.pa03,b.pa04 from pridate,patent b" & _
            " where pd01='" & cp(1) & "' and pd02='" & cp(2) & "' and pd03='" & cp(3) & "' and pd04='" & cp(4) & "'" & _
            " and pd07='020' and b.pa11(+)=pd06 and b.pa09=pd07 and b.pa57 is null)"
         cnnConnection.Execute strSql, intI
      End If
      
      'Add by Morgan 2010/5/26
      '輸入申請案號時新增補文件提申期限
      If Text6 <> "" Then
         strExc(0) = "select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "'" & _
            " and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='202' and cp27>0 and cp47 is null" & _
            " and not exists(select * from nextprogress where np01=cp09 and np06 is null and np07 in ('994','995','996','998'))"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Added by Lydia 2025/10/29 因為是專業部管制，不用NP23
            If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
               strExc(1) = PUB_GetPOurDeadline(strSrvDate(1), pa(9))
            Else
            'end 2025/10/29
                strExc(1) = CompWorkDay(5, strSrvDate(1))
            End If 'Added by Lydia 2025/10/29
            With RsTemp
            Do While Not .EOF
               strSql = "insert into nextprogress(np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
                  "select cp09,cp01,cp02,cp03,cp04,'998'," & strExc(1) & "," & strExc(1) & ",'" & strUserNum & "',np22" & _
                  " from caseprogress,(select max(np22)+1 np22 from nextprogress) where cp09='" & .Fields(0) & "'"
               cnnConnection.Execute strSql, intI
               .MoveNext
            Loop
            End With
         End If
      End If
   End If
   
   'Add by Morgan 2010/5/26
   '更新保密審查核准日
   If m_430CP09 <> "" Then
      If m_430CP25 <> "" Then
         strSql = "update caseprogress set cp24='1',cp25=" & m_430CP25 & " where cp09='" & m_430CP09 & "' and cp24 is null"
         cnnConnection.Execute strSql, intI
         PUB_430OkInform pa 'Add by Morgan 2010/6/2
      End If
      
      'Added by Morgan 2015/9/14 上已提申
      If m_430CP25 <> "" Or m_bolUpd430CP47 Then
         strSql = "update caseprogress set cp47=" & stCP47 & " where cp09='" & m_430CP09 & "' and cp47 is null"
         cnnConnection.Execute strSql, intI
         
         strSql = "update nextprogress set np06='Y' where np01='" & m_430CP09 & "' and np06 is null and np07='998'"
         cnnConnection.Execute strSql, intI
      End If
      'end 2015/9/14
   End If
     
    'Added by Lydia 2015/09/09 大陸發明案先掛下一程序999公開期限
    If pa(1) = "P" And pa(9) = "020" And pa(8) = "1" Then
       If InStr(NewCasePtyList, Text14) > 0 And Trim("" & pa(16) & pa(57) & pa(12) & pa(21)) = "" Then
          If PUB_GetOpenLimit020(pa(1), pa(2), pa(3), pa(4), strLDate, strLtitle) Then
             '公開期限的所限(工作天)=法限
             tmpCp06 = PUB_GetWorkDay1(strLDate, True)
             strExc(0) = "select * from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07='999' "
             intI = 1: strExc(7) = ""
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
               If IsNull(RsTemp.Fields("NP06")) Then
                    '未處理,更新期限
                   'Modified by Lydia 2016/1/14
'                  strSql = "UPDATE NEXTPROGRESS SET NP01='" & tmpCp09 & "', NP08=" & tmpCp06 & " , NP09=" & strLDate & ",NP15='" & strLtitle & "' " & _
'                           "WHERE  NP22=" & RsTemp.Fields("NP22") & " AND NP07='999' "
                  strSql = "UPDATE NEXTPROGRESS SET NP01='" & CP09 & "', NP08=" & tmpCp06 & " , NP09=" & strLDate & ",NP15='" & strLtitle & "' " & _
                           "WHERE  NP22=" & RsTemp.Fields("NP22") & " AND NP07='999' "
                  cnnConnection.Execute strSql, intI
               Else '已處理,另外新增
                  strExc(7) = "A"
               End If
             Else
               strExc(7) = "A"
             End If
            '掛下一程序公開999
             If strExc(7) = "A" Then
                  'Modified by Lydia 2016/1/14
               '  strSql = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22)" & _
                       " SELECT '" & tmpCp09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','999'" & _
                       "," & tmpCp06 & "," & strLDate & ",'" & strUserNum & "'" & _
                       ",'" & strLtitle & "',NP22 FROM (SELECT MAX(NP22)+1 NP22 FROM NEXTPROGRESS) X"
                 strSql = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22)" & _
                       " SELECT '" & CP09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','999'" & _
                       "," & tmpCp06 & "," & strLDate & ",'" & strUserNum & "'" & _
                       ",'" & strLtitle & "',NP22 FROM (SELECT MAX(NP22)+1 NP22 FROM NEXTPROGRESS) X"
                 cnnConnection.Execute strSql, intI
             End If
          End If
       End If
       '若該大陸案有香港案之關聯時，同時更新大陸案之公開期限至香港案之標準記錄請求的法定期限;
       Call PUB_UpdCP07by020(pa, m_bolFMP, "4")
    End If
    'end 2015/09/09
    
    'Added by Lydia 2018/07/09 為配合外專新案命名作業，針對香港案和澳門案寫入相應大陸案之名稱
    If pa(1) = "P" And pa(9) = "020" And InStr(NewCasePtyList, Text14) > 0 And Trim("" & pa(16) & pa(57) & pa(12) & pa(21)) = "" Then
         '香港
         strExc(1) = "": strExc(2) = ""
         tmpBol = ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), strExc(1), strExc(2), strExc(3), strExc(4), strExc(5), , "4")
         If tmpBol = True And strExc(1) <> "" And strExc(2) <> "" Then
              strSql = "update patent set pa05=" & CNULL(pa(5)) & ", pa06=" & CNULL(pa(6)) & ", pa07=" & CNULL(pa(7)) & _
                          " where PA01='" & strExc(1) & "' AND PA02='" & strExc(2) & "' AND PA03='" & strExc(3) & "' AND PA04='" & strExc(4) & "' "
              Pub_SeekTbLog strSql
              cnnConnection.Execute strSql, intI
         End If
         '澳門
         strExc(1) = "": strExc(2) = ""
         tmpBol = ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), strExc(1), strExc(2), strExc(3), strExc(4), strExc(5), , "5")
         If tmpBol = True And strExc(1) <> "" And strExc(2) <> "" Then
              strSql = "update patent set pa05=" & CNULL(pa(5)) & ", pa06=" & CNULL(pa(6)) & ", pa07=" & CNULL(pa(7)) & _
                          " where PA01='" & strExc(1) & "' AND PA02='" & strExc(2) & "' AND PA03='" & strExc(3) & "' AND PA04='" & strExc(4) & "' "
              Pub_SeekTbLog strSql
              cnnConnection.Execute strSql, intI
         End If
    End If
    'end 2018/07/09
    
   'Add by Sindy 2016/9/21
   If m_strIR01 <> "" Then
      'Add by Sindy 2018/9/7 信件自動歸至卷宗區
      'FMP案件
      'If pa(1) = "P" And Left(strCP12New, 1) = "F" And pa(9) <> 台灣國家代號 Then
      'Modify By Sindy 2018/12/4 不限FMP或P
      If pa(9) <> 台灣國家代號 Then
         '下載信件檔,上傳卷宗區(外來郵件)
         'Add By Sindy 2018/9/17 通知申請日輸入提申日，整封郵件存入卷宗區後，代理人來函匯入的地方就不需再管制ALTR缺檔
         'RX==>ALTR
         'Modified by Sindy 2019/1/21
         'If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, strOldcp09, "ALTR") = False Then
         If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, strOldcp09, IIf(Pub_StrUserSt03 = "F22", "ALTR", "PAT")) = False Then
         'end 2019/1/21
            GoTo ErrorHandler
         End If
'      '非台灣案通知申請日
'      ElseIf Trim(Text6) = "" And pa(9) <> 台灣國家代號 Then
'         '下載信件檔,上傳卷宗區(代理人來函)
'         If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, strOldcp09, "ALTR") = False Then
'            GoTo ErrorHandler
'         End If
      End If
      '2018/9/7 END
      
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", strOldcp09, "")
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010401", IIf(Pub_StrUserSt03 = "F22", strOldcp09, "")
   End If
   '2016/9/21 END
   
   'Added by Lydia 2022/04/28  工程師命名作業收文FMP主修和告代的承辦期限：(告代)通知申請日或通知申請案號起算6個工作天
   If m_bolFMP = True And (m_strCP10 = 通知申請日 Or m_strCP10 = 通知申請案號) Then
       Call Pub_GetFMPbCP48("2", cp, "901")
   End If
   'end 2022/04/28
   
   'Modified by Morgan 2017/9/12 從Transaction外移進來
   'Add by Morgan 2009/8/17
   '若歐盟設計的其他多國皆有申請日時提醒該案已可發文
   chk103in239OK cp
         
   cnnConnection.CommitTrans
   
   'Add by Sindy 2018/1/2
   If m_strIR01 <> "" And strBillNo <> "" Then
      MsgBox "已新增帳單【 " & strBillNo & " 】。", vbInformation
   End If
   '2018/1/2 END
   
   Screen.MousePointer = vbDefault
   FormSave = True
   
   'Add by Morgan 2009/7/15
   If st307Msg <> "" Then
      MsgBox st307Msg
   End If
   
   If m_1201CP09 <> "" Then g_PrtForm001.PrintCForm m_1201CP09, st1201CP64, stCP48Desc  'Added by Morgan 2014/3/5 FMP才會印
   
ErrorHandler:
   If FormSave = False Then
      cnnConnection.RollbackTrans
      If strErrMsg <> "" Then MsgBox strErrMsg, vbCritical 'Added by Morgan 2016/6/30
   End If
   Screen.MousePointer = vbDefault
End Function

Public Sub SetDefault()
   '若資料正確且申請國為台灣且申請日有資料, 則游標設定在申請案號欄
   If Len(Me.Text5.Text) > 0 And pa(9) = 台灣國家代號 _
      And Me.Text6.Visible And Me.Text6.Enabled Then
      Text6.SetFocus
   Else
        'Modify By Cheng 2003/01/24
        '設定輸入起始欄位
'      Text5.SetFocus
        If Me.Text5.Enabled Then
            Me.Text5.SetFocus
        Else
            If Me.Text6.Enabled And Me.Text6.Visible Then
                Me.Text6.SetFocus
            End If
        End If
   End If
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, i As Integer, strTmp As String, bolTmp As Boolean
'Add By Cheng 2002/11/01
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

   Screen.MousePointer = vbHourglass
   For Each Lbl In Label4
      Lbl = ""
   Next
   'Modify by Morgan 2009/11/19
   'Text15.Visible = False
   'Label14(1).Visible = False
   Frame1.Visible = False
   cmdDeadLine.Enabled = False 'Add by Morgan 2010/3/25
   Me.Height = Me.Height - 570
   'end 2009/11/19
   
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   Select Case pa(1)
      Case "P"
         If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
            If pa(8) <> "" Then
               If pa(9) = 台灣國家代號 Then
                  bolTmp = False
                  ' 90.07.10 modify by louis (非台灣國家才可修改案件性質)
                  EnableTextBox Text14, False
               Else
                  bolTmp = True
                  ' 90.07.10 modify by louis (非台灣國家才可修改案件性質)
                  EnableTextBox Text14, True
               End If
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetPatentTrademarkKind(專利, pA(8), strTmp, BolTmp, pA(9)) = 1 Then
               If ClsPDGetPatentTrademarkKind(專利, pa(8), strTmp, bolTmp, pa(9)) = 1 Then
                  Label4(2) = strTmp
               End If
            End If
            Text11 = pa(5)
            Text12 = pa(6)
            Text13 = pa(7)
            If pa(9) <> "" Then
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetNation(pA(9), strExc(0)) Then
               If ClsPDGetNation(pa(9), strExc(0)) Then
                  Label4(0) = strExc(0)
               End If
            End If
            
            'Modify by Morgan 2005/3/22 大陸案之申請日入改以西元年輸入
            'Text5 = pa(10)
            If pa(9) = 台灣國家代號 Then
               Text5.MaxLength = 7
               Text5 = pa(10)
            Else
               Text5.MaxLength = 8
               Text5 = TransDate(pa(10), 2)
            End If
            '2005/3/22 end

            Text6 = pa(11)
            
            If pa(46) = "Y" Then
               'Modify by Morgan 2009/11/19
               Text15.Visible = True
               Label14(1).Visible = True
               Frame1.Visible = True
               cmdDeadLine.Enabled = True 'Add by Morgan 2010/3/25
               Me.Height = Me.Height + 570
               'end 2009/11/19
            End If
            'Add By Cheng 2002/11/29
            Me.Text18.Text = "" & pa(46)
            'Removed by Morgan 2012/2/10 改抓CP 因為非新案的案件性質也可能有副本
            'Text20 = pa(140) 'Add by Morgan 2007/8/16
         End If
      Case "PS"
         If ClsPDReadServicePracticeDatabase(pa, intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadServicePracticeDatabase(pA, intWhere) Then
            Text11 = pa(5)
            Text12 = pa(6)
            Text13 = pa(7)
            If pa(8) <> "" Then
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetCustomerNation(pA(8), strTmp) Then
               '   If objPublicData.GetNation(strTmp, strExc(0)) Then
               If ClsPDGetCustomerNation(pa(8), strTmp) Then
                  If ClsPDGetNation(strTmp, strExc(0)) Then
                     Label4(0) = strExc(0)
                  End If
               End If
            End If
            If pa(9) = 台灣國家代號 Then
               ' 90.07.10 modify by louis (非台灣國家才可修改案件性質)
               EnableTextBox Text14, False
            Else
               ' 90.07.10 modify by louis (非台灣國家才可修改案件性質)
               EnableTextBox Text14, True
            End If
            Text5 = pa(10)
            'Modify by Morgan 2005/3/22 大陸案之申請日入改以西元年輸入
            If pa(9) = 台灣國家代號 Then
               Text5.MaxLength = 7
               Text5 = pa(10)
            Else
               Text5.MaxLength = 8
               Text5 = TransDate(pa(10), 2)
            End If
            '2005/3/22 end
            Text6 = pa(11)
         End If
   End Select
   
   cp(9) = CP09
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.ReadCaseProgressDatabase(cp, intWhere) Then
   If ClsPDReadCaseProgressDatabase(cp, intWhere) Then
      If cp(10) <> "" Then
         Text14 = cp(10)
         ChgType 14
      End If
      'Add by Morgan 2008/5/13
      text8 = cp(35)
      Text24 = cp(117)
      'end 2008/5/13
      
      'Added by Morgan 2012/2/10 改抓CP 因為非新案的案件性質也可能有副本
      Text20 = cp(145)
   End If
   
   If pa(1) = "P" And (IsNull(Text5) Or Text5 = "") Then
      'Modify by Morgan 2004/11/26 分割案要抓母案申請日
      'Text5 = strSrvDate(2)
      If cp(10) = "307" Then
         'Modify by Morgan 2005/3/22 大陸案之申請日入改以西元年輸入
         'Text5 = TransDate(PUB_DivAppDate(pa(1), pa(2), pa(3), pa(4), True), 1)
         If pa(9) = 台灣國家代號 Then
            Text5 = TransDate(PUB_DivAppDate(pa(1), pa(2), pa(3), pa(4), True), 1)
         Else
            Text5 = TransDate(PUB_DivAppDate(pa(1), pa(2), pa(3), pa(4), True), 2)
         End If
      Else
         'Modify by Morgan 2005/3/22 大陸案之申請日入改以西元年輸入
         'Text5 = strSrvDate(2)
         If pa(9) = 台灣國家代號 Then
            Text5 = strSrvDate(2)
         Else
            Text5 = strSrvDate(1)
         End If
         '2005/3/22 end

      End If
   End If
   
   strFirstPriDate = "": lblPriDate = ""
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   StrSQLa = "Select PD05 From PriDate Where PD01='" & pa(1) & "' AND PD02='" & pa(2) & "' AND PD03='" & pa(3) & "' AND PD04='" & pa(4) & "' AND PD05 IS NOT NULL ORDER BY PD05 "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       strFirstPriDate = "" & rsA("PD05").Value
       lblPriDate = ChangeWStringToWDateString(rsA("PD05").Value)
       rsA.MoveNext
       While Not rsA.EOF
           lblPriDate = lblPriDate.Caption & ", " & ChangeWStringToWDateString(rsA("PD05").Value)
           rsA.MoveNext
       Wend
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

   '若專利種類為發明時
   m_strSitu = ""
   'Modify by Morgan 2009/7/10 非發明案也會有實審
   'If pa(1) = "P" And pa(8) = "1" Then
   If pa(1) = "P" Then
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      '檢查本案是否有案件性質為416的資料
      'Modify by Morgan 2008/5/30 +判斷未取消收文
      'StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10='416' "
      StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10='416' and cp57 is null"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount <= 0 Then
          Me.Text16.Text = "N"
          Me.Text16.Enabled = False
          m_strSitu = "C"
      Else
          While Not rsA.EOF
              If "" & rsA("CP27").Value = "" Then
                  m_strSitu = "A"
              End If
              rsA.MoveNext
          Wend
          If m_strSitu = "" Then m_strSitu = "B"
      End If
   
      If pa(8) = "1" Then
        If rsA.State <> adStateClosed Then rsA.Close
        'Add by Morgan 2004/6/7
        '是否為分割案
        StrSQLa = "Select np08 From CaseProgress,nextprogress Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10='307' and np01(+)=cp09 and np07='416'"
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            m_bol307Exist = True
            m_st416NP08 = "" & rsA.Fields("np08")
        End If
        Set rsA = Nothing
      End If
    End If
   
   Screen.MousePointer = vbDefault

    '彼所案號
    If pa(9) <> 台灣國家代號 Then
        Me.Label23.Visible = True
        Me.Label23.Enabled = True
        Me.Text19.Visible = True
        Me.Text19.Enabled = True
        Me.Text19.Text = "" & cp(45)
    End If
    'add by nickc 2005/06/09 香港案標準專利紀錄請求 申請日=cm10='4' 的 申請日
    'edit by nickc 2005/07/21 香港時就要秀香港
    'If pa(9) = "013" And cp(10) = "110" Then
    'Modify by Morgan 2006/11/2 發明才要 --玲玲
    'If pA(9) = "013" Then
    If pa(9) = "013" And pa(8) = "1" Then
         Label20.Caption = "香港提交日:"
         If rsA.State <> adStateClosed Then rsA.Close
         StrSQLa = "Select * From casemap,patent Where cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "' and cm10='4' and cm05=pa01(+) and cm06=pa02(+) and cm07=pa03(+) and cm08=pa04(+) "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            Text5.Text = CheckStr(rsA.Fields("pa10"))
         'Add by Morgan 2010/8/20 沒有大陸案則帶基本檔申請日(不可預設系統日)
         Else
            Text5 = TransDate(pa(10), 2)
         End If
         Set rsA = Nothing
    Else
         Label20.Caption = "PCT提交日:"
    End If
    'Add by Morgan 2004/3/15
    SplitCheck
    '讀取優先權資料
    'Modify by Morgan 2007/4/25 加strPriority(4)
    'Modify by Amy 2014/04/11 +strPriority(5)
    ClsPDReadPriority pa, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)
    
    'Add by Morgan 2008/4/9 +控制台灣案7碼,其他8碼 P86935
    If pa(9) = "000" Then
      Text15.MaxLength = 7
    Else
      Text15.MaxLength = 8
    End If
    
   'Modified by Morgan 2021/1/21 從 Formsave 移來以便共用
   'Add by Morgan 2006/7/7
   '新增CP or NP時智權人員一律Call Function 抓最新收文智權人員
   strCP13New = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
   strCP12New = GetSalesArea(strCP13New)
   
    'Add by Morgan 2009/11/23
    'Modified by Morgan 2021/1/21
    'If Left(cp(12), 1) = "F" And pa(10) <> "000" Then
    'Modified by Lydia 2023/06/20 pa(10)=> pa(9)
    If Left(strCP12New, 1) = "F" And pa(9) <> "000" Then
    'end 2021/1/21
      m_bolFMP = True
    Else
      m_bolFMP = False
    End If
    
    
   'Added by Morgan 2012/4/24
   'Modified by Morgan 2012/5/16 +非台灣
   If pa(140) <> "" And pa(9) <> "000" Then
      lblFavDt.Visible = True
      txtFavDt.Visible = True
   Else
      lblFavDt.Visible = False
      txtFavDt.Visible = False
   End If
   'end 2012/4/24
End Sub

Private Sub Text5_GotFocus()
   InverseTextBox Text5
End Sub

Private Sub Text6_GotFocus()
   InverseTextBox Text6
   'edit by nickc 2007/07/11 切換輸入法改用API
   'Text6.IMEMode = 2
   CloseIme
End Sub

Private Sub Text15_GotFocus()
   InverseTextBox Text15
End Sub

Private Sub Text15_Validate(Cancel As Boolean)
   If Text15 <> "" Then
      'Add by Morgan 2008/4/9
      If pa(9) <> "000" And Len(Text15) <> 8 Then
         MsgBox "非台灣案補文件期限必須輸入西元年！"
         Cancel = True
         Exit Sub
      End If
      'end 2008/4/9
      If ChkDate(Text15) Then
         'Remove by Morgan 2009/11/23 改存檔檢查
         'If Val(DBDATE(Text15)) < Val(strSrvDate(1)) Then
         '   MsgBox "補文件期限不可小於系統日，請重新輸入 !", vbCritical
         '   Cancel = True
         'End If
         'end 2009/11/23
      Else
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse Text15
End Sub

Private Sub Text9_GotFocus()
   InverseTextBox Text9
End Sub

Private Sub Text10_GotFocus()
   InverseTextBox Text10
End Sub

Private Sub Text11_GotFocus()
   InverseTextBox Text11
End Sub

Private Sub Text12_GotFocus()
   InverseTextBox Text12
End Sub

Private Sub Text13_GotFocus()
   InverseTextBox Text13
End Sub

Private Sub Text14_GotFocus()
   InverseTextBox Text14
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim oBox As TextBox

TxtValidate = False

   'Added by Morgan 2021/12/16 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/16
   
If Me.Text13.Enabled = True Then
   Cancel = False
   Text13_Validate Cancel
   If Cancel = True Then
      Me.Text13.SetFocus
      Text13_GotFocus
      Exit Function
   End If
End If

If Me.Text14.Enabled = True Then
   Cancel = False
   Text14_Validate Cancel
   If Cancel = True Then
      Me.Text14.SetFocus
      Text14_GotFocus
      Exit Function
   End If
End If

If Me.Text5.Enabled = True Then
   Cancel = False
   Text5_Validate Cancel
   If Cancel = True Then
      Me.Text5.SetFocus
      Text5_GotFocus
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

If Me.Text17.Enabled = True Then
   Cancel = False
   Text17_Validate Cancel
   If Cancel = True Then
      Me.Text17.SetFocus
      Text17_GotFocus
      Exit Function
   End If
End If

'Add by Morgan 2004/3/15
If txtSplitDate.Visible = True Then
    If txtSplitDate = "" Then
        MsgBox "分割案提交日不可空白！", vbCritical
        txtSplitDate.SetFocus
        Exit Function
    Else
        Cancel = False
        txtSplitDate_Validate Cancel
        If Cancel = True Then
            Me.txtSplitDate.SetFocus
            txtSplitDate_GotFocus
            Exit Function
        'Added by Morgan 2023/12/12
        '台灣分割案提交日應該要等於發文日
        ElseIf pa(9) = "000" And txtSplitDate <> cp(27) Then
            If MsgBox("分割案提交日與發文日不同！是否確定？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
               txtSplitDate.SetFocus
               Exit Function
            End If
        'end 2023/12/12
        End If
    End If
End If
'2004/3/15

'Add by Morgan 2003/11/28
If (Text17.Text <> "") And (pa(9) = "020") And (Text18.Text <> "Y") Then
   MsgBox "必須為PCT案！", vbCritical
   'Remove by Morgan 2006/11/8
   'Me.Text18.SetFocus
   'Text18_GotFocus
   'end 2006/11/8
   Exit Function
End If
'End

'Add By Cheng 2003/12/22
'若有輸入代理人D/N No.或帳單日期
If Me.Text21.Text <> "" Or Me.Text22.Text <> "" Then

   'Modify By Sindy 2009/06/17 若為專利處只須以代理人+代理人D/N No.做重覆檢核
   If Text1 = "P" And Left(Trim(GetStaffDepartment(strUserNum)), 2) = "P1" Then
      '若有輸入代理人D/N No.
      If Me.Text21.Text <> "" Then
         If PUB_ChkDNDup("", ChangeCustomerL(cp(44)), Text21.Text) = True Then
            Text21.SetFocus
            Text21_GotFocus
            Exit Function
         End If
      End If
   Else
   'Modify by Morgan 2006/4/26 改Call共用函數
   '    StrSQLa = "Select * From ACC150 Where  A1502=" & (Val(DBDATE(Me.Text22.Text)) - 19110000) & " And A1503='" & ChangeCustomerL(cp(44)) & "' And A1504='" & Me.Text21.Text & "' "
   '    rsA.CursorLocation = adUseClient
   '    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   '    If rsA.RecordCount > 0 Then
   '        MsgBox "此帳單資料重覆，請確認!!!", vbExclamation + vbOKOnly
   '        If rsA.State <> adStateClosed Then rsA.Close
   '        Set rsA = Nothing
   '        Exit Function
   '    End If
   '    If rsA.State <> adStateClosed Then rsA.Close
   '    Set rsA = Nothing
      If PUB_ChkDNDup(Text22.Text, ChangeCustomerL(cp(44)), Text21.Text) = True Then
         Text21.SetFocus
         Text21_GotFocus
         Exit Function
      End If
   '2006/4/26 end
   End If
   
   'Added by Morgan 2016/6/30 非臺灣案電子化
   'Removed by Morgan 2025/8/13 帳單已全部都電子化
   'If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
   'end 2025/8/13
      '匯入該案帳單電子檔
      If Not PUB_ImportInvoice(pa(1), pa(2), pa(3), pa(4)) Then
         Exit Function
      End If
   'End If
   'end 2016/6/30
End If

'add by nickc 2005/06/29
'Modify by Morgan 2006/11/2 只有發明--玲玲
'If pA(9) = "013" Then
If pa(9) = "013" And pa(8) = "1" Then
    'edit by nickc 2005/07/21 只控制標準專利請求
    'If Text17 = "" Then
    If Text17 = "" And Text14 = "110" Then
        MsgBox "香港提交日不可空白！", vbCritical
        Text17.SetFocus
        Exit Function
    Else
        Cancel = False
        Text17_Validate Cancel
        If Cancel = True Then
            Me.Text17.SetFocus
            Text17_GotFocus
            Exit Function
        End If
    End If
End If

'Add by Morgan 2005/11/4
If Me.Text19.Visible And Me.Text19.Enabled Then
   If pa(9) <> "000" And Text19 = "" Then
      MsgBox "非台灣案時彼所案號不可空白！", vbExclamation
      Text19.SetFocus
      Exit Function
   End If
End If
'2005/11/4

'Add by Morgan 2009/11/19
If Frame1.Visible = True Then
   Cancel = False
   Text25_Validate Cancel
   If Cancel = True Then
      Text25.SetFocus
      Text25_GotFocus
      Exit Function
   End If
   Cancel = False
   Text27_Validate Cancel
   If Cancel = True Then
      Text27.SetFocus
      Text27_GotFocus
      Exit Function
   End If
   
   If Text27 <> "" Then
      If Val(Text27) <= Val(strSrvDate(1)) Then
         MsgBox "補文件法定期限必須大於系統日！"
         Text25.SetFocus
         Text25_GotFocus
         Exit Function
      End If
   End If
   
   Cancel = False
   Text15_Validate Cancel
   If Cancel = True Then
      Text15.SetFocus
      Text15_GotFocus
      Exit Function
   End If
   
   If Text15 <> "" Then
      If Val(Text15) < Val(strSrvDate(1)) Then
         MsgBox "補文件本所期限不可小於系統日！"
         Text15.SetFocus
         Text15_GotFocus
         Exit Function
      End If
   End If
   
   If Text15 <> "" And m_si880017 = 0 Then
      MsgBox "有補件期限時必須點選【補件資料】輸入補件資料並按確定！"
      Exit Function
   End If
End If
'end 2009/11/19

'Add by Morgan 2010/5/26
m_430CP09 = ""
'Modified by Morgan 2015/9/14
'If pa(9) = "020" And Text6 <> "" Then
m_bolUpd430CP47 = False
If pa(9) = "020" Then
'end 2015/9/14
   strExc(0) = "select cp09,cp47 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='430' and cp24 is null and cp27>0 and cp57 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_430CP09 = RsTemp(0)
      'Added by Morgan 2015/9/14 +保密審查是否已提交檢查
      If Text6 = "" Then
         If IsNull(RsTemp("cp47")) Then
            intI = MsgBox("請確認是否有同時提交保密審查？", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
            If intI = vbCancel Then
               Exit Function
            ElseIf intI = vbYes Then
               m_bolUpd430CP47 = True
            End If
         End If
      Else
      'end 2015/9/14
         'Modified by Morgan 2013/5/13 +取消選項--敏惠
         intI = MsgBox("請確認是否有保密審查通知書？", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
         If intI = vbCancel Then
            Exit Function
         ElseIf intI = vbYes Then
            m_430CP09 = RsTemp(0)
            m_430CP25 = "?"
            Do While m_430CP25 = "?"
               m_430CP25 = InputBox("請輸入發文日：", "保密審查通知書", strSrvDate(1))
               If m_430CP25 = "" Then
                  Exit Function
               Else
                  If Not ChkDate(m_430CP25) Then
                     m_430CP25 = "?"
                  Else
                     m_430CP25 = DBDATE(m_430CP25)
                  End If
               End If
            Loop
         End If
      End If 'Added by Morgan 2015/9/14
   End If
End If
'end 2010/5/26

   'Added by Morgan 2012/4/24
   If txtFavDt.Visible = True Then
      If txtFavDt = "" Then
         MsgBox "請輸入優惠期日期！", vbExclamation
         txtFavDt.SetFocus
         Exit Function
      ElseIf DBDATE(txtFavDt) <> DBDATE(pa(140)) Then
         MsgBox "優惠期日期與分案不同！", vbExclamation
         txtFavDt.SetFocus
         Exit Function
      End If
   End If
   'end 2012/4/24
   
   'Added by Morgan 2012/5/3
   If cp(46) <> "" Then
      If txtSplitDate <> "" Then
         strExc(1) = txtSplitDate
         Set oBox = txtSplitDate
      ElseIf Text17 <> "" Then
         strExc(1) = Text17
         Set oBox = Text17
      Else
         strExc(1) = Text5
         Set oBox = Text5
      End If
      If DBDATE(strExc(1)) < DBDATE(cp(46)) Then
         MsgBox "提申日不可早於收達日！", vbExclamation
         If oBox.Visible And oBox.Enabled Then
            oBox.SetFocus
            TextInverse oBox
         End If
         Exit Function
      End If
   End If
   'end 2012/5/3
   
   'Added by Morgan 2020/1/16
   m_bolNoCP27 = False
   'Modified by Morgan 2020/3/10 +有申請號--茹曣
   'Removed by Morgan 2024/1/30 取消--郭
   'If pa(9) = "020" And Text9 <> "N" And Text6 <> "" Then
   '   If PUB_GetCustomerValue(pa(26), "CU182") = "Y" Then
   '      If MsgBox("請確認是否已收到公文正本？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
   '         m_bolNoCP27 = True
   '      End If
   '   End If
   'End If
   'end 2020/1/16
   
TxtValidate = True
End Function

Private Sub txtFavDt_GotFocus()
   TextInverse txtFavDt
   CloseIme
End Sub

Private Sub txtFavDt_Validate(Cancel As Boolean)
   If txtFavDt <> "" Then
      Cancel = Not ChkDate(txtFavDt)
   End If
End Sub

Private Sub txtSplitDate_GotFocus()
     TextInverse txtSplitDate
End Sub

Private Sub txtSplitDate_Validate(Cancel As Boolean)
   
   'Modify by Morgan 2005/4/21 加控制非台灣輸西元
   If pa(9) = 台灣國家代號 Then
      If CheckIsTaiwanDate(txtSplitDate) = False Then
          txtSplitDate_GotFocus
          Cancel = True
      End If
   Else
      If CheckIsDate(txtSplitDate) = False Then
          txtSplitDate_GotFocus
          Cancel = True
      End If
   End If
   
   'Add by Morgan 2005/11/2 修改時雙重檢查
   If Cancel = False Then
      If CheckReKey(txtSplitDate) = True Then
         txtSplitDate.Tag = txtSplitDate
      Else
         Cancel = True
      End If
   End If
End Sub
'Add by Morgan 2003/3/30
'檢查母案是否存在
Private Function CheckDivCase(ByRef stPA10) As Boolean

On Error GoTo flgErr

   Dim stSQL As String, rsQuery As New ADODB.Recordset
   
   stSQL = "select pa10 from patent, divisioncase where pa01(+)=dc05 and pa02(+)=dc06 and pa03(+)=dc07 and pa04(+)=dc08" & _
      " and dc01='" & Text1 & "' and dc02='" & Text2 & "' and  dc03='" & Text3 & "' and dc04='" & Text4 & "'"
   
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      stPA10 = "" & rsQuery.Fields(0)
      CheckDivCase = True
   End If
   
flgErr:
   Set rsQuery = Nothing
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
    
End Function

'2009/3/24 add by sonia
'檢查是否有未發文之補文件期限
Private Function Check202()

On Error GoTo flgErr

   Dim stSQL As String, rsQuery As New ADODB.Recordset
   
   stSQL = "select cp09,cp06,cp64 from caseprogress where cp01='" & Text1 & "' and cp02='" & Text2 & "' and cp03='" & Text3 & "' and cp04='" & Text4 & "' and cp10='202' and cp27 is null and cp57 is null "
   
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      If rsQuery.RecordCount > 1 Then
         MsgBox ("未發文補文件期限有一筆以上，輸入申請案號後請自行維護補文件期限 !"), vbCritical
      ElseIf "" & rsQuery.Fields("cp09") <> "" Then
         If MsgBox("是否更新未發文補文件期限？進度備註為 " & rsQuery.Fields("cp64"), vbYesNo + vbDefaultButton1) = vbYes Then
            m_has202CP09 = rsQuery.Fields("cp09")
         End If
      End If
   End If
   
flgErr:
   Set rsQuery = Nothing
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
    
End Function
'2009/3/24 end

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   If Combo1 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo1, Label25(0)) = False Then
      Cancel = True
      Combo1.SetFocus
   End If
End Sub
