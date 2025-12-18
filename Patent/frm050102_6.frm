VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050102_6 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文（補件及其他案件質或CPS案件）"
   ClientHeight    =   5748
   ClientLeft      =   336
   ClientTop       =   1008
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   8580
   Begin VB.TextBox txtAppDate 
      Height          =   270
      Left            =   7500
      MaxLength       =   7
      TabIndex        =   5
      Top             =   1830
      Width           =   885
   End
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   7455
      MaxLength       =   4
      TabIndex        =   73
      Top             =   2715
      Width           =   540
   End
   Begin VB.CommandButton CmdFav 
      Caption         =   "優惠期日期"
      Height          =   270
      Left            =   6420
      TabIndex        =   70
      Top             =   2125
      Width           =   1095
   End
   Begin VB.TextBox txtFavDt 
      Height          =   270
      Left            =   7515
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   8
      Top             =   2130
      Width           =   885
   End
   Begin VB.CheckBox chkChoose 
      Caption         =   "一併繳指定費"
      Height          =   180
      Index           =   2
      Left            =   7065
      TabIndex        =   69
      Top             =   1320
      Width           =   1395
   End
   Begin VB.TextBox txtChkRltDate 
      Height          =   270
      Left            =   7110
      MaxLength       =   8
      TabIndex        =   14
      Top             =   3030
      Width           =   975
   End
   Begin VB.CheckBox chkChoose 
      Caption         =   "一併提實審"
      Height          =   180
      Index           =   3
      Left            =   7065
      TabIndex        =   65
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdPriority 
      Caption         =   "輸入(&P)"
      Height          =   270
      Left            =   6135
      TabIndex        =   63
      Top             =   2430
      Width           =   972
   End
   Begin VB.TextBox txtPA57 
      Enabled         =   0   'False
      Height          =   270
      Left            =   4560
      MaxLength       =   1
      TabIndex        =   13
      Top             =   3015
      Width           =   405
   End
   Begin VB.CommandButton cmdCountry 
      Caption         =   "指定國家"
      Height          =   300
      Left            =   5310
      TabIndex        =   60
      Top             =   1170
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   4965
      MaxLength       =   1
      TabIndex        =   11
      Top             =   2715
      Width           =   375
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   4965
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2130
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   405
      Index           =   3
      Left            =   4416
      TabIndex        =   22
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   6468
      TabIndex        =   24
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5640
      TabIndex        =   23
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   7692
      TabIndex        =   25
      Top             =   70
      Width           =   800
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   14
      Left            =   5040
      TabIndex        =   2
      Top             =   1500
      Width           =   825
      VariousPropertyBits=   671107097
      Size            =   "1455;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   15
      Left            =   6480
      TabIndex        =   3
      Top             =   1500
      Width           =   330
      VariousPropertyBits=   671107097
      Size            =   "582;529"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   4
      Left            =   945
      TabIndex        =   12
      Top             =   2970
      Width           =   2430
      VariousPropertyBits=   671107099
      MaxLength       =   20
      Size            =   "4286;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   5
      Left            =   1680
      TabIndex        =   15
      Top             =   3285
      Width           =   6735
      VariousPropertyBits=   671107099
      MaxLength       =   100
      Size            =   "11880;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   6
      Left            =   1680
      TabIndex        =   16
      Top             =   3570
      Width           =   6735
      VariousPropertyBits=   671107099
      MaxLength       =   100
      Size            =   "11880;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   7
      Left            =   1680
      TabIndex        =   17
      Top             =   3855
      Width           =   6735
      VariousPropertyBits=   671107099
      MaxLength       =   100
      Size            =   "11880;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   8
      Left            =   1680
      TabIndex        =   18
      Top             =   4140
      Width           =   6735
      VariousPropertyBits=   671107099
      MaxLength       =   600
      Size            =   "11880;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   9
      Left            =   1680
      TabIndex        =   19
      Top             =   4425
      Width           =   6735
      VariousPropertyBits=   671107099
      MaxLength       =   600
      Size            =   "11880;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   10
      Left            =   1680
      TabIndex        =   20
      Top             =   4710
      Width           =   6735
      VariousPropertyBits=   671107099
      MaxLength       =   600
      Size            =   "11880;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   13
      Left            =   1680
      TabIndex        =   10
      Top             =   2715
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   12
      Left            =   3390
      TabIndex        =   1
      Top             =   1515
      Visible         =   0   'False
      Width           =   945
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1667;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   3
      Left            =   960
      TabIndex        =   9
      Top             =   2430
      Width           =   255
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "450;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   2
      Left            =   1680
      TabIndex        =   6
      Top             =   2130
      Width           =   375
      VariousPropertyBits=   671107099
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   300
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   1485
      Width           =   870
      VariousPropertyBits=   671107099
      MaxLength       =   8
      Size            =   "1535;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCaseField 
      Height          =   705
      Index           =   11
      Left            =   1665
      TabIndex        =   21
      Top             =   5010
      Width           =   6765
      VariousPropertyBits=   -1467987941
      ScrollBars      =   2
      Size            =   "11933;1244"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblAppDate 
      AutoSize        =   -1  'True
      Caption         =   "約定期限："
      Height          =   180
      Left            =   6570
      TabIndex        =   75
      Top             =   1890
      Width           =   900
   End
   Begin VB.Label lblCP113 
      AutoSize        =   -1  'True
      Caption         =   "工作時數："
      Height          =   180
      Index           =   18
      Left            =   6510
      TabIndex        =   74
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label12 
      Alignment       =   1  '靠右對齊
      Caption         =   "費用："
      Height          =   195
      Index           =   6
      Left            =   4455
      TabIndex        =   72
      Top             =   1545
      Width           =   555
   End
   Begin VB.Label Label12 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "點數："
      Height          =   180
      Index           =   8
      Left            =   5940
      TabIndex        =   71
      Top             =   1545
      Width           =   540
   End
   Begin VB.Label lblChkRltDate 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   6300
      TabIndex        =   67
      Top             =   3045
      Width           =   765
   End
   Begin VB.Label lblCaseFee 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   8100
      TabIndex        =   66
      Tag             =   "Y"
      Top             =   2970
      Width           =   255
   End
   Begin VB.Label Label37 
      Caption         =   "優先權資料："
      Height          =   180
      Left            =   4995
      TabIndex        =   64
      Top             =   2460
      Width           =   1080
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "是否閉卷：            （Y：閉卷）"
      Height          =   180
      Left            =   3600
      TabIndex        =   62
      Top             =   3060
      Width           =   2460
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "EPC："
      Height          =   180
      Left            =   4320
      TabIndex        =   61
      Top             =   1230
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "案件性質："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   59
      Top             =   1230
      Width           =   900
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   6
      Left            =   1080
      TabIndex        =   58
      Top             =   1230
      Width           =   495
   End
   Begin MSForms.Label lblCasePropertyName 
      Height          =   180
      Left            =   1635
      TabIndex        =   57
      Top             =   1230
      Width           =   2580
      VariousPropertyBits=   27
      Size            =   "1879;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否列印通知函：           （N:不印）"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   56
      Top             =   2760
      Width           =   2820
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "是否修改通知函內容：           (Y : Word)"
      Height          =   180
      Index           =   2
      Left            =   3105
      TabIndex        =   55
      Top             =   2760
      Width           =   3075
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "新案指示信日期："
      Height          =   180
      Left            =   1950
      TabIndex        =   54
      Top             =   1560
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "是否修改指示信內容：           (Y : Word)"
      Height          =   180
      Index           =   1
      Left            =   3105
      TabIndex        =   53
      Top             =   2175
      Width           =   3075
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "進度備註："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   52
      Top             =   5010
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "補件內容          (1.優先權證明文件 2.委任狀 3.讓渡書 4.其他)"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   51
      Top             =   2475
      Width           =   4665
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "對造號數:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   50
      Top             =   3060
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "對造名稱(中)："
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   49
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "對造名稱(英)："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   48
      Top             =   4485
      Width           =   1200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "對造名稱(日)："
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   47
      Top             =   4770
      Width           =   1200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "對造案件名稱(日)："
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   46
      Top             =   3915
      Width           =   1560
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "對造案件名稱(英)："
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   45
      Top             =   3630
      Width           =   1560
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "對造案件名稱(中)："
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   44
      Top             =   3345
      Width           =   1560
   End
   Begin VB.Label lblAgent 
      Height          =   255
      Left            =   2220
      TabIndex        =   43
      Top             =   1830
      Width           =   4155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否列印指示信：           （N:不印）"
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   42
      Top             =   2160
      Width           =   2850
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   41
      Top             =   1830
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Left            =   120
      TabIndex        =   40
      Top             =   1530
      Width           =   720
   End
   Begin VB.Label lblTrademarkKind 
      Height          =   180
      Left            =   5880
      TabIndex        =   33
      Top             =   570
      Width           =   2535
   End
   Begin MSForms.Label lblSalesName 
      Height          =   180
      Left            =   5955
      TabIndex        =   32
      Top             =   780
      Width           =   2460
      VariousPropertyBits=   27
      Size            =   "1879;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   5
      Left            =   5280
      TabIndex        =   31
      Top             =   990
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   4
      Left            =   5280
      TabIndex        =   30
      Top             =   780
      Width           =   615
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   3
      Left            =   5280
      TabIndex        =   29
      Top             =   570
      Width           =   495
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   2
      Left            =   1080
      TabIndex        =   28
      Top             =   990
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   1
      Left            =   1080
      TabIndex        =   27
      Top             =   780
      Width           =   3135
   End
   Begin VB.Label lblCaseField 
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   26
      Top             =   570
      Width           =   3135
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   39
      Top             =   780
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   38
      Top             =   990
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   37
      Top             =   570
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Top             =   780
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      Height          =   180
      Left            =   120
      TabIndex        =   35
      Top             =   570
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   34
      Top             =   990
      Width           =   900
   End
   Begin VB.Label lblCaseFees 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   8145
      TabIndex        =   68
      Top             =   3030
      Width           =   255
   End
End
Attribute VB_Name = "frm050102_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/6 改成Form2.0 (txtCaseField,lblSalesName...)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，是0:結束  1:回上一畫面
Dim intLeaveKind As Integer
'Add By Cheng 2003/09/16
Dim strCountry As String '存放EPC指定國家
Dim strAssignCountry As String 'Add by Morgan 2004/11/10 EPC指定國家
'Add by Morgan 2005/11/17
Dim strPriority(1 To 5) As String '優先權資料
Dim m_bolRePriDate As Boolean '優先權資料需重新輸入
Dim m_bolEPC7Up As Boolean '是否超過7個成員國
Dim m_990CP09 As String 'Added by Morgan 2016/12/5
'Add By Sindy 2018/1/8
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2018/1/8 END
Dim strOldFileName As String, strNewFileName As String 'Added by Morgan 2018/7/19 工程師上傳的客戶函檔名, 系統預設的客戶函檔名
Dim m_strAF01 As String, m_strLD18 As String 'Added by Morgan 2018/8/22
Dim m_bolEngLetter As Boolean, m_strSubject As String 'Added by Morgan 2018/9/6
Dim bolPost As Boolean 'Added by Morgan 2018/9/12 是否寄發客戶函
Dim bolRegMail As Boolean 'Added by Morgan 2018/12/6 是否掛號直寄
'Dim bolAddLP As Boolean 'Added by Morgan 2018/10/9 過渡期C類來函是否要新增LP 'Removed by Morgan 2022/6/16
Dim strCP09List  As String 'Added by Morgan 2020/8/14 子案指示信-子案總收文號
Dim strNP22 As String 'Added by Morgan 2021/9/2

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String, Optional ByVal ET02 As String)
 Dim strTxt(1 To 5) As String, intStep As Integer, i As Integer, j As Integer
 'Add by Morgan 2004/11/10
 Dim strEPCCountry As String
 Dim iPos As Integer '字元搜尋位置
 
   If ET02 = "" Then ET02 = cp(9)
   
   j = 0
   EndLetter ET01, ET02, ET03, strUserNum
   If txtCaseField(12) <> "" Then
      j = j + 1
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','其他日期','" & DBDATE(txtCaseField(12)) & "')"
   End If
   
   'Add by Morgan 2004/11/10
   If strAssignCountry <> "" Then
      strEPCCountry = PUB_GetNationName(strAssignCountry, 2)
      strEPCCountry = Replace(strEPCCountry, ",", ", ")
      i = 0: iPos = 0
      Do
         iPos = i
         i = i + 1
         i = InStr(i, strEPCCountry, ",")
      Loop While i > 0
      If iPos > 0 Then strEPCCountry = Left(strEPCCountry, iPos - 1) & " and" & Mid(strEPCCountry, iPos + 1)
      j = j + 1
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','EPC指定國家','" & strEPCCountry & "')"
   End If
   '2004/11/10 end
   
   'Add by Morgan 2009/7/17
   '一併提實審
   If chkChoose(3).Value = 1 Then
      j = j + 1
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','一併提實審','file a request for substantive examination and ')"
   End If
   
   'Add by Morgan 2011/4/14
   'EPC檢索報告218發文若有勾選一併提實審或指定費時要帶入信函內
   If field(9) = "221" And cp(10) = "218" Then
      strExc(1) = ""
      If chkChoose(3).Value = vbChecked And chkChoose(2).Value = vbChecked Then
         strExc(1) = "、請求實體審查及繳交指定費"
      ElseIf chkChoose(3).Value = vbChecked Then
         strExc(1) = "及請求實體審查"
      ElseIf chkChoose(2).Value = vbChecked Then
         strExc(1) = "及繳交指定費"
      End If
      If strExc(1) <> "" Then
         j = j + 1
         strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','同時辦理事項','" & strExc(1) & "')"
      End If
   End If
   
   'Added by Morgan 2014/1/7
   'modify by sonia 2024/7/18 印度040發明商業使用聲明改三年呈報一次，故印度不跑此句
   'If cp(10) = "930" Then
   If cp(10) = "930" And field(9) <> "040" Then
      j = j + 1
      strTxt(j) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','聲明年度','" & (strSrvDate(1) \ 10000 - 1) & "')"
   End If
   'end 2014/1/7
   
   If j > 0 Then
      'edit by nickc 2007/02/05 不用 dll 了
      'If Not objLawDll.ExecSQL(j, strTxt) Then
      If Not ClsLawExecSQL(j, strTxt) Then
         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
      End If
   End If
End Sub

Private Sub cmdCountry_Click()
   'Modified by Morgan 2020/8/14 +"1"
   'Modified by Morgan 2023/3/7 改傳入案件性質
   ModifyLicenceCountry strCountry, strAssignCountry, field(10), cp(10)
End Sub

Private Sub cmdOK_Click(Index As Integer)
 Dim i As Integer, bolTmp As Boolean, strTmp As String
 Dim stLetter As String 'Add by Morgan 2005/2/2
 Dim rsTmp As New ADODB.Recordset 'Added by Lydia 2017/06/06
 Dim strPA(1 To 4) As String 'Added by Lydia 2017/07/18
 Dim arrAF01() As String 'Added by Morgan 2010/8/14 子案指示信-子案總收文號
 Dim b930Ask As Boolean 'Added by Morgan 2020/12/11 是否有使用(商業使用聲明)
 
   Select Case Index
      Case 0, 3 '確定, 同時發文
      
         'Added by Morgan 2015/8/7
         If PUB_ChkFileNP(cp(9)) Then
            MsgBox "下一程序已有提申或收達期限，不可發文！"
            Exit Sub
         End If
         'end 2015/8/7
   
         Screen.MousePointer = vbHourglass
         For i = 0 To 11
            If i <> 1 Then
                If txtCaseField(i).Enabled Then
                      If CheckKeyIn(i) <> 1 Then
                         txtCaseField(i).SetFocus
                         txtCaseField_GotFocus (i)
                            'Modify By Cheng 2003/04/16
'                         Exit For
                            Screen.MousePointer = vbDefault
                            Exit Sub
                      End If
                End If
            'Add By Cheng 2002/08/19
            Else
               If CheckKeyIn(i) <> 1 Then
                  Me.Combo1.SetFocus
                    'Modify By Cheng 2003/04/16
'                  Exit For
                    Screen.MousePointer = vbDefault
                    Exit Sub
               End If
            End If
         Next i
         If i = 12 Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
            
            'Add by Morgan 2005/11/8
            If cmdPriority.Enabled = True Then
               If strPriority(1) = "" Or m_bolRePriDate = False Then
                  MsgBox "本案有優先權資料,請重新輸入以便與原資料檢核！", vbCritical
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            
         '2006/5/26 ADD BY SONIA 自請撤回413,放棄專利權429詢問是否要閉卷
         'Modified by Morgan 2023/3/3 自請撤回413改核准才問/自動閉卷--郭
         'If (cp(10) = "413" Or cp(10) = "429") And txtPA57 <> "Y" Then
         If cp(10) = "429" And txtPA57 <> "Y" Then
         'end 2023/3/3
         
            'Added by Lydia 2015/09/10 若自請撤回413的相關總收文號為新申請案時, 不必再詢問直接閉卷.
            'Removed by Morgan 2023/3/3 自請撤回413改核准才問/自動閉卷--郭
            'If cp(10) = "413" Then
            '   strExc(0) = "select cp27 from caseprogress where cp09='" & cp(43) & "' and cp10 in (" & NewCasePtyList & ") "
            '   intI = 1
            '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            '   If intI = 1 Then txtPA57 = "Y"
            'End If
            'end 2023/3/3
            
            If txtPA57 = "" Then
                'Modified by Morgan 2023/3/3 自請撤回413改核准才問/自動閉卷--郭
                'If MsgBox("發文案件性質為 自請撤回 或 放棄專利權,請問是否要閉卷 ？", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
                If MsgBox("發文案件性質為放棄專利權,請問是否要閉卷 ？", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
                'end 2023/3/3
                   txtPA57 = "Y"
                Else
                   txtPA57 = ""
                End If
            End If
            'end 2015/09/10
         End If
         '2006/5/26 END
         
         'Add by Morgan 2007/12/27
         If field(9) = EPC指定國家 And cp(10) = "224" And strAssignCountry = "" Then
            MsgBox "未輸入領證之指定國家 !", vbCritical
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         
         'Added by Morgan 2020/12/11
         '商業使用聲明指示信內容確認
         If cp(10) = "930" And txtCaseField(2) <> "N" Then
            b930Ask = False
            intI = MsgBox("是否有使用？", vbQuestion + vbYesNoCancel + vbDefaultButton3, "商業使用聲明指示信內容確認")
            If intI = vbCancel Then
               Screen.MousePointer = vbDefault
               Exit Sub
            ElseIf intI = vbYes Then
               b930Ask = True
            End If
         End If
         'end 2020/12/11
         
            If SaveDatabase Then
            
               'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
               PUB_CheckEMail cp(44), cp(116)
               If field(1) = "CFP" Then 'Added by Morgan 2018/6/12 Ex.CPS-00098
                  PUB_CheckEMail field(75), field(144)
                  If field(145) <> "" Then
                     PUB_CheckEMail field(75), field(145)
                  End If
                  
               Else
                  PUB_CheckEMail field(26), field(76)
                  If field(77) <> "" Then
                     PUB_CheckEMail field(26), field(77)
                  End If
               End If
               'end 2008/2/20
               
               '指示信
               If txtCaseField(2) <> "N" Then
                  'Modify by Morgan 2006/4/7 改預設27,00留給通知函用
                  'strTmp = "00" 'Add by Morgan 2004/11/10 預設值
                  strTmp = "27"
                  If Text1(0) = "Y" Then
                     bolTmp = True
                  Else
                     bolTmp = False
                  End If
                  Select Case cp(10)
                     'Added by Morgan 2022/1/28
                     Case "232"
                        strTmp = "31"
                     'end 2022/1/28
                     Case 補文件
                        Select Case txtCaseField(3)
                           Case "1" '優先權證明文件
                              strTmp = "31"
                           Case "2" '委任狀
                              strTmp = "32"
                           Case "3" '讓渡書
                              strTmp = "33"
                        End Select
                     Case "216"
                        strTmp = "31"
                     'Add by Morgan 2008/2/12 加判斷是否一併提實審，是否指定7個以上會員國
                     Case "215"
                        'Add by Morgan 2009/7/17 2009.4.1以後改單一費用
                        'Modify by Morgan 2010/10/26 改判斷申請案提交日
                        'If Val(DBDATE(field(10))) >= 20090401 Then
                        If ChkEpcNewLetter Then
                           strTmp = "31"
                        Else
                        'end 2009/7/17
                           'Add by Morgan 2010/8/23
                           '若沒指定則與分案相同
                           If strAssignCountry = "" Then
                              ClsPDReadCountry intCaseKind, cp(), strAssignCountry, True, False
                           End If
                           'end 2010/8/23
                           m_bolEPC7Up = PUB_CheckEPC(strAssignCountry, cp(1) & cp(2) & cp(3) & cp(4))
                           If m_bolEPC7Up Then
                              If chkChoose(3).Value = 0 Then
                                 strTmp = "29" '7個以上
                              Else
                                 strTmp = "30" '7個以上且一併提實審
                              End If
                           Else
                              If chkChoose(3).Value = 1 Then
                                 strTmp = "28" '一般且一併提實審
                              End If
                           End If
                        End If
                     'Added by Morgan2018/6/21 新增子案指示信--禧佩
                     Case "224" '指定國註冊費
                        If field(9) <> EPC指定國家 Then strTmp = "28"
                     
                     'Added by Morgan 2025/5/8
                     Case "249"
                        strTmp = "28"
                        
                     'Added by Morgan 2020/12/11
                     Case "930"
                        If b930Ask = True Then strTmp = "28"
                        
                        'add by sonia 2024/7/18 印度040發明商業使用聲明改三年呈報一次，指示信獨立
                        If field(9) = "040" Then
                           If b930Ask = True Then
                              strTmp = "30"   '有使用
                           Else
                              strTmp = "29"   '未使用
                           End If
                        End If
                        'end 2024/7/18
                        
                     'Added by Morgan 2021/6/3
                     Case "217"
                        If field(9) = "019" Then strTmp = "28"
                  End Select
                  
                  'Add by Morgan 2004/10/18 要印傳真封面
                  'Removed by Morgan 2018/10/22 取消傳真封面--慧汶
                  'Select Case cp(10)
                  '   '自請撤回 (413)
                  '   'Modify by Morgan 2005/5/20 加放棄專利權(429)
                  '   'Modify by Morgan 2006/4/13 加指定費(215)
                  '   'Modify by Morgan 2007/10/30 加公開費(217)
                  '   'Modify by Morgan 2007/12/29 加指定國註冊費(224)
                  '   'Modified by Morgan 2014/1/7 +930商業使用聲明--禧佩
                  '   Case "413", "429", "215", "217", "224", "930"
                  '      If bolTmp = True Then
                  '         NowPrint cp(9), "01", "99", False, strUserNum, , , True, stLetter, , , , , , , , , m_strAF01
                  '      Else
                  '         NowPrint cp(9), "01", "99", False, strUserNum, , , , , , , , , , , , , m_strAF01
                  '      End If
                  '      If m_strAF01 <> "" Then Sleep 1000 '等1秒以確保letterdemand不會發生dupe錯誤 Added by Morgan 2018/8/20
                  'End Select
                  'end 2018/10/22
                  
                  'Added by Morgan 2020/8/14 EPC指定國註冊費子案指示信
                  If field(9) = EPC指定國家 And cp(10) = "224" And Len(strCP09List) > 0 Then
                     arrAF01 = Split(strCP09List, ",")
                     For i = 0 To UBound(arrAF01)
                        '沒有例外欄位
                        NowPrint arrAF01(i), "01", "28", False, strUserNum, , , , , , , , , , , , , arrAF01(i)
                     Next
                     MsgBox "本程序不需指示信，子案指示信請至待處理區作業！", vbExclamation
                     
                  'Added by Morgan 2024/12/3
                  ElseIf field(9) = EPC指定國家 Then
                     '要帶子案的本所號
                     NowPrint m_strAF01, "01", strTmp, bolTmp, strUserNum, 0, stLetter, , , , , , , , , , , m_strAF01
                  Else
                  'end 2020/8/14
                  
                     StartLetter "01", strTmp
                     NowPrint cp(9), "01", strTmp, bolTmp, strUserNum, 0, stLetter, , , , , , , , , , , m_strAF01
                     
                  End If 'Added by Morgan 2020/8/14
                  
                        
                  'Added by Morgan 2018/8/22 CFP電子化
                  If bolTmp = True And m_strAF01 <> "" Then
                     frm1105_1.m_RecNo = m_strAF01
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".DATA.PDF"
                     frm1105_1.Show
                     If Text1(1).Text = "Y" Then
                        MsgBox "指示信編輯中，客戶函請至定稿維護修改！", vbExclamation
                        Text1(1).Text = ""
                     End If
                  End If
                  'end 2018/8/22
                  
                  
'Removed by Morgan 2012/3/7 不必詢問,需要時程序自行列印--甄妮
'                  If Combo1 <> "" Then
'                     strExc(0) = "Select FA04,FA05||' '||FA63||' '||FA64||' '||FA65,FA06,FA01||FA02 From FAGENT WHERE FA01||FA02='" & Left(Combo1 & "000", 9) & "'"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                       With RsTemp
'                       If MsgBox("代理人名稱(中)：" & .Fields(0).Value & Chr(10) & Chr(13) & _
'                                 "　　　　　(英)：" & .Fields(1).Value & Chr(10) & Chr(13) & _
'                                 "　　　　　(日)：" & .Fields(2).Value & Chr(10) & Chr(13) & Chr(10) & Chr(13) & _
'                                 "是否列印代理人小信封？", vbExclamation + vbYesNo) = vbYes Then
'                           PUB_PrintCase "" & .Fields(3).Value
'                       End If
'                       End With
'                     End If
'                  End If

               'Added by Morgan 2018/9/6 +有工程師的指示信
               ElseIf m_bolEngLetter Then
                  PUB_SendOrderLetterP m_strAF01, m_strSubject
               'end 2018/9/6
               
               End If
               '通知函
               If txtCaseField(13) <> "N" Then
                  If Text1(1) = "Y" Then
                     bolTmp = True
                  Else
                     bolTmp = False
                  End If
                  
                  StartLetter "01", "00", cp(9)
                  NowPrint cp(9), "01", "00", bolTmp, strUserNum, 0, , , , , , , , , , , , m_strLD18
                  'Added by Morgan 2018/8/22 CFP電子化
                  If Text1(1).Text = "Y" And m_strLD18 <> "" Then
                     frm1105_1.m_RecNo = m_strLD18
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".CUS.PDF"
                     frm1105_1.Show
                  End If
                  'end 2018/8/22
               End If
                       
               'Added by Lydia 2017/06/06 韓國案件若主張台灣案優先權，於分案及韓國案主張優先權106發文時檢查台灣案是否收文"優先權電子交換"(437)，若未收文請e-mail提醒業務收文。
               'Modified by Lydia 2021/09/13 排除設計案 + And field(8) <> "3"  ; ex. CFP-32691是韓國案，主張台灣案P-126542優先權
               If field(9) = "012" And cp(10) = "106" And cp(13) <> "" And field(8) <> "3" Then
                  'Modified by Lydia 2017/08/28 被主張之優先權案若尚未發文,優先權為案號
                  'strExc(0) = "SELECT PA01,PA02,PA03,PA04 From PRIDATE, PATENT WHERE PD01='" & field(1) & "' AND PD02='" & field(2) & "' AND PD03='" & field(3) & "' AND PD04='" & field(4) & "' AND PD07='000' AND PD06=PA11(+) AND PD05=PA10(+) AND PD07=PA09(+) "
                  strExc(0) = "SELECT NVL(P1.PA01,P2.PA01) PA01,NVL(P1.PA02,P2.PA02) PA02,NVL(P1.PA03,P2.PA03) PA03,NVL(P1.PA04,P2.PA04) PA04,PD06 From PRIDATE, PATENT P1,PATENT P2 " & _
                              "WHERE PD01='" & field(1) & "' AND PD02='" & field(2) & "' AND PD03='" & field(3) & "' AND PD04='" & field(4) & "' AND PD07='000' " & _
                              "AND PD06=P1.PA11(+) AND PD05=P1.PA10(+) AND PD07=P1.PA09(+) " & _
                              "AND SUBSTR(PD06,1,1)=P2.PA01(+) AND SUBSTR(PD06,-9,6)=P2.PA02(+) AND SUBSTR(PD06,-3,1)=P2.PA03(+) AND SUBSTR(PD06,-2)=P2.PA04(+) "
                  intI = 1
                  Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     rsTmp.MoveFirst
                     strExc(1) = ""
                     Do While Not rsTmp.EOF
                        'Modified by Lydia 2017/07/18 debug:要傳台灣案號
                        'If PUB_ChkCPExist(field, "437") = False Then
                        
                        'Added by Lydia 2017/08/28 中途轉本所案,一律發信
                        If "" & rsTmp.Fields("PD06") <> "" And Trim("" & rsTmp.Fields("PA01")) = "" Then
                            strExc(1) = strExc(1) & IIf(strExc(1) <> "", "、", "") & rsTmp.Fields("PD06")
                        Else
                        'end 2017/08/28
                            strPA(1) = rsTmp.Fields("PA01"): strPA(2) = rsTmp.Fields("PA02")
                            strPA(3) = rsTmp.Fields("PA03"): strPA(4) = rsTmp.Fields("PA04")
                            If PUB_ChkCPExist(strPA, "437") = False Then
                            'end 2017/07/18
                               strExc(1) = strExc(1) & IIf(strExc(1) <> "", "、", "") & IIf(rsTmp.Fields("PA03") & rsTmp.Fields("PA04") = "000", rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02"), rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02") & "-" & rsTmp.Fields("PA03") & "-" & rsTmp.Fields("PA04"))
                            End If
                        End If 'end 2017/08/28
                        rsTmp.MoveNext
                     Loop
                     If strExc(1) <> "" Then
                         strExc(1) = IIf(field(3) & field(4) = "000", field(1) & "-" & field(2), field(1) & "-" & field(2) & "-" & field(3) & "-" & field(4)) & "主張國際優先權已發文,被主張之台灣案" & strExc(1) & "尚未收文優先權電子交換 !"
                         Call PUB_SendMail("", cp(13), "", strExc(1), "同主旨")
                     End If
                  End If
                  Set rsTmp = Nothing
               End If
               'end 2017/06/06
               
               bolLeave = True
               intLeaveKind = 1
               'Add By Cheng 2002/04/30
               '若有未發文資料顯示警告
               PUB_GetCPunIssueDatas "" & Me.lblCaseField(1).Caption
                'Add By Cheng 2003/11/27
                ' 發文回前畫面時
                Select Case Index
                   Case 0:
                      ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
                     'Add By Sindy 2013/5/28
                     If frm050102_1.bolIsEMPFlow = True Then
                        intLeaveKind = 0
                        'Unload frm050102_1
                        frm090202_4.Show
                        frm090202_4.QueryData
                     '2013/5/28 End
                     'Add By Sindy 2018/1/8
                     ElseIf Me.m_strIR01 <> "" Then
                        intLeaveKind = 0
                        'Modify By Sindy 2022/5/20
                        'frm04010519.GoNext
                        Forms(0).Tmpfrm04010519.GoNext
                        Set Forms(0).Tmpfrm04010519 = Nothing
                        '2022/5/20 END
                     '2018/1/8 END
                     Else
                        frm050102_1.Show
                        frm050102_1.Clear
                     End If
                   Case 3:
                        '若尚有未發文資料
                        If PUB_ChkUnissueDatas(Me.lblCaseField(1).Caption) = True Then
                            ' 90.07.12 modify by louis (回發文主畫面並重新查詢)
                           'Add By Sindy 2013/5/28
                           If frm050102_1.bolIsEMPFlow = True Then
                              frm090202_4.QueryData
                           'End If
                           '2013/5/28 End
                           'Add By Sindy 2018/1/8
                           ElseIf Me.m_strIR01 <> "" Then
                              'intLeaveKind = 0
                              'Modify By Sindy 2022/5/20
                              'frm04010519.GoNext
                              Forms(0).Tmpfrm04010519.GoNext
                              Set Forms(0).Tmpfrm04010519 = Nothing
                              '2022/5/20 END
                           '2018/1/8 END
                           End If
                           frm050102_1.Show
                           frm050102_1.ReQuery
                        '若無未發文資料
                        Else
                            ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
                           'Add By Sindy 2013/5/28
                           If frm050102_1.bolIsEMPFlow = True Then
                              intLeaveKind = 0
                              'Unload frm050102_1
                              frm090202_4.Show
                              frm090202_4.QueryData
                           '2013/5/28 End
                           'Add By Sindy 2018/1/8
                           ElseIf Me.m_strIR01 <> "" Then
                              intLeaveKind = 0
                              'Modify By Sindy 2022/5/20
                              'frm04010519.GoNext
                              Forms(0).Tmpfrm04010519.GoNext
                              Set Forms(0).Tmpfrm04010519 = Nothing
                              '2022/5/20 END
                           '2018/1/8 END
                           Else
                              frm050102_1.Show
                              frm050102_1.Clear
                           End If
                        End If
                End Select
                'End
               Unload Me

            '911202 nick
            Else
                MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
            End If
         End If
         Screen.MousePointer = vbDefault
      Case 1, 2
         'Add By Sindy 2013/5/28
         If frm050102_1.bolIsEMPFlow = True Then
            intLeaveKind = 0
            'Unload frm050102_1
            frm090202_4.Show
            frm090202_4.QueryData
         '2013/5/28 End
         'Add By Sindy 2018/1/8
         ElseIf Me.m_strIR01 <> "" Then
            intLeaveKind = 0
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         '2018/1/8 END
         Else
            If Index = 2 Then
               intLeaveKind = 0
            Else
               intLeaveKind = 1
            End If
         End If
         Unload Me
   End Select
    'Modify By Cheng 2003/11/27
    '本段程式往上移
'   ' 發文回前畫面時
'   Select Case Index
'      Case 0:
'         ' 90.07.12 modify by louis (回發文主畫面並清除畫面)
'         frm050102_1.Clear
'      Case 3:
'         ' 90.07.12 modify by louis (回發文主畫面並重新查詢)
'         frm050102_1.ReQuery
'   End Select
End Sub

Private Function SaveDatabase() As Boolean
Dim strTxt(1 To 10) As String, iStep As Integer
Dim strUpdate As String, i As Integer
Dim Tmp As String                        '2008/9/22 add by sonia
Dim strRecNo As String 'Add by Morgan 2009/5/13
Dim strDateS(3) As String
Dim strLetterJudge As String '指示信判發人/主旨 Added by Morgan 2018/8/22
Dim strAgentList As String, bolEPC224 As Boolean 'Added by Morgan 2020/8/14
Dim bolAdd250NP As Boolean 'Added by Morgan 2023/3/9 是否管制 UPC選擇退出

'911106 nick transation
SaveDatabase = True
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   cp(27) = txtCaseField(0)
   
   'Modify by Morgan 2008/2/20
   'cp(44) = Combo1
   intI = InStr(Combo1, "-")
   If intI > 0 Then
      cp(44) = Left(Combo1, intI - 1)
      cp(116) = Mid(Combo1, intI + 1)
   Else
      cp(44) = Combo1
      cp(116) = ""
   End If
   'end 2008/2/20
   cp(44) = ChangeCustomerL(cp(44))
   
   cp(36) = txtCaseField(4)
   cp(37) = txtCaseField(5)
   cp(38) = txtCaseField(6)
   cp(39) = txtCaseField(7)
   cp(40) = txtCaseField(8)
   cp(41) = txtCaseField(9)
   cp(42) = txtCaseField(10)
   cp(64) = txtCaseField(11)
   
   'Added by Morgan 2018/11/8
   If Val(txtCaseField(14)) > 0 Then
      cp(64) = ChangeWStringToWDateString(strSrvDate(1)) & " 報價：" & txtCaseField(14) & "(" & txtCaseField(15) & "P); " & cp(64)
   End If
   'end 2018/11/8
   
   cp(113) = txtCP113 'Added by Lydia 2021/05/25 工作時數
   
   'Modified by Morgan 2012/2/15 改呼叫共用函數
   'strExc(0) = "select cp45 from caseprogress where cp01=" + CNULL(cp(1)) + _
   '   " and cp02=" + CNULL(cp(2)) + " and cp03=" + CNULL(cp(3)) + _
   '   " and cp04=" + CNULL(cp(4)) + " and cp44=" + CNULL(cp(44)) + " order by cp27 desc"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   'cp(45) = ""
   'If intI = 1 And Not IsNull(RsTemp.Fields("CP45")) Then cp(45) = RsTemp.Fields("CP45")
   If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   'end 2012/2/15
   
   strTxt(1) = GetCPSQL(cp())
   
   
   '911106 nick transation
   cnnConnection.Execute strTxt(1)
   
   'Added by Lydia 2025/02/24 TIPS分配比例管制：與ACS案有關之智財協作發文時一併產生TIPS案請款階段分配比例
   If field(1) = "CPS" Then
      Call PUB_InsertACS_TIPS_Rate(field(1), field(2), field(3), field(4), cp(9), cp(10))
   End If
   'end 2025/02/24
   
   iStep = 2
   If intCaseKind = 專利 Then
      'Add By Cheng 2002/07/30
      If cp(10) = 異議_專 Or cp(10) = 舉發 Then
         field(5) = txtCaseField(5)
         field(6) = txtCaseField(6)
         field(7) = txtCaseField(7)
         strUpdate = strUpdate & ",pa05='" & ChgSQL(field(5)) & "'"
         strUpdate = strUpdate & ",pa06='" & ChgSQL(field(6)) & "'"
         strUpdate = strUpdate & ",pa07='" & ChgSQL(field(7)) & "'"
      End If
      
      Select Case cp(10)
         Case "301", "302", "303"
            field(8) = Mid(cp(10), 3, 1)
            strUpdate = strUpdate & ",pa08='" & field(8) & "'"
         Case "305"
            field(8) = "3"
            strUpdate = strUpdate & ",pa08='" & field(8) & "'"
         '2008/8/21 add by sonia 414回復原狀發文時專用期仍有效者才更新專利權是否存在pa1
         Case "414"
            If Val(field(25)) >= Val(strSrvDate(1)) Then
               strUpdate = strUpdate & ",pa17='Y'"
            End If
         '2008/8/21 end
      End Select
      
      '2008/9/22 add by sonia 改請時案件備註加註改請
      If Left(cp(10), 1) = "3" And cp(10) <> "307" Then
         Tmp = ""
         '先抓原申請案件性質
         strExc(0) = "SELECT CPM03 FROM caseprogress,CASEPROPERTYMAP WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & "and instr('" & NewCasePtyList & "',cp10)>0 and cp01=CPM01(+) and cp10=CPM02(+)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI > 0 Then
            Tmp = RsTemp.Fields(0)
         End If
         '再抓改請案件性質
         strExc(0) = "SELECT CPM03 FROM CASEPROPERTYMAP WHERE CPM01=" & CNULL(cp(1)) & " and CPM02=" + CNULL(cp(10))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI > 0 Then
            Tmp = ";" + ChangeTStringToTDateString(txtCaseField(0)) + Tmp + RsTemp.Fields(0)
         End If
         strUpdate = strUpdate & ",pa91=pa91||'" & Tmp & "'"
      End If
      '2008/9/22 end

'Remove by Morgan 2007/4/16 合併到下面一起做,因為程式有錯以前也都沒做過
'      strTxt(iStep) = GetPASQL(field())
'      '911106 nick transation
'      cnnConnection.Execute strTxt(1)
'      iStep = iStep + 1
'end 2007/4/16
   End If
   'end 2007/4/16
   
   '91106 nick transation
   'SaveDatabase = objLawDll.ExecSQL(iStep - 1, strTxt)
    'Add By Cheng 2003/09/16
    '若有ECP指定國家, 則新增案件進度檔資料
    If field(9) = EPC指定國家 And strCountry <> "" Then
         'Add by Morgan 2007/12/27
         If cp(10) = "224" Then
            'Modified by Morgan 2020/8/14
            'If PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strAssignCountry) Then
            strAgentList = PUB_GetAgentList(field(1), field(2), field(3), strAssignCountry)
            bolEPC224 = True
            If PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strAssignCountry, strAgentList, strCP09List, cp(10)) Then
            'end 2020/8/14
               Dim varTmp As Variant, pa04 As String
               varTmp = Split(strCountry, ",")
               For i = 0 To UBound(varTmp)
                  '未繳指定國註冊費的上閉卷
                  'Modified by Morgan 2023/3/8 224ＵＰ除外
                  If Format(varTmp(i)) <> "224" And InStr(strAssignCountry, Format(varTmp(i))) = 0 Then
                     pa04 = GetPA04(field(1), field(2), field(3), Format(varTmp(i)))
                     strSql = "UPDATE PATENT SET PA57='Y' WHERE PA01='" & field(1) & "' AND PA02='" & field(2) & "' AND PA03='" & field(3) & "' AND PA04='" & pa04 & "'"
                     cnnConnection.Execute strSql, intI
                  End If
               Next
               
               'Added by Morgan 2020/8/14 EPC指定國註冊費子案指示信
               Dim ArrCP09() As String, strSubject As String, arrCP(4) As String
               varTmp = Split(strAssignCountry, ",")
               ArrCP09 = Split(strCP09List, ",")
               For i = 0 To UBound(varTmp)
                  'Added by Morgan 2023/3/9
                  If InStr(UPMember, Format(varTmp(i))) > 0 Then
                     bolAdd250NP = True
                  End If
                  'end 2023/3/9
                  pa04 = GetPA04(field(1), field(2), field(3), Format(varTmp(i)))
                  strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), cp(10), Format(varTmp(i)))
                  strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), pa04, cp(10), field(11), , Format(varTmp(i)))
                  PUB_AddAppForm ArrCP09(i), True, strLetterJudge, strSubject
                  
                  arrCP(1) = cp(1): arrCP(2) = cp(2): arrCP(3) = cp(3): arrCP(4) = pa04
                  '催審期限
                  If txtChkRltDate <> "" Then
                     PUB_UpdateChkResultDate txtChkRltDate, arrCP, ArrCP09(i), "224"
                  End If
                  '提申期限
                  strExc(3) = ""
                  If cp(7) = "" Then
                     strExc(2) = PUB_Get224CtrlDate(2, strExc(3), cp)
                  Else
                     strExc(2) = cp(7)
                  End If
                  strExc(1) = PUB_Get224CtrlDate(3, strExc(3), cp)
                  PUB_SetApplyDate cp(1), cp(2), cp(3), pa04, strExc(2), ArrCP09(i), "224", txtCaseField(0), Format(varTmp(i)), strExc(1)
                  '收達
                  PUB_SetArriveDate ArrCP09(i)
               Next
               'end 2020/8/14
               
               'Added by Morgan 2023/3/9
               '管制 UPC選擇退出
               If bolAdd250NP Then
                  If DBDATE(txtCaseField(0)) <= 20300501 Then
                      '法限=2030/5/1
                      strExc(3) = "20300501"
                      '所限=2030/4/1
                      strExc(4) = "20300401"
                      If strExc(4) < strSrvDate(1) Then strExc(4) = strSrvDate(1)
                      strExc(4) = PUB_GetWorkDay1(strExc(4), True)
                      '智權人員
                      strExc(5) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
                      '流水號
                      strExc(6) = GetNextProgressNo()
                     
                      strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                         "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','250'," & strExc(4) & "," & strExc(3) & ",'" & strExc(5) & "'," & strExc(6) & ") "
                      cnnConnection.Execute strSql, intI
                  End If
               End If
               'end 2023/3/9
            End If
         Else
         'end 2007/12/27
            'Modify by Morgan 2006/4/7
            'If Not objPublicData.SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strCountry) Then
            If Not PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strCountry) Then
                  GoTo CheckingErr
            End If
         End If
    'Added by Morgan 2023/3/14
    'UP註冊 收文/催審/提申/收達
    ElseIf cp(10) = "249" Then
      '子案收文
      If Not PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), "224", , m_strAF01, cp(10)) Then
          GoTo CheckingErr
      End If
      pa04 = GetPA04(field(1), field(2), field(3), "224")
      strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), cp(10), "224")
      strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), pa04, cp(10), field(11), , "224")
      PUB_AddAppForm m_strAF01, True, strLetterJudge, strSubject
      
      arrCP(1) = cp(1): arrCP(2) = cp(2): arrCP(3) = cp(3): arrCP(4) = pa04
      '催審期限
      If txtChkRltDate <> "" Then
         PUB_UpdateChkResultDate txtChkRltDate, arrCP, m_strAF01, "224"
      End If
      '提申期限
      strExc(3) = ""
      If cp(7) = "" Then
         strExc(2) = PUB_Get224CtrlDate(2, strExc(3), cp, True)
      Else
         strExc(2) = cp(7)
      End If
      PUB_SetApplyDate cp(1), cp(2), cp(3), pa04, strExc(2), m_strAF01, cp(10), txtCaseField(0), "224"
      '收達
      PUB_SetArriveDate m_strAF01
      
   'UPC選擇退出
   ElseIf cp(10) = "250" Then
      strExc(1) = ""
      strExc(0) = "select pa09 from patent a where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & cp(3) & "' and pa04<>'00' and instr('" & UPMember & "',pa09)>0 and pa57 is null order by 1"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Do While Not RsTemp.EOF
            strExc(1) = strExc(1) & "," & RsTemp.Fields("pa09")
            RsTemp.MoveNext
         Loop
         strExc(1) = Mid(strExc(1), 2)
         '子案收文
         If Not PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), cp(9), strExc(1)) Then
             GoTo CheckingErr
         End If
      End If
   End If
   ' 90.12.05 modify by louis 若案件國家收費表存在代理人收達天數則新增一筆收達的下一程序檔
   If SaveDatabase = True Then
      'Modify by Morgan 2015/8/7 發文收達期限管控改呼叫公用函式
      'Modified by Morgan 2023/3/14 249 UP註冊也管制子案
      If bolEPC224 = False And cp(10) <> "249" Then 'Added by Morgan 2020/8/17 EPC指定國註冊費直接管制子案
         PUB_SetArriveDate cp(9)
      End If
      'end 2015/8/7
   End If
   
   'Add by Morgan 2007/4/16
   If txtPA57 = "Y" Then
      strUpdate = strUpdate & ",PA57='Y',PA58=" & strSrvDate(1) & ",PA59='09'"
   End If
   If strUpdate <> "" Then
      strSql = "UPDATE PATENT SET " & Mid(strUpdate, 2) & " WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'"
      cnnConnection.Execute strSql
   End If
   'end 2007/4/16
   
   'Add by Morgan 2005/2/2　閉卷
   If txtPA57 = "Y" Then
   
'Remove by Morgan 2007/4/16 移到上面
'      strSQL = "UPDATE PATENT SET PA57='Y',PA58=" & strSrvDate(1) & ",PA59='09' WHERE PA01='" & cp(1) & "' AND PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'"
'      cnnConnection.Execute strSQL
'end 2007/4/16

      strExc(0) = AutoNo("B", 6) 'B類總收文號
      strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp43,cp44,cp45,cp46,cp57,cp58) values " & _
         " ('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & _
         ",'" & strExc(0) & "','913','" & cp(12) & "','" & cp(13) & "','" & strUserNum & "','N','N'," & strSrvDate(1) & _
         ",'N','" & cp(9) & "','" & cp(44) & "','" & cp(45) & "','" & cp(46) & "'," & strSrvDate(1) & ",'09')"
      cnnConnection.Execute strSql
   End If
   
   'Add by Morgan 2005/11/17 有優先權資料且有重輸才要
   If strPriority(1) <> Empty And m_bolRePriDate = True Then
      'Modify by Amy 2014/04/17 +, strPriority(5)
      ClsPDSavePriority field, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)
   End If
   
   'Add by Morgan 2009/5/12
   '回覆代理人發文檢查若相關總收文號的相關總收文號已發文未提申時
   '以本程序發文日重算提申期限更新之
   If cp(10) = "902" Then
      strExc(0) = "select cp09,cp10 from caseprogress a where cp09=(select b.cp43 from caseprogress b where b.cp09='" & cp(43) & "') and cp27>0 and cp47 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strRecNo = RsTemp(0)
         If ClsPDGetCaseDelayDay(field(1), field(9), RsTemp("cp10"), , , Tmp) Then
            If Tmp <> "" Then
               '提申期限
               strExc(1) = CompDate(2, Val(Tmp), DBDATE(txtCaseField(0))) '法
               strExc(2) = PUB_GetWorkDay1(strExc(1), True) '所
               strExc(0) = "select * from nextprogress where np01='" & strRecNo & "' and np07='998' and np06 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & _
                     " where np01='" & RsTemp("np01") & "' and np07='998' and np22=" & RsTemp("np22")
                  cnnConnection.Execute strSql, intI
               Else
                  strSql = " insert into nextprogress a (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
                     " values('" & strRecNo & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'" & _
                     ",'998'," & strExc(2) & "," & strExc(1) & ",'" & strUserNum & "',GETNP22)"
                  cnnConnection.Execute strSql, intI
               End If
            End If
         End If
      End If
   End If
   
   'Add by Morgan 2009/8/18
   '催審
   If txtChkRltDate <> "" Then
      'Added by Morgan 2020/8/17 ECP指定國註冊費發文母案不管制催審，子案才要
      'Modified by Morgan 2023/3/14 249 UP註冊也管制子案
      If bolEPC224 = False And cp(10) <> "249" Then 'Added by Morgan 2020/8/17 EPC指定國註冊費直接管制子案
         PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9), cp(10), cp(43)
      End If 'Added by Morgan 2020/8/17
   End If
   
   'Add by Morgan 2013/5/29
   '提申管制
   'Modified by Morgan 2015/8/7 改呼叫共用
   'Added by Morgan 2020/8/17 EPC指定國註冊費直接管制子案
   If cp(10) = "224" Then
      '子案發文(正常不需要,因發母案自動新增子案)
      If bolEPC224 = False Then
         strExc(3) = ""
         If cp(7) = "" Then
            strExc(2) = PUB_Get224CtrlDate(2, strExc(3), cp)
         Else
            strExc(2) = cp(7)
         End If
         strExc(1) = PUB_Get224CtrlDate(3, strExc(3), cp)
         PUB_SetApplyDate cp(1), cp(2), cp(3), cp(4), strExc(2), cp(9), cp(10), txtCaseField(0), field(9), strExc(1)
      End If
   'Modified by Morgan 2023/3/14 249 UP註冊也管制子案
   'Else
   ElseIf cp(10) <> "249" Then
   'end 2023/3/14
      PUB_SetApplyDate cp(1), cp(2), cp(3), cp(4), cp(7), cp(9), cp(10), txtCaseField(0), field(9)
   End If 'Added by Morgan 2020/8/17
   'end 2015/8/7
   'end 2013/5/29
   
   'Add by Morgan 2010/6/3
   '主張優先權發文若需直譯本且尚未收文其他翻譯時掛期限=發文日+2個月
   If cp(10) = "106" And cp(71) = "Y" Then
      If PUB_ChkCPExist(cp, "927") = False Then
         strExc(1) = CompDate(1, 2, txtCaseField(0))
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01='" & cp(9) & "' and np06 is null and np07='927'"
         cnnConnection.Execute strSql, intI
         If intI = 0 Then
            strSql = "insert into nextprogress(np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
                  "select cp09,cp01,cp02,cp03,cp04,'927'," & strExc(2) & "," & strExc(1) & ",'" & cp(13) & "',np22" & _
                  " from caseprogress,(select max(np22)+1 np22 from nextprogress) where cp09='" & cp(9) & "'"
            cnnConnection.Execute strSql, intI
         End If
      End If
   End If
   'end 2010/6/3
   
   
   'Added by Morgan 2021/9/2 更新約定期限
   If strNP22 <> "" And txtAppDate.Tag <> txtAppDate.Text Then
      strSql = "update nextprogress set np23=" & DBDATE(txtAppDate) & " where np01='" & cp(9) & "' and np22=" & strNP22
      cnnConnection.Execute strSql, intI
   End If
   'end 2021/9/2

   'Add By Sindy 2015/8/3 發文時,若工程師各項日期未輸入者,自動更新為發文日
   Call PUB_UpdEmpDate(cp(9), cp(1), cp(10), DBDATE(cp(27)))
   
   'Added by Morgan 2016/12/5
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql, intI
   End If
   'end 2016/12/5
   
   'Add by Sindy 2018/1/8
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm050102_1"
   End If
   '2018/1/8 END
   
   'Added by Morgan 2018/2/9 +土耳其發明案實審法限先以申請檢索報告發文日+3個月管制--甄妮
'Removed by Morgan 2020/2/3 發文管控取消(只需以檢索報告通知日管控) --甄妮 Ex:CFP-030467
'   If field(9) = "235" And field(8) = "1" And cp(10) = "421" Then
'      strExc(1) = CompDate(1, 3, txtCaseField(0))
'
'      strDateS(1) = field(1)
'      strDateS(2) = field(9)
'      strDateS(3) = strExc(1)
'      GetCtrlDT strDateS
'      '所限
'      strExc(2) = strDateS(0)
'      strExc(2) = PUB_GetWorkDay1(strExc(2), True)
'
'      strSql = "Select cp09,cp27 From caseprogress Where cp01='" & cp(1) & "' AND cp02='" & cp(2) & "' AND cp03='" & cp(3) & "' AND cp04='" & cp(4) & "' AND cp10='416' and cp57 is null"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         If IsNull(RsTemp("cp27")) Then
'            strSql = "update caseprogress set cp06=" & strExc(2) & ",cp07=" & strExc(1) & " where cp09='" & RsTemp("cp09") & "'"
'            cnnConnection.Execute strSql
'         End If
'      Else
'         strSql = "Select NP01,NP07,NP22 From Nextprogress Where NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP07='416' AND NP06 IS NULL  ORDER BY NP22 DESC"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'         If intI = 1 Then
'            strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " WHERE NP01='" & RsTemp("NP01") & "' AND NP22=" & RsTemp("NP22")
'         Else
'            strSql = "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               " select '" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','416'," & strExc(2) & "," & strExc(1) & ",'" & cp(13) & "',newNP22 from dual,(select nvl(max(np22),0)+1 newNP22 from nextprogress)"
'         End If
'         cnnConnection.Execute strSql, intI
'      End If
'   End If
'end 2020/2/3
   'end 2018/2/9
   
   'Added by Lydia 2018/06/05 CFP案商業使用聲明發文時，判斷下一程序是否已掛隔年之商業使用聲明期限，若沒有則補掛一年。
   If field(9) = "040" And cp(10) = "930" Then
      'Modified by Morgan 2020/12/4 2020/10/20印度新法:法限改為9/30,所限=法限-1個月--禧佩
      'strExc(1) = Left(CompDate(0, 1, TransDate(txtCaseField(0), 2)), 4) & "0331" '次年商業使用聲明的法限
      'modify by sonia 2024/7/18 印度修改商業使用聲明每年呈報改為每三年呈報一次
      'strExc(1) = CompDate(0, 1, cp(7)) '次年商業使用聲明的法限
      strExc(1) = CompDate(0, 3, cp(7)) '次年商業使用聲明的法限
      'end 2020/12/4
      
      'modify by sonia 2024/7/18 2024/3修法每年呈報改為每三年呈報一次，且新期限大於專利檔的專用止日時則掛專用期止日當年或隔年之9/30
      '補商業使用聲明的法限不可大於專利檔的專用期限。(ex.CFP-019423年費結案不辦，但商業使用聲明又給我們辦。)
      'If strExc(1) <= field(25) And strExc(1) > strSrvDate(1) Then '若計算結果的法定期限<系統日則該期限不產生
      '    strSql = "select np01,np22,np06,np07 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' " & _
      '                "and np07='930' and np06 is null and np09=" & CNULL(strExc(1), True)
      '    intI = 1
      '    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      '    If intI = 0 Then
      '        'Modified by Morgan 2020/12/4 2020/10/20印度新法:法限改為9/30,所限=法限-1個月--禧佩
      '        'strExc(2) = PUB_GetWorkDay1(Mid(strExc(1), 1, 4) & "0131", True) '所限=明年1/31(抓工作天)
      '        strExc(2) = PUB_GetWorkDay1(CompDate(1, -1, strExc(1)), True) '所限=法限-1個月
      '        'end 2020/12/4
      '        strSql = "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
      '                    " select '" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','930'," & strExc(2) & "," & strExc(1) & ",'" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "',newNP22 from dual,(select nvl(max(np22),0)+1 newNP22 from nextprogress)"
      '        cnnConnection.Execute strSql, intI
      '    End If
      'End If
      If strExc(1) > field(25) Then   '新期限大於專利檔的專用止日時則掛專用期止日當年或隔年之9/30
         If cp(7) > field(25) Then    '若發文進度之法定已大於專利檔的專用期限則不必再掛下次期限
            strExc(1) = ""
         Else
            If Right(field(25), 4) <= "0930" Then       '判斷專用期止日是否<=當年9/30
               strExc(1) = Left(DBDATE(field(25)), 4)         '掛當年之9/30
            Else
               strExc(1) = Left(DBDATE(field(25)), 4) + 1     '掛次年之9/30
            End If
            strExc(1) = strExc(1) & "0930"
         End If
      End If
      If strExc(1) <> "" Then
         strSql = "select np01,np22,np06,np07 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' " & _
                  "and np07='930' and np06 is null and np09=" & CNULL(strExc(1), True)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 0 Then
            strExc(2) = PUB_GetWorkDay1(CompDate(1, -1, strExc(1)), True) '所限=法限-1個月
            strSql = "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     " select '" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','930'," & strExc(2) & "," & strExc(1) & ",'" & PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4)) & "',newNP22 from dual,(select nvl(max(np22),0)+1 newNP22 from nextprogress)"
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2024/7/18
   End If
   'end 2018/06/05
   
   'Added by Morgan 2018/8/22 CFP電子化
   If strSrvDate(1) >= CFP指示信電子化啟用日 Then
      'Modified by Morgan 2018/9/6 +有工程師的指示信
      If txtCaseField(2) <> "N" Or m_bolEngLetter = True Then
         'Added by Morgan 2020/8/14 母案不用指示信
         If field(9) = EPC指定國家 And cp(10) = "224" And Len(strCP09List) > 0 Then
            m_strAF01 = ""
         'Added by Morgan 2023/3/14
         '249 UP註冊在上面跑子案指示信
         ElseIf cp(10) = "249" Then
            m_strAF01 = m_strAF01
         'end 2023/3/14
         Else
         'end 2020/8/14
         
            If m_bolEngLetter Then
               strLetterJudge = strUserNum
            Else
               strLetterJudge = PUB_GetLetterJudgeNew("2", cp(1), cp(10), field(9))
            End If
            m_strSubject = PUB_GetSubject(cp(1), cp(2), cp(3), cp(4), cp(10), field(11), cp(45), field(9))
            PUB_AddAppForm cp(9), True, strLetterJudge, m_strSubject
            m_strAF01 = cp(9)
            
         End If 'Added by Morgan 2020/8/14
      End If
   End If
   'end 2018/8/22
   
   'Added by Morgan 2018/7/19 CFP電子化
   If strSrvDate(1) >= CFP第一階段電子化啟用日 Then
      If txtCaseField(13) = "N" And Left(cp(12), 1) <> "F" Then
         
         '工程師寫的客戶函都在歷程判發,信函進度直接上自行判發
         If cp(9) > "C" Then
            strExc(2) = ""
            'Modified by Morgan 2022/4/28
            'strExc(1) = PUB_SetLP11(field(26), field(75), strExc(2))
            strExc(1) = ""
            If bolRegMail Then
               strExc(1) = PUB_SetLP11(field(26), field(75), strExc(2))
            End If
            'end 2022/4/28
            
            If bolPost Then
               'Modified by Morgan 2018/10/12 若有齊備日lp03時要上判發日lp05
               'Modified by Morgan 2022/4/21 +LP52='Y'
               strSql = "update letterprogress set lp04=null,lp05=decode(lp03,0,0," & strSrvDate(1) & "),lp08='" & cp(14) & "',lp09=" & strSrvDate(1) & ",lp10='Y',lp11='" & strExc(1) & "',LP31='" & strExc(2) & "',LP52='" & IIf(bolRegMail, "Y", "") & "' where lp01='" & cp(9) & "'"
               cnnConnection.Execute strSql, intI
               
               'Added by Morgan 2022/4/21
               '掛號直寄若為E化案件(非全E)，已上發文室發文(QPGMR)要清除
               'Modified by Morgan 2022/11/11
               'If strExc(1) = "Y" Then
               '   strSql = "update caseprogress set cp154=null,cp127=null,cp128=null where cp09='" & cp(9) & "' and cp154='QPGMR' and exists(select * from letterprogress where lp01=cp09 and lp26='Y')"
               '   cnnConnection.Execute strSql, intI
               'End If
               PUB_ECaseAuotPost cp(9), cp(1)
               'end 2022/11/11
               'end 2022/4/21
            Else
               If txtCaseField(0) = "111111" Then
                  strExc(0) = "不通知客戶;"
               Else
                  strExc(0) = "不寄紙本通知函給客戶;"
               End If
               strSql = "update letterprogress set lp04=null,lp05=decode(lp03,0,0," & strSrvDate(1) & "),lp06='" & strUserNum & "',lp07=" & strSrvDate(1) & ",lp08='" & cp(14) & "',lp09=" & strSrvDate(1) & ",lp10='N',lp11='" & strExc(1) & "',lp12='" & strExc(0) & "',LP31='" & strExc(2) & "' where lp01='" & cp(9) & "'"
               cnnConnection.Execute strSql, intI
            End If
         Else
            cnnConnection.Execute "delete LetterProgress where lp01='" & cp(9) & "'", intI 'Added by Morgan 2018/12/6 可能會重新發文
            If bolPost Then
               'Modified by Morgan 2018/12/6
               'PUB_AddLetterProgress cp(9), 0, True, , IIf(cp(10) = "941", True, False), field(26), cp(10), field(75)
               PUB_AddLetterProgress cp(9), 0, True, , bolRegMail, field(26), cp(10), field(75)
               'end 2018/12/6
            'Added by Morgan 2018/10/12 詢問代理人(957)需發後補看也要新增LP
            ElseIf cp(10) = "957" Then
               PUB_AddLetterProgress cp(9), 0, False
            End If
         End If
      ElseIf txtCaseField(13) <> "N" Then
         strLetterJudge = PUB_GetLetterJudgeNew("1", field(1), cp(10), field(9))
         PUB_AddLetterProgress cp(9), 0, True, strLetterJudge, False, field(26), cp(10), field(75)
         m_strLD18 = cp(9)
      End If
   End If
   'end 2018/7/19

   'Added by Morgan 2018/12/6
   '有輸報價時要更新信函進度報價
   If Val(txtCaseField(14)) > 0 Then
      PUB_UpdateLP2930 cp(9), txtCaseField(14), txtCaseField(15)
   End If
   'end 2018/12/6
   
   cnnConnection.CommitTrans
     Exit Function
CheckingErr:
   SaveDatabase = False
   cnnConnection.RollbackTrans

End Function
Private Sub ReadAllData()
Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String, strTemp1 As String, j As Integer
Dim adoRecord As Object, strSameName As String

On Error GoTo ErrHnd

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'Modify by Morgan 2006/10/19 改不Call Dll
'If objPublicData.ReadAllData(frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.Row, 5),cp(), field(), intCaseKind, intPWhere) Then
ReDim cp(TF_CP) As String
cp(9) = frm050102_1.grdDataList.TextMatrix(frm050102_1.grdDataList.row, 5)
'Debug.Print "1:" & cp(9)
If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
'end 2006/10/19
   lblCaseField(0) = cp(9)
   lblCaseField(1) = cp(1) + " - " + cp(2) + _
      IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
      IIf(cp(4) = "00", "", " - " + cp(4))
   lblCaseField(2) = TransDate(cp(6), 1)
   lblCaseField(4) = cp(13)
   lblCaseField(5) = TransDate(cp(7), 1)
   'Add by Morgan 2004/11/10 加案件性質
   lblCaseField(6).Caption = cp(10)
   'edit by nickc 2007/02/02 不用 dll 了
   'Call objPublicData.GetCaseProperty(cp(1), lblCaseField(6), strTemp)
   Call ClsPDGetCaseProperty(cp(1), lblCaseField(6), strTemp)
   lblCasePropertyName.Caption = strTemp
   '2004/11/10
   
   If intCaseKind = 專利 Then
      lblCaseField(3) = field(8)
   End If
   'Modify By Cheng 2002/08/19
'   If objPublicData.GetCasePreAgent(cp(), strTemp) Then
'      txtCaseField(1) = strTemp
'      CheckKeyIn 1
'   End If
   Set adoRecord = CreateObject("ADODB.Recordset")
   lblCaseField(6).Caption = cp(10)
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.SelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
   '2007/4/23 MODIFY BY SONIA 加發文日降冪排序
   'If ClsPDSelectTable("select cp44 from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "'", adoRecord) Then
   'Modify by Morgan 2008/2/20 加聯絡人
   'Added by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設
   If cp(31) = "Y" Then
      AddAgent Combo1, cp, , , , cp(9), field(9), field(26)
      If Combo1 <> "" Then CheckKeyIn 1
      
   Else '非新案照原本
        If ClsPDSelectTable("select cp44||decode(cp116,null,null,'-'||cp116) from caseprogress where cp01 = '" & cp(1) & "' and cp02 = '" & cp(2) & "' and cp03 = '" & cp(3) & "' and cp04 = '" & cp(4) & "' and cp09<'C' and cp44 is not null order by cp27 desc", adoRecord) Then
        '2007/4/23 END
           Do While adoRecord.EOF = False
              If IsNull(adoRecord.Fields(0).Value) = False Then
                 If strSameName <> adoRecord.Fields(0).Value Then
                    Combo1.AddItem adoRecord.Fields(0).Value
                    strSameName = adoRecord.Fields(0).Value
                 End If
              End If
              adoRecord.MoveNext
           Loop
           Combo1 = Combo1.List(0)
        End If
        
      'Added by Morgan 2023/10/30 已有設定時不必再重新設定(IDS分案會先設,且抓預設代理人時也會剔除)
      If cp(44) <> "" Then
         Combo1 = cp(44) & IIf(cp(116) <> "", "-" & cp(116), "")
         CheckKeyIn 1
      Else
      'end 2023/10/30
      
        'edit by nickc 2007/02/02 不用 dll 了
        'If objPublicData.GetCasePreAgent(cp(), strTemp) Then
        If ClsPDGetCasePreAgent(cp(), strTemp) Then
           Combo1 = strTemp
           CheckKeyIn 1
        End If
        
         'Added by Morgan 2023/3/9 UP註冊要預設核准報價時指定的代理人
         If cp(10) = "249" Then
             strExc(0) = PUB_GetAgentList(cp(1), cp(2), cp(3), "224")
             Combo1 = strExc(0)
             If strExc(0) <> "" Then
                  CheckKeyIn 1
             Else
                  lblAgent = ""
             End If
         End If
         'end 2023/3/9
         
      End If 'Added by Morgan 2023/10/30
   End If
   'end 2016/10/27
   
   Text1(0) = "Y": Text1(1) = ""
   'Modify by Morgan 2007/10/30 加 公開費 217 --禧佩
   'Modified by Morgan 2014/1/7 +930商業使用聲明--禧佩
   'Modified by Morgan 2015/12/29 +126期末拋棄--禧佩
   'Modified by Morgan 2018/6/20 +224指定國註冊費(領證)--禧佩
   'Modified by Morgan 2018/10/12 +957詢問任代理人--慧汶
   'Modified by Morgan 2025/5/9 +417公開通知--玫音
   If cp(10) <> 補文件 And cp(10) <> "216" And cp(10) <> "217" And cp(10) <> "930" And cp(10) <> "126" And cp(10) <> "224" And cp(10) <> "957" And cp(10) <> "417" Then
      txtCaseField(2) = "N"
      
   'Added by Morgan 2018/10/31
   '非程序承辦的詢問任代理人預設N(工程師跑歷程有指示信 Ex:CFP-28773)--玫音
   ElseIf cp(10) = "957" And PUB_GetST03(cp(14)) <> "P12" Then
      txtCaseField(2) = "N"
      
   '2010/7/8 ADD BY SONIA C類也不要指示信
   ElseIf cp(9) > "C" Then
      txtCaseField(2) = "N"
   '2010/7/8 END
   End If

   'Modify by Morgan 2004/12/2 加 回覆檢索報告 218
   'Modify by Morgan 2006/3/7 加 申請檢索報告 421--慧汶
   'Modify by Morgan 2006/9/29 加 指定費 215 --禧佩
   'Modify by Morgan 2007/10/30 加 公開費 217 --禧佩
   'Modified by Morgan 2013/10/23 +408 面詢 --禧佩
   'Modified by Morgan 2013/12/25 +805 復審 --禧佩 Ex.CFP-024428
   'Modified by Morgan 2014/1/7 +930商業使用聲明--禧佩
   'MODIFY BY SONIA 2014/5/13 +423申請技術評價書--禧佩 CFP-026117
   'Modified by Morgan 2017/11/24 +906異同分析--禧佩 CFP-023700
   'Modified by Morgan 2018/2/9 +803舉發--禧佩 CFP-030116
   'Modified by Morgan 2018/3/13 +PPH431--禧佩 CFP-030373
   'Modified by Morgan 2023/2/7 +422加速審查及802異議答辯 --禧佩
   'Modified by Morgan 2023/10/26 +402變更 --禧佩
   'Modified by Morgan 2025/5/9 +417提早公開--玫音
   Select Case cp(10)
      Case 訴願, "424", "307", "218", "421", "215", "217", "408", "805", "930", "423", "906", "803", "431", "422", "802", "402", "417"
         
      Case Else
         txtCaseField(13) = "N"
   End Select
   txtCaseField(4) = cp(36)
   txtCaseField(5) = cp(37)
   txtCaseField(6) = cp(38)
   txtCaseField(7) = cp(39)
   txtCaseField(8) = cp(40)
   txtCaseField(9) = cp(41)
   txtCaseField(10) = cp(42)
   txtCaseField(11) = cp(64)
    'Add By Cheng 2003/02/20
    '若案件性質為補文件, 顯示新案指示信日期的輸入欄位
    If cp(10) = 補文件 Then
        Me.Label3.Visible = True
        Me.txtCaseField(12).Visible = True
    End If
    
    'Add By Cheng 2003/09/16
    '讀取ECP指定國家
    'Modified by Morgan 2018/9/12
    'Modified by Morgan 2023/3/16 249UP註冊、250UPC選擇退出除外
    If field(9) = EPC指定國家 And cp(9) < "C" And cp(10) <> "249" And cp(10) <> "250" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'objPublicData.ReadCountry intCaseKind, cp(), strCountry, True, False
      ClsPDReadCountry intCaseKind, cp(), strCountry, True, False
      'Add by Morgan 2004/11/10
      Label16.Visible = True
      cmdCountry.Visible = True
      'txtCaseField(2).Text = ""
   Else
      Label16.Visible = False
      cmdCountry.Visible = False
      '2004/11/10 end
   End If
   
   'Added by Lydia 2025/10/27 各專業部的智財協作發文，都預設不出定稿
   If field(1) = "CPS" And cp(10) = "967" Then
      txtCaseField(13) = "N"
   End If
   'end 2025/10/27
   
   'Add by Morgan 2005/2/2 新案的自請撤回預設閉卷
   txtPA57 = ""
   txtPA57.Enabled = False
'Removed by Morgan 2023/3/3 自請撤回413改核准才問/自動閉卷--郭
'   If cp(10) = 自請撤回 Then
'      strSql = "select CP31 from CASEPROGRESS WHERE CP09='" & cp(43) & "'"
'      intI = 0
'      'edit by nickc 2007/02/05 不用 dll 了
'      'Set RsTemp = objLawDll.ReadRstMsg(intI, strSQL)
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         If "" & RsTemp.Fields(0) = "Y" Then
'            txtPA57.Enabled = True
'            txtPA57 = "Y"
'            txtCaseField(2) = ""
'         End If
'      End If
'   End If
'end 2023/3/3
   '2005/2/2 end
   
   'Add by Morgan 2011/4/14
   'EPC回覆檢索報告218發文時若有實審或指定費當日或未發文時預設一併送件
   chkChoose(2).Enabled = False
   If field(9) = "221" Then
      If cp(10) = "218" Then
         chkChoose(2).Enabled = True
         strExc(0) = "select distinct cp10 from caseprogress where " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
            " and cp57 is null and cp10 in ('416','215') and (cp27 is null or cp27=" & strSrvDate(1) & ")" & _
            " order by 1"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.RecordCount = 2 Then
               chkChoose(2).Value = vbChecked
               chkChoose(3).Value = vbChecked
            ElseIf RsTemp(0) = "416" Then
               chkChoose(3).Value = vbChecked
            ElseIf RsTemp(0) = "215" Then
               chkChoose(2).Value = vbChecked
            End If
         End If
      ElseIf cp(10) = "215" Then
         strExc(0) = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & _
            " and cp57 is null AND CP10 IN ('218','416') and (cp27 is null or cp27=" & strSrvDate(1) & ")"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            txtCaseField(2) = "N"
            txtCaseField(13) = "N"
         End If
      End If
   End If
   
   'Add by Morgan 2009/8/18
   If txtCaseField(0).Tag <> txtCaseField(0) Then
      'Added by Morgan 2020/8/18
      '指定國註冊費催審期限不同
      If cp(10) = "224" Then
         txtChkRltDate.Enabled = False
         lblCaseFee.Enabled = False
         strExc(3) = ""
         strExc(1) = PUB_Get224CtrlDate(1, strExc(3), cp)
         txtChkRltDate = TransDate(strExc(1), 1)
      'Added by Morgan 2023/3/14
      'UP註冊
      ElseIf cp(10) = "249" Then
         txtChkRltDate.Enabled = False
         lblCaseFee.Enabled = False
         strExc(3) = ""
         strExc(1) = PUB_Get224CtrlDate(1, strExc(3), cp, True)
         txtChkRltDate = TransDate(strExc(1), 1)
      Else
      'end 2020/8/18
         PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8)
         
         'Added by Morgan 2023/6/13 俄羅斯申請紙本專利證書若未設定催審天數時再以領證的天數預設
         If field(9) = "023" And cp(10) = "443" And txtChkRltDate = "" Then
            PUB_SetChkResultDate cp(1), field(9), "601", txtCaseField(0), txtChkRltDate, cp, field(8)
         End If
         'end 2023/6/13
      End If
      txtCaseField(0).Tag = txtCaseField(0)
   End If
   
   'Added by Morgan 2012/4/24
   If cp(1) = "CFP" And cp(10) = "123" Then
     ' lblFavDt.Visible = True
      txtFavDt.Visible = True
      CmdFav.Visible = True 'Add by Lydia 2015/02/02
   Else
     ' lblFavDt.Visible = False
      txtFavDt.Visible = False
      CmdFav.Visible = False 'Add by Lydia 2015/02/02
   End If
   'end 2012/4/24
   
   'Add by Lydia 2014/12/30 EPC提供其他國家檢索報告之控管-回覆檢索報告(218)發文時,再次檢查所主張之案件狀態
   If field(9) = "221" And cp(10) = "218" Then
        strExc(5) = "": strExc(6) = "": strExc(7) = "": strExc(8) = ""
        strExc(0) = " select * from pridate where " & ChgPriDate(cp(1) & cp(2) & cp(3) & cp(4))
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            With RsTemp
            Do While Not .EOF
               strExc(1) = PUB_GetEPCPriDLatest(RsTemp!pd06, RsTemp!pd07, strExc(5), strExc(6), strExc(7), "2")
               If strExc(1) <> "" Then
                  strExc(8) = IIf(IsNull(strExc(8)), strExc(1), strExc(8) & ", " & strExc(1))
               End If
              .MoveNext
            Loop
            End With
          If Len(strExc(8)) > 0 Then
            strExc(0) = "本案之基礎案 " & strExc(8) & " 已有審查結果，請確認此發文是否有一併提交相關檢索報告！"
            MsgBox strExc(0), vbInformation
          End If
        End If
   End If
   'end 2014/12/30
   
   'Added by Morgan 2018/11/8
   'C類來函的特定案件性質要報價
   If cp(9) > "C" And InStr("1002,1006,1201,1202,1205,1206,1209,1231,1307,1401,1402,1801,1802,1812,1902", cp(10)) > 0 Then
      txtCaseField(14).Enabled = True
      txtCaseField(15).Enabled = True
   End If
   'end 2018/11/8
   
   
   'Added by Morgan 2024/10/21
   If field(9) = "221" And cp(10) = "1209" Then
      MsgBox "請確認檢索報告是否已公開，並確認通知函上的期限是否正確！", vbExclamation
   End If
   'end 2024/10/21
   
   'Added by Morgan 2021/9/2
   'C類來函檢查下一程序有約定期限的要帶出並於存檔加提醒"請確認約定期限是否須更改"
   strNP22 = ""
   txtAppDate.Enabled = False
   If cp(9) > "C" Then
      strExc(0) = "select np22,np23,cpm04 from nextprogress,casepropertymap" & _
         " where np01='" & cp(9) & "' and np23 is not null and cpm01(+)=np02 and cpm02(+)=np07 and np06 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strNP22 = RsTemp("np22")
         txtAppDate.Text = TransDate(RsTemp("np23"), 1)
         txtAppDate.Tag = txtAppDate.Text
         txtAppDate.Enabled = True
         strExc(0) = "本來函有【" & RsTemp("cpm04") & "】的約定期限 " & txtAppDate.Text & "，請確認是否須更改！"
         MsgBox strExc(0), vbInformation
      End If
   End If
   'end 2021/9/2

  
    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25
Else
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If

ErrHnd:
ErrorMsg
Screen.MousePointer = varSaveCursor
End Sub
'Add by Morgan 2005/11/17
Private Sub cmdPriority_Click()
   'Modify by Amy 2014/04/17 +, strPriority(5)
   If m_bolRePriDate = True Then
      '第二次不再檢查
      ModifyPriority strPriority(1), strPriority(2), strPriority(3), , , field(1) & field(2) & field(3) & field(4), , , strPriority(4), strPriority(5)
   Else
      m_bolRePriDate = True
      ModifyPriority strPriority(1), strPriority(2), strPriority(3), , m_bolRePriDate, field(1) & field(2) & field(3) & field(4), , , strPriority(4), strPriority(5)
   End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
   Dim strNo As String, iPos As Integer
   If Combo1.Text <> "" Then
      If CheckKeyIn(1) = -1 Then
         Cancel = True
      End If
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If Cancel = False Then
         strNo = Combo1.Text
         'Add by Morgan 2008/2/20 加聯絡人判斷
         iPos = InStr(Combo1.Text, "-")
         If iPos > 0 Then
            strNo = Left(Combo1.Text, iPos - 1)
         End If
         'end 2008/2/20
         
         If PUB_CheckStatus(strNo) = False Then
            Cancel = True
         'Added by Morgan 2012/3/7 發文都要顯示代理人備註--甄妮
         Else
            strExc(0) = "select FA29 from Fagent where " & ChgFagent(strNo) & " and FA29 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "" & RsTemp(0), vbExclamation, "代理人備註"
            End If
         'end 2012/3/7
         End If
      End If
      
      If Cancel Then Combo1.SetFocus
   End If
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String

Select Case Index
   Case 3
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
      If ClsPDGetPatentTrademarkKind(專利, lblCaseField(Index), strTemp, , 台灣國家代號) = 1 Then
         lblTrademarkKind = strTemp
      End If
   Case 4
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetStaff(lblCaseField(Index), strTemp) Then
      If ClsPDGetStaff(lblCaseField(Index), strTemp) Then
         lblSalesName = strTemp
      Else
         lblSalesName = ""
      End If
End Select
End Sub
Private Sub Form_Activate()
   Static bolActivated As Boolean 'Added by Morgan 2015/8/7

   'Added by Morgan 2015/8/7
   If bolActivated = True Then Exit Sub
   bolActivated = True
   'end 2015/8/7

   txtCaseField(0) = strSrvDate(2)
   ReadAllData
   
   If PUB_ChkFileNP(cp(9)) Then MsgBox "下一程序已有提申或收達期限，若為重新發文時需要先刪除後才可作業！" 'Added by Morgan 2015/8/7
   
   'Add By Cheng 2003/03/04
   'DoEvents 'Removed by Morgan 2023/9/5
   'Remove by Morgan 2008/1/2
   'txtCaseField(0).SetFocus
   'end 2008/1/2
   
   'Add by Morgan 2005/11/16 優先權輸入控制
   m_bolRePriDate = False
   If cp(10) = "106" Or cp(10) = "121" Then
      cmdPriority.Enabled = True
      '讀取優先權資料
      'Modify by Amy 2014/04/17 +, strPriority(5)
      ClsPDReadPriority field, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)
   Else
      cmdPriority.Enabled = False
   End If
   
   'Added by Morgan 2023/9/5
   If Left(cp(9), 1) = "B" Then
      'Added by Morgan 2023/9/11
      If txtCaseField(13) = "N" Then
         strExc(0) = PUB_AskBKindLetter(cp(1), cp(9), cp(10), 1)
      Else
      'end 2023/9/11
         strExc(0) = PUB_AskBKindLetter(cp(1), cp(9), cp(10))
         If txtCaseField(13) <> strExc(0) Then
            txtCaseField(13) = strExc(0)
            If txtCaseField(0).Enabled Then txtCaseField(0).SetFocus
         End If
      End If
   End If
   'end 2023/9/5

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   
   'Add By Sindy 2018/1/8
   m_strIR01 = frm050102_1.m_strIR01
   m_strIR02 = frm050102_1.m_strIR02
   m_strIR03 = frm050102_1.m_strIR03
   m_strIR04 = frm050102_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2018/1/8 END
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2009/8/18
   If intLeaveKind = 1 Then
      frm050102_1.Show
   ElseIf intLeaveKind = 0 Then
     Unload frm050102_1
   End If
   ShowEditForm 'Added by Morgan 2018/8/22
   'Set frm050102_6 = Nothing 'Removed by Morgan 2021/12/10 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If

End Sub


Private Sub txtAppDate_GotFocus()
   TextInverse txtAppDate
End Sub

Private Sub txtAppDate_Validate(Cancel As Boolean)
   If txtAppDate <> "" Then
      Cancel = Not ChkDate(txtAppDate)
   End If
End Sub

Private Sub txtCaseField_Change(Index As Integer)
   Select Case Index
      Case 1
         lblAgent = ""
        'Add By Cheng 2003/12/11
        Case 3 '補件內容
            If Me.txtCaseField(3).Text = "4" Then
                Me.txtCaseField(2).Text = "N"
                Me.Text1(0).Text = ""
                Me.txtCaseField(13).Text = "N"
                Me.Text1(1).Text = ""
            End If
        'End
   End Select
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 1, 2, 4, 13
         KeyAscii = UpperCase(KeyAscii)
      Case 3
         KeyAscii = UpperCase(KeyAscii)
'         If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
         If (KeyAscii > 52 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Added by Morgan 2018/11/8
      '費用,點數
      Case 14, 15
         If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub
Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
   If CheckKeyIn(Index) = -1 Then
      Cancel = True
   End If
   ' 90.12.07 modify by louis (加檢查字串欄位的長度)
   If CheckFieldLength(Index) = False Then
      Cancel = True
   End If
   If Cancel Then txtCaseField_GotFocus (Index)
End Sub
Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strCusTemp As String

CheckKeyIn = -1
Select Case intIndex
             Case 0
                         If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                            CheckKeyIn = 1
                           'Add by Morgan 2009/8/18
                           If txtCaseField(0).Tag <> txtCaseField(0) Then
                              If txtChkRltDate.Enabled Then 'Added by Morgan 2020/8/18
                                 PUB_SetChkResultDate field(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8)
                              End If
                              txtCaseField(0).Tag = txtCaseField(0)
                           End If
                        End If
             Case 1 '代理人
                        lblAgent = ""
                        If Combo1.Text = "" Then
                           MsgBox "代理人欄不可空白!!!", vbExclamation
                        Else
                           strCusTemp = Combo1
                           'Add by Morgan 2008/2/20 加判斷是否為聯絡人
                           If InStr(strCusTemp, "-") > 0 Then
                              If ClsPDGetContact(strCusTemp, strTemp) Then
                                 Combo1 = strCusTemp
                                 lblAgent.Caption = strTemp
                                 CheckKeyIn = 1
                              End If
                           
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetAgent(strCusTemp, strTemp) Then
                           ElseIf ClsPDGetAgent(strCusTemp, strTemp) Then
                              Combo1 = strCusTemp
                              lblAgent.Caption = strTemp
                              CheckKeyIn = 1
                           End If
                        End If
             Case 2, 13
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 3
                        If cp(10) = 補文件 And txtCaseField(2) <> "N" Then
                           If txtCaseField(intIndex).Text = "" Then
                              MsgBox "補件內容欄不可空白!!!", vbExclamation
                           Else
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
             Case 12
                        If cp(10) = 補文件 And txtCaseField(2) <> "N" Then
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              CheckKeyIn = 1
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
             Case Else
                        CheckKeyIn = 1
End Select
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
txtCaseField(Index).SelStart = 0
txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
End Sub

Private Function CheckFieldLength(ByVal nIndex) As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckFieldLength = True
   Select Case nIndex
      Case 5, 6, 7:
         If StrLength(txtCaseField(nIndex)) > 100 Then
            CheckFieldLength = False
         End If
      Case 8, 9, 10:
         If StrLength(txtCaseField(nIndex)) > 600 Then
            CheckFieldLength = False
         End If
      Case 11:
         If StrLength(txtCaseField(nIndex)) > 2000 Then
            CheckFieldLength = False
         End If
      Case Else:
         CheckFieldLength = True
   End Select
   
   If CheckFieldLength = False Then
      strTit = "檢核資料"
      strMsg = "輸入的資料內容太長"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
   End If
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

   'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   'add by nickc 2008/05/01
   If IsDebt(field(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        Exit Function
   End If
For Each objTxt In Me.txtCaseField
   If objTxt.Enabled = True Then
      Cancel = False
      txtCaseField_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

   'Added by Morgan 2012/4/24
   If txtFavDt.Visible = True Then
      If txtFavDt = "" Then
         MsgBox "請輸入優惠期日期！", vbExclamation
         txtFavDt.SetFocus
         Exit Function
      ElseIf DBDATE(txtFavDt) <> DBDATE(field(140)) Then
         MsgBox "優惠期日期與分案不同！", vbExclamation
         txtFavDt.SetFocus
         Exit Function
      End If
   End If
   'end 2012/4/24
   
'Add by Morgan 2004/9/14
If Combo1.Enabled = True Then
   Cancel = False
   Combo1_Validate Cancel
   If Cancel = True Then
      Combo1.SetFocus
      Exit Function
   End If
End If

'Added by Morgan 2018/11/8
If txtCaseField(14).Enabled And txtCaseField(14) = "" Then
   MsgBox "請輸入費用！" & vbCrLf & vbCrLf & "若無報價請輸 0！", vbExclamation
   txtCaseField(14).SetFocus
   Exit Function
End If
If txtCaseField(15).Enabled And txtCaseField(15) = "" And Val(txtCaseField(14)) > 0 Then
   MsgBox "請輸入點數！", vbExclamation
   txtCaseField(15).SetFocus
   Exit Function
End If
'end 2018/11/8

'Added by Morgan 2016/12/5
'檢查有設定副本收受人需提醒並新增信函副本B類收文
m_990CP09 = ""
bolPost = False
bolRegMail = False 'Added by Morgan 2018/12/6
'案件性質先比照P案--郭
'Modified by Morgan 2018/9/12  其他翻譯927不用--慧汶
'Modified by Morgan 2018/10/8 假發文除外
'Modified by Morgan 2018/12/6 +報告客戶956--郭 Ex:CFP-29087
'modify by sonia 2021/9/16 +檢視核准版本614--甄妮CFP-032130
If Left(cp(12), 1) <> "F" And txtCaseField(13) = "N" And (cp(9) > "C" Or InStr("941,903,956,614", cp(10)) > 0) And txtCaseField(0) <> "111111" Then
   
   'Modified by Morgan 2022/5/20 判斷是否全E化調整訊息
   strExc(1) = "是否寄通知函給客戶?"
   strExc(2) = "通知函是否掛號直寄?"
   'Modified by Morgan 2022/6/16 改用個案函數判斷
   If PUB_ChkECustCase(field(1), field(2), field(3), field(4)) = True Then
      strExc(1) = "是否要 EMail 通知函給客戶?" & vbCrLf & vbCrLf & "※本案為全E化客戶案件!!"
      strExc(2) = "通知函是否要確收?" & vbCrLf & vbCrLf & "※原紙本掛號直寄則為要確收!!"
   End If
   
   'C類都要報告(掛號)
   If cp(9) > "C" Then
      bolPost = True
      bolRegMail = True 'Added by Morgan 2018/12/6
      'Modified by Morgan 2018/10/17
      '代理人通知修正、其他來函、依職權電話通知修正,發文時詢問user要不要紙本發文--郭
      If InStr("1224,1225,1902", cp(10)) > 0 Then
         If MsgBox(strExc(1), vbYesNo + vbQuestion + vbDefaultButton1) = vbNo Then
            bolPost = False
            
         'Added by Morgan 2022/4/28
         '其他來函無期限不一定要掛號，改用問的--禧佩
         ElseIf cp(10) = "1902" And cp(7) = "" Then
            If MsgBox(strExc(2), vbYesNo + vbQuestion + vbDefaultButton1) = vbNo Then
               bolRegMail = False
            End If
         'end 2022/4/28
         
         End If
      End If
      'end 2018/10/17
      
   'AB類
   Else
'Modified by Morgan 2022/9/2 改一律不掛號並由智權同仁自行決定寄送給客戶的方式--郭
'      If MsgBox(strExc(1), vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
'         bolPost = True
'         If MsgBox(strExc(2), vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
'            bolRegMail = True
'         End If
'      End If
      bolPost = True
'end 2022/9/2
   End If
   'end 2022/5/20
   
   If bolPost Then
      If PUB_ChkCC(cp(1), cp(2), cp(3), cp(4), cp(9), m_990CP09) = False Then
         Exit Function
      End If
      
      'Added by Morgan 2018/9/12 CFP電子化
      If strSrvDate(1) >= CFP第一階段電子化啟用日 And Left(cp(12), 1) <> "F" Then
      
      
         'Added by Morgan 2022/6/17
         '全E化客戶案件要用原始檔(.CUS.DOC or .CUS.DOCX)轉有發文日的PDF檔上傳到卷宗區
         strOldFileName = ""
         'Modified by Morgan 2025/11/11 +半E化客戶
         If PUB_ChkECustCase(field(1), field(2), field(3), field(4), True) = True Then
            strExc(1) = ""
            If PUB_MakeCusPdf(cp(9), strExc(1)) = True Then
               strOldFileName = strExc(1)
            Else
               Exit Function
            End If
         End If
         If strOldFileName = "" Then
         'end 2022/6/17
         
            strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & cp(9) & "' and substr(upper(cpp02),-8)='.CUS.PDF' and cpp10<>'D'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strOldFileName = RsTemp(0)
            Else
               MsgBox "卷宗區無 "".CUS.PDF"" 的檔案，不可發文！", vbCritical
               Exit Function
            End If
            
         End If 'Added by Morgan 2022/6/17
         
         If m_990CP09 <> "" Then
            strExc(0) = "select cpp02 from casepaperpdf where cpp01='" & m_990CP09 & "' and substr(upper(cpp02),-8)='.CUS.PDF' and cpp10<>'D'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               strExc(1) = " and (substr(upper(cpf02),-8)='.CUS.DOC'"
               If cp(10) = "941" Or cp(10) = "1002" Then
                  strExc(1) = strExc(1) & " or substr(upper(cpf02),-8)='.RJN.DOC'"
               ElseIf cp(10) = "903" Then
                  strExc(1) = strExc(1) & " or (instr(upper(cpf02),'.903.SER')>0 and  and substr(UPPER(cpf02),-4)='.DOC')"
               End If
               strExc(1) = strExc(1) & ")"
               
               strExc(0) = "select cpf13 from casepaperfile where cpf01='" & cp(9) & "'" & strExc(1)
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(2) = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & ".990.CUS.PDF"
                  If PUB_Make990Pdf(cp(1), m_990CP09, RsTemp(0), strExc(2), IIf(cp(9) > "C", True, False)) = False Then
                     Exit Function
                  End If
               Else
                  MsgBox "信函副本產生失敗，無法讀取信函原始檔(.CUS.DOC)！", vbCritical
                  Exit Function
               End If
            End If
            
         End If
         
         'Removed by Morgan 2022/6/16
         'If DBDATE(cp(5)) < CFP第一階段電子化啟用日 Then
         '   'Modified by Morgan 2018/11/12 改都要補發文號
         '   intI = MsgBox("客戶通知函是否有印發文號：" & Right(cp(9), 6) & "？" & vbCrLf & vbCrLf & "若沒有請補上並重印客戶函後再發文!!", vbYesNo + vbQuestion + vbDefaultButton2)
         '   If intI = vbYes Then
         '      bolAddLP = True
         '   Else
         '      Exit Function
         '   End If
         'End If
         'end 2022/6/16
         
      End If
      'end 2018/9/12
   End If
End If
'end 2016/12/5

'Added by Morgan 2018/9/6
'若系統不出指示信時判斷是否有工程師的指示信要寄送
m_bolEngLetter = False
If txtCaseField(2) = "N" And cp(9) < "C" And InStr("941,903,956", cp(10)) = 0 Then
   If PUB_EngLtrChk(cp(9), txtCaseField(0).Text, m_bolEngLetter) = False Then
      Exit Function
   End If
End If
'end 2018/9/6

'Added by Morgan 2018/9/12 CFP電子化-接洽單檢查
If strSrvDate(1) >= CFP第一階段電子化啟用日 And Left(cp(12), 1) <> "F" Then
   If cp(9) < "B" Then
      If PUB_CheckPDF3(cp(1), cp(2), cp(3), cp(4)) = False Then
         Exit Function
      End If
   End If
End If
'end 2018/9/12

'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
If Pub_ChkACS112isNull(field(1), field(2), field(3), field(4), txtCP113) = True Then
    txtCP113.SetFocus
    txtCP113_GotFocus
    Exit Function
End If
'end 2021/05/25

'Added by Lydia 2023/11/17 針對PS及CPS之智財協作967，請一定要輸入工作時數才可過
If field(1) = "CPS" And cp(10) = "967" And Val(Trim(txtCP113)) <= 0 Then
    MsgBox "智財協作，請一定要輸入工作時數"
    txtCP113.SetFocus
    txtCP113_GotFocus
    Exit Function
End If
'end 2023/11/17
   
'Added by Morgan 2021/9/2
If txtAppDate.Enabled = True And Trim(txtAppDate) = "" Then
   MsgBox "約定期限不可空白！", vbExclamation
   txtAppDate.SetFocus
   Exit Function
End If
'end 2021/9/2

'Added by Morgan 2023/6/13
If field(9) = "023" And cp(10) = "443" And txtChkRltDate = "" Then
   MsgBox "催審期限不可空白！", vbExclamation
   txtChkRltDate.SetFocus
   Exit Function
End If
'end 2023/6/13

TxtValidate = True
End Function

'Added by Lydia 2021/05/25
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/05/25
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

Private Sub txtFavDt_GotFocus()
   TextInverse txtFavDt
   CloseIme
End Sub

Private Sub txtFavDt_Validate(Cancel As Boolean)
   If txtFavDt <> "" Then
      Cancel = Not ChkDate(txtFavDt)
   End If
End Sub

Private Sub txtPA57_GotFocus()
   TextInverse txtPA57
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtPA57.IMEMode = 2
   CloseIme
End Sub

Private Sub txtPA57_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Chr(KeyAscii) <> "Y" Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add by Morgan 2009/8/18
Private Sub lblCaseFee_Click()
   frm12040102_2.txtCF(1) = cp(1)
   frm12040102_2.txtCF(2) = field(9)
   frm12040102_2.txtCF(3) = cp(10)
   frm12040102_2.Show vbModal
   If Val(txtCaseField(0)) > 0 Then
      PUB_SetChkResultDate cp(1), field(9), cp(10), txtCaseField(0), txtChkRltDate, cp, field(8)
   End If
End Sub

Private Sub lblCaseFee_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblCaseFee, lblCaseFees
End Sub

Private Sub lblCaseFee_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblCaseFee, lblCaseFees
End Sub

Private Sub txtChkRltDate_Validate(Cancel As Boolean)
   If txtChkRltDate <> "" Then
      If ChkDate(txtChkRltDate) = False Then
         Cancel = True
      End If
   End If
End Sub
'Add by Morgan 2010/10/26
Private Function ChkEpcNewLetter() As Boolean
   If Val(DBDATE(field(10))) >= 20090401 Then
      ChkEpcNewLetter = True
      
   ElseIf field(46) = "Y" Then
      strExc(0) = "select cp47 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='101' and cp47>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) >= 20090401 Then
            ChkEpcNewLetter = True
         End If
      End If
   End If
   
End Function

'Add by Lydia 2015/02/02 輸入新穎性優惠期公開事實 (多筆)
Private Sub CmdFav_Click()
If cp(10) = "123" Then
   Set frm880020.m_PrevF = Me
   frm880020.m_dbCheck = True 'Modified by Lydia 2015/02/25  發文DoubleCheck
   frm880020.strFPD01 = field(1):   frm880020.strFPD02 = field(2)
   frm880020.strFPD03 = field(3):   frm880020.strFPD04 = field(4)
   frm880020.strNation = field(9)
   frm880020.strPA140 = IIf(txtFavDt.Text = "", IIf(field(140) <> "", Val(field(140)) - 19110000, ""), txtFavDt.Text)
   frm880020.Show
End If
End Sub
