VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm081031_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "創新業務－分案"
   ClientHeight    =   6744
   ClientLeft      =   240
   ClientTop       =   984
   ClientWidth     =   8856
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6744
   ScaleWidth      =   8856
   Begin VB.TextBox TextLC48 
      Height          =   270
      Left            =   6156
      MaxLength       =   1
      TabIndex        =   12
      Top             =   3312
      Width           =   255
   End
   Begin VB.CommandButton cmdCP06 
      Caption         =   "其他期限"
      Height          =   345
      Left            =   7860
      TabIndex        =   77
      Top             =   4590
      Width           =   945
   End
   Begin VB.TextBox txtKind 
      Height          =   300
      Left            =   1650
      MaxLength       =   1
      TabIndex        =   21
      Top             =   4590
      Width           =   375
   End
   Begin VB.CheckBox Check11 
      Caption         =   "急件"
      ForeColor       =   &H00000000&
      Height          =   200
      Left            =   3540
      TabIndex        =   74
      Top             =   2700
      Width           =   765
   End
   Begin VB.TextBox txtF0301 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   2970
      Locked          =   -1  'True
      TabIndex        =   73
      Text            =   "txtF0301"
      Top             =   330
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "檢視接洽單"
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1800
      TabIndex        =   72
      Top             =   0
      Width           =   1065
   End
   Begin VB.CommandButton cmdPFrate 
      Caption         =   "專業分配比例"
      Height          =   300
      Left            =   6990
      TabIndex        =   71
      Top             =   2985
      Width           =   1425
   End
   Begin VB.CommandButton ComBack 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   300
      Left            =   8040
      TabIndex        =   30
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdPrePic 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   300
      Left            =   6884
      TabIndex        =   29
      Top             =   0
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   300
      Left            =   6053
      TabIndex        =   28
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton Command3 
      Caption         =   "案件進度(&C)"
      Default         =   -1  'True
      Height          =   300
      Left            =   4950
      TabIndex        =   27
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "相關卷號(&F)"
      Height          =   300
      Left            =   3825
      TabIndex        =   26
      Top             =   0
      Width           =   1100
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一筆(&N)"
      Height          =   300
      Left            =   2895
      TabIndex        =   25
      Top             =   0
      Width           =   900
   End
   Begin VB.TextBox txtcp04 
      Height          =   300
      Left            =   3150
      MaxLength       =   2
      TabIndex        =   18
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtcp03 
      Height          =   300
      Left            =   2835
      MaxLength       =   1
      TabIndex        =   17
      Top             =   3960
      Width           =   255
   End
   Begin VB.TextBox txtcp02 
      Height          =   300
      Left            =   1890
      MaxLength       =   6
      TabIndex        =   16
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox txtcp01 
      Height          =   300
      Left            =   1290
      MaxLength       =   3
      TabIndex        =   15
      Top             =   3960
      Width           =   550
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1065
      Left            =   750
      TabIndex        =   22
      Top             =   4980
      Width           =   7935
      _ExtentX        =   13991
      _ExtentY        =   1884
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      AllowUserResizing=   1
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label lblLC48 
      AutoSize        =   -1  'True
      Caption         =   "特殊出名公司:          (J:智權公司 空白:系統預設)"
      Height          =   180
      Left            =   4920
      TabIndex        =   78
      Top             =   3396
      Width           =   3696
   End
   Begin VB.Label lblKind 
      Caption         =   "101(自行申請首次驗證)輸入:A當年度申請驗證,B隔年度申請驗證,C其他；1012(再驗證)輸入：D前一年度有101,E前一年度沒有101"
      Height          =   450
      Index           =   1
      Left            =   2070
      TabIndex        =   76
      Top             =   4590
      Width           =   5760
   End
   Begin VB.Label lblKind 
      Caption         =   "TIPS自動內部收文："
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   75
      Top             =   4650
      Width           =   1620
   End
   Begin MSForms.ComboBox Combo2 
      Height          =   315
      Left            =   1200
      TabIndex        =   6
      Top             =   2190
      Width           =   7185
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "12674;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   4
      Left            =   1500
      TabIndex        =   4
      Top             =   1575
      Width           =   6840
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "12065;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   3
      Left            =   1500
      TabIndex        =   3
      Top             =   1245
      Width           =   6840
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "12065;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   0
      Left            =   5376
      TabIndex        =   0
      Top             =   300
      Width           =   1215
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "2143;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   2
      Left            =   1500
      TabIndex        =   2
      Top             =   930
      Width           =   6840
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "12065;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   6
      Left            =   5370
      TabIndex        =   1
      Top             =   600
      Width           =   3000
      VariousPropertyBits=   671105051
      MaxLength       =   50
      Size            =   "5292;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   1890
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   585
      Index           =   18
      Left            =   765
      TabIndex        =   23
      Top             =   6120
      Width           =   3600
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "6350;1023"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   585
      Index           =   19
      Left            =   5040
      TabIndex        =   24
      Top             =   6120
      Width           =   3600
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "6350;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   15
      Left            =   5370
      TabIndex        =   14
      Top             =   3630
      Width           =   2175
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "3836;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   16
      Left            =   5370
      TabIndex        =   19
      Top             =   3960
      Width           =   495
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "873;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   14
      Left            =   1290
      TabIndex        =   13
      Top             =   3630
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "661;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   10
      Left            =   5370
      TabIndex        =   8
      Top             =   2640
      Width           =   975
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1720;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   13
      Left            =   3576
      TabIndex        =   11
      Top             =   3312
      Width           =   1092
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   12
      Left            =   1290
      TabIndex        =   10
      Top             =   3315
      Width           =   1095
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1931;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   9
      Left            =   1290
      TabIndex        =   9
      Top             =   2985
      Width           =   615
      VariousPropertyBits=   671105051
      MaxLength       =   4
      Size            =   "1085;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   7
      Left            =   1290
      TabIndex        =   7
      Top             =   2640
      Width           =   975
      VariousPropertyBits=   671105051
      MaxLength       =   6
      Size            =   "1720;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text 
      Height          =   300
      Index           =   17
      Left            =   6765
      TabIndex        =   20
      Top             =   4275
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "656;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(英)："
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   70
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(日)："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   69
      Top             =   1650
      Width           =   1200
   End
   Begin VB.Label Label26 
      Caption         =   "案件屬性："
      Height          =   180
      Left            =   240
      TabIndex        =   68
      Top             =   2280
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收  文  日："
      Height          =   180
      Index           =   0
      Left            =   4416
      TabIndex        =   66
      Top             =   348
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收  文  號： "
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   65
      Top             =   330
      Width           =   945
   End
   Begin VB.Label lbePaperNum 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """#-##-######"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   64
      Top             =   330
      Width           =   1695
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(中)："
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   63
      Top             =   990
      Width           =   1200
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "分所案號："
      Height          =   180
      Index           =   0
      Left            =   4440
      TabIndex        =   62
      Top             =   660
      Width           =   900
   End
   Begin VB.Label lbeNumber 
      Height          =   252
      Left            =   1200
      TabIndex        =   61
      Top             =   624
      Width           =   1572
   End
   Begin MSForms.Label lbe 
      Height          =   300
      Index           =   1
      Left            =   2340
      TabIndex        =   60
      Top             =   1890
      Width           =   6000
      VariousPropertyBits=   27
      Size            =   "10583;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "當  事  人："
      Height          =   180
      Left            =   240
      TabIndex        =   59
      Top             =   1950
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   58
      Top             =   660
      Width           =   900
   End
   Begin VB.Label lbeMoney 
      Height          =   255
      Left            =   4245
      TabIndex        =   57
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "後金："
      Height          =   180
      Left            =   3645
      TabIndex        =   56
      Top             =   4320
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "轉本所案號："
      Height          =   180
      Left            =   240
      TabIndex        =   55
      Top             =   4020
      Width           =   1080
   End
   Begin VB.Label lbeCloseDate 
      Height          =   300
      Left            =   5370
      TabIndex        =   54
      Top             =   2985
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "案件備註"
      Height          =   375
      Left            =   4530
      TabIndex        =   53
      Top             =   6225
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "進度備註"
      Height          =   375
      Left            =   270
      TabIndex        =   52
      Top             =   6225
      Width           =   375
   End
   Begin VB.Label Label29 
      Caption         =   "(Y:取消閉卷)"
      Height          =   180
      Left            =   5940
      TabIndex        =   51
      Top             =   4020
      Width           =   1095
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "取消收文日期："
      Height          =   180
      Left            =   4050
      TabIndex        =   50
      Top             =   3045
      Width           =   1260
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "相關總收文號："
      Height          =   180
      Left            =   4050
      TabIndex        =   49
      Top             =   3690
      Width           =   1260
   End
   Begin MSForms.Label lbe 
      Height          =   300
      Index           =   9
      Left            =   2010
      TabIndex        =   48
      Top             =   2985
      Width           =   1935
      VariousPropertyBits=   27
      Size            =   "3413;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label21 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "是否取消閉卷："
      Height          =   180
      Index           =   1
      Left            =   4050
      TabIndex        =   47
      Top             =   4020
      Width           =   1260
   End
   Begin VB.Label Label5 
      Caption         =   "(N：不算)"
      Height          =   180
      Index           =   0
      Left            =   1875
      TabIndex        =   46
      Top             =   3690
      Width           =   1215
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "是否算案件數："
      Height          =   180
      Left            =   75
      TabIndex        =   45
      Top             =   3690
      Width           =   1260
   End
   Begin VB.Label lbePointNum 
      Height          =   255
      Left            =   2565
      TabIndex        =   44
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "點數："
      Height          =   180
      Left            =   1965
      TabIndex        =   43
      Top             =   4320
      Width           =   540
   End
   Begin VB.Label lbeCost 
      Height          =   255
      Left            =   810
      TabIndex        =   42
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "費用："
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   41
      Top             =   4320
      Width           =   540
   End
   Begin MSForms.Label lbe 
      Height          =   300
      Index           =   10
      Left            =   6450
      TabIndex        =   40
      Top             =   2640
      Width           =   1905
      VariousPropertyBits=   27
      Size            =   "3360;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   4410
      TabIndex        =   39
      Top             =   2700
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Left            =   2592
      TabIndex        =   38
      Top             =   3396
      Width           =   900
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Left            =   240
      TabIndex        =   37
      Top             =   3375
      Width           =   900
   End
   Begin MSForms.Label lbe 
      Height          =   300
      Index           =   7
      Left            =   2310
      TabIndex        =   36
      Top             =   2640
      Width           =   1200
      VariousPropertyBits=   27
      Size            =   "2117;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   35
      Top             =   3045
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   240
      TabIndex        =   34
      Top             =   2700
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   1995
      Left            =   120
      Top             =   600
      Width           =   8535
   End
   Begin VB.Label Label19 
      Caption         =   "本案期限"
      Height          =   495
      Left            =   270
      TabIndex        =   33
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "是否向客戶收款："
      Height          =   180
      Index           =   0
      Left            =   5325
      TabIndex        =   32
      Top             =   4320
      Width           =   1440
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "(N:不收)"
      Height          =   180
      Left            =   7245
      TabIndex        =   31
      Top             =   4320
      Width           =   645
   End
   Begin VB.Label lblClose 
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
      Height          =   255
      Left            =   6690
      TabIndex        =   67
      Top             =   330
      Width           =   1605
   End
End
Attribute VB_Name = "frm081031_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/04/26 Form2.0已修改 Text(index)、lbe(index) ;  MSHFlexGrid1改字型=新細明體-ExtB
'Create by sonia 2019/7/24
Option Explicit

Dim strCP09() As String, t As Integer
Dim strDate As String, LcTmp As String
Dim intLastRow As Integer, intCols As Integer, strOldLc As String
Dim blnIsNew As Boolean, blnIsSave As Boolean, blnOKtoShow As Boolean
Dim lc01 As String, lc02 As String, lc03 As String, lc04 As String
Dim stName(1 To 3) As String
Dim m_ODate As String '本所期限
Dim m_LDate As String '法定期限
Dim m_CurrSel As Integer
Dim m_CPCount As Integer
Dim m_Cpindex As Integer
Dim m_CP60 As String, m_LC11 As String
'Dim m_strCust1 As String 'Mark by Lydia 2024/06/13
Dim m_CP98 As String
Dim m_CP101 As String
Dim m_CP104 As String
Dim m_CP65 As String
Dim strTemp As Variant
Dim m_CP27 As String
Dim m_CP31 As String
Public intCP09Col As String 'Add by Amy 2021/07/09 記錄前一畫面Grid cp09欄號
Dim m_bolUpdCP27 As Boolean 'Added by Lydia 2022/12/02 是否上發文日
Dim strCP122 As String 'Add by Amy 2022/11/17
Public strBCP06List As String  'Added by Lydia 2023/04/14 TIPS自動內部收文：本所期限
Dim strFirstCP06 As String 'Added by Lydia 2023/04/14 TIPS自動內部收文：首次驗證所限
Dim m_CP156 As String 'Added by Lydia 2025/03/28 TIPS請款階段(原本是輸入紙本接洽單數量,自2024/03/18改成TIPS請款階段)

'Add by Amy 2022/12/07 檢視接洽單
Private Sub cmdFile_Click()
    frm090801_Q.SetParent Me
    frm090801_Q.m_blnCallPrint = True
    frm090801_Q.Text5 = txtF0301
    Call frm090801_Q.cmdok_Click(4)
    frm090801_Q.Show
End Sub

Private Sub cmdNext_Click()
Dim i As Integer
  
  ClearForm
  'Add by Amy 2022/12/07 判斷接洽單已開啟就關閉
  If PUB_CheckFormExist("frm090801_Q") = True Then
    Unload frm090801_Q
  End If
  m_Cpindex = m_Cpindex + 1
  If m_Cpindex = m_CPCount - 1 Then
     cmdNext.Enabled = False
  ElseIf m_Cpindex = m_CPCount Then
     Exit Sub
  End If
  GetData (m_Cpindex)

End Sub

Private Sub cmdok_Click()
   Dim bolHadPoMsg As Boolean 'Add by Amy 2021/07/09
   Dim tmpBol As Boolean 'Added by Lydia 2023/04/14
   
   If AllTextBeforeSaveCheck Then Exit Sub
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Sub
   End If
   
   'Add by Amy 2021/07/09 控管轉本所案號時,不檢查接洽單PDF檔(同frm020101_02)
   If Not (txtcp01 <> "" And txtcp02 <> "") Then
      If lbePaperNum < "B" Then
         'Modify By Sindy 2022/12/16 電子收文不用檢查
         If Not (txtF0301 <> "" And Len(txtF0301) = 10) Then
         '2022/12/16 END
            If PUB_CheckPDF2(lbePaperNum, 0, True, strExc(0), Text(9), , bolHadPoMsg) = False Then
                If bolHadPoMsg = False Then
                    MsgBox "無接洽單PDF檔,不可分案!", vbCritical
                End If
                Exit Sub
            End If
         End If
      End If
   End If
   'end 2021/07/09
   
   If txtcp01 <> "" And txtcp02 <> "" Then
      strExc(1) = txtcp01
      strExc(2) = txtcp02
      strExc(3) = txtcp03
      If strExc(3) = "" Then strExc(3) = "0"
      strExc(4) = txtcp04
      If strExc(4) = "" Then strExc(4) = "00"
      
      strExc(5) = Text(9).Text '案件性質
      strExc(6) = lbe(9) '案件性質名稱
      strExc(7) = Text(0) '收文日
      strExc(8) = lbePaperNum '總收文號
      strExc(9) = m_LC11
      If Not ClsLawChkSameCase(strExc) Then Exit Sub
      'Added by Lydia 2020/08/18 更新相關卷號前,先檢查是否有重複
      If m_CP31 = "Y" Then
          If PUB_ChkUpdCR(lc01, lc02, lc03, lc04, strExc(1), strExc(2), strExc(3), strExc(4)) = False Then
              Exit Sub
          End If
      End If
      'end 2020/08/18
   End If
   
   'Added by Lydia 2021/04/28 ACS智財顧問專業分配比例管制：檢查智財顧問112一定要有ACSPFRate資料，若沒有則顯示訊息
   If cmdPFrate.Visible = True And cmdPFrate.Enabled = True Then
      strSql = "select count(*) cnt from ACSPFrate where ar01='" & lbePaperNum.Caption & "' and ar03 > 0 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Val("" & RsTemp.Fields("cnt")) = 0 Then
              MsgBox "智財顧問案件，請先輸入各專業部門分配比例！", vbCritical, "檢核資料"
              Exit Sub
         End If
      End If
   End If
   'end 2021/04/28
   
   m_bolUpdCP27 = False 'Added by Lydia 2022/12/02
   If Me.txtcp01.Text <> "" And Me.txtcp02.Text <> "" Then
      MsgBox "本次只更新轉本所案號的資料，不更新其他欄位的修改!!!", vbCritical
   'Added by Lydia 2022/12/02 若輸入的承辦人與智權人員相同，自動上發文日並且工作時數預設為 0。ex.ACS-000155
   Else
      '判斷是否上發文日
      'Modified by Lydia 2023/03/23 排除112智財顧問; 因為案件在分案階段通常是沒有還沒有執行完畢 ---黃教威
      If Text(9) <> "112" And Text(7) = Text(10) And Text(7) <> "" And Val(m_CP27) = 0 Then
          If MsgBox("因承辦人為智權人員系統將自動上發文日，確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
             Exit Sub
          End If
          m_bolUpdCP27 = True
      End If
   'end 2022/12/02
   End If
   'Added by Lydia 2020/01/21 所有系統別分案，若該收文號存在於下一程序(NP24)，若修改案件性質則彈訊息
   If Me.Text(9).Tag <> Me.Text(9).Text Then
       If Pub_CheckNP24Exists(lbePaperNum.Caption) = True Then
       End If
   End If
   'end 2020/01/21
   
   'Added by Lydia 2023/04/14 TIPS自動內部收文
   If txtKind.Visible = True And txtKind <> "" Then
       strFirstCP06 = ""
       txtKind_Validate tmpBol
       If tmpBol = True Then
          Exit Sub
       End If
       'Modifieby Lydia 2023/06/21 改成分案日和分案當年度  DBDATE(Text(0))=>strSrvDate(1)
       'Modified by Lydia 2023/12/22
       'If (txtKind = "A" Or txtKind = "E") And strSrvDate(1) > Left(strSrvDate(1), 4) & "0630" Then
       '    If MsgBox("收文日已超過當年度5月31日，是否繼續存檔？" & vbCrLf & "P.S.自動收文的本所期限從當年度5月31日開始", vbInformation + vbYesNo + vbDefaultButton2, "TIPS自動內部收文") = vbNo Then
       If txtKind = "A" And strSrvDate(1) > Left(strSrvDate(1), 4) & "0630" Then
           If MsgBox("分案日已超過當年度5月31日，是否繼續存檔？" & vbCrLf & "P.S.自動收文的本所期限從當年度5月31日開始", vbInformation + vbYesNo + vbDefaultButton2, "TIPS自動內部收文") = vbNo Then
              txtKind.SetFocus
              txtkind_GotFocus
              Exit Sub
           End If
       End If
       'Modified by Lydia 2023/06/21
       'If txtKind = "C" And strBCP06List = "" Then
       'Modified by Lydia 2023/12/22
       'If (txtKind = "C" Or txtKind = "F") And strBCP06List = "" Then
       If txtKind = "C" And strBCP06List = "" Then
           MsgBox "請輸入自動收文的本所期限！", vbExclamation, "TIPS自動內部收文"
           txtKind.SetFocus
           txtkind_GotFocus
           Exit Sub
       End If
       'Mark by Lydia 2023/06/21 改用前一年度的同一案件性質的所限+1年；若無則比照方案A(分案日和分案當年度)
'       If txtKind = "D" Then
'          strExc(1) = "select cp09,cp06 from caseprogress where cp01='" & lc01 & "' and cp02='" & lc02 & "' and cp03='" & lc03 & "' and cp04='" & lc04 & "' and cp159=0 and cp10='101' "
'          intI = 1
'          Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
'          If intI = 1 Then
'              strFirstCP06 = "" & RsTemp.Fields("cp06")
'          End If
'          If Val(strFirstCP06) = 0 Then
'JumpReInput:
'              strExc(2) = UCase(InputBox("請輸入首次驗證日期(民國年月日)，若要放棄請輸入空白：", "輸入首次驗證日期"))
'              If strExc(2) = "" Then
'                 txtKind.SetFocus
'                 txtkind_GotFocus
'                 Exit Sub
'              Else
'                 '檢查日期
'                 If Len(strExc(2)) = 7 Then
'                     If ChkDate(strExc(2)) = False Then
'                         GoTo JumpReInput
'                     End If
'                     If strExc(2) > strSrvDate(2) Then
'                         GoTo JumpReInput
'                     End If
'                     strFirstCP06 = DBDATE(strExc(2))
'                 Else
'                     GoTo JumpReInput
'                 End If
'              End If
'          End If 'If Val(strFirstCP06) = 0 Then
'       End If 'If txtKind = "D" Then
        'end ---- Mark by Lydia 2023/06/21
   End If
   'end 2023/04/14
   
   If Not SaveData Then DataErrorMessage (3)

   Screen.MousePointer = vbDefault
   ' 當轉本所案號時檢查原本所案號是否還有案件進度的資料
   If IsEmptyText(txtcp01) = False Then
      strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(lc01 & lc02 & lc03 & lc04)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If RsTemp.Fields(0) < 1 Then
         MsgBox "原本所案號 " & lc01 & lc02 & lc03 & lc04 & "已無案件進度資料，請通知收文人員刪號！", vbInformation
      Else
         MsgBox "原本所案號為 " & lc01 & lc02 & lc03 & lc04 & "，請自行更新原本所案號之下一程序資料 !", vbInformation
      End If
   End If

    If m_Cpindex = m_CPCount - 1 Then
      cmdOK.Enabled = False
      intForm = 0
      intNowRec = 0
      blnIsFormBack = True
      Unload Me
      frm081031.Show
      Exit Sub
   End If
   cmdNext_Click
End Sub

Private Sub cmdPrePic_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   intForm = 0
   intNowRec = 0
   blnIsFormBack = True
   Unload Me
   frm081031.Show
End Sub

Private Sub ComBack_Click()
   If blnIsSave = False Then
      If MsgBox("你並未存檔,確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   intForm = 0
   intNowRec = 0
   Unload Me
   Unload frm081031
End Sub

Private Sub Combo2_Click()
   Combo2.Tag = Combo2
End Sub

'Modified by Lydia 2023/04/14
'Private Sub Combo2_DropDown()
'   PUB_SetCasePTM "4", , Me.Combo2
'End Sub
Private Sub Combo2_DropButtonClick()
    PUB_SetCasePTM "4", , Me.Combo2
End Sub

Private Sub Command1_Click()
Dim i As Integer
Dim strNum As String
Dim strTmp As String
   
   strTmp = lbeNumber.Caption
   If strTmp = "" Then
      Exit Sub
   End If
   i = InStr(strTmp, "-")
   If i <> 0 Then
      strNum = Left(strTmp, i - 1)
      strTmp = Mid(strTmp, i + 1)
   End If
   frm1103_2.intWhereComeFrom = 1
   Set frm1103_2.m_form = Me
   frm1103_2.lblSystem = strNum
   i = InStr(strTmp, "-")
   If i <> 0 Then
      strNum = Left(strTmp, i - 1)
      If strTmp <> "" Then
         strTmp = Mid(strTmp, i + 1)
      End If
      frm1103_2.lblCode(0) = strNum
   Else
      frm1103_2.lblCode(0) = strTmp
      strTmp = ""
   End If
   If i <> 0 Then
      i = InStr(strTmp, "-")
      If i <> 0 Then
         strNum = Left(strTmp, i - 1)
         If strTmp <> "" Then
            strTmp = Mid(strTmp, i + 1)
         End If
         frm1103_2.lblCode(1) = strNum
      Else
         frm1103_2.lblCode(1) = strTmp
      End If
   Else
         frm1103_2.lblCode(1) = "0"
   End If
   
   If strTmp <> "" Then
      frm1103_2.lblCode(2) = strTmp
   Else
      frm1103_2.lblCode(2) = "00"
   End If
   
   frm1103_2.Show
   Me.Hide

End Sub

Private Sub Command3_Click()
   frm081031_2.Show
   If IsNoExistData Then Unload frm081031_2
End Sub

Private Sub Form_Load()
Dim i As Integer, n As Integer
 
   m_CPCount = 0
   MoveFormToCenter Me
   ClearForm 'Added by Lydia 2023/04/14
   t = 0
   blnIsSave = False
   With frm081031.MSHFlexGrid1
      n = 0
      For i = 1 To .Rows - 1
         .row = i
         .col = 0
         If .Text = "v" Then
            .col = intCP09Col 'Modify by Amy 2021/07/09 原:2
            ReDim Preserve strCP09(n)
            strCP09(n) = .Text
            m_CPCount = m_CPCount + 1
            n = n + 1
         End If
      Next
   End With
   GetData (0)
   If m_CPCount = 1 Then Me.cmdNext.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Amy 2022/12/07
   If PUB_CheckFormExist("frm090801_Q") = True Then
        Unload frm090801_Q
   End If
   PUB_SendMailCache
   intCP09Col = "" 'Add by Amy2021/07/09
   Set frm081031_1 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
Dim intRow As Integer
Dim i As Integer
  
   If MSHFlexGrid1.Rows > 1 Then
      If MSHFlexGrid1.row > 0 Then
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "v" Then
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = Empty
         Else
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "v"
         End If
      End If
   End If

   If MSHFlexGrid1.Rows < 2 Then Exit Sub
      With MSHFlexGrid1
             For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = "v" Then
                   Text(12).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(i, 2))
                   Text(13).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(i, 3))
                   Text(15).Text = MSHFlexGrid1.TextMatrix(i, 8)
                   Text(18).Text = IIf(Val(Me.Text(18).Tag) > 0, Me.Text(18).Tag & IIf(Len(Me.MSHFlexGrid1.TextMatrix(i, 13)) > 0, "，" & Me.MSHFlexGrid1.TextMatrix(i, 13), ""), Me.MSHFlexGrid1.TextMatrix(i, 13))
                   Exit For
                Else
                   Text(12).Text = ""
                   Text(13).Text = ""
                   Text(15).Text = ""
                   Text(18).Text = "" & Me.Text(18).Tag
                End If
            Next i
      End With
End Sub

Private Sub MSHFlexGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      If MSHFlexGrid1.row > 0 Then
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "V" Then
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = Empty
         Else
            MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "V"
         End If
      End If
   End If

End Sub

' 將MSHFlexGrid1所選取的列反白, 並將未選取的列設成一般顏色
Private Sub MSHFlexGrid1_ShowSelection()
Dim nCurrSel As Integer
Dim nCol As Integer

   nCurrSel = MSHFlexGrid1.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = MSHFlexGrid1.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < MSHFlexGrid1.Rows Then
      MSHFlexGrid1.row = m_CurrSel
      MSHFlexGrid1.col = 1
      If MSHFlexGrid1.CellBackColor <> &H80000005 Then
         For nCol = 1 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.col = nCol
            If MSHFlexGrid1.CellBackColor <> &H80000005 Then: MSHFlexGrid1.CellBackColor = &H80000005
            If MSHFlexGrid1.CellForeColor <> &H80000008 Then: MSHFlexGrid1.CellForeColor = &H80000008
         Next nCol
      End If
      MSHFlexGrid1.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < MSHFlexGrid1.Rows Then
      MSHFlexGrid1.row = m_CurrSel
      MSHFlexGrid1.col = 1
      For nCol = 1 To MSHFlexGrid1.Cols - 1
         MSHFlexGrid1.col = nCol
        MSHFlexGrid1.CellBackColor = &HFFC0C0
         MSHFlexGrid1.CellForeColor = &H80000008
      Next nCol
      MSHFlexGrid1.col = 0
   End If
EXITSUB:
End Sub

Private Sub MSHFlexGrid1_SelChange()
  MSHFlexGrid1_ShowSelection
End Sub

Private Sub Text_Change(Index As Integer)
Dim i As Integer
   
   Select Case Index
      Case 1, 7, 9, 10
         If Text(Index) = "" Then lbe(Index) = ""
   End Select
End Sub

Private Sub Text_GotFocus(Index As Integer)
   Select Case Index
      Case 12
         If Text(Index) <> "" Then strDate = Text(Index)
         TextInverse Text(Index)
         CloseIme
      Case 2, 4
         TextInverse Text(Index)
         OpenIme
      Case 18, 19
         OpenIme
      Case Else
         CloseIme
         TextInverse Text(Index)
   End Select
End Sub

'Modified by Lydia 2021/04/26
'Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case Index
         KeyAscii = UpperCase(KeyAscii)
         If Index = 16 Then
            If KeyAscii <> 89 And KeyAscii <> 8 Then
               KeyAscii = 0
            End If
         End If
   End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
   CloseIme
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim strTempName As String, i As Integer, blnIsEmpty As Boolean
Dim m_SalesST15 As String '畫面上智權人員的收文部門
Dim m_Tuser As String '創新業務部預設收文人員

   strTempName = ""
   Select Case Index
      Case 0
         If Text(Index) <> "" Then
            If CheckIsTaiwanDate(Text(Index)) Then
               If Val(GetTaiwanTodayDate) - Val(Text(Index)) < 0 Then
                  DataErrorMessage 2, "收文日期"
                  Cancel = True
               End If
            Else
               Cancel = True
            End If
         Else
            MsgBox "收文日不可空白", vbCritical
            Cancel = True
         End If
      Case 1 '當事人
         If Text(Index) <> "" Then
            Text(Index) = UCase(Text(Index))
            If Left(Text(Index), 1) = "Y" Then
               Text(Index) = "X" & Mid(Text(Index), 2)
            ElseIf Left(Text(Index), 1) <> "X" Then
               MsgBox "當事人代碼輸入錯誤!", vbCritical
               TextInverse Text(Index)
               Cancel = True
               Exit Sub
            End If
            
            If ClsPDGetCustomer(Text(Index), strTempName) Then
               lbe(Index) = strTempName
               If m_CP60 <> "" And InStr(ChangeCustomerL(m_LC11), ChangeCustomerL(Text(Index))) = 0 Then
                  strExc(1) = lc01
                  strExc(2) = lc02
                  strExc(3) = lc03
                  strExc(4) = lc04
                  strExc(5) = m_CP60
                  strExc(6) = Text(Index)
                  strExc(7) = strTempName
                  strExc(8) = m_LC11
                  If Not ClsLawUpdAcc0k0(strExc()) Then
                     lbe(Index) = ""
                     Cancel = True
                  End If
               End If
            
            Else
               Cancel = True
               lbe(Index) = ""
            End If
         Else
            MsgBox "當事人不可空白", vbCritical
            Cancel = True
         End If
         If Cancel = False Then
            'Modified by Lydia 2024/06/13
            'If m_strCust1 <> Me.Text(1).Text Then
            If m_LC11 <> ChangeCustomerL(Me.Text(1).Text) Then
               If Not PUB_EditCustOk(Me.lbePaperNum.Caption, lc01, lc02, lc03, lc04) Then Cancel = True
            End If
         End If
      Case 6
         If Text(Index) <> "" Then
            Text(Index) = UCase(Text(Index))
         End If
      Case 7, 10
         If Text(Index) <> "" Then
            Text(Index) = UCase(Text(Index))
            If ClsPDGetStaff(Text(Index), strTempName) Then lbe(Index) = strTempName Else Cancel = True: lbe(Index) = ""
              If Index = 10 Then
                 m_SalesST15 = GetST15(Text(Index))
                 If PUB_ChkIsT10T20("2", Text(Index).Text, m_Tuser, strTempName) = True Then
                     Text(Index).Text = m_Tuser
                     lbe(Index).Caption = strTempName
                     Text(Index).SetFocus
                     Call Text_GotFocus(Index)
                     Cancel = True
                     Exit Sub
                 End If
              End If
              'Added by Lydia 2023/06/21 101、1012、1013請在分案時自動設定本所期限為:分案當年度12月1日。
              'Modified by Lydia 2023/12/05 1013=>抽驗L(1013)，新增:抽驗S(1014) +Or Text(9) = "1014"
              If Index = 7 Then
                 'Modified by Lydia 2024/04/12 +131前置自行申請首次驗證,141諮詢再驗證,142諮詢抽驗Ｌ
                 'If Text(7).Tag = "" And Text(12) = "" And (Text(9) = "101" Or Text(9) = "1012" Or Text(9) = "1013" Or Text(9) = "1014") Then
                 If Text(7).Tag = "" And Text(12) = "" And (Text(9) = "101" Or Text(9) = "1012" Or Text(9) = "1013" Or Text(9) = "131" Or Text(9) = "141" Or Text(9) = "142") Then
                    Text(12) = TransDate(PUB_GetWorkDay1(Left(strSrvDate(1), 4) & "1201", True), 1)
                 End If
              End If
              'end 2023/06/21
         End If
      Case 9
           If Text(Index) = "" Then
              MsgBox "案件性質不可空白", vbCritical
              lbe(Index) = ""
              Cancel = True
           Else
               If ClsPDGetCaseProperty(CheckCaseNum, Text(Index), strTempName, False) Then
                   lbe(Index) = strTempName
                   'Added by Lydia 2023/04/14 TIPS自動內部收文：顯示欄位
                   If Text(Index).Tag <> Text(Index).Text Then
                      Call GetKindType
                   End If
                   'end 2023/04/14
               Else
                   Cancel = True
               End If
           End If
      Case 12
           If Text(Index) <> "" Then
              If CheckIsTaiwanDate(Text(Index)) Then
                  '若本所期限非工作天則直接調整至最近的工作天
                  Text(Index) = TransDate(PUB_GetWorkDay1(Text(Index), True), 1)
                  If Text(13) <> "" Then
                     If Val(Text(13)) - Val(Text(Index)) < 0 Then DataErrorMessage 13: Cancel = True
                  End If
              Else
                 Cancel = True
              End If
           End If
       Case 13
            If Text(Index) <> "" Then
              If CheckIsTaiwanDate(Text(Index)) Then
                  If Text(12) <> "" Then
                     If Val(Text(Index)) - Val(Text(12)) < 0 Then DataErrorMessage 12: Cancel = True
                  End If
              Else
                 Cancel = True
              End If
           End If
      Case 14
         Text(Index) = UCase(Text(Index))
         If Text(Index) <> "" And Text(Index) <> "N" Then
            Cancel = True
            DataErrorMessage 1, "是否算案件數"
         End If
      Case 15
          If Text(Index) <> "" Then
             Text(Index) = UCase(Text(Index))
             If Text(Index) = lbePaperNum Then
                MsgBox "且不可為本身之收文號", vbCritical
                Cancel = True
             End If
             If Not ClsLawGetRelation(LcTmp, lbePaperNum, Text(Index)) Then Cancel = True
          End If
      Case 16
          If Text(Index) <> "" Then
                Text(Index) = UCase(Text(Index))
              If Text(Index) = "Y" Then
                 i = MsgBox("確定閉卷?", vbYesNo, "詢問")
                 If i = 7 Then Text(Index) = ""
              Else
                 DataErrorMessage 1, "是否閉卷": Cancel = True
              End If
          End If
      Case 17
         Text(Index) = UCase(Text(Index))
         If Text(Index) <> "" And Text(Index) <> "N" Then
            Cancel = True
            DataErrorMessage 1, "是否向客戶收款"
         End If
   End Select
   If Cancel Then TextInverse Text(Index)
End Sub

'Add by Amy 2023/08/18 特殊出名公司
Private Sub TextLC48_GotFocus()
   TextInverse TextLC48
   CloseIme
End Sub

Private Sub TextLC48_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("J") Then
      KeyAscii = 0
      Beep
   End If
End Sub
'end 2028/08/18

Private Sub txtcp01_GotFocus()
   TextInverse txtcp01
End Sub

Private Sub txtcp01_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtcp01_LostFocus()
   If Len(txtcp01) > 0 Then txtcp01 = UCase(txtcp01)
End Sub

Private Sub txtcp01_Validate(Cancel As Boolean)
   If txtcp01 <> "" Then
     txtcp01 = UCase(txtcp01)
     If txtcp01 <> GetCaseNumSysKind(lbeNumber) Then
     DataErrorMessage 1, "本所案號"
     Cancel = True
     End If
   End If
   If Cancel Then TextInverse txtcp01
End Sub

Private Sub txtcp02_GotFocus()
   TextInverse txtcp02
End Sub

Private Sub txtcp02_Validate(Cancel As Boolean)
Dim strTemp As String, i As Integer, yn As Integer, strlcTemp As String
   
   If txtcp02 <> "" Then
      If Len(txtcp02) = 6 Then
         If ClsPDChkCaseNum(txtcp01, txtcp02) Then
            TextInverse txtcp02
            Cancel = True
         Else
            If txtcp03 = "" Then
               strlcTemp = txtcp01 + txtcp02 + "000"
            Else
               strlcTemp = txtcp01 + txtcp02 + txtcp03 + txtcp04
            End If
         End If
      Else
         DataErrorMessage 1, "本所案號"
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtcp02
End Sub

Private Sub txtcp03_GotFocus()
   TextInverse txtcp03
End Sub

Private Sub txtcp03_Validate(Cancel As Boolean)
  If txtcp02 <> "" And txtcp03 = "" Then txtcp03 = "0"
  If txtcp03 <> "" Then
      If Len(txtcp03) > 1 Then
         DataErrorMessage 1, "本所案號"
         Cancel = True
         Exit Sub
      End If
   End If
   If Cancel Then TextInverse txtcp03
End Sub

Private Sub GetData(ByVal intI As Integer)
Dim yn As Boolean, i As Integer, j As Integer
Dim rsR1 As New ADODB.Recordset
Dim St(33) As String
    
   'Modify by Amy 2022/12/07 +CP122/cp140
   'Modify by Amy 2023/08/18 +LC48
   'Modified by Lydia 2025/03/28 +CP156
   strCP122 = ""
   strExc(1) = "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp13," & _
      "cp14,cp16,cp18,cp19,cp20,cp26,cp29,cp43,LC22,LC23,cp49,cp57,cp64," & _
      "lc05,lc06,lc07,lc08,lc11,lc13,lc14,lc15,lc16,lc27,CP60,CP65,LC47,cp31,cp27,cp122,cp140,LC48,CP156 " & _
      "from lawcase, caseprogress " & _
      "where cp09='" + strCP09(intI) + "' AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 " & _
      " order by LC01,LC02,LC03,LC04"
   intI = 0
   Set rsR1 = ClsLawReadRstMsg(intI, strExc(1))
   If intI = 1 Then
      For j = 1 To 33
         If IsNull(rsR1.Fields(j - 1).Value) = False Then
            St(j) = rsR1.Fields(j - 1).Value
         Else
            St(j) = ""
         End If
      Next
      
      If Not IsNull(rsR1.Fields("CP60")) Then
         m_CP60 = rsR1.Fields("CP60")
      Else
         m_CP60 = ""
      End If
      
      m_CP65 = ""
      If IsNull(rsR1.Fields("CP65")) = False Then
         m_CP65 = rsR1.Fields("CP65")
      End If
      
      m_CP27 = Empty
      If IsNull(rsR1.Fields("CP27")) = False Then
         m_CP27 = rsR1.Fields("CP27")
      End If
      
      If Not IsNull(rsR1.Fields("LC11")) Then
         m_LC11 = rsR1.Fields("LC11")
      Else
         m_LC11 = ""
      End If
      'Add by Amy 2022/12/07
      strCP122 = "" & rsR1.Fields("CP122") '是否急件
      If IsNull(rsR1.Fields("CP140")) = False Then: txtF0301 = rsR1.Fields("CP140") '接洽單編號
      'end2022/12/07
      m_CP156 = "" & rsR1.Fields("CP156") 'Added by Lydia 2025/03/28
      
      lc01 = St(1)
      lc02 = St(2)
      lc03 = St(3)
      lc04 = St(4)
      lbeNumber = GiveSymbol(St(1), St(2), St(3), St(4), LcTmp)
      lbeNumber.Tag = LcTmp
      Text(0) = ChangeWStringToTString(St(5))
      Text(12) = ChangeWStringToTString(St(6))
      Text(13) = ChangeWStringToTString(St(7))
      m_ODate = ChangeWStringToTString(St(6))
      m_LDate = ChangeWStringToTString(St(7))
      
      lbePaperNum = St(8)
      Text(9) = St(9): ChgType (9)
      Text(9).Tag = Text(9).Text 'Added by Lydia 2020/01/21
      Text(10) = St(10): ChgType (10)
      Text(7) = St(11): ChgType (7)
      Text(7).Tag = Text(7).Text 'Add By Sindy 2023/3/2
      lbeCost = St(12)
      lbePointNum = St(13)
      lbeMoney = St(14)
      Text(17) = UCase(St(15))
      Text(14) = UCase(St(16))
      Text(15) = UCase(St(18))
      lbeCloseDate = ChangeWStringToTDateString(St(22))
      'Added by Lydia 2021/04/28 ACS智財顧問專業分配比例管制： 沒有取消收文
      If lc01 = "ACS" And Text(9).Text = "112" And Trim(St(22)) = "" And strSrvDate(1) >= ACS_PFrateStart Then
         cmdPFrate.Visible = True
      Else
         cmdPFrate.Visible = False
      End If
      'end 2021/04/28
      Text(18) = St(23)
      Me.Text(18).Tag = Me.Text(18).Text
      If UCase(St(27)) = "Y" Then
         Me.lblClose.Caption = "已閉卷"
         Me.Text(16).Visible = True
         Me.Label21(1).Visible = True
         Me.Label29.Visible = True
      Else
         Me.lblClose.Caption = ""
         Me.Text(16).Visible = False
         Me.Label21(1).Visible = False
         Me.Label29.Visible = False
      End If
      
      Text(1) = ChangeCustomerS(St(28)): ChgType (1)
      '案件屬性
      Combo2 = "" & rsR1.Fields("lc47")
      Combo2.Tag = Combo2
      Text(6) = St(32)
      Text(19) = St(33)
      For i = 1 To 3
         stName(i) = ""
         stName(i) = St(23 + i)
      Next
      Text(2) = St(24)
      Text(3) = St(25)
      Text(4) = St(26)
      'm_strCust1 = "" & Me.Text(1).Text 'Mark by Lydia 2024/06/13
      
      Getrs
      
      m_CP31 = "" & rsR1.Fields("CP31")
      'Modify by Amy 2023/08/18 +特殊出名公司,判斷新案且收文日在一個月之內才可修改 特殊出名公司
      TextLC48 = "" & rsR1.Fields("LC48")
      TextLC48.Visible = False
      lblLC48.Visible = False
      'CP31為Y時,Shape1內的欄位才可修改,否則鎖住
      If "" & rsR1.Fields("CP31") = "Y" Then
         Text(2).Locked = False
         Text(3).Locked = False
         Text(4).Locked = False
         Text(1).Locked = False
         Text(6).Locked = False
         TextLC48.Visible = True
         lblLC48.Visible = True
         TextLC48.Enabled = False
         If Val(strSrvDate(1)) >= Val(St(5)) And Val(strSrvDate(1)) <= Val(DBDATE(DateAdd("m", 1, Format(Val(St(5)), "####/##/##")))) Then
            TextLC48.Enabled = True
         End If
      'end 2023/08/18
         'Combo2.Locked = False 'Remove by Lydia 2020/11/06 取消ACS的案件屬性欄
      Else
         Text(2).Locked = True
         Text(3).Locked = True
         Text(4).Locked = True
         Text(1).Locked = True
         Text(6).Locked = True
         'Combo2.Locked = True 'Remove by Lydia 2020/11/06 取消ACS的案件屬性欄
      End If
      
      'Remove by Lydia 2020/11/06 取消ACS的案件屬性欄
      'If Combo2.Locked = False Then
      '   PUB_SetCasePTM "4", , Me.Combo2
      'End If
      'end 2020/11/06
      
      CheckCaseDestroy lc01, lc02, lc03, lc04
      
      'C類來函或已發文案件須鎖住轉本所案號欄位, 若為併號請以聯絡單通知電腦中心處理
      'Modified by Lydia 2025/03/28 已有請款階段設定不可轉本所案號 +And m_CP156 = ""
      If Trim(lbePaperNum.Caption) < "C" And Val(m_CP27) = 0 And m_CP156 = "" Then
          Me.txtcp01.Enabled = True
          Me.txtcp02.Enabled = True
          Me.txtcp03.Enabled = True
          Me.txtcp04.Enabled = True
      Else
          Me.txtcp01.Enabled = False
          Me.txtcp02.Enabled = False
          Me.txtcp03.Enabled = False
          Me.txtcp04.Enabled = False
      End If
   End If
   'Modify by Amy 2022/12/07 +接洽單電子收文才顯示「檢視接洽單」鈕和急件
   cmdFile.Visible = False
   Check11.Visible = False '急件
   Check11.Value = 0 'Add By Sindy 2023/1/10 要先清欄位值,再後續判斷是否急件
   'Modify by Amy 2023/01/03 8碼(結案單)不可開接洽單會錯: + And Len(txtF0301) = 10
   If strSrvDate(1) >= 接洽單電子收文啟用日 And txtF0301 <> MsgText(601) And Len(txtF0301) = 10 Then
        cmdFile.Visible = True
        Check11.Visible = True
        If strCP122 = "Y" Then Check11.Value = 1
        'Add by Amy 2023/01/07 直接開啟接洽單
        frm090801_Q.SetParent Me
        frm090801_Q.m_blnCallPrint = True
        frm090801_Q.Text5 = txtF0301
        Call frm090801_Q.cmdok_Click(4)
        frm090801_Q.Show
        'end 2023/01/07
   End If
   
   Call GetKindType 'Added by Lydia 2023/04/14 TIPS自動內部收文

   blnIsSave = False
   
   'Add By Sindy 2024/1/30 各部門分案時，若本所期限與法定期限與接洽單的本所期限與法定期限不同時，要提醒
   Call PUB_ChkCRLdtCP06CP07(St(8))
   
   Set rsR1 = Nothing
End Sub

Private Function SaveData() As Boolean
Dim i As Integer, blnIsChange As Boolean
Dim cp01 As String, cp02 As String, cp03 As String, cp04 As String
Dim strTmp As String, iStep As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim sLC() As String
ReDim sLC(1 To TF_LC) As String
Dim strApply As Variant
Dim strCP122_Now As String 'Add by Amy 2022/12/07 急件
Dim arrTmp1, arrTmp2    'Added by Lydia 2023/04/14

On Error GoTo CheckingErr
   SaveData = True
   cnnConnection.BeginTrans

   iStep = 1
   '若有輸入轉本所案號
   If txtcp01 <> "" Then
      cp01 = txtcp01
      cp02 = txtcp02
      cp03 = Left(txtcp03 & "0", 1)
      cp04 = Left(txtcp04 & "00", 2)
      blnIsChange = True
      
'cancel by sonia 2024/11/26 已不立卷不必再通知分所收文人員
'      '若為分所收文案件則發Mail通知收文人員
'      strExc(0) = PUB_GetST06(m_CP65)
'      If strExc(0) > "1" Then
'         strExc(1) = "原本所案號 " & lc01 & "-" & lc02 & IIf(lc03 & lc04 = "000", "", "-" & lc03 & "-" & lc04)
'         strExc(1) = strExc(1) & " 已更改為 " & cp01 & "-" & cp02 & IIf(cp03 & cp04 = "000", "", "-" & cp03 & "-" & cp04) & " 。"
'         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'            " values ('" & strUserNum & "','" & m_CP65 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'            ",'" & ChgSQL(strExc(1)) & "','總收文號：" & lbePaperNum & " 改本所案號如主旨')"
'         cnnConnection.Execute strSql
'      End If
'end 2024/11/26
      
      GoTo ProcChgCaseNum
   '若未輸入轉本所案號
   Else
     cp01 = lc01
     cp02 = lc02
     cp03 = lc03
     cp04 = lc04
     blnIsChange = False
   End If
   If cp03 = "" Then cp03 = "0"
   If cp04 = "" Then cp04 = "00"
   LcTmp = cp01 & cp02 & cp03 & cp04
   
Dim strTmp1(1 To 3) As String
   If Text(1) <> "" Then
      If ClsPDGetCustomerNameAndAddress(Text(1).Text, strExc(0), strTmp1(1), strTmp1(2), strTmp1(3)) Then
         '修改當事人時
         If InStr(ChangeCustomerL(m_LC11), ChangeCustomerL(Text(1))) = 0 Then
            If m_CP60 <> "" Then
               strExc(1) = lc01
               strExc(2) = lc02
               strExc(3) = lc03
               strExc(4) = lc04
               strExc(5) = m_CP60
               strExc(6) = Text(1)
               strExc(7) = strExc(0)
               strExc(8) = m_LC11
               If Not ClsLawUpdAcc0k0(strExc(), True) Then
                  Text(1).SetFocus
                  GoTo CheckingErr
               End If
            End If
         End If
      End If
   End If

   If Me.Text(16).Text = "Y" Then
      strExc(1) = " , LC08=NULL, LC09=NULL, LC10=NULL "
   Else
      strExc(1) = " "
   End If
   
   'Modify by Amy 2024/07/19 +LC48
   strExc(1) = "update lawcase set lc05=" & CNULL(ChgSQL(Text(2))) & ",lc06=" & CNULL(ChgSQL(Text(3))) & _
      ",lc07=" & CNULL(ChgSQL(Text(4))) & strExc(1) & ",lc11=" & CNULL(ChangeCustomerL(Text(1))) & _
      ",lc16=" & CNULL(Text(6)) & ",lc27=" & CNULL(ChgSQL(Text(19))) & _
      ",lc47=" & CNULL(ChgSQL(Combo2)) & ",lc48=" & CNULL(ChgSQL(TextLC48)) & " where " & ChgLawcase(LcTmp)
      
   '有修改lc47寫log
   If Combo2.Visible = True Then
      If Combo2.Tag <> Combo2 Then Pub_SeekTbLog strExc(1)
   End If
   cnnConnection.Execute strExc(1)

   'Add By Sindy 2021/10/14 分案輸入承辦人,直接上齊備日
   If Text(7) <> "" Then
      strSql = " Update engineerprogress Set EP06=" & strSrvDate(1) & " WHERE EP02='" & Me.lbePaperNum.Caption & "' and EP06 is null"
      cnnConnection.Execute strSql
   End If
   '2021/10/14
   'Added by Lydia 2022/12/02 若輸入的承辦人與智權人員相同，自動上發文日並且工作時數預設為 0。ex.ACS-000155
   'Modify By Sindy 2023/3/2 ACS案件分案時，案件性質為706代收代付時，存檔時同時上發文日為系統日
   If (m_bolUpdCP27 = True Or Text(9) = "706") And Text(7) <> "" And Val(m_CP27) = 0 Then
      'Mark by Lydia 2024/03/29 配合TIPS請款階段發文產生定稿，所以不上發文日
      'strSql = "Update CaseProgress set cp27=" & CNULL(strSrvDate(1)) & ", cp113=0 where cp09='" & Me.lbePaperNum.Caption & "' and cp27 is null"
      'cnnConnection.Execute strSql, intI
      'end 2024/03/29
      'Modify By Sindy 2023/3/2 發EMAIL給系統特殊設定「財務處總帳人員」。
      If Text(9) = "706" Then
         'modify by sonia 2024/7/19 修改文字，原：提供匯款證明、請款單予智權同仁
         strExc(1) = "請依卷宗區內接洽單內容於指定期限內完成匯款，並提供匯款證明予顧服組(acs01)與智權同仁。"
         'Added by Lydia 2024/07/31 與婉莘副理再次討論後，為確保匯款作業資訊正確，還請電腦中心於系統通知財務處匯款時，一併帶入接洽單中的案件說明事項處理頁籤內容。by 黃教威
         strExc(3) = Pub_GetField("ConsultRecordList", "CRL01='" & txtF0301 & "'", "CRL57")
         If strExc(3) <> "" Then
            strExc(3) = vbCrLf & vbCrLf & "案件說明事項處理：" & vbCrLf & strExc(3)
         End If
         'end 2024/07/31
         
         'Modified by Lydia 2024/07/31 + strExc(3)
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " values ('" & strUserNum & "','" & Pub_GetSpecMan("財務處總帳人員") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & cp01 & "-" & cp02 & IIf(cp03 & cp04 = "000", "", "-" & cp03 & "-" & cp04) & "總收文號：" & lbePaperNum & "(" & lbe(9) & ")-->匯款通知','" & strExc(1) & strExc(3) & "')"
         cnnConnection.Execute strSql
      End If
   End If
   'end 2022/12/02
   
   'Add by Amy 2022/12/07 +CP122 急件
   If Check11.Visible = True Then
       strCP122_Now = "N"
       If Check11.Value = 1 Then strCP122_Now = "Y"
       'Memo DB資料若為null,回存N,避免與內商混淆
       If strCP122 <> strCP122_Now Then
           strSql = "Update CaseProgress Set CP122=" & CNULL(strCP122_Now) & " Where cp09='" & Me.lbePaperNum.Caption & "' "
           cnnConnection.Execute strSql
       End If
   End If
   'end 2022/12/07
   
ProcChgCaseNum:
   
   If blnIsChange Then
      If Me.txtcp01.Text <> "" And Me.txtcp02.Text <> "" And SaveData = True Then
         StrSQLa = "SELECT * FROM LAWCASE WHERE " & ChgLawcase(cp01 & cp02 & cp03 & cp04)
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount <= 0 Then
            If PUB_ReadLawCaseData(sLC(), lc01, lc02, lc03, lc04) Then
               sLC(1) = Me.txtcp01.Text
               sLC(2) = Me.txtcp02.Text
               sLC(3) = Left(Me.txtcp03.Text & "0", 1)
               sLC(4) = Left(Me.txtcp04.Text & "00", 2)
               If PUB_AddNewLawCase(sLC()) Then
               Else
                  GoTo CheckingErr
               End If
            End If
         '若基本檔有資料, 若是否新案欄為'Y'更新為Null
         Else
            strSql = " Update CaseProgress Set CP31=NULL WHERE CP09='" & Me.lbePaperNum.Caption & "'"
            cnnConnection.Execute strSql
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
   
      strExc(iStep) = "update caseprogress set cp01=" & CNULL(cp01) & _
         ",cp02=" & CNULL(cp02) & ",cp03=" & CNULL(cp03) & _
         ",cp04=" & CNULL(cp04) & " WHERE CP09='" & Me.lbePaperNum.Caption & "'"
      Pub_SeekTbLog strExc(iStep) 'Added by Lydia 2025/09/25 L-006989誤輸入轉入ACS案
      cnnConnection.Execute strExc(iStep)
      iStep = iStep + 1
      
      strExc(iStep) = "UPDATE CASEPROGRESS SET CP43='' WHERE CP09='" & Me.lbePaperNum.Caption & "'"
      cnnConnection.Execute strExc(iStep)
      iStep = iStep + 1
      
      'Added by Lydia 2020/08/18 更新CaseRelation1和DivisionCase
      If m_CP31 = "Y" Then
          Call PUB_UpdateCaseRelation1(lc01, lc02, lc03, lc04, cp01, cp02, cp03, cp04)
      End If
      'end 2020/08/18
      
      '更正財務相關資料
      PUB_UpdateAccData Trim(lbePaperNum), lc01 & lc02 & lc03 & lc04
      
   Else
      'Modified by Lydia 2023/04/14 + TIPS自動內部收文
      strTmp = "cp05=" & CNULL(TransDate(Text(0), 2)) & _
               ",cp06=" & CNULL(TransDate(Text(12), 2)) & _
               ",cp07=" & CNULL(TransDate(Text(13), 2)) & _
               ",cp10=" & CNULL(Text(9)) & ",cp12=" & CNULL(GetST15(Text(10))) & ",cp13=" & CNULL(Text(10)) & ",cp14=" & CNULL(Text(7)) & _
               ",cp20=" & CNULL(Text(17)) & ",cp26=" & CNULL(Text(14)) & _
               ",cp43=" & CNULL(Text(15)) & ",cp64=" & CNULL(ChgSQL(IIf(txtKind <> "", ChangeWStringToWDateString(strSrvDate(1)) & "TIPS自動內部收文：" & txtKind & ";", "") & Text(18))) & _
               " where cp09=" & CNULL(lbePaperNum)
      strExc(iStep) = "update caseprogress set " & strTmp
      cnnConnection.Execute strExc(iStep)
      iStep = iStep + 1
      
      If SaveNextProgress = False Then GoTo CheckingErr
   End If
      
   'Added by Lydia 2023/04/14 TIPS自動內部收文
   If txtKind.Visible = True And Trim(txtKind) <> "" Then
       'Added by Lydia 2023/06/21 【方案F】自行輸入期限(和方案C相同)
       'Modified by Lydia 2023/12/22 1013(抽驗L)不收213內部稽核,215管理審查
       'If txtKind = "F" Then
       If Text(9) = "1013" Then
          arrTmp1 = Split("208,2091,2092,211,213,215", ",")
       Else
       'end 2023/06/21
          arrTmp1 = Split("208,2091,2092,211,213,215,216,218", ",")
       End If 'Added by Lydia 2023/06/21
       Select Case txtKind
           'Modified by Lydia 2023/12/22 簡化自動收文為【方案A~C】
           'Case "A", "E"
           '   '【方案A】:當年度收文+當年度申請驗證
           Case "A" '【方案A】:分案當年度申請驗證
           'end 2023/12/22
              strBCP06List = ""
              'Modified by Lydia 2023/06/21 改成分案日和分案當年度 DBDATE(Text(0))=>strSrvDate(1)
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1(CompDate(2, 30, strSrvDate(1)), True)   '208 啟始會議:收文日+30日
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1(Left(strSrvDate(1), 4) & "0531", True)   '2091 權責人員教育訓練:系統日(原本是收文日的當年度)5月31日
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1(Left(strSrvDate(1), 4) & "0630", True)   '2092 全體員工教育訓練:系統日(原本是收文日的當年度)6月30日
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1(Left(strSrvDate(1), 4) & "0531", True)   '211 文件修制訂:系統日(原本是收文日的當年度)5月31日
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1(Left(strSrvDate(1), 4) & "0715", True)   '213 內部稽核:系統日(原本是收文日的當年度)7月15日
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1(Left(strSrvDate(1), 4) & "0815", True)   '215 管理審查:系統日(原本是收文日的當年度)8月15日
              If Text(9) <> "1013" Then 'Added by Lydia 2023/12/22 判斷1013(抽驗L)不收213內部稽核,215管理審查
                 strBCP06List = strBCP06List & "," & PUB_GetWorkDay1(Left(strSrvDate(1), 4) & "0820", True)   '216 自評報告:系統日(原本是收文日的當年度)8月20日
                 strBCP06List = strBCP06List & "," & PUB_GetWorkDay1(Left(strSrvDate(1), 4) & "0831", True)   '218 驗證申請:系統日(原本是收文日的當年度)8月31日
              End If
              strBCP06List = Mid(strBCP06List, 2)
           'Modified by Lydia 2023/12/22
           Case "B"   '【方案B】:分案隔年度申請驗證
              strBCP06List = ""
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1(GetLastDay((Left(DBDATE(strSrvDate(1)), 4) + 1) & "0215"), True)   '208 啟始會議:分案隔年度2月15日
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(strSrvDate(1)), 4) + 1) & "0531", True)   '2091 權責人員教育訓練:分案隔年度5月31日
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(strSrvDate(1)), 4) + 1) & "0630", True)   '2092 全體員工教育訓練:分案隔年度6月30日
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(strSrvDate(1)), 4) + 1) & "0531", True)   '211 文件修制訂:分案隔年度5月31日
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(strSrvDate(1)), 4) + 1) & "0715", True)   '213 內部稽核:分案隔年度7月15日
              strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(strSrvDate(1)), 4) + 1) & "0815", True)   '215 管理審查:分案隔年度8月15日
              If Text(9) <> "1013" Then '判斷1013(抽驗L)不收213內部稽核,215管理審查
                 strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(strSrvDate(1)), 4) + 1) & "0820", True)   '216 自評報告:分案隔年度8月20日
                 strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(strSrvDate(1)), 4) + 1) & "0831", True)   '218 驗證申請:分案隔年度8月31日
              End If
              strBCP06List = Mid(strBCP06List, 2)
           'Mark by Lydia 2023/12/22
           'Case "B"
           '   '【方案B】:收文當年度作智財報告書+隔年度申請驗證
           '   strBCP06List = ""
           '   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1(GetLastDay((Left(DBDATE(Text(0)), 4) + 1) & "0201"), True)   '208 啟始會議:收文隔年度2月
           '   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(Text(0)), 4) + 1) & "0531", True)   '2091 權責人員教育訓練:收文隔年度5月31日
           '   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(Text(0)), 4) + 1) & "0630", True)   '2092 全體員工教育訓練:收文隔年度6月30日
           '   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(Text(0)), 4) + 1) & "0531", True)   '211 文件修制訂:收文隔年度5月31日
          '    strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(Text(0)), 4) + 1) & "0715", True)   '213 內部稽核:收文隔年度7月15日
           '   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(Text(0)), 4) + 1) & "0815", True)   '215 管理審查:收文隔年度8月15日
           '   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(Text(0)), 4) + 1) & "0820", True)   '216 自評報告:收文隔年度8月20日
           '   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(DBDATE(Text(0)), 4) + 1) & "0831", True)   '218 驗證申請:收文隔年度8月31日
           '   strBCP06List = Mid(strBCP06List, 2)
           'Case "D"
           '   '【方案D】:前一年度有自行申請首次驗證：改抓前一年度的同一案件性質的所限+1年
           '   'Mark by Lydia 2023/06/21 改用前一年度的同一案件性質的所限+1年；若無則比照方案A(分案日和分案當年度)
           '   'If strFirstCP06 = "" Then
           '   '   strExc(0) = PUB_GetWorkDay1(GetLastDay((Left(DBDATE(Text(0)), 4) + 1) & "0201"), True)
           '   'Else
           '   '   strExc(0) = CompDate(0, 1, strFirstCP06)
           '   'End If
           '   'strBCP06List = ""
            '  'end 2023/06/21
            '  For i = 0 To UBound(arrTmp1)
            '      '改抓前一年度的同一案件性質的所限+1年
            '      'Mark by Lydia 2023/06/21
            '      'If i = 0 Then
            '      '   strBCP06List = strBCP06List & "," & strExc(0)
            '      'Else
             '     'end 2023/06/21
            '         strSql = "select cp06 from caseprogress where cp01='" & lc01 & "' and cp02='" & lc02 & "' and cp03='" & lc03 & "' and cp04='" & lc04 & "' and cp10='" & arrTmp1(intI) & "' and cp159=0 "
            '         strSql = strSql & " and cp05 < " & Left(strSrvDate(1), 4) & "0101" 'Added by Lydia 2023/06/21 前一年度
            '         intI = 1: strExc(1) = ""
            '         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            '         If intI = 1 Then
            '            strExc(1) = "" & RsTemp.Fields("cp06")
            '         End If
            '         If strExc(1) = "" Then
            '             'Modified by Lydia 2023/06/21 若無前一年度收文則比照方案A
            '             'If strFirstCP06 <> "" Then
            '             '   strExc(1) = strFirstCP06
            '             'Else
            '             '   strExc(1) = DBDATE(Text(0))
             '            'End If
             '            strExc(1) = CompDate(0, -1, strSrvDate(1))
             '            'end 2023/06/21
             '            Select Case Trim(arrTmp1(i))
             '                'Added by Lydia 2023/06/21
             '                Case "208"
             '                   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(strExc(1), 4) + 1) & "0201", True)
             '                'end 2023/06/21
             '                Case "2091"
             '                   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(strExc(1), 4) + 1) & "0531", True)
              '               Case "2092"
             '                   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(strExc(1), 4) + 1) & "0630", True)
             '                Case "211"
             '                   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(strExc(1), 4) + 1) & "0531", True)
             '                Case "213"
             '                   strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(strExc(1), 4) + 1) & "0715", True)
             '                Case "215"
              '                  strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(strExc(1), 4) + 1) & "0815", True)
              '               Case "216"
              '                  strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(strExc(1), 4) + 1) & "0820", True)
              '               Case "218"
              '                  strBCP06List = strBCP06List & "," & PUB_GetWorkDay1((Left(strExc(1), 4) + 1) & "0831", True)
              '           End Select
              '       Else
              '           strBCP06List = strBCP06List & "," & CompDate(0, 1, strExc(1))
              '       End If
              '    'End If 'Mark by Lydia 2023/06/21
              'Next i
              'strBCP06List = Mid(strBCP06List, 2)
       End Select
       
       arrTmp2 = Split(strBCP06List, ",")
       strExc(1) = GetST15(Text(10))
       For intI = 0 To UBound(arrTmp2)
           strExc(2) = AutoNo("B", 6)
           'Modified by Lydia 2023/06/21 +CP43
           strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp11,cp12,cp13,cp14,cp20,cp32,cp43) " & _
                        "values ('" & lc01 & "','" & lc02 & "','" & lc03 & "','" & lc04 & "'," & strSrvDate(1) & "," & CNULL("" & arrTmp2(intI), True) & ",'" & strExc(2) & "' " & _
                        ",'" & arrTmp1(intI) & "','90'," & CNULL(strExc(1)) & "," & CNULL(Text(10)) & "," & CNULL(Text(7)) & ",'N','N','" & lbePaperNum.Caption & "' )"
           cnnConnection.Execute strSql
       Next intI
   End If
   'end 2023/04/14
   
   If SaveData Then blnIsSave = True
   frm081031.SetDataComplete lbePaperNum.Caption
   
'   'add by nickc 2005/03/17 加入加乘註記及寄件值
'   m_CP98 = "": m_CP101 = "": m_CP104 = ""
'   If PUB_GetFlagValue(Me.lbePaperNum.Caption, m_CP98, m_CP101, m_CP104) = True Then
'      strSql = "update caseprogress set cp98=" & m_CP98 & ",cp101=" & m_CP101 & ",cp104=" & m_CP104 & " WHERE CP09 = '" & Me.lbePaperNum.Caption & "' "
'      cnnConnection.Execute strSql
'   End If
   
   'Add By Sindy 2022/12/15
   If Me.Text(9).Tag <> Me.Text(9).Text Then
      '有修改案件性質
      'Modify By Sindy 2023/9/22 + , Text(9).Tag
      If PUB_ModCrLCRCData(lbePaperNum, txtF0301, Text(9), Text(9).Tag, "000", Text(18)) = False Then
         GoTo CheckingErr
      End If
   End If
   '2022/12/15 END
   
   cnnConnection.CommitTrans
   Exit Function
   
CheckingErr:
   SaveData = False
   cnnConnection.RollbackTrans
End Function

Private Sub Getrs()
   strExc(1) = "select '',decode(np02||np07,cpm01||CPM02,CPM03,CPM04)," + _
      "decode(np08,null,'',SUBSTR(np08,1,4)-1911||'/'||SUBSTR(np08,5,2)||'/'||SUBSTR(np08,7,2))," + _
      "decode(np09,null,'',SUBSTR(np09,1,4)-1911||'/'||SUBSTR(np09,5,2)||'/'||SUBSTR(np09,7,2))," + _
      "np13,np14,decode(np11,null,'',SUBSTR(np11,1,4)-1911||'/'||SUBSTR(np11,5,2)||'/'||" + _
      "SUBSTR(np11,7,2)),np06,np01,np07,np16,np17,np18,np15,np22 from nextprogress,CASEPROPERTYMAP where " + _
      "" + ChgNextProgress(LcTmp) + " and (np02=cpm01(+) and np07=cpm02(+)) and (np06='N' or np06 is null)"
   intI = 1
   Set MSHFlexGrid1.Recordset = ClsLawReadRstMsg(intI, strExc(1))
   GridHead
End Sub

Private Sub GridHead()
Dim i As Integer
   
   With MSHFlexGrid1
      blnOKtoShow = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 900: .Text = "下一程序"
      .col = 2: .ColWidth(2) = 1000: .Text = "本所期限"
      .col = 3: .ColWidth(3) = 900: .Text = "法定期限"
      .col = 4: .ColWidth(4) = 900: .Text = "機關文號"
      .col = 5: .ColWidth(5) = 900: .Text = "相關人"
      .col = 6: .ColWidth(6) = 1500: .Text = "解除期限日期"
      .col = 7: .ColWidth(7) = 0
      .col = 8: .ColWidth(8) = 0
      .col = 9: .ColWidth(9) = 0
      .col = 10: .ColWidth(10) = 0
      .col = 11: .ColWidth(11) = 0
      .col = 12: .ColWidth(12) = 0
      .col = 13: .ColWidth(13) = 1500: .Text = "備註"
      .col = 14: .ColWidth(14) = 0
      intLastRow = 0
      blnOKtoShow = True
   End With
End Sub

Private Function SaveNextProgress() As Boolean
Dim i As Integer, n As Integer, NP07 As String, np08 As String
Dim np16 As String, np17 As String, np18 As String, np06 As String, np01 As String
Dim np22 As String
   
   With MSHFlexGrid1
      n = 0
         For i = 1 To .Rows - 1
          .row = i
          .col = 0
           If .Text = "v" Then
               .col = 2
                np08 = ChangeTStringToWString(Replace(.Text, "/", ""))
               .col = 8
               np01 = .Text
               .col = 9
               NP07 = .Text
               .col = 10
               np16 = .Text
               .col = 11
               np17 = .Text
               .col = 12
               np18 = .Text
               .col = 14
               np22 = .Text
               np06 = "Y"
               strExc(1) = "update nextprogress set np06=" & CNULL(np06) & _
                  " where np01=" & CNULL(np01) & " and " & ChgNextProgress(LcTmp) & _
                  " and np07=" & CNULL(NP07) & " and np08=" & CNULL(np08) & _
                  " and np16=" & CNULL(np16) & " and np17=" & CNULL(np17) & _
                  " and np18=" & CNULL(np18) & " and np22=" & CNULL(np22)
               cnnConnection.Execute strExc(1)
           End If
         Next
   End With
   SaveNextProgress = True
End Function

Private Function ChangText() As Boolean
Dim i As Integer, strTemp As String, yn As Integer, strlcTemp As String

   strlcTemp = GiveSymbol(txtcp01, txtcp02, txtcp03, txtcp04)
   If lbeNumber = strlcTemp Then
      MsgBox "此本所案號與原本所案號相同", vbCritical
      ChangText = True
      txtcp01 = ""
      txtcp02 = ""
      txtcp03 = ""
      txtcp04 = ""
      Exit Function
   End If
   
   If ClsPDChkCaseNum(txtcp01, txtcp02) Then
      TextInverse txtcp02
      ChangText = True
   Else
      If txtcp01 = "LA" Then i = 2 Else i = 1
      If Not ClsPDCheckIsExistCaseNum(i, Replace(strlcTemp, "-", ""), strTemp) Then
         yn = MsgBox("" + strTemp + ",是否要轉入此本所案號", vbYesNo)
         Select Case yn
            Case 6
              strOldLc = lbeNumber
               blnIsNew = True
              Getrs
            Case 7
              txtcp01 = ""
              txtcp02 = ""
              txtcp03 = ""
              txtcp04 = ""
         End Select
      Else
        MsgBox "" + strTemp + "", vbCritical
        ChangText = True
      End If
   End If
End Function

Private Function AllTextBeforeSaveCheck() As Boolean
Dim i As Integer
Dim strTempName As String
Dim blnIsEmpty As Boolean

   strTempName = ""
   AllTextBeforeSaveCheck = True
   If Text(0) = "" Then
      MsgBox "收文日不可空白", vbCritical
      Text(0).SetFocus
      AllTextBeforeSaveCheck = True
      Exit Function
   End If
  
   If Text(2) = "" And Text(3) = "" And Text(4) = "" Then
      MsgBox "案件名稱不可全部空白!", vbCritical
      Text(2).SetFocus
      AllTextBeforeSaveCheck = True
      Exit Function
   End If

   If Text(1) = "" Then
   Else
      Text(1) = UCase(Text(1))
      If Left(Text(1), 1) = "Y" Then
         Text(1) = "X" & Mid(Text(1), 2)
      ElseIf Left(Text(1), 1) <> "X" Then
             MsgBox "當事人代碼輸入錯誤!", vbCritical
             TextInverse Text(1)
             AllTextBeforeSaveCheck = True
             Exit Function
      End If
      If ClsPDGetCustomer(Text(1), strTempName) Then
         lbe(1) = strTempName
      Else
          lbe(1) = ""
          AllTextBeforeSaveCheck = True
          Exit Function
      End If
   End If
   
   If Text(7) <> "" Then
      Text(7) = UCase(Text(7))
      If ClsPDGetStaff(Text(7), strTempName) Then
         lbe(7) = strTempName
      Else
         lbe(7) = ""
         Text(7).SetFocus
         TextInverse Text(7)
         AllTextBeforeSaveCheck = True
         Exit Function
      End If
   End If

   If Text(9) = "" Or IsNull(Text(9)) Then
      MsgBox "案件性質不可空白", vbCritical
      Text(9).SetFocus
      AllTextBeforeSaveCheck = True
      Exit Function
   End If

   strTempName = ""
   If Text(10) <> "" Then
      Text(10) = UCase(Text(10))
      If ClsPDGetStaff(Text(10), strTempName) Then
         lbe(10) = strTempName
      Else
         lbe(10) = ""
         Text(10).SetFocus
         TextInverse Text(10)
         AllTextBeforeSaveCheck = True
         Exit Function
      End If
   End If

   If Text(12) <> "" Then
      If CheckIsTaiwanDate(Text(12)) Then
          If Text(13) <> "" Then
             If Val(Text(13)) - Val(Text(12)) < 0 Then DataErrorMessage 13
          End If
      Else
        Text(12).SetFocus
        TextInverse Text(12)
        AllTextBeforeSaveCheck = True
        Exit Function
      End If
    End If
      If m_ODate <> "" Then
         If Text(12) <> m_ODate Then
             If MsgBox("是否要修改本所期限?", vbYesNo, "修改") = vbNo Then
                 Text(12) = m_ODate
             End If
         End If
       End If

      If Text(13) <> "" Then
        If CheckIsTaiwanDate(Text(13)) Then
            If Text(12) <> "" Then
               If Val(Text(13)) - Val(Text(12)) < 0 Then DataErrorMessage 12
            End If
        Else
           Text(13).SetFocus
           TextInverse Text(13)
           AllTextBeforeSaveCheck = True
           Exit Function
        End If
     End If
   If m_LDate <> "" Then
      If Text(13) <> m_LDate Then
         If MsgBox("是否要修改法定期限?", vbYesNo, "修改") = vbNo Then
            Text(13) = m_LDate
         End If
      End If
   End If
   
   Text(14) = UCase(Text(14))
   If Text(14) <> "" And Text(14) <> "N" Then
       DataErrorMessage 1, "是否算案件數"
       AllTextBeforeSaveCheck = True
       Text(14).SetFocus
       TextInverse Text(14)
       Exit Function
   End If
   
   If Text(15) <> "" Then
      Text(15) = UCase(Text(15))
      If Text(15) = lbePaperNum Then
         MsgBox "且不可為本身之收文號", vbCritical
         AllTextBeforeSaveCheck = True
         Text(15).SetFocus
         TextInverse Text(15)
         Exit Function
      End If
      If Not ClsLawGetRelation(LcTmp, lbePaperNum, Text(15)) Then
         AllTextBeforeSaveCheck = True
         Text(15).SetFocus
         TextInverse Text(15)
         Exit Function
      End If
   End If

   If Text(16) <> "" Then
         Text(16) = UCase(Text(16))
       If Text(16) = "Y" Then
          i = MsgBox("確定取消閉卷?", vbYesNo, "詢問")
          If i = 7 Then Text(16) = ""
       Else
          DataErrorMessage 1, "是否閉卷"
          AllTextBeforeSaveCheck = True
          Text(16).SetFocus
          TextInverse Text(16)
          Exit Function
       End If
   End If
   
   Text(17) = UCase(Text(17))
   If Text(17) <> "" And Text(17) <> "N" Then
      DataErrorMessage 1, "是否向客戶收款"
      AllTextBeforeSaveCheck = True
      Text(17).SetFocus
      TextInverse Text(17)
      Exit Function
   End If

   If txtcp02 <> "" And txtcp03 = "" Then txtcp03 = "0"
   If txtcp02 <> "" And txtcp04 = "" Then txtcp04 = "00"

   If txtcp02 <> "" Then
      If txtcp01 = lc01 And txtcp02 = lc02 And txtcp03 = lc03 And txtcp04 = lc04 Then
         MsgBox "轉本所案號不可與原本所案號相同 !", vbCritical
         AllTextBeforeSaveCheck = True
         txtcp02.SetFocus
         Exit Function
      End If
   End If

   AllTextBeforeSaveCheck = False
End Function

Private Function CheckCaseNum() As String
Dim strKind As String, i As Integer
   
   For i = 1 To 4
      If Mid(lbeNumber, i, 1) = "-" Then
         CheckCaseNum = Left(lbeNumber, i - 1)
         Exit For
      End If
   Next
End Function

Private Sub ChgType(i As Integer)
Dim strTempName As String, blnIsEmpty As Boolean
   
   lbe(i) = ""
   Select Case i
      Case 1
         If Text(i) <> "" Then
            If ClsPDGetCustomer(Text(i), strTempName) Then
               lbe(i) = strTempName
            End If
         Else
             MsgBox "當事人不可空白", vbCritical
         End If
      Case 7, 10
         If Text(i) <> "" Then
            If ClsPDGetStaff(Text(i), strTempName) Then
               lbe(i) = strTempName
            End If
         End If
      Case 9
         If Text(i) = "" Then
            MsgBox "案件性質不可空白", vbCritical
         Else
            If ClsPDGetCaseProperty(CheckCaseNum, Text(i), strTempName, False) Then
               lbe(i) = strTempName
            End If
         End If
      Case 20
         If Text(i) <> "" Then
            If ClsPDGetNation(Text(i), strTempName) = True Then
               lbe(i).Caption = strTempName
            End If
         End If
   End Select
End Sub

Private Sub txtcp04_GotFocus()
   TextInverse txtcp04
End Sub

Private Sub txtcp04_Validate(Cancel As Boolean)
   If txtcp02 <> "" And txtcp04 = "" Then txtcp04 = "00"
   If txtcp04 <> "" Then
      If Len(txtcp04) <> 2 Then
         DataErrorMessage 1, "本所案號"
         Cancel = True
      Else
         If ChangText Then TextInverse txtcp02
      End If
   End If
   
   If Cancel = False Then
      If Me.txtcp01.Text <> "" And Me.txtcp02.Text <> "" Then
         MsgBox "不可同時做" & """轉本所案號""" & "及" & """修改其他欄位""" & "，否則資料可能有不一致的情形產生!!!", vbExclamation + vbOKOnly
      End If
   End If
   If Cancel Then TextInverse txtcp04
End Sub

Private Sub ClearForm()
Dim i As Integer
  
   For i = 0 To 19
      If i <> 5 And i <> 8 And i <> 11 Then
         Text(i).Text = ""
      End If
   Next i
   MSHFlexGrid1.Clear
   MSHFlexGrid1.Rows = 2
   txtcp01.Text = ""
   txtcp02.Text = ""
   txtcp03.Text = ""
   txtcp04.Text = ""
   lbePaperNum.Caption = ""
   lbeNumber.Caption = ""
   lbeCloseDate.Caption = ""
   lbeCost.Caption = ""
   lbePointNum.Caption = ""
   lbeMoney.Caption = ""
   txtF0301 = Empty 'Add by Amy 2022/12/17
   'Added by Lydia 2023/04/14
   txtKind = ""
   strBCP06List = ""
   strFirstCP06 = ""
   'end 2023/04/14
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim bolData As Boolean, strMCTF(0) As String, strTmp(0) As String

   TxtValidate = False
   For Each objTxt In Text
      If objTxt.Enabled = True Then
         Cancel = False
         Text_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   If Me.txtcp01.Enabled = True Then
      Cancel = False
      txtcp01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtcp02.Enabled = True Then
      Cancel = False
      txtcp02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.txtcp03.Enabled = True Then
      Cancel = False
      txtcp03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add by Amy 2023/08/18 智權公司的判斷
   If TextLC48.Visible = True And TextLC48 = "J" Then
      '檢查客戶是否有設不開發票
      'Modify By Sindy 2024/9/25 增加傳入公司別做判斷
      If PUB_ChkCU144isN(Mid(m_LC11, 1, 8), Mid(m_LC11, 9, 1), "", TextLC48, False) = True Then
         MsgBox m_LC11 & "此客戶為不開發票，因此特殊出名公司不可輸智權公司 !", vbCritical
         TextLC48.SetFocus
         Exit Function
      End If
      '收文進度若有專利布局分析,不可開智權公司
      strExc(0) = ""
      strExc(1) = lc01
      strExc(2) = lc02
      strExc(3) = lc03
      strExc(4) = lc04
      If PUB_ChkCPExist(strExc, "113", , strExc(0)) = True Then
         MsgBox "ACS案有收文113專利布局分析，收據公司別不可選擇智權公司!!!"
         TextLC48.SetFocus
         Exit Function
      End If
   End If
   'end 2023/08/18
   
   'add by sonia 2023/11/9
   If Text(9) = "706" And Text(15) = "" Then
      MsgBox lbe(9) & "必須輸入相關總收文號！可由 案件進度 按鈕選取！", vbExclamation
      Text(15).SetFocus
      Exit Function
   End If
   'end 2023/11/9
   'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
   strExc(1) = ChangeCustomerL(Text(1))
   strExc(2) = ChangeCustomerL(m_LC11)
   If strExc(1) <> "" And strExc(1) <> strExc(2) Then
      If GetCustomerAndState(strExc(1), strExc(3), , , , lc01, strExc(8), False, Me.Name, lc02, lc03, lc04) = False Then
         Text(1).SetFocus
         Text_GotFocus 1
         Exit Function
      End If
   End If
   'end 2024/06/13

   'Remove by Lydia 2020/11/06 取消ACS的案件屬性欄; 物件直接隱藏
   'If Combo2.Text = "" Then
   '   MsgBox "請勾選案件屬性!", vbExclamation + vbOKOnly
   '   Exit Function
   'End If
   'end 2020/11/06
     
   TxtValidate = True
End Function

'Added by Lydia 2021/04/28 ACS智財顧問專業分配比例管制
Private Sub cmdPFrate_Click()
    If PUB_CheckFormExist("frm081031_3") Then
        MsgBox "請先關閉〔智財顧問專業分配比例〕畫面！"
        Exit Sub
    End If
        
    Call frm081031_3.SetParent(Me, lbePaperNum.Caption, "M")
    Me.Hide
    frm081031_3.Show
End Sub

'Added by Lydia 2023/04/14
Private Sub txtkind_GotFocus()
   TextInverse txtKind
End Sub

Private Sub txtKind_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtKind_Validate(Cancel As Boolean)
   If txtKind <> "" Then
      'Modified by Lydia 2023/12/22 簡化自動收文為【方案A~C】
      'If Text(9) = "101" Then
      'Modified by Lydia 2024/04/12 +131前置自行申請首次驗證,141諮詢再驗證,142諮詢抽驗Ｌ
      If Text(9) = "101" Or Text(9) = "1012" Or Text(9) = "1013" Or Text(9) = "131" Or Text(9) = "141" Or Text(9) = "142" Then
         If txtKind <> "A" And txtKind <> "B" And txtKind <> "C" Then
             MsgBox "請輸入A~C！", vbCritical
             Cancel = True
             txtKind.SetFocus
             Exit Sub
         End If
      'Modified by Lydia 2023/12/22 簡化自動收文為【方案A~C】
      'ElseIf Text(9) = "1012" Then
      '   If txtKind <> "D" And txtKind <> "E" Then
      '       MsgBox "請輸入D~E！", vbCritical
      '       Cancel = True
      '       txtKind.SetFocus
      '       Exit Sub
      '   End If
      ''Modified by Lydia 2023/12/05 1013=>抽驗L(1013)，新增:抽驗S(1014) +Or Text(9) = "1014"
      'ElseIf Text(9) = "1013" Or Text(9) = "1014" Then
       '  If txtKind <> "F" Then
       '      MsgBox "請輸入F！", vbCritical
       '      Cancel = True
       '      txtKind.SetFocus
       '      Exit Sub
       '  End If
      Else
         If txtKind <> "C" Then
             MsgBox "請輸入C！", vbCritical
             Cancel = True
             txtKind.SetFocus
             Exit Sub
         End If
      End If
   End If
End Sub

'Added by Lydia 2023/04/14 TIPS自動內部收文：顯示欄位
Private Sub GetKindType()
   lblKind(0).Visible = False: lblKind(1).Visible = False: txtKind.Visible = False
   'strExc(1) = lc01: strExc(2) = lc02: strExc(3) = lc03: strExc(4) = lc04 'Mark by Lydia 2023/06/21
   cmdCP06.Visible = False
   txtKind = ""
   strBCP06List = ""
   'Modified by Lydia 2023/06/21 + Or Text(9) = "1012"
   'Modified by Lydia 2023/12/05 1013=>抽驗L(1013)，新增:抽驗S(1014) +Or Text(9) = "1014"
   'Modified by Lydia 2023/12/22 拿掉1014
   'Modified by Lydia 2024/04/12 +131前置自行申請首次驗證,141諮詢再驗證,142諮詢抽驗Ｌ
   If Left(lbePaperNum, 1) = "A" And (Text(9) = "101" Or Text(9) = "1012" Or Text(9) = "1013" Or Text(9) = "131" Or Text(9) = "141" Or Text(9) = "142") Then
      'Modified by Lydia 2023/06/21 判斷相關總收文號
      'If PUB_ChkCPExist(strExc, "208", 1) = False Then
      strExc(0) = "select cp09 from caseprogress where cp01='" & lc01 & "' and cp02='" & lc02 & "' and cp03='" & lc03 & "' and cp04='" & lc04 & "' and cp10='208' and cp57 is null and cp27 is null and cp43='" & lbePaperNum.Caption & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
      'end 2023/06/21
         If InStr(Text(18), "TIPS自動內部收文：") = 0 Then
           lblKind(0).Visible = True: lblKind(1).Visible = True: txtKind.Visible = True
           
           'Modified by Lydia 2023/06/21 區分顯示輸入
           'If Text(9) = "101" Then cmdCP06.Visible = True
           'Modified by Lydia 2023/12/22 簡化自動收文為【方案A~C】
           'If Text(9) = "101" Then
           '   cmdCP06.Visible = True
           '   lblKind(1).Caption = "101(自行申請首次驗證)輸入:A當年度申請驗證,B隔年度申請驗證,C其他；"
           'ElseIf Text(9) = "1012" Then
           '   lblKind(1).Caption = "1012(再驗證)輸入：D前一年度有101,E前一年度沒有101"
           ''Modified by Lydia 2023/12/05 1013=>抽驗L(1013)，新增:抽驗S(1014) +Or Text(9) = "1014"
           'ElseIf Text(9) = "1013" Or Text(9) = "1014" Then
           '   cmdCP06.Visible = True
           '   cmdCP06.Caption = "輸入期限"
           '   'Modified by Lydia 2023/12/05
           '   'lblKind(1).Caption = "1013(抽驗)輸入:F自行輸入期限；"
           '   lblKind(1).Caption = IIf(Text(9) = "1013", "1013(抽驗L)", "1014(抽驗S)") & "輸入:F自行輸入期限；"
           'End If
           cmdCP06.Visible = True
           cmdCP06.Caption = "輸入期限"
           'Modified by Lydia 2024/04/12 +131前置自行申請首次驗證,141諮詢再驗證,142諮詢抽驗Ｌ
           If Text(9) = "101" Or Text(9) = "1012" Or Text(9) = "1013" Or Text(9) = "131" Or Text(9) = "141" Or Text(9) = "142" Then
              lblKind(1).Caption = Text(9) & lbe(9).Caption & "：A分案當年度申請驗證,B分案隔年度申請驗證,C自行輸入期限；"
           Else
              lblKind(1).Caption = Text(9) & lbe(9).Caption & "：C自行輸入期限；"
           End If
           'end 2023/06/21
         End If
      End If
   End If
End Sub

'Added by Lydia  2023/04/14 TIPS自動內部收文：輸入本所期限
Private Sub cmdCP06_Click()

    If PUB_CheckFormExist("frm081031_4") Then
        MsgBox "請先關閉〔TIPS自動內部收文-本所期限〕畫面！"
        Exit Sub
    End If
    'Modified by Lydia 2023/06/21 + txtKind
    'Modified by Lydia 2023/12/22
    'Call frm081031_4.SetParent(Me, lbePaperNum.Caption, txtKind)
    Call frm081031_4.SetParent(Me, lbePaperNum.Caption, IIf(Text(9) = "1013", "F", ""))
    Me.Hide
    frm081031_4.Show
End Sub
