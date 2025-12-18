VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040104_7 
   BorderStyle     =   1  '單線固定
   Caption         =   "內專發文-領證及繳年費"
   ClientHeight    =   5520
   ClientLeft      =   -216
   ClientTop       =   1356
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   9348
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   4380
      MaxLength       =   4
      TabIndex        =   13
      Top             =   3180
      Width           =   540
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   285
      Left            =   2910
      TabIndex        =   87
      Top             =   5190
      Width           =   6075
      Begin VB.TextBox txtCP118 
         Height          =   264
         Left            =   1860
         MaxLength       =   1
         TabIndex        =   25
         Top             =   3
         Width           =   255
      End
      Begin VB.TextBox txtPayToday 
         Height          =   264
         Left            =   5295
         MaxLength       =   1
         TabIndex        =   26
         Top             =   3
         Width           =   255
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "是否電子送件:         (Y:是)"
         Height          =   180
         Index           =   2
         Left            =   690
         TabIndex        =   89
         Top             =   45
         Width           =   1995
      End
      Begin VB.Label lblPayToday 
         AutoSize        =   -1  'True
         Caption         =   "電子送件是否當日扣款:         (Y/N)"
         Height          =   180
         Left            =   3360
         TabIndex        =   88
         Top             =   45
         Width           =   2655
      End
   End
   Begin VB.TextBox txtChkRltDate1 
      Height          =   270
      Left            =   8280
      TabIndex        =   86
      Top             =   4372
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtChkRltDate 
      Height          =   270
      Left            =   1110
      MaxLength       =   8
      TabIndex        =   24
      Top             =   5190
      Width           =   975
   End
   Begin VB.TextBox txtCP71 
      Height          =   270
      Left            =   7230
      MaxLength       =   7
      TabIndex        =   22
      Top             =   4372
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.TextBox txtFee 
      Height          =   270
      Left            =   7575
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   5130
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1110
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   5130
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1380
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   900
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1380
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   5
      Left            =   900
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1650
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   900
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1110
      Width           =   240
   End
   Begin VB.CheckBox chk412 
      Caption         =   "延緩公告發文："
      Enabled         =   0   'False
      Height          =   195
      Left            =   4275
      TabIndex        =   73
      Top             =   4410
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   5205
      MaxLength       =   1
      TabIndex        =   17
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Index           =   0
      Left            =   6870
      MaxLength       =   1
      TabIndex        =   20
      Top             =   4080
      Width           =   255
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   990
      TabIndex        =   18
      Top             =   3780
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   1476
      MaxLength       =   1
      TabIndex        =   16
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   1
      Left            =   6255
      MaxLength       =   1
      TabIndex        =   14
      Top             =   3180
      Width           =   255
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Index           =   1
      Left            =   2100
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2865
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   2
      Left            =   8412
      TabIndex        =   31
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "變更事項(&R)"
      Height          =   405
      Index           =   4
      Left            =   5136
      Style           =   1  '圖片外觀
      TabIndex        =   28
      Top             =   45
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   6360
      TabIndex        =   29
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   7188
      TabIndex        =   30
      Top             =   45
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   405
      Index           =   3
      Left            =   3912
      TabIndex        =   27
      Top             =   45
      Width           =   1200
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   1080
      MaxLength       =   9
      TabIndex        =   21
      Top             =   4365
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   990
      MaxLength       =   8
      TabIndex        =   12
      Top             =   3180
      Width           =   1095
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Index           =   0
      Left            =   2220
      MaxLength       =   1
      TabIndex        =   19
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   4935
      MaxLength       =   1
      TabIndex        =   10
      Top             =   2865
      Width           =   255
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   270
      Index           =   0
      Left            =   996
      MaxLength       =   2
      TabIndex        =   8
      Text            =   "1"
      Top             =   2865
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   35
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   34
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   33
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   32
      Top             =   720
      Width           =   375
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   864
      Left            =   7572
      TabIndex        =   15
      Top             =   3180
      Width           =   1500
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;1524"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   2070
      Width           =   7560
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "13335;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text12 
      Height          =   465
      Left            =   1110
      TabIndex        =   23
      Top             =   4680
      Width           =   7995
      VariousPropertyBits=   -1467987941
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "14102;820"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   2310
      Width           =   7560
      VariousPropertyBits=   671107099
      MaxLength       =   250
      Size            =   "13335;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text15 
      Height          =   300
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   2550
      Width           =   7560
      VariousPropertyBits=   671107099
      MaxLength       =   160
      Size            =   "13335;529"
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
      Left            =   3540
      TabIndex        =   90
      Top             =   3225
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
      Left            =   2100
      TabIndex        =   84
      Tag             =   "Y"
      Top             =   5130
      Width           =   255
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   150
      TabIndex        =   83
      Top             =   5205
      Width           =   765
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人:"
      Height          =   180
      Left            =   6570
      TabIndex        =   82
      Top             =   3225
      Width           =   945
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "延緩月數/日期"
      Height          =   180
      Left            =   6030
      TabIndex        =   81
      Top             =   4410
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblFee 
      AutoSize        =   -1  'True
      Caption         =   "變更規費:"
      Height          =   180
      Left            =   6570
      TabIndex        =   80
      Top             =   2910
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "申請人2:"
      Height          =   180
      Left            =   4380
      TabIndex        =   79
      Top             =   1155
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人1:"
      Height          =   180
      Left            =   180
      TabIndex        =   78
      Top             =   1155
      Width           =   675
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   3
      Left            =   1200
      TabIndex        =   77
      Top             =   1155
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5186;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   4
      Left            =   5445
      TabIndex        =   76
      Top             =   1155
      Width           =   3120
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5503;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCP81 
      AutoSize        =   -1  'True
      Caption         =   "lblCP81"
      Height          =   180
      Left            =   8685
      TabIndex        =   75
      Top             =   930
      Width           =   480
   End
   Begin VB.Label lblCP81C 
      AutoSize        =   -1  'True
      Caption         =   "本案最新減免狀態："
      Height          =   180
      Left            =   7020
      TabIndex        =   74
      Top             =   930
      Width           =   1620
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否修改通知函內容:       (Y:Word)"
      Height          =   180
      Index           =   4
      Left            =   3525
      TabIndex        =   72
      Top             =   3525
      Width           =   2670
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(日):"
      Height          =   180
      Left            =   165
      TabIndex        =   71
      Top             =   2580
      Width           =   1065
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(英):"
      Height          =   180
      Index           =   1
      Left            =   165
      TabIndex        =   70
      Top             =   2340
      Width           =   1065
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱(中):"
      Height          =   180
      Left            =   165
      TabIndex        =   69
      Top             =   2100
      Width           =   1065
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   150
      X2              =   9120
      Y1              =   2010
      Y2              =   2010
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   9120
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "是否修改指示信或申請書內容:              (Y:Word)"
      Height          =   180
      Left            =   4290
      TabIndex        =   68
      Top             =   4125
      Width           =   3705
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   13
      Left            =   7860
      TabIndex        =   67
      Top             =   720
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2328;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "代理人:"
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   66
      Top             =   3825
      Width           =   585
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否列印指示信或申請書:        (N:不印)"
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   65
      Top             =   4125
      Width           =   3750
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "是否列印通知函:       (N:不印)"
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   64
      Top             =   3525
      Width           =   2925
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   1
      Left            =   7020
      TabIndex        =   63
      Top             =   720
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   11
      Left            =   2700
      TabIndex        =   62
      Top             =   3825
      Width           =   4815
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "8493;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家:"
      Height          =   180
      Index           =   1
      Left            =   4380
      TabIndex        =   61
      Top             =   510
      Width           =   765
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   12
      Left            =   5220
      TabIndex        =   60
      Top             =   510
      Width           =   1740
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "3069;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   10
      Left            =   2340
      TabIndex        =   59
      Top             =   4410
      Width           =   1560
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2752;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   9
      Left            =   7860
      TabIndex        =   58
      Top             =   510
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2328;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   7
      Left            =   1200
      TabIndex        =   57
      Top             =   1695
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5186;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   6
      Left            =   5445
      TabIndex        =   56
      Top             =   1425
      Width           =   3120
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5503;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   5
      Left            =   1200
      TabIndex        =   55
      Top             =   1425
      Width           =   2940
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "5186;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   2
      Left            =   5220
      TabIndex        =   54
      Top             =   720
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2117;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   1
      Left            =   5220
      TabIndex        =   53
      Top             =   930
      Width           =   1215
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "2143;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label2 
      Height          =   180
      Index           =   0
      Left            =   1020
      TabIndex        =   52
      Top             =   510
      Width           =   480
      VariousPropertyBits=   27
      Caption         =   "Label2"
      Size            =   "847;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   150
      TabIndex        =   51
      Top             =   4740
      Width           =   765
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "年費通知人:"
      Height          =   180
      Left            =   150
      TabIndex        =   50
      Top             =   4410
      Width           =   945
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   7020
      TabIndex        =   49
      Top             =   510
      Width           =   585
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   120
      TabIndex        =   48
      Top             =   3225
      Width           =   585
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "(Y:雙倍)"
      Height          =   180
      Left            =   5295
      TabIndex        =   47
      Top             =   2910
      Width           =   645
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "費用是否要雙倍:"
      Height          =   180
      Left            =   3555
      TabIndex        =   46
      Top             =   2910
      Width           =   1305
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "年 年費"
      Height          =   180
      Left            =   2820
      TabIndex        =   45
      Top             =   2910
      Width           =   585
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "至"
      Height          =   180
      Index           =   0
      Left            =   1710
      TabIndex        =   44
      Top             =   2910
      Width           =   180
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "繳納第:"
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   43
      Top             =   2910
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   180
      TabIndex        =   42
      Top             =   510
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   0
      Left            =   4380
      TabIndex        =   41
      Top             =   930
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   40
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   0
      Left            =   4380
      TabIndex        =   39
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "申請人3:"
      Height          =   180
      Left            =   180
      TabIndex        =   38
      Top             =   1425
      Width           =   675
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "申請人4:"
      Height          =   180
      Left            =   4380
      TabIndex        =   37
      Top             =   1425
      Width           =   675
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "申請人5:"
      Height          =   180
      Left            =   180
      TabIndex        =   36
      Top             =   1695
      Width           =   675
   End
   Begin VB.Label lblCaseFees 
      BackColor       =   &H80000010&
      Height          =   255
      Left            =   2145
      TabIndex        =   85
      Top             =   5190
      Width           =   255
   End
End
Attribute VB_Name = "frm040104_7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/14 改成Form2.0 (Text15,Text12,lstNameAgent,Label2)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'整理 by Morgan 2005/7/15
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/7/14 改用動態陣列
'Dim pa(1 TO T_PA) As String, cp(T_CP) As String
Dim pa() As String, cp() As String

Dim intWhere As Integer
Dim strCaseFee(1 To 2) As String 'strCaseFee(1) 國家檔中繳費年度，strCaseFee(2) 國家檔中起算日
Dim m_CurrFee As String
'Add By Cheng 2003/04/15
Dim m_strOfficalFee  As String
Dim m_strServiceFee  As String
Dim m_strPoints  As String
'Add By Cheng 2003/04/24
Dim m_strNP09 As String
'Add By Cheng 2003/09/01
Dim m_strNP09_1 As String
'92.7.7 ADD BY SONIA
Dim m_EndDate As String
'Add By Cheng 2003/10/06
Dim m_blnClkChgEvnBtn As Boolean '是否按下變更事項按鈕

'Add by Morgan 2004/6/24
Dim m_str412CP09 As String '延緩公告收文號
Dim m_str412CP71 As String '延緩月數
Dim m_bolNew As Boolean '是否用新法
Public m_bol412 As Boolean '是否有收延緩公告 Modify By Sindy 2020/3/26 + Public
Public m_strOfficalFee1  As String  '證書費
Public m_lngOfficalFee1Year As Long '第一年年費 Add By Sindy 2020/3/27
Public m_lngFee1  As Long  '證書費+第一年年費(未減免)
Public m_lngFee2  As Long  '第二年以後年費(未減免)
Public m_lngDisc As Long '減免金額 Modify By Sindy 2020/3/26 + Public
Public m_lngDisc1Year As Long '第一年減免金額 Add By Sindy 2020/4/8
Dim m_lngSub As Long  '抵減金額

'Add by Morgan 2004/7/22
Dim m_DiscType As String   '減免身分
Dim m_bolActive As Boolean 'Active事件是否已觸發
'Add by Morgan 2004/9/3
Public m_lngFinalFee As Long '總規費(含變更事項) Modify By Sindy 2020/3/26 + Public
Dim m_bolChanged As Boolean '是否有變更事項
'Add by Morgan 2006/10/5
Dim m_strPA14 As String '預估公告日
Dim m_bolNeedReasign As Boolean '是否需要重新委任
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/23 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
Dim m_bolFMP As Boolean 'Add by Morgan 2009/11/13
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/06/20 是否為寰華
'Add by Morgan 2011/12/15
Public m_bolBeCalled As Boolean
Public m_CP01 As String
Public m_CP02 As String
Public m_CP03 As String
Public m_CP04 As String
Public m_CP09 As String
Dim m_Subject As String
Dim m_bolAutoMail As Boolean 'Added by Morgan 2012/3/22
Dim m_AD1516(5, 3) As String 'Added by Morgan 2013/3/25 中小企業減免資格
Dim m_UseNewForm As Boolean 'Added by Morgan 2013/3/25 使用新申請書
Dim m_str414CP09 As String 'Added by Morgan 2103/6/25 回復原狀收文號
Dim m_bolNonTwCase605Alert As Boolean 'Added by Morgan 2017/1/13 非臺灣案年費期限提醒
Dim m_PrevForm As Form '前一畫面 Add By Sindy 2020/3/26
Dim strCP13New As String, strCP12New As String
Dim m_str412AddCP64 As String 'Added by Morgan 2022/12/30

'Add By Sindy 2020/3/26
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 50) As String, strTmp As String, iItemNo As String
'Add By Cheng 2003/01/13
Dim ii As Integer
Dim intJ As Integer, strTmp2 As String
    
    ii = 0
    EndLetter ET01, strReceiveNo, ET03, strUserNum
    If Text7(0).Text = Text7(1).Text Then
        strTmp = "第 " & Text7(0) & " 年年費"
    Else
        strTmp = "第 " & Text7(0) & " 至 " & Text7(1) & " 年年費"
    End If
    
   'Modified by Morgan 2012/3/22
   'If m_bolBeCalled Then
   If m_bolBeCalled Or m_bolAutoMail Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','自動發文','♀')"
   End If
    
    If m_lngFee1 > 0 Then
          
      If m_lngSub > 0 Then
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
          "','抵減金額','" & Format(m_lngSub) & "')"
      End If
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
         ii = ii + 1
         'Modify by Morgan 2004/7/22
         '自然人=1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選5','" & IIf(InStr(m_DiscType, "1") > 0, "■ ", "□ ") & "')"
         ii = ii + 1
         '學校=2
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選6','" & IIf(InStr(m_DiscType, "2") > 0, "■ ", "□ ") & "')"
         ii = ii + 1
         '中小企業=3
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選7','" & IIf(InStr(m_DiscType, "3") > 0, "■ ", "□ ") & "')"
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選2','□ ')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選5','□ ')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選6','□ ')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選7','□ ')"
      End If
      
      '出名否
      If Text8(1) = "N" Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選8','■ ')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選9','□ ')"
      Else
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選8','□ ')"
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
             "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
             "','勾選9','■ ')"
      End If
      
   End If
    
    
   'Added by Morgan 2013/3/25
   '中小企業可減免
   If InStr(m_DiscType, "3") > 0 And m_UseNewForm = True Then
      For intI = 1 To 5
         If pa(25 + intI) <> "" Then
            If txtAD(intI) = "1" Then
               strTmp = "■ 自然人"
            ElseIf txtAD(intI) = "2" Then
               strTmp = "■ 學校"
            ElseIf txtAD(intI) = "3" Then
               strTmp = "■ 中小企業"
            End If
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','減免身分" & intI & "','" & strTmp & "')"
            If txtAD(intI) = "3" Then
               'Modify By Sindy 2020/7/27
               'For intJ = 1 To 4
               For intJ = 5 To 6
               '2020/7/27 END
                  If Val(m_AD1516(intI, 1)) = intJ Then
                     strTmp = "■ "
                     strTmp2 = Format(m_AD1516(intI, 2), "#,###")
                  Else
                     strTmp = "□ "
                     strTmp2 = "　　　"
                  End If
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','勾選" & intI & intJ & "','" & strTmp & "')"
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','數值" & intI & intJ & "','" & strTmp2 & "')"
               Next
            End If
         End If
      Next
   End If
   'end 2013/3/25
   
    
    ii = ii + 1
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
        "','第幾年至幾年費','" & strTmp & "')"
    'Add By Cheng 2003/04/15
    ii = ii + 1
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
        "','規費','" & m_strOfficalFee & "')"
       
    'Add by Morgan 2004/6/25
    ii = ii + 1
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
        "','證書費','" & m_strOfficalFee1 & "')"
   
   'Add by Morgan 2004/6/24
   If m_bolNew Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
        "','機關文書','" & IIf(pa(8) = "2", "處分書", "審定書") & "')"
        
      'Modify by Morgan 2005/11/1 '智慧局又改回合併申請--郭
      'Remove by Morgan 2004/9/3 '延緩公告另外出申請書
'      If m_bol412 = True Then
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'           "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'           "','延緩公告','三、本案因向國外申請專利需要，故同時申請延緩公告" & PUB_ChgNumber2Chinese(txtCP71) & "個月，謹請　鈞局准予辦理，至感德便。" & "')"
'      End If
      iItemNo = 1
      If m_bol412 = True Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
           "','延緩公告申請','（暨申請延緩公告）" & "')"
         
         'Modify by Morgan 2006/10/5 都印延緩日期--95.9.28請作單
         'Modified by Morgan 2019/2/14 改格式
         'iItemNo = iItemNo + 1
         'If m_strPA14 = "" Then
         '   strTmp = PUB_ChgNumber2Chinese(iItemNo) & "、申請延緩公告" & vbCrLf & "　　　1.延緩公告期限：" & IIf(Len(txtCP71) = 1, txtCP71 & "個月", "至民國 " & PUB_DBYEAR(txtCP71) - 1911 & " 年 " & Val(PUB_DBMONTH(txtCP71)) & " 月 " & Val(PUB_DBDAY(txtCP71)) & " 日") & vbCrLf & "　　　2.延緩公告理由：向國外申請專利需要"
         'Else
         '   strTmp = PUB_ChgNumber2Chinese(iItemNo) & "、申請延緩公告" & vbCrLf & "　　　1.延緩公告期限：至民國 " & PUB_DBYEAR(m_strPA14) - 1911 & " 年 " & Val(PUB_DBMONTH(m_strPA14)) & " 月 " & Val(PUB_DBDAY(m_strPA14)) & " 日" & vbCrLf & "　　　2.延緩公告理由：向國外申請專利需要"
         'End If
         If m_strPA14 = "" Then
            strTmp = "　　1.延緩公告期限：" & IIf(Len(txtCP71) = 1, txtCP71 & "個月", "至民國 " & PUB_DBYEAR(txtCP71) - 1911 & " 年 " & Val(PUB_DBMONTH(txtCP71)) & " 月 " & Val(PUB_DBDAY(txtCP71)) & " 日") & vbCrLf & "　　　2.延緩公告理由：向國外申請專利需要"
         Else
            strTmp = "　　1.延緩公告期限：至民國 " & PUB_DBYEAR(m_strPA14) - 1911 & " 年 " & Val(PUB_DBMONTH(m_strPA14)) & " 月 " & Val(PUB_DBDAY(m_strPA14)) & " 日" & vbCrLf & "　　　2.延緩公告理由：向國外申請專利需要"
         End If
         'end 2019/2/14
         
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
              "','延緩公告','" & strTmp & "')"
      End If

      'Add by Morgan 2004/9/3 變更事項
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
      
      'Add by Morgan 2007/6/23
      If m_bolNeedReasign = True Then
         strExc(0) = GetAgentName
         If strExc(0) <> "" Then
            iItemNo = iItemNo + 1
            strTmp = PUB_ChgNumber2Chinese(iItemNo) & "、本案同時辦理重新委任「" & strExc(0) & "」為專利代理人，檢附委任書。"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','同時辦理事項','" & strTmp & "')"
         End If
      End If
   End If
   
   'Add by Morgan 2006/8/22
   '大陸未逾期才有
   If pa(9) = "020" And Val(DBDATE(cp(7))) > Val(strSrvDate(1)) Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
        "','法定期限','" & DBDATE(cp(7)) & "')"
   End If
   
   'Add by Morgan 2013/6/25
   If Text6 = "Y" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','有回復原狀要印','♀')"
   End If
   'end 2013/6/25
   
   'edit by nickc 2007/02/05 不用 dll 了
    'If Not objLawDll.ExecSQL(ii, strTxt) Then
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

      If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
      
   End With
   
   CheckOC
End Function

Public Function Process(Index As Integer) As Boolean
Dim strTmp As String, bolChk As Boolean
Dim strContent As String, strMailTo As String 'Add by Morgan 2010/12/16
Dim strPath As String, strFile As String 'Added by Morgan 2016/3/29
'Added by Lydia 2020/04/07
Dim strFilePath As String '記錄智慧局收文文號
Dim bolUp As Boolean '是否需要上傳檔案到卷宗區
Dim strNewCP64 As String '保留進度備註
Dim str412FilePath As String 'Added by Morgan 2022/12/27

   If Text9.Text = "" Then MsgBox "發文日不可空白，請重新輸入 !", vbCritical: Exit Function
   ' 90.07.10 modify by louis (先Remark)
   'If IsEmptyText(Text7(0)) = True Or IsEmptyText(Text7(1)) = True Then MsgBox "發文日不可空白，請重新輸入 !", vbCritical: Exit Sub
   'Add By Cheng 2002/03/08
   '檢查輸入資料的完整性
   If CheckDataIntegrity = False Then Exit Function
   'Add By Cheng 2002/05/22
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Function
      'Modify By Cheng 2003/04/16
      '若申請國家為台灣
      If pa(9) = 台灣國家代號 Then
         'Add By Cheng 2003/04/15
         '檢查計算出的規費與進度檔的規費是否相同
         If ChkPatentYearFee(pa(9), pa(8), "Y00000001", cp(10), Me.Text7(0).Text, Me.Text7(1).Text, IIf(Me.Text6.Text = "Y", True, False)) = False Then Exit Function
      End If
   
   strNewCP64 = Text12.Text  'Added by Lydia 2020/04/07 保留進度備註
   'Added by Lydia 2020/04/08 電子送件要在發文前，先產生申請書；所以發文不用印
   If txtCP118 = "Y" Then Text8(0) = "N"

   'Add by Morgan 2009/3/23 設定是否算發文室案件
   If pa(9) = 台灣國家代號 Then
      'Added by Lydia 2020/04/07 年費發文開放可電子送件
      m_CP09s = "": m_CP123s = ""
      If Frame1.Visible = True And txtCP118 = "Y" Then
             '電子送件也要記錄主管機關
             'If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text9) = False Then
             If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text9, , True) = False Then
                Exit Function
             End If
    
             strExc(0) = InputBox("請輸入智慧局收文文號!!")
             If strExc(0) = "" Then
                Exit Function
             Else
                strFilePath = strExc(0)  '記錄智慧局收文文號
                strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & Text12.Text  '保留進度備註
             End If
             
            'Added by Morgan 2022/12/27
            If strSrvDate(1) >= "20230101" And m_str412CP09 <> "" Then
JumpToReInput:  'Added by Lydia 2024/10/17
               strExc(0) = InputBox("請再輸入【延緩公告】智慧局收文文號!!")
               If strExc(0) = "" Then
                  Exit Function
               Else
                  'Added by Lydia 2024/10/17 增加延緩公告和領證-智慧局收文文號，不可相同
                  If strExc(0) = strFilePath Then
                     MsgBox "【延緩公告】智慧局收文文號不可與【領證】智慧局收文文號相同!!", vbExclamation
                     strExc(0) = ""
                     GoTo JumpToReInput
                  End If
                  'end 2024/10/17
                  str412FilePath = strExc(0)
                  m_str412AddCP64 = "智慧局收文文號:" & strExc(0) & ";"
               End If
            End If
            'end 2022/12/27
      Else
      'end 2020/04/07
            'Add by Morgan 2009/4/28
            If ModifyDispatchCp130(cp(9), m_CP09s, m_CP123s, m_CP130, Text9) = False Then
               Exit Function
            End If
            If m_CP123s = "Y" Then
            'end 2009/4/28
               'modify by sonia 2014/6/23 加傳發文規費, P-108903
               If ModifyDispatch(cp(9), m_CP09s, m_CP123s, m_strOfficalFee, Text9) = False Then
                   Exit Function
               End If
            End If
      End If 'Added by Lydia 2020/04/07
      
      'Add by Amy 2014/10/14 P台灣案發文控制
      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
        'Modified by Morgan 2016/6/29 不限制P--玲玲確認
        'If pa(1) = "P" And cp(9) < "C" Then
        If cp(9) < "C" Then
        'end 2016/6/29
            If cp(9) < "B" Then
                'A類一定要有接洽單才可發文
                'Modify by Amy 2014/11/27 取消ChkOneDayHasCP27判斷,接洽單改檢查,因考慮可能同時發文其他案件性質情形
                'If PUB_CheckPDF2(cp(9), 0, True, strExc(0)) = False And ChkOneDayHasCP27(pa(1), pa(2), pa(3), pa(4), cp(5) + 19110000) = False Then
                If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
                    Exit Function
                End If
            End If
            'AB類申請書確認檢查,符合條件才可發文
            'Modified by Morgan 2015/3/17
            'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And Text8(0) = "N" And PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            'Modified by Lydia 2020/04/08 排除電子送件
            'If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And Text8(0) = "N" Then
            If PUB_GetST03(cp(14)) = "P12" And Left(m_CP123s, 1) = "Y" And Text8(0) = "N" And txtCP118 <> "Y" Then
               If PUB_CheckPDF2(cp(9), 1, True, strExc(0)) = False Then
            'end 2015/3/17
                  MsgBox "無申請書PDF檔 ,不可發文!", vbInformation
                  Exit Function
               End If 'Added by Morgan 2015/3/17
            End If
        End If
      End If
      'end 2014/10/14
   'Added by Morgan 2016/6/29 非臺灣案電子化
   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      If cp(9) < "B" And Left(cp(12), 1) <> "F" Then
          If PUB_CheckPDF3(Text1, Text2, Text3, Text4) = False Then
              Exit Function
          End If
      End If
   'end 2016/6/29
   End If
 
   
   'Added by Morgan 2012/3/22
   m_bolAutoMail = False
   'Modified by Morgan 2014/6/4 + 改與整批發文相同判斷非台灣
   'If pa(9) = "020" And Not m_bolBeCalled Then
   'Modified by Lydia 2015/05/26 外專人員不彈訊息也不發email
   If Left(Pub_StrUserSt03, 1) <> "F" Then
        'Modified by Morgan 2016/3/30 +修改指示信不必問
        If pa(9) <> "000" And Not m_bolBeCalled And Text5(0) <> "Y" Then
        'end 2014/6/4
           'Modified by Morgan 2016/5/20 指示信電子化,不直接寄送舊開啟編輯畫面
           If MsgBox("是否直接發 E-Mail 給代理人??" & vbCrLf & vbCrLf & "(選否將開啟編輯畫面)", vbYesNo + vbDefaultButton2) = vbYes Then
              m_bolAutoMail = True
           Else
              Text5(0) = "Y"
           End If
        End If
        'end 2012/3/22
   End If
   
   
    'Added by Lydia 2020/04/07 檢查是否有電子送件的檔案
    bolUp = False
    If Frame1.Visible = True And txtCP118.Text = "Y" And strFilePath <> "" Then
        strExc(1) = cp(82)
        If Val(cp(82)) > 0 Then
            If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                 strExc(1) = ""
            End If
        End If
        If Val(strExc(1)) = 0 Then
           '先判斷是否上傳檔案
           If Pub_AutoEsetToCppByP(True, pa(1), pa(2), pa(3), pa(4), pa(8), "", cp(10), strFilePath, Text9.Text) = False Then
                 Exit Function
           End If
      
            'Added by Morgan 2022/12/30
            '延緩公告申請書
            If str412FilePath <> "" Then
              If Pub_AutoEsetToCppByP(True, pa(1), pa(2), pa(3), pa(4), pa(8), "", "412", str412FilePath, Text9.Text) = False Then
                    Exit Function
              End If
            End If
            'end 2022/12/30
                
           bolUp = True
        End If
    End If
    Text12.Text = strNewCP64 '檢查完畢，更新備註欄位
    'end 2020/04/07

   If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical:  Exit Function
   
   Process = True
   
   If m_bolNonTwCase605Alert Then MsgBox "下次年費期限小於下次催年費的期限區間起始日，請確是否需通知下一年度年費期限！", vbExclamation 'Added by Morgan 2017/1/13
   
   'Add by Morgan 2008/2/20 檢查代理人Email(需考慮可能為FF案件)
   PUB_CheckEMail Combo2
   PUB_CheckEMail pa(75), pa(144)
   If pa(145) <> "" Then
      PUB_CheckEMail pa(75), pa(145)
   End If
   'end 2008/2/20
   
   If Text5(1) = "Y" Then
      bolChk = True
   Else
      bolChk = False
   End If
   
   'Add by Morgan 2007/6/14
   m_bolNeedReasign = False
   If pa(9) = "000" Then
      m_bolNeedReasign = PUB_IsLatestAgent(pa(1), pa(2), pa(3), pa(4), strReceiveNo)
      If m_bolNeedReasign = True Then
         PUB_ReAsignInform pa(1), pa(2), pa(3), pa(4), strReceiveNo, True
      End If
   End If
   
   '2012/7/23 add by sonia 檢查計算出的規費與進度檔的規費不同,仍繼續發文時發mail給智權人員
   If pa(9) = "000" Then
      If Val(m_strOfficalFee) <> Val(cp(17)) Then
         '2013/7/2 modify by sonia 改用共用module
         'Modified by Lydia 2020/04/07 +是否電子送件
         'PUB_ChkOfficialFee cp(9), m_strOfficalFee
         PUB_ChkOfficialFee cp(9), m_strOfficalFee, IIf(txtCP118 = "Y", "A", "")
      End If
   End If
   '2012/7/23 end
         
   If Text8(2) <> "N" Then '通知函
      If pa(9) = 台灣國家代號 Then '台灣 00
         'Modify by Amy 2014/09/16
         'strTmp = "00"
         'StartLetter "02", strTmp
         'Modify by Amy 2014/09/19 +if 大對台定稿
         If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then
            strTmp = "31"
         Else
            strTmp = "30"
         End If
         'end 2014/09/19
         StartLetter2 "02", strTmp
         'end 2014/09/16
      ElseIf pa(9) <> 台灣國家代號 Then
         strTmp = "01"
         StartLetter "02", strTmp
      End If
      'StartLetter "02", strTmp 'Modify by Amy 2014/09/09 往上搬
      'Modify by Amy 2014/09/09 +傳strLetterRecNo
      NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
   End If
   If Text5(0) = "Y" Then
      bolChk = True
   Else
      bolChk = False
   End If
   If Text8(0) <> "N" Then '指示信
      If pa(9) = 台灣國家代號 Then '台灣申請書 06
         'Modify by Morgan 2004/7/9
         '93.7.1以後核准用官方申請書
         If m_bolNew = True Then
            '判斷可否抵減
            If m_lngSub > 0 Then
               strTmp = "15"
            Else
               'Modified by Morgan 2013/3/26 若有中小企業則跑新申請書(要印減免資格),檢查系統已無可抵減案件
               'strTmp = "18"
               If InStr(m_DiscType, "3") > 0 And m_UseNewForm = True Then
                  '多人申請
                  If pa(27) <> "" Then
                     strTmp = "17"
                  Else
                     strTmp = "16"
                  End If
               Else
                  strTmp = "18"
               End If
               'end 2013/3/26
            End If
            
'Remove by Morgan 2005/10/1 智慧局申請書又合併--郭
'                  'Add by Morgan 2004/8/17
'                  If m_bol412 = True Then
'                     StartLetter1 m_str412CP09, "01", "00"
'                     NowPrint m_str412CP09, "01", "00", bolChk, strUserNum, 0
'                  End If
            'Add end 2004/8/17
         Else
            'Modify By Cheng 2003/01/13
            '預設未逾期的處理狀況
            strTmp = "06"
            '若有法定期限
            'Modify By Cheng 2003/05/07
'               If Me.Label2(13).Caption <> "" Then
            If m_strNP09 <> "" Then
                '若發文日大於法定期限(逾期)
'                   If DBDATE(Text9.Text) > DBDATE(Me.Label2(13).Caption) Then
                'Modify By Cheng 2003/09/01
                '若法定期限為假日則用大於法定期限最近的工作日與發文日比較
'                   If DBDATE(Text9.Text) > DBDATE(m_strNP09) Then
                If DBDATE(Text9.Text) > IIf(DBDATE(m_strNP09) >= DBDATE(m_strNP09_1), DBDATE(m_strNP09), DBDATE(m_strNP09_1)) Then
                    strTmp = "07"
                End If
            End If
            'Modify by Morgan 2004/7/5
            If cp(81) = "Y" Then
               strTmp = Format(Val(strTmp) + 4, "00")
            End If
         End If
         StartLetter "01", strTmp
         'Modify by Amy 2014/09/09 +傳strLetterRecNo 及台灣案申請書修改改開1105_1
         NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum, , , , , , , , , , , , , strReceiveNo
         If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
            If pa(9) = 台灣國家代號 And Text8(0) <> "N" And Text5(0) = "Y" Then
             frm1105_1.m_RecNo = strReceiveNo
             frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".DATA.PDF"
             frm1105_1.Show
            End If
         End If
         'end 2014/09/09
         'Add by Morgan 2010/12/16
         If m_bolBeCalled Then
            PUB_PrintLetter strReceiveNo, True
         End If
      
      ElseIf pa(9) = "020" Then
         strExc(0) = "SELECT COUNT(*) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
            " AND NP11 IS NOT NULL AND NP07='" & cp(10) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If RsTemp.Fields(0) > 0 Then '大陸曾結案 35
            strTmp = "35"
         Else '大陸未曾結案 34
            strTmp = "34"
         End If
         StartLetter "02", strTmp
         
         'Modify by Morgan 2010/12/16
         'NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, 0
         'Modified by Morgan 2012/3/22
         'If m_bolBeCalled Then
         'Modified by Morgan 2016/3/30
         '指示信電子化
         'If m_bolBeCalled Or m_bolAutoMail Then
         '   strMailTo = PUB_GetCFAgentEMail(strReceiveNo)
         '   'Modify by Amy 2014/09/09 +傳strLetterRecNo
         '   NowPrint strReceiveNo, "02", strTmp, False, strUserNum, 0, , True, strContent, 1, , , , , True, , , strReceiveNo
         '   'Modified by Morgan 2014/5/19 mail 內容 word 格式未轉換(Ex 底線...),改加轉tag函數並用HTML方式寄發
         '   'PUB_SendMail strUserNum, strMailTo, "", m_Subject, strContent, , , , , , , "patent"
         '   PUB_SendMail strUserNum, strMailTo, "", m_Subject, ChgHTMLFormat(strContent), , , True, , , , "patent"
         'Else
         '   'Modify by Amy 2014/09/09 +傳strLetterRecNo
         '   NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, , , , , , , , , , , , , strReceiveNo
         'End If
         If m_bolBeCalled Or m_bolAutoMail Then
            strPath = "": strFile = ""
            '指示信轉pdf檔
            If PUB_PrintDocAsPdf(strReceiveNo, "02", strTmp, strReceiveNo, strPath, strFile) Then
               'FMP案紙本還是要印
               'Removed by Morgan 2016/5/20 不印指示信,只印寄件備份就好--玲玲
               'If m_bolFMP Then
               '   PUB_PrintPDF strPath & "\" & strFile
               'End If
               'end 2016/5/20
               
               '上傳卷宗區
               If PUB_UploadOrderLetter(strPath & "\" & strFile, strReceiveNo) Then
                  'EMail
                  PUB_SendOrderLetterP strReceiveNo, m_Subject, True
               End If
            End If
         Else
            If Left(Pub_StrUserSt03, 1) = "F" Then
               NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum
            Else
               NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, , , , , , , , , , , , , strReceiveNo
               If bolChk Then
                  frm1105_1.m_RecNo = strReceiveNo
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".DATA.PDF"
                  frm1105_1.m_Subject = m_Subject
                  frm1105_1.Show
               End If
            End If
         End If
         'end 2016/3/30
      End If
   End If

   'Added by Lydia 2020/04/07 是否可以上傳檔案,前面已判斷
   If bolUp = True Then
      If Pub_AutoEsetToCppByP(False, pa(1), pa(2), pa(3), pa(4), pa(8), cp(9), cp(10), strFilePath, Text9) = False Then
           Exit Function
      End If
      
      'Added by Morgan 2022/12/30
      '延緩公告申請書
      If str412FilePath <> "" Then
        If Pub_AutoEsetToCppByP(True, pa(1), pa(2), pa(3), pa(4), pa(8), m_str412CP09, "412", str412FilePath, Text9.Text) = False Then
              Exit Function
        End If
      End If
      'end 2022/12/30
            
   End If
   'end 2020/04/07
   
End Function

Private Sub cmdok_Click(Index As Integer)
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   Select Case Index
      Case 0, 3 '確定
         'Modify by Morgan 2010/2/10 改呼叫函數方式以便鎖定按鍵
         cmdOK(Index).Enabled = False
         If Not Process(Index) Then
            cmdOK(Index).Enabled = True
         Else
            ' 90.07.11 modify by louis (回第一個畫面清除)
            If Index = 3 Then
               'Add By Sindy 2013/5/20
               If frm040104_1.bolIsEMPFlow = True Then
                  frm090202_4.QueryData
               End If
               '2013/5/20 End
               frm040104_1.Show
               frm040104_1.ReQuery
            Else
'               'Add By Cheng 2002/04/30
'               '若有未發文資料顯示警告
'               PUB_GetCPunIssueDatas "" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)
'               frm040104_1.Show
'               frm040104_1.Clear
               '若有未發文資料顯示警告
               If PUB_GetCPunIssueDatas("" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)) Then
                  'Add By Sindy 2013/5/20
                  If frm040104_1.bolIsEMPFlow = True Then
                     frm090202_4.QueryData
                  End If
                  '2013/5/20 End
                  frm040104_1.Show
                  frm040104_1.Command1_Click
               Else
                  'Add By Sindy 2013/5/20
                  If frm040104_1.bolIsEMPFlow = True Then
                     Unload frm040104_1
                     frm090202_4.Show
                     frm090202_4.QueryData
                  Else
                  '2013/5/20 End
                     frm040104_1.Show
                     ' 90.07.11 modify by louis (回第一個畫面清除)
                     frm040104_1.Clear
                  End If
               End If
            End If
            Unload Me
         End If
         
      Case 1
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            frm040104_1.Show
         End If
         Unload Me
         
      Case 2
         'Add By Sindy 2013/5/20
         If frm040104_1.bolIsEMPFlow = True Then
            Unload frm040104_1
            frm090202_4.Show
            frm090202_4.QueryData
         Else
         '2013/5/20 End
            Unload frm040104_1
         End If
         Unload Me

      Case 4
         Me.Hide
         Set frm06010303_1.oParent = Me 'Add by Morgan 2011/10/5
         frm06010303_1.LoadMe strReceiveNo, pa(1), pa(2), pa(3), pa(4), 45
         frm06010303_1.Text41 = "N"
         frm06010303_1.Caption = "內專發文-變更事項"
         m_blnClkChgEvnBtn = True
         m_bolChanged = True 'Add by Morgan 2004/9/3
   End Select
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
End Sub

Private Function FormSave() As Boolean
Dim i As Integer, varTmp As Variant, iMax As Long
Dim strTmp(0 To 5) As String, strTmp1(0 To 5) As String, strTxt(1 To 20) As String
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
Dim ii As Integer
'Add by Morgan 2005/10/17 更新國外案期限用
Dim strCP06 As String, strCP07 As String
Dim strMsg As String
Dim stCP118 As String, stCP152 As String 'Added by Lydia 2020/04/07 電子送件和自動扣款日
Dim stNP23 As String 'Added by Lydia 2025/10/29

On Error GoTo ErrHnd
   'Add By Cheng 2002/11/06
   FormSave = True
cnnConnection.BeginTrans

   'Added by Morgan 2013/6/7 自 lstNameAgent_Validate 移來,否則若觸發 Form_Activate 事件會跑 ReadPatent 導致 cp(110) 被清除
   cp(110) = ""
   If lstNameAgent.Visible = True Then
      For ii = 0 To lstNameAgent.ListCount - 1
         If lstNameAgent.Selected(ii) = True Then
            'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
            'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
            'Modified by Morgan 2021/12/14f Forms2.0 改用模組
            'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
            cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         End If
      Next
      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   End If
   'end 2013/6/7

   ' 90.07.18 modify by louis
   If IsEmptyText(Text11) = False Then
      Text11 = Text11 & String(9 - Len(Text11), "0")
   End If
    'Add By Cheng 2002/11/08
    '補逗號
    If "" & pa(72) = "" Then
        If Val(Me.Text7(0).Text) > 1 Then
            For ii = 1 To Val(Me.Text7(0).Text) - 1
                pa(72) = pa(72) & ","
            Next ii
            If "" & pa(73) = "" Then pa(73) = pa(72)
            If "" & pa(74) = "" Then pa(74) = pa(72)
        End If
    End If
          
   ' 90.07.25 modify by louis
   ' 計算逗號的總數(幾格)
   nDot = 0
   For nPos = 1 To Len(pa(72))
      If Mid(pa(72), nPos, 1) = "," Then nDot = nDot + 1
   Next nPos
   
   aryAll = Split(strCaseFee(2), ",")
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
   ' 填入新資料
   For nPos = nPosBegin To nPosEnd
      aryCurr(nPos) = DBDATE(Text9)
   Next nPos
   ' 讀取新資料
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
      'Modified by Morgan 2012/10/2 只有第一年加倍
      'If Text6 = "Y" Then
      If Text6 = "Y" And nPos = nPosBegin Then
         aryCurr(nPos) = "Y"
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
   strPA74 = strFLD
   strTxt(1) = "UPDATE PATENT SET PA05=" & CNULL(ChgSQL(Text15(0))) & ",PA06=" & CNULL(ChgSQL(Text15(1))) & _
      ",PA07=" & CNULL(ChgSQL(Text15(2))) & ",PA76=" & CNULL(Text11) & "," & _
      "PA72=" & CNULL(strPA72) & ",PA73=" & CNULL(strPA73) & "," & _
      "PA74=" & CNULL(strPA74) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
    cnnConnection.Execute strTxt(1)
   
   If Combo2 <> "" Then
      'Modify by Morgan 2008/2/22
      'cp(44) = ChangeCustomerL(Combo2)
      intI = InStr(Combo2, "-")
      If intI > 0 Then
         cp(44) = Left(Combo2, intI - 1)
         cp(116) = Mid(Combo2, intI + 1)
      Else
         cp(44) = Combo2
         cp(116) = ""
      End If
      cp(44) = ChangeCustomerL(cp(44))
      'end 2008/2/22
      
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.GetCaseThatCode(cp) Then cp(45) = ""
      If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
   Else
      cp(44) = ""
      cp(116) = ""
      cp(45) = ""
   End If
   
   'Added by Morgan 2011/11/28
   cp(53) = Text7(0)
   cp(54) = Text7(1)
   'end 2011/11/28
   
   'Added by Lydia 2020/04/07 電子送件和自動扣款日
   If Frame1.Visible = True Then
        stCP118 = txtCP118
        stCP152 = ""
        If txtCP118 = "Y" Then
           If txtPayToday <> "" Then
              stCP118 = "A"
              If txtPayToday = "Y" Then
                 stCP152 = CompWorkDay(2, DBDATE(Text9))
              Else
                 stCP152 = CompWorkDay(3, DBDATE(Text9))
              End If
           End If
        End If
   End If
   'end 2020/04/07
   
   'Modify by Morgan 2005/7/15 加 cp110
   'Modify by Morgan 2008/2/22 +cp116
   'Modify by Morgan 2011/11/28 +CP53,CP54
   'Modified by Morgan 2013/3/22 -CP14
   'Modified by Lydia 2020/04/07+CP118,CP152 電子送件和自動扣款日
   'Modified by Lydia 2021/05/25 +CP113工作時數
   'Modified by Lydia 2023/06/20 +CP14
   strTxt(2) = "UPDATE CASEPROGRESS SET CP27=" & CNULL(TransDate(Text9, 2)) & ",CP14=" & CNULL(cp(14)) & _
      ",CP44=" & CNULL(cp(44)) & ",CP116=" & CNULL(cp(116)) & _
      ",CP45=" & CNULL(cp(45)) & ",cp64=" & CNULL(ChgSQL(Text12)) & _
      ",CP22=" & CNULL(Text8(1)) & ",cp81=" & CNULL(cp(81)) & ", cp84=" & Format(m_lngFinalFee) & _
      ",cp110=" & CNULL(cp(110)) & ",cp53='" & cp(53) & "',cp54='" & cp(54) & "' " & _
      ",cp118=" & CNULL(stCP118) & ", cp152=" & CNULL(stCP152) & " ,cp113=" & CNULL(txtCP113, True) & _
      " WHERE CP09='" & strReceiveNo & "'"
    cnnConnection.Execute strTxt(2)

   'Add by Morgan 2004/7/22
   '設定客戶減免身分
   If pa(9) = "000" Then
      For ii = 1 To 5
         If txtAD(ii).Enabled = True Then
            '身分有變更才要做
            If txtAD(ii).Tag <> txtAD(ii).Text Then
               '不可減免
               If txtAD(ii).Text = "N" Then
                  strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "N")
               '自然人
               'Modify by Morgan 2004/7/29
               '學校也不用證明
               'ElseIf txtAD(ii).Text = "1" Then
               ElseIf (txtAD(ii).Text = "1" Or txtAD(ii).Text = "2") Then
                  strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "Y", txtAD(ii).Text)
               '公司
               Else
                  '原來沒有減免資料或不可減免
                  If txtAD(ii).Tag = "" Or txtAD(ii).Tag = "N" Then
                     strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "Y", txtAD(ii).Text, pa(1), pa(2), pa(3), pa(4))
                  '修改減免身分別
                  Else
                     strSql = PUB_GetADSQL(pa(25 + ii), pa(9), "Y", txtAD(ii).Text)
                  End If
               End If
               cnnConnection.Execute strSql
            End If
         End If
      Next
   End If
   
   
   strTmp1(0) = strReceiveNo
   For i = 1 To 4
      strTmp1(i) = pa(i)
   Next
   
   m_strPA14 = ""
   'Modify by Morgan 2004/6/24   '台灣無公告日的用發文日預估期限
   If m_bolNew Then
      'Modified by Morgan 2014/11/20 +系統別參數
      PUB_Get605NP pa(1), Text9.Text, Text7(1).Text, strTmp, Val(txtCP71.Text)
      m_strPA14 = strTmp(3) 'Add by Morgan 2006/10/5 申請書要用
   Else
      If GetMoneyDate(Val(pa(8)) + 10, pa(9), strTmp1, strTmp(1), strTmp(2), m_EndDate) Then   '抓專用期起止日
      End If
      If GetMoneyDate(pa(8), pa(9), strTmp1, strTmp(1), strTmp(2), strTmp(3)) = True Then
         varTmp = Split(strTmp(2), ",")
'Modify by Morgan 2007/1/11 不再算到月底
'         '先加一天
'         If strTmp(3) <> "" Then strTmp(3) = CompDate(2, 1, TransDate(strTmp(3), 2))
         '加年度
         strTmp(1) = CompDate(0, Val(varTmp(Val(Text7(1).Text) - 1)), TransDate(strTmp(3), 2))
'         '減一天
'         If strTmp(1) <> "" Then strTmp(1) = CompDate(2, -1, TransDate(strTmp(1), 2))
'end 2007/1/11
         'Added by Lydia 2025/10/29
         If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
            strTmp(2) = PUB_GetPOurDeadline(strTmp(1), pa(9))
         Else
         'end 2025/10/29
            If pa(9) = 台灣國家代號 Then
               'Added by Morgan 2014/10/28
               If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                  strTmp(2) = PUB_GetOurDeadline(strTmp(1))
               Else
               'end 2014/10/28
                  strTmp(2) = CompDate(2, -2, strTmp(1))
               End If 'Added by Morgan 2014/10/28
            Else
               'Modify by Morgan 2010/1/7 FMP案所限改法限-10天
               'Modified by Morgan 2018/10/3 非FMP也改10天
               'If m_bolFMP Then
                  strTmp(2) = CompDate(2, -10, strTmp(1))
               'Else
               '   strTmp(2) = CompDate(1, -1, strTmp(1))
               '   strTmp(2) = CompDate(2, -5, strTmp(2))
               'End If
               'end 2018/10/3
            End If
         End If 'Added by Lydia 2025/10/29
      End If
      
   End If
   
'Modify by Morgan 2010/4/23 改呼叫公用函數
   
'   'Add by Morgan 2005/10/17
'   '若有國外案且為未發文無期限之新案時則以國內的預估公告日-10天更新國外新案的本所期限.
'   'Modify by Morgan 2006/4/14 改抓收文號
'   'strSQL = "SELECT CM01,CM02,CM03,CM04 FROM CASEMAP,PATENT WHERE " & ChgCaseMap(pA(1) & pA(2) & pA(3) & pA(4), 0, 1) & " AND PA01(+)=CM01 AND PA02(+)=CM02 AND PA03(+)=CM03 AND PA04(+)=CM04 AND PA57 IS NULL"
'   'Modify by Morgan 2006/6/1 加判斷未收文主張國際優先權
'   strSql = "SELECT CM01,CM02,CM03,CM04,CP06,CP07,CP09" & _
'      " FROM CASEMAP,PATENT,CASEPROGRESS" & _
'      " WHERE " & ChgCaseMap(pa(1) & pa(2) & pa(3) & pa(4), 0, 1) & " AND CM10='0'" & _
'      " AND PA01(+)=CM01 AND PA02(+)=CM02 AND PA03(+)=CM03 AND PA04(+)=CM04 AND PA57 IS NULL" & _
'      " AND CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP27 IS NULL AND CP31='Y' AND CP57 IS NULL" & _
'      " AND NOT EXISTS(SELECT * FROM CASEPROGRESS WHERE CP01=CM01 AND CP02=CM02 AND CP03=CM03 AND CP04=CM04 AND CP10='106' AND CP57 IS NULL)"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
'   If intI = 1 Then
'      'Modify by Morgan 2006/4/18
'      '法定期限=預估公告日
'      If pa(9) = "000" Then
'         strCP07 = strTmp(3)
'      Else
'         '大陸預估公告日=發文日+2個月
'         strCP07 = CompDate(1, 2, Text9)
'      End If
'
'      '本所期限=法定期限-10天
'      strCP06 = PUB_GetWorkDay1(CompDate(2, -10, strCP07), True)
'      If strCP06 < strSrvDate(1) Then
'         strCP06 = strSrvDate(1)
'      End If
'      With RsTemp
'         Do While Not .EOF
'            strExc(1) = "" & .Fields("CM01")
'            strExc(2) = "" & .Fields("CM02")
'            strExc(3) = "" & .Fields("CM03")
'            strExc(4) = "" & .Fields("CM04")
'            'Modify by Morgan 2006/4/17 大陸案也要
'            'If strExc(1) = "CFP" Then
'               'Modify by Morgan 2006/4/14 加判斷原來沒期限或期限較晚才更新
'               'strSQL = "Update caseprogress set CP06=" & strCP06 & ",CP07=" & strCP07 & _
'                     " WHERE CP06 IS NULL AND CP01='" & strExc(1) & "' AND CP02='" & strExc(2) & "' AND CP03='" & strExc(3) & "' AND CP04='" & strExc(4) & "'" & _
'                     " AND CP27 IS NULL AND CP31='Y' AND CP57 IS NULL"
'               If IsNull(.Fields("CP06")) Or strCP06 < "" & .Fields("CP06") Then
'                  strSql = "Update caseprogress set CP06=" & strCP06 & ",CP07=" & strCP07 & _
'                     " WHERE CP09='" & .Fields("CP09") & "'"
'                  cnnConnection.Execute strSql
'               End If
'            'End If
'            .MoveNext
'         Loop
'      End With
'   End If
'   '2005/10/17 end
   If pa(9) = "000" Then
      strCP07 = strTmp(3)
   Else
      '大陸預估公告日=發文日+2個月
      strCP07 = CompDate(1, 2, Text9)
   End If
   
   PUB_UpdCP07byPA14 pa, strCP07, , , True 'Modified by Morgan 2015/9/16 +更新期限要EMail給工程師
   
'end 2010/4/23
   
   m_bolNonTwCase605Alert = False 'Added by Morgan 2017/1/13
   
   'Modify by Morgan 2005/7/19 加判斷已繳年度<>專用年度才要新增年費期限
   'If m_EndDate = "" Or Val(strTmp(1)) < Val(m_EndDate) Then
   If strCaseFee(2) <> strPA72 And (m_EndDate = "" Or Val(strTmp(1)) < Val(m_EndDate)) Then
      iMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了 iMax = objPublicData.GetNextProgressNo
        '重抓智權人員
        '若本所期限非工作天則抓最近的工作天
        'Modified by Lydia 2025/10/29
       'strTxt(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22) " & _
         "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
         "','" & pa(4) & "','" & strCP13New & "'," & 年費 & "," & CNULL(PUB_GetWorkDay1(strTmp(2), True)) & "," & CNULL(strTmp(1)) & _
         "," & iMax & ")"
        stNP23 = "NULL"
        If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
           strExc(8) = PUB_GetPOurDeadline(strTmp(1), pa(9), stNP23, pa(1), 年費)
        Else
           strExc(8) = PUB_GetWorkDay1(strTmp(2), True)
        End If
       strTxt(3) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP10,NP07,NP08,NP09,NP22,NP23) " & _
         "VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & _
         "','" & pa(4) & "','" & strCP13New & "'," & 年費 & "," & CNULL(strExc(8)) & "," & CNULL(strTmp(1)) & _
         "," & iMax & "," & stNP23 & ")"
       'end 2025/10/29
       cnnConnection.Execute strTxt(3)
       
      'Added by Morgan 2017/1/13
      '非臺灣案若下一程序年費期限早於下一次催年費的期限區間時提醒程序
      If Left(Pub_StrUserSt03, 1) <> "F" Then
          'Added by Lydia 2022/02/23 和碩案提早為法定期限前6個月
          'Modified by Lydia 2022/08/12 傳入變數
          'If Pub_Getfrm040303Except(pa(26)) = True Then
          '     Get605InformPeriod4NonTwCase CompDate(2, 1, strSrvDate(1)), m_bolFMP, strExc(1), , "X70017000", True
          'Modified by Lydia 2025/07/25 改模組
          'If Pub_Getfrm040303Except(pa(26), , strExc(3)) = True Then
          If cntFrm040303New = "Y" Then
             If Pub_Getfrm040303ExceptNew(pa(1), "", "", pa(26), , , strExc(3)) = True Then
                strExc(0) = "Y"
             Else
                strExc(0) = ""
             End If
          Else
             If Pub_Getfrm040303Except(pa(26), , strExc(3)) = True Then
                strExc(0) = "Y"
             Else
                strExc(0) = ""
             End If
          End If
          If strExc(0) = "Y" Then
          'end 2025/07/25
               Get605InformPeriod4NonTwCase CompDate(2, 1, strSrvDate(1)), m_bolFMP, strExc(1), , strExc(3), True
          'end 2022/08/12
               If Val(strTmp(1)) < Val(strExc(1)) Then
                  m_bolNonTwCase605Alert = True
               End If
          Else
          'end 2022/02/23
             Get605InformPeriod4NonTwCase CompDate(2, 1, strSrvDate(1)), m_bolFMP, strExc(1), , pa(75), True
            'FMP抓法限
            If m_bolFMP Then
               If Val(strTmp(1)) < Val(strExc(1)) Then
                  m_bolNonTwCase605Alert = True
               End If
            '非FMP抓所限
            Else
               If Val(PUB_GetWorkDay1(strTmp(2), True)) < Val(strExc(1)) Then
                  m_bolNonTwCase605Alert = True
               End If
            End If
         End If 'Added by Lydia 2022/02/23
      End If
      'end 2017/1/13
   End If
   iMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了 iMax = objPublicData.GetNextProgressNo
   
   
'Modify by Morgan 2011/8/19 改都要管制收達且管制比照其他案件性質模式--玲玲
'   If m_bolBeCalled Then 'Add by Morgan 2010/12/17 自動發文設定3天的收達管制(原來收費表沒設管制天數)
'      i = 3
'      strExc(0) = "SELECT CF23 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & cp(10) & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If Not IsNull(RsTemp.Fields(0)) Or RsTemp.Fields(0) <> 0 Then
'            i = 4
'           '若本所期限非工作天則抓最近的工作天
'            strTxt(i) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'               "NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & _
'               "','" & pa(3) & "','" & pa(4) & "'," & 收達 & "," & _
'               PUB_GetWorkDay1(CompDate(2, RsTemp.Fields(0), TransDate(Text9, 2)), True) & "," & _
'               CompDate(2, RsTemp.Fields(0), TransDate(Text9, 2)) & ",'" & _
'               strUserNum & "'," & iMax & ")"
'           cnnConnection.Execute strTxt(i)
'         End If
'      End If
'   End If
   If pa(9) <> "000" Then PUB_SetArriveDate strReceiveNo
'end 2011/8/19
   
   'Add by Morgan 2004/6/21 延緩公告發文
   If m_str412CP09 <> "" Then
      'Added by Morgan 2022/12/27
      If strSrvDate(1) >= "20230101" Then
         strSql = "UPDATE CASEPROGRESS SET CP27=" & DBDATE(Text9) & ",CP64='" & m_str412AddCP64 & "'||CP64,CP110=" & CNULL(cp(110)) & ",CP118='" & stCP118 & "',CP130='" & m_CP130 & "' WHERE CP09='" & m_str412CP09 & "'"
      Else
      'end 2022/12/27
         'Modify by Morgan 2005/9/6 加cp110
         strSql = "UPDATE CASEPROGRESS SET CP27=" & DBDATE(Text9) & ",CP110=" & CNULL(cp(110)) & " WHERE CP09='" & m_str412CP09 & "'"
      End If
      cnnConnection.Execute strSql
   End If
   
   'Added by Morgan 2013/6/25
   If m_str414CP09 <> "" Then
      strSql = "UPDATE CASEPROGRESS SET CP27=" & DBDATE(Text9) & ",CP110=" & CNULL(cp(110)) & " WHERE CP09='" & m_str414CP09 & "'"
      cnnConnection.Execute strSql
   End If
   'end 2013/6/25
   
   'Add by Morgan 2006/11/3
   '大陸發明案若有香港案時更新批准紀錄請求111期限
   'Modify by Morgan 2009/11/13 FMP案領證發文不必更新香港案期限以使 Reminder 與資料庫的期限一致(真實期限是用發證日起算)
   'If pa(9) = "020" And pa(8) = "1" Then
   If pa(9) = "020" And pa(8) = "1" And Not m_bolFMP Then
      'Modified by Morgan 2014/9/23 香港標準專利(發明)才要--郭
      If ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), strTmp(1), strTmp(2), strTmp(3), strTmp(4), , "1") = True Then
         strExc(9) = CompDate(1, 6, Text9) '法限=發文日+6個月
         'Added by Lydia 2025/10/29
         stNP23 = ""
         If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
            strExc(8) = PUB_GetPOurDeadline(strExc(9), pa(9), stNP23, pa(1), "111")
         Else
         'end 2025/10/29
            strExc(0) = ""
            strExc(1) = pa(1)
            strExc(2) = pa(9)
            strExc(3) = strExc(9)
            GetCtrlDT strExc
            strExc(8) = PUB_GetWorkDay1(strExc(0), True) '所限
         End If 'Added by Lydia 2025/10/29
         If PUB_ChkCPExist(strTmp, "111", 0, strExc(7)) Then
            strSql = "Update CaseProgress Set CP06=" & strExc(8) & ",CP07=" & strExc(9) & " Where CP09='" & strExc(7) & "' and cp27 is null"
         ElseIf PUB_ChkNPExist(strTmp, "111", 0, strExc(6), strExc(5)) Then
            'Modified by Lydia 2025/10/29 +NP23
            strSql = "Update NextProgress Set NP08=" & strExc(8) & ",NP09=" & strExc(9) & ",NP23=" & IIf(stNP23 = "", "NP23", stNP23) & " Where NP22=" & strExc(6) & " and NP01='" & strExc(5) & "'"
         Else
            strSql = "declare intMax number;begin select max(np22)+1 into intMax from nextprogress;IF intMax IS NULL THEN intMax:=1; END IF;"
            'Modified by Lydia 2025/10/29 +NP23
            strSql = strSql & "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22,NP23) " & _
               " Values ('" & cp(9) & "','" & strTmp(1) & "','" & strTmp(2) & "','" & strTmp(3) & "','" & strTmp(4) & "','111'," & strExc(8) & "," & strExc(9) & ",'" & strCP13New & "',intMax," & CNULL(stNP23, True) & "); "
            strSql = strSql & " end;"
         End If
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2006/11/3
   
   'Add by Amy 2014/09/09 for 台灣案電子化
   If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
      If pa(9) = "000" Then
            cnnConnection.Execute "delete LetterProgress where lp01='" & strReceiveNo & "'", intI 'Added by Morgan 2016/2/26 可能會重新發文
            '*沒出客戶通知函
            If Text8(2) = "N" Then
                'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'                'Modify by Amy 2015/02/13 原:判斷同一天 沒有其他有規費的發文
'                strExc(1) = PUB_GetLetterJudge(pa(1), cp(10))
'                '1.    電子送件且規費>0,有收據(此一定有規費)
'                '2.非電子送件且經發文室要計件,有回執
'                If cp(118) = "Y" Then
'                    PUB_AddLetterProgress strReceiveNo, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                Else
'                    If Left(m_CP123s, 1) = "Y" Then
'                        PUB_AddLetterProgress strReceiveNo, 1, False, strExc(1), False, pa(26), cp(10), pa(75), True
'                    End If
'                End If
               
            '*有出客戶通知函
            Else
                'Modified by Morgan 2018/8/1
                'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , , pa(1), pa(2), pa(3), pa(4))
                strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10))
                'Modify by Amy 2015/02/13 修改、整理判斷條件
                  '1.　電子送件有規費的有收據(一定有規費,無規費的不考慮)
                  '2.非電子送件要計件的有回執；不計件的無回執
                'Mark by Amy 2015/03/06 回執改至PUB_UpdateLP19做
'                If cp(118) = "Y" Then
'                    PUB_AddLetterProgress strReceiveNo, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
'                Else
                     If Left(m_CP123s, 1) = "Y" Then
                        PUB_AddLetterProgress strReceiveNo, 1, True, strExc(1), False, pa(26), cp(10), pa(75), True
                     Else
                        PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75), False
                     End If
'                End If
                'end 2015/03/06
            End If
            '*有申請書
            If Text8(0) <> "N" Then
                If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
                     '新增申請書轉檔記錄
                    PUB_AddAppForm strReceiveNo
                End If
            End If
            
      'Added by Morgan 2016/3/30
      ElseIf Left(Pub_StrUserSt03, 1) <> "F" Then
         '指示信電子化
          If Text8(0) <> "N" Then
            'Modified by Morgan 2016/12/13 若已有指示信表示有異常(Ex.P100930)
            'If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
            If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) Then
               Err.Raise 999, , "指示信記錄(Appform)已存在，若為之前有誤操作請通知電腦中心刪除該筆記錄！"
            Else
            'end 2016/12/13
               'Added by Morgan 2016/5/19 主旨要寫到Appfrom自StartLetter移來
               If Text7(0).Text = Text7(1).Text Then
                  strExc(1) = "第 " & Text7(0) & " 年年費"
               Else
                  strExc(1) = "第 " & Text7(0) & " 至 " & Text7(1) & " 年年費"
               End If
               m_Subject = "請代為繳納專利證書費及" & strExc(1) & " Y/R:" & cp(45) & "; O/R:" & Text1 & "-" & Text2 & IIf(Text3 & Text4 = "000", "", "-" & Text3 & "-" & Text4)
               'end 2016/5/19
               'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
               strExc(2) = PUB_GetLetterJudgeNew("2", pa(1), cp(10), pa(9))
               PUB_AddAppForm strReceiveNo, True, strExc(2), m_Subject '不轉檔,自行判發
            End If
          End If
      'end 2016/3/30
         
         'Added by Morgan 2016/5/26
         '客戶通知函
         If 內專全面電子化啟用日 <= Val(strSrvDate(1)) Then
            If Text8(2) <> "N" Then
               'Modified by Morgan 2018/8/1
               'strExc(1) = PUB_GetLetterJudge(pa(1), cp(10), , pa(9), pa(1), pa(2), pa(3), pa(4))
               strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), cp(10), pa(9), , , m_bolFMP)
               PUB_AddLetterProgress strReceiveNo, 0, True, strExc(1), False, pa(26), cp(10), pa(75)
            End If
         End If
         'end 2016/5/26
      End If
    End If
    'end 2014/09/09
      
   'Add by Morgan 2009/3/23
   If pa(9) = 台灣國家代號 Then
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130
      'Add by Amy 2015/02/13 更新收據/回執設定
      'Modify by Amy 2015/03/06 +發文日參數
      PUB_UpdateLP19 cp(1), cp(2), cp(3), cp(4), m_CP09s, m_CP123s, Text9
   End If
   
   'Add by Morgan 2009/8/13
   If txtChkRltDate <> "" Then
      'Modified by Lydia 2016/10/13 + pa09
      PUB_UpdateChkResultDate txtChkRltDate, cp, cp(9), cp(10), cp(43), pa(9)
      'Add by Morgan 2009/9/21
      '同時發文延緩公告
      'Modified by Morgan 2019/2/13 +判斷有收文延緩公告(chk412.Visible = True)
      If chk412.Value = 1 And chk412.Visible = True Then
         PUB_SetChkResultDate pa(1), pa(9), "412", Text9, txtChkRltDate1, cp, pa(8)
         If txtChkRltDate1 <> "" Then
            PUB_UpdateChkResultDate txtChkRltDate1, cp, m_str412CP09, "412"
         End If
      End If
   End If
   
   'Added by Morgan 2020/3/6
   '非台灣案領證要管制最終提申--玲玲
   If pa(9) <> "000" Then
      strExc(1) = DBDATE(cp(7))
      strExc(2) = PUB_GetWorkDay1(strExc(1), True)
      strSql = " insert into nextprogress a (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
         " values('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','996'" & _
         "," & strExc(2) & "," & strExc(1) & ",'" & strUserNum & "',GETNP22)"
      cnnConnection.Execute strSql, intI
   End If
   'end 2020/3/6
      
   cnnConnection.CommitTrans
   If strMsg <> "" Then MsgBox strMsg 'Add by Morgan 2010/4/23
   Exit Function
ErrHnd:
    cnnConnection.RollbackTrans
   FormSave = False
   
   MsgBox Err.Description, vbCritical 'Added by Morgan 2016/12/13
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

Private Sub Combo2_Click()
   Combo2_Validate False
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
Dim strNo As String, iPos As Integer
   
   If Combo2.Text = "" Then
      If pa(9) <> 台灣國家代號 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Cancel = True
         Exit Sub
      End If
      
   ElseIf Not ChgType(12) Then
      Cancel = True
      
   Else
      strNo = Combo2.Text
      
      'Add by Morgan 2008/2/22 加聯絡人判斷
      iPos = InStr(strNo, "-")
      If iPos > 0 Then
         strNo = Left(strNo, iPos - 1)
      End If
      'end 2008/2/22
      
      'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
      If PUB_CheckStatus(strNo) = False Then Cancel = True
   End If
End Sub

Private Sub Form_Activate()
   If m_bolBeCalled Then Exit Sub 'Added by Morgan 2016/7/1 當Form被呼叫且有彈出訊息對話框時,若同時觸發Form_Activate 跑SetFocus指令會發生執行階段錯誤
   
   'Add By Cheng 2003/10/06
   '若有按下變更事項按鈕, 則重新讀取資料
   If m_blnClkChgEvnBtn = True Then
       ReadPatent
       'Add by Morgan 2004/7/22
       m_bolActive = False
       Label2(0) = strReceiveNo
       m_blnClkChgEvnBtn = False
   End If
    
   'Add by Morgan 2004/7/22
   '若沒有客戶減免身分需輸入則游標預設在繳年費起始年欄
   If m_bolActive = True Then Exit Sub
   m_bolActive = True
   
   If pa(9) = 大陸國家代號 Then
      Me.Text7(0).Enabled = True
      If Me.Enabled Then 'Added by Morgan 2021/5/4 可能會有強制視窗被開啟 Ex:P-127221
         Text7(0).SetFocus
      End If
   Else
      If pa(9) = "000" Then
         Dim i As Integer
         For i = 1 To 5
            If txtAD(i).Enabled = True And txtAD(i).Text = "" Then
               txtAD(i).SetFocus
               Exit Sub
            End If
         Next
      End If
      If Text7(1).Enabled = True And Text7(1).Visible = True Then
         If Me.Enabled Then 'Added by Morgan 2021/5/4 可能會有強制視窗被開啟 Ex:P-127221
            Text7(1).SetFocus
         End If
      End If
   End If
   
End Sub

Private Sub Form_Initialize()
   'Add by Morgan 2005/7/14
   ReDim pa(1 To TF_PA)
   ReDim cp(TF_CP)
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   
   'Add by Morgan 2010/12/15
   If m_bolBeCalled Then
      Text1 = m_CP01
      Text2 = m_CP02
      Text3 = m_CP03
      Text4 = m_CP04
      strReceiveNo = m_CP09
   Else
   'end 2010/12/15
      'Add By Sindy 2020/3/26
      If UCase(TypeName(m_PrevForm)) = UCase("frm04010310_1") Then
         With m_PrevForm
            Text1 = .Text1
            Text2 = .Text2
            Text3 = .Text3
            Text4 = .Text4
            strReceiveNo = .strReceiveNo
         End With
      Else
      '2020/3/26 END
         With frm040104_1
            Text1 = .Text1
            Text2 = .Text2
            Text3 = .Text3
            Text4 = .Text4
            strReceiveNo = .Tag
         End With
      End If
   End If
    
   ReadPatent
   Text8(2).Text = "N"   'Add by Amy 2014/09/09 由ReadPatent 搬過來國內或大陸不印通知函
   
   cp(110) = "" '要清空,否則若重新發文會殘留前次發文資料,當新案有改出名人而本程序未改選將會造成不一致 Added by Morgan 2012/9/7
   
   'Add by Morgan 2005/7/14
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   Text8(1).Visible = False
   lstNameAgent.Clear
   If pa(9) = "000" Then
      PUB_SetOurAgent lstNameAgent, pa(), cp(110), , True  'Modified by Morgan 2021/12/14 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2005/7/14 END
   
   Label2(0) = strReceiveNo
   m_blnClkChgEvnBtn = False
   
   'Modify by Amy 2014/09/16 年費定稿上線 是否列印通知函開放可輸
   Text8(2) = ""
   Text8(2).Enabled = True
   'end 2014/09/16
   
   'Added by Morgan 2021/1/27 從 Formsave 移來以便共用
   'Mark by Lydia 2023/06/20 改在ReadPatent
   'strCP13New = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
   'strCP12New = GetSalesArea(strCP13New)
   'end 2021/1/27
   
   'Add by Morgan 2009/11/13
   'Modified by Morgan 2021/1/27
   'If Left(cp(12), 1) = "F" And pa(10) <> "000" Then
   'Modified by Lydia 2023/06/20 pa(10) => pa(9)
   If Left(strCP12New, 1) = "F" And pa(9) <> "000" Then
   'end 2021/1/27
      'm_bolFMP = True 'Mark by Lydia 2023/06/20 改在ReadPatent
      'FMP案領證/年費發文不出定稿，收到收據後才要
      Text8(2) = "N"
      Text8(2).Enabled = False
   Else
      'm_bolFMP = False 'Mark by Lydia 2023/06/20 改在ReadPatent
   End If
   
   '2010/4/8 ADD BY SONIA 大陸案B類收文預設不印通知函
   If cp(9) > "B" And pa(9) = "020" Then
      Text8(2) = "N"
   End If
   '2010/4/8 END
  
   'Add by Amy 2014/09/09
   'Modified by Morgan 2016/6/22 非臺灣案電子化
    'If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
    If Left(Pub_StrUserSt03, 1) <> "F" Then
    'end 2016/6/22
        '通知函不可修改
        Text5(1).Enabled = False
    End If
   'end 2014/09/09
   
   'Added by Morgan 2017/1/11
   '專利處人員操作時年費通知人欄位鎖住以避免不小心改到(目前只有外專人員會設定)
   If Left(Pub_StrUserSt03, 1) = "P" Then
      Text11.Locked = True
   End If
   'end 2017/1/11
   
   Frame1.BackColor = &H8000000F 'Added by Lydia 2020/04/07
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2007/3/26
   'Set frm040104_7 = Nothing 'Removed by Morgan 2021/12/14 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub ReadPatent()
Dim Lbl As Object, txt As Object, i As Integer
Dim strTmp1(1 To 5) As String
'Add By Cheng 2003/04/24
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Add by Morgan 2004/7/22
Dim strAD10 As String, strCU15 As String
Dim m_Fee As String         '銷帳服務費 2012/8/1 add by sonia
Dim m_Official As String    '銷帳規費   2012/8/1 add by sonia

   For Each Lbl In Label2
      Lbl = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
         
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      For i = 3 To 7
         If pa(i + 23) <> "" Then ChgType (i)
      Next
      
      Text15(0) = pa(5)
      Text15(1) = pa(6)
      Text15(2) = pa(7)
      
      If pa(76) <> "" Then Text11 = pa(76): ChgType (11)
      
      'Text8(2).Text = "N"   '國內或大陸不印通知函 'Mark by Amy 2014/09/09 搬至FormLoad 避免通知函設空又觸發Active導致通知函會再被設為N
      If pa(9) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(pA(9), strExc(0)) Then Label2(12) = strExc(0)
         If ClsPDGetNation(pa(9), strExc(0)) Then Label2(12) = strExc(0)
      End If
      
      m_CurrFee = pa(72)
   End If
   
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(13) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(13), strExc(0)) Then Label2(1) = strExc(0)
         If ClsPDGetStaff(cp(13), strExc(0)) Then Label2(1) = strExc(0)
      End If
      Label2(2) = cp(6)
      Label2(13) = cp(7)
      If cp(27) = "" Then
         Text9 = strSrvDate(2)
      Else
         Text9 = cp(27)
      End If
      'Added by Lydia 2023/06/20
      '從Form_Load移過來
      strCP13New = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
      strCP12New = GetSalesArea(strCP13New)
      If Left(strCP12New, 1) = "F" And pa(9) <> "000" Then
         m_bolFMP = True
      Else
         m_bolFMP = False
      End If
      If m_bolFMP = True Then
         m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4))
         '寰華案:承辦人為外專程序時,改為操作人員
         If m_bolFMP2 = True Then
            cp(14) = GetFCPUser(cp(14))
         End If
      End If
      'end 2023/06/20
      If cp(14) <> "" Then
         'Modified by Morgan 2013/3/22
         'Text10 = cp(14): ChgType (10)
         If ClsPDGetStaff(cp(14), strExc(0)) Then Label2(9) = strExc(0)
      End If
      Text12 = cp(64)
      Text8(1) = cp(22)
      'Modify by Morgan 2008/10/16 +若進度檔已有代理人則預設
      'Modified by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設 => cp(9), pa(9), pa(26)
      AddAgent Combo2, cp, , cp(44), cp(116), cp(9), pa(9), pa(26)
      ' 90.07.10 modify by louis (讀取繳年費記錄)
      'strTmp1(0) = strReceiveNo
      For i = 1 To 4
         strTmp1(i) = pa(i)
      Next
      If GetMoneyDate(pa(8), pa(9), strTmp1, strCaseFee(1), strCaseFee(2)) = True Then
      End If
   End If
   
   '2012/8/1 add by sonia 若有銷帳則要扣除銷帳規費
   If Val(cp(77)) > 0 Then
      If GetCP77Detail(cp(9), m_Fee, m_Official) = True Then
         cp(17) = cp(17) - m_Official
      End If
   End If
   '2012/8/1 end

    'Add By Cheng 2003/04/24
    '取得下一程序的法定期限ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4))
    StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " And Np07='" & cp(10) & "' Order By NP09 Desc "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        m_strNP09 = "" & ChangeWStringToTString(rsA("NP09").Value)
    Else
        m_strNP09 = "" & cp(7)
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    'Add By Cheng 2003/09/01
    '若法定期限為假日時, 抓大於法定期限最近的工作天
    m_strNP09_1 = ""
    If m_strNP09 <> "" Then
        m_strNP09_1 = DBDATE(PUB_GetLawDay(DBDATE(m_strNP09)))
    End If
    
   'Add by Morgan 2004/6/24'檢查是否有延緩公告未發文
   chk412.Visible = False
   txtCP71.Visible = False
   m_bolNew = False: m_bol412 = False
    
   'Add by Morgan 2004/6/23
   '台灣可設定申請人年費減免身分
   lblCP81.Visible = False
   lblCP81C.Visible = False

   If pa(9) = 台灣國家代號 Then
   
      lblCP81C.Visible = True
      lblCP81.Visible = True
      lblCP81.Caption = PUB_GetCP81(pa)
      
      'Add by Morgan 2004/7/21
      '減免身分
      For i = 1 To 5
         txtAD(i).Enabled = False
         txtAD(i).Tag = ""
         txtAD(i).Text = ""
         If pa(25 + i) <> "" Then
            txtAD(i).Text = PUB_GetAD03(pa(25 + i), pa(9), strAD10, strCU15)
            txtAD(i).Tag = txtAD(i).Text
            '個人只可設定自然人(1)
            If strCU15 = "0" Then
               txtAD(i).Text = "1"
            'Added by Morgan 2014/7/15 學校也預設--玲玲
            ElseIf strCU15 = "2" Then
               txtAD(i).Text = "2"
            'end 2014/7/15
            '公司
            Else
               If txtAD(i).Text = "Y" Then
                  txtAD(i).Text = strAD10
                  txtAD(i).Tag = txtAD(i).Text
               End If
               txtAD(i).Enabled = True
            End If
         End If
      Next

      'Add by Morgan 2004/6/24'檢查是否有延緩公告未發文
      If Val(pa(14)) = 0 Or Val(pa(14)) >= 930701 Then
         m_bolNew = True
         m_bol412 = PUB_Get412Data(pa, m_str412CP09, m_str412CP71)
         If m_bol412 = True Then
            chk412.Enabled = False
            chk412.Visible = True
            chk412.Value = 1 'Added by Morgan 2019/2/13
            Label9.Visible = True
            txtCP71.Visible = True
         End If
      End If
      
      'Add by Morgan 2004/9/3
      txtFee.Visible = True
      lblFee.Visible = True
   End If

   'Add by Morgan 2009/8/5
   If Text9 <> "" Then
      'Add by Morgan 2009/9/21
      '台灣領證催審期限=預定公告日+30天
      If pa(9) = "000" Then
         'Modified by Morgan 2014/11/20 +系統別參數
         PUB_Get605NP pa(1), Text9.Text, Text7(0).Text, strExc, txtCP71
         txtChkRltDate = CompDate(2, 30, strExc(3))
         lblCaseFee.Enabled = False
      Else
         PUB_SetChkResultDate pa(1), pa(9), cp(10), Text9, txtChkRltDate, cp, pa(8)
      End If
      Text9.Tag = Text9
   End If
   
   'Modify By Sindy 2020/4/15 + if
   If pa(9) = 台灣國家代號 Then
   '2020/4/15 END
      'Add By Sindy 2020/4/13
      cp(81) = "" '設定案件是否可減免
      If PUB_GetCaseDiscStat(pa(1) & pa(2) & pa(3) & pa(4), m_DiscType) = "Y" Then
         cp(81) = "Y"
      Else
         cp(81) = "N"
      End If
      '2020/4/13 END
   End If
   
   'Add by Morgan 2011/8/17
   Text7(0) = cp(53)
   Text7(1) = cp(54)
   'end 2011/8/17
   
   'Added by Lydia 2020/04/07 是否電子送件
   txtCP118 = ""
   Frame1.Visible = False
   If pa(9) = 台灣國家代號 And m_bolBeCalled = False Then '排除整批發文
        If cp(118) <> "" Then txtCP118 = "Y"
        txtPayToday = "" '自動扣款
        Frame1.Visible = True
        If txtCP118 = "Y" Then
            If Val(ServerTime) <= 153000 Then '自動扣款(A)若超過3點半發文則須人工輸入是否當日扣款
               txtPayToday = "Y"
            End If
        End If
   End If
   'end 2020/04/07
   
    'Added by Lydia 2021/05/25
    txtCP113 = ""
    If cp(113) <> "" Then txtCP113 = cp(113)
    'end 2021/05/25
End Sub

Private Function ChgType(i As Integer) As Boolean
Dim strTempName As String
   
   ChgType = False
   Select Case i
      Case 0
         'Modify by Morgan 6/24 暫改可大於系統日但不可大於 930709
         'If Not ChkDate(Text9) Or Val(Text9.Text) > Val(strSrvDate(2)) Then
         '2011/12/8 MODIFY BY SONIA 發文日可輸系統日的下一個工作日,並取消930709控管
         'If Not ChkDate(Text9) Or (Val(Text9.Text) > Val(strSrvDate(2)) And Val(Text9.Text) > 930709) Then
         '   MsgBox "發文日期不正確或發文日大於系統日，請重新輸入 !", vbCritical
         If Not ChkDate(Text9) Or DBDATE(Val(Text9.Text)) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
            MsgBox "發文日期不正確或發文日大於系統日下一個工作日，請重新輸入 !", vbCritical
         '2011/12/8 END
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
         
'Removed by Morgan 2013/3/22 已經沒有再用
'      Case 10
'         'edit by nickc 2007/02/02 不用 dll 了
'         'If objPublicData.GetStaff(Text10, strTempName) Then
'         If ClsPDGetStaff(Text10, strTempName) Then
'            Label2(9) = strTempName
'            ChgType = True
'         End If

      Case 11
         If Text11 <> "" Then
            'edit by nickc 2007/02/05 不用 dll 了
            'If objLawDll.LawGetName(Text11, strTempName) Then
            If ClsLawLawGetName(Text11, strTempName) Then
               Label2(10) = strTempName
               ChgType = True
            End If
         Else
            ChgType = True
         End If
      Case 12 '代理人
         strExc(1) = Combo2.Text
         'Add by Morgan 2008/2/22 加判斷是否為聯絡人
         If InStr(strExc(1), "-") > 0 Then
            If ClsPDGetContact(strExc(1), strTempName) Then
               Combo2 = strExc(1)
               Label2(11) = strTempName
               ChgType = True
            End If
         
         '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
         ElseIf PUB_GetAgentName(pa(1), strExc(1), strTempName) = True Then
            Combo2.Text = strExc(1)
            Label2(11).Caption = strTempName
            ChgType = True
         Else
            Label2(11).Caption = ""
         End If
   End Select
End Function

'Removed by Morgan 2013/3/22 已經沒有再用
'Private Sub Text10_GotFocus()
'  TextInverse Text10
'End Sub
'
'Private Sub Text10_Validate(Cancel As Boolean)
'   If Text10 <> "" Then
'      If Not ChgType(10) Then Cancel = True
'   End If
'End Sub

Private Sub Text11_GotFocus()
  TextInverse Text11
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Not ChgType(11) Then Cancel = True
   'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
   If Cancel = False Then
      If PUB_CheckStatus(Text11.Text) = False Then Cancel = True
   End If
End Sub

Private Sub Text12_GotFocus()
  TextInverse Text12
End Sub

Private Sub Text15_GotFocus(Index As Integer)
    'Modify By Cheng 2002/10/28
    Select Case Index
    Case 0
        Me.Text15(Index).SelStart = 0
        Me.Text15(Index).SelLength = 0
    Case Else
        TextInverse Text15(Index)
    End Select
End Sub

Private Sub Text5_GotFocus(Index As Integer)
  TextInverse Text5(Index)
  'edit by nickc 2007/07/11 切換輸入法改用API
  'Text5(Index).IMEMode = 2
  CloseIme
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
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
    '若發文日及法定期限有值時, 才比較
    If Me.Text9.Text <> "" And m_strNP09 <> "" Then
        'Modify By Cheng 2003/09/01
        '若法定期限為假日則用大於法定期限最近的工作日與發文日比較
'        If DBDATE(Text9) > DBDATE(m_strNP09) And pa(9) = 台灣國家代號 Then
        If DBDATE(Text9) > IIf(DBDATE(m_strNP09) >= DBDATE(m_strNP09_1), DBDATE(m_strNP09), DBDATE(m_strNP09_1)) And pa(9) = 台灣國家代號 Then
            If Text6 <> "Y" Then
               If Not m_bolBeCalled Then 'Add by Morgan 2010/12/15
                  If m_PrevForm Is Nothing Then 'Added by Morgan 2022/3/17
                     MsgBox "發文日大於法定期限則【費用是否雙倍】必須為 Y !", vbCritical
                  End If
               End If
               Text6 = "Y"
            End If
        Else
            If Text6 = "Y" Then
                MsgBox "費用是否雙倍錯誤 !", vbCritical
                Cancel = True
            End If
        End If
    Else
        If Text6 = "Y" Then
            MsgBox "費用是否雙倍錯誤 !", vbCritical
            Cancel = True
        End If
    End If
    If Cancel = True Then TextInverse Me.Text6
End Sub

Private Sub Text7_GotFocus(Index As Integer)
  TextInverse Text7(Index)
End Sub

Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
   'Add by Morgan 2007/10/15 不可輸入空白
   'Modify by Morgan 2009/1/10 改控制不可輸入非數字(輸入"."會造成繳費年度更新有問題,Ex P-88839)
   'If KeyAscii = vbKeySpace Then
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
   'end 2007/10/15
End Sub

Private Sub Text7_LostFocus(Index As Integer)
    'Add By Cheng 2002/11/08
    Dim i As Integer, varTmp As Variant
    Dim bolChk As Boolean
    
    'Add By Cheng 2002/11/08
    If Index = 1 Then
         
    End If
End Sub

Private Sub Text7_Validate(Index As Integer, Cancel As Boolean)
Dim i As Integer, bolChk As Boolean, varTmp As Variant
 
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
               Exit Sub
            '92.7.7 ADD BY SONIA
            Else
               varTmp = Split(strCaseFee(2), ",")
               If Text7(1) > UBound(varTmp) + 1 Then
                  MsgBox "繳費年度大於應繳年度，請查明後再輸入 !", vbCritical
                  Cancel = True
                  Exit Sub
                  
               'Add by Morgan 2011/7/1
               Else
                  If cp(81) = "Y" And pa(8) = "3" And Val(Text7(1)) < 3 And Val(Text7(1)) <> UBound(varTmp) + 1 Then
                     If UBound(varTmp) + 1 < 3 Then
                        strExc(1) = UBound(varTmp) + 1
                     Else
                        strExc(1) = 3
                     End If
                     MsgBox "繳費年度請輸入 " & strExc(1) & " 以上(可減免客戶1~3年免繳年費)!!"
                     Cancel = True
                     Exit Sub
                  End If
               
               End If
            '92.7.7 END
            End If
         Else
            Cancel = True
            Exit Sub
         End If
      End If
   Else
      MsgBox "年度不可空白 !", vbCritical
      TextInverse Text7(Index)
      Cancel = True
   End If
   'Add By Cheng 2002/11/08
   If Cancel Then TextInverse Me.Text7(Index)
End Sub

Private Sub Text8_GotFocus(Index As Integer)
  TextInverse Text8(Index)
  'edit by nickc 2007/07/11 切換輸入法改用API
  'Text8(Index).IMEMode = 2
  CloseIme
End Sub

Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Text9 <> "" Then
      If Not ChgType(0) Then
         Cancel = True
      Else
         'Add by Morgan 2009/8/5
         If Text9.Tag <> Text9 Then
            'Add by Morgan 2009/9/21
            '台灣領證催審期限=預定公告日+30天
            If pa(9) = "000" Then
               'Modified by Morgan 2014/11/20 +系統別參數
               PUB_Get605NP pa(1), Text9.Text, Text7(0).Text, strExc, txtCP71
               txtChkRltDate = CompDate(2, 30, strExc(3))
               'Added by Lydia 2020/04/07 當發文日有改時,電子送件案要人工輸入是否當日扣款
               If Frame1.Visible = True And txtCP118 = "Y" Then
                  txtPayToday.Text = ""
               End If
               'end 2020/04/07
            Else
               PUB_SetChkResultDate pa(1), pa(9), cp(10), Text9, txtChkRltDate, cp, pa(8)
            End If
            Text9.Tag = Text9
         End If
      End If
   Else
      MsgBox "發文日不可空白 !", vbCritical
      Cancel = True
   End If
   
End Sub

'Add By Cheng 2002/03/08
Private Function CheckDataIntegrity() As Boolean
Dim Cancel As Boolean
   'add by nickc 2008/05/01
   If IsDebt(pa(9), cp(9)) Then
        MsgBox "未收款且無 預定收款日 請轉告智權同仁！！", vbOKOnly, "警告！禁止發文！"
        GoTo IntegrityOrNot
   End If
   
   Cancel = False
   
   '檢查代理人欄位
   Combo2_Validate Cancel
   If Cancel = True Then
      Me.Combo2.SetFocus
      GoTo IntegrityOrNot
   End If
   
   CheckDataIntegrity = True
   Exit Function
   
IntegrityOrNot:
   CheckDataIntegrity = False
End Function

'Add By Cheng 2002/05/22
'Modified by Morgan 2022/3/17
'Private Function TxtValidate() As Boolean
Public Function TxtValidate() As Boolean
'end 2022/3/17
   Dim objTxt As Object
   Dim ii As Integer
   Dim Cancel As Boolean
   'Add By Cheng 2002/11/08
   Dim bolChk As Boolean, varTmp As Variant
   Dim i As Integer

   TxtValidate = False
   
   'Added by Morgan 2021/12/14 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/14
   
   If Me.Text6.Enabled = True Then
      Cancel = False
      Text6_Validate Cancel
      If Cancel = True Then
         Me.Text6.SetFocus
         Text6_GotFocus
         Exit Function
      End If
   End If
   
If m_PrevForm Is Nothing Then 'Added by Morgan 2022/3/21 以下出申請書時先不檢查，發文才要

'Removed by Morgan 2013/3/22 已經沒有再用
'   If Me.Text10.Enabled = True Then
'      Cancel = False
'      Text10_Validate Cancel
'      If Cancel = True Then
'         Me.Text10.SetFocus
'         Text10_GotFocus
'         Exit Function
'      End If
'   End If
    
   If Me.Text11.Enabled = True Then
      Cancel = False
      Text11_Validate Cancel
      If Cancel = True Then
         Me.Text11.SetFocus
         Text11_GotFocus
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
   
    'Add by Morgan 2005/7/14
    If lstNameAgent.Visible = True Then
       Cancel = False
       lstNameAgent_Validate Cancel
       If Cancel = True Then
          lstNameAgent.SetFocus
          Exit Function
       End If
    End If
    '2005/7/14 END
   
    'Added by Lydia 2020/04/07
    If Frame1.Visible = True And Me.txtCP118 = "Y" Then
       If Me.Text8(1) = "N" Then
           MsgBox "電子送件不可不出名！", vbCritical
           Exit Function
       End If
      
      If Me.txtPayToday = "" Then
        MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
        Me.txtPayToday.SetFocus
        Exit Function
      End If
    End If
    'end 2020/04/07
   
         'Add by Morgan 2004/7/22
         If pa(9) = "000" Then
            m_DiscType = ""
            For i = 1 To 5
               m_DiscType = m_DiscType & txtAD(i).Text
               If txtAD(i).Enabled = True Then
                  If txtAD(i).Text = "" Then
                     MsgBox "申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】減免身分不可空白", vbInformation
                     txtAD(i).SetFocus
                     txtAD_GotFocus i
                     Exit Function
                  '公司可減免
                  'Modify by Morgan 2004/7/29
                  '學校不需證明
                  'ElseIf (txtAD(i).Text = "2" Or txtAD(i).Text = "3") Then
                  '學校
                  ElseIf (txtAD(i).Text = "2") Then
                     '變更
                     If (txtAD(i).Tag <> "2" And txtAD(i).Tag <> "") Then
                        If MsgBox("確定要變更申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】減免身分為【學校】？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                           txtAD(i).SetFocus
                           txtAD_GotFocus i
                           Exit Function
                        End If
                     End If
                  '公司
                  ElseIf (txtAD(i).Text = "3") Then
                     '新增或變更
                     If (txtAD(i).Tag <> "3") Then
                        If MsgBox("申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】的減免身分將設定為【中小企業】，確定有【證明文件】存放於本卷？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                           txtAD(i).SetFocus
                           txtAD_GotFocus i
                           Exit Function
                        End If
                     End If
                  '不可減免
                  ElseIf (txtAD(i).Text = "N") Then
                     '身分變更
                     If (txtAD(i).Tag <> "N" And txtAD(i).Tag <> "") Then
                        If MsgBox("確定要變更申請人【" & pa(25 + i) & "-" & Label2(i + 2) & "】減免身分為【不可減免】？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                           txtAD(i).SetFocus
                           txtAD_GotFocus i
                           Exit Function
                        End If
                     End If
                  End If
               End If
            Next
            If InStr(m_DiscType, "N") > 0 Then
               cp(81) = "N"
            Else
               cp(81) = "Y"
            End If
         End If
   
   
   'Memo by Morgan 2011/7/1
   '年度檢查要在減免身份設定後
   For Each objTxt In Text7
      objTxt = Trim(objTxt)
      If objTxt.Enabled = True Then
         Cancel = False
         Text7_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Me.Text7(objTxt.Index).SetFocus
            Text7_GotFocus objTxt.Index
            Exit Function
         End If
      End If
   Next
   
   'Add by Morgan 2004/6/21
   If m_bol412 = True Then
      If txtCP71 = "" Then
         MsgBox "必須輸入延緩月份！", vbCritical
         txtCP71.SetFocus
         txtCP71_GotFocus
         Exit Function
      Else
         txtCP71_Validate Cancel
         If Cancel = True Then
            Me.txtCP71.SetFocus
            txtCP71_GotFocus
            Exit Function
         End If
      End If
   End If
   
   'Add by Morgan 2004/9/14
   If Combo2.Enabled = True Then
      Cancel = False
      Combo2_Validate Cancel
      If Cancel = True Then
         Combo2.SetFocus
         Exit Function
      End If
   End If
   
   'Added by Morgan 2013/3/26
   m_UseNewForm = False
   Erase m_AD1516
   
   If DBDATE(cp(7)) >= "20130718" Then
      If pa(9) = "000" And cp(81) = "Y" Then
         For ii = 1 To 5
            If txtAD(ii) = "3" Then
               strExc(0) = "select ad15,ad16 from applicantdiscount where ad01='" & Left(pa(25 + ii) & "00", 8) & "' and ad02='000' and ad15 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  m_AD1516(ii, 1) = "" & RsTemp("ad15")
                  m_AD1516(ii, 2) = "" & RsTemp("ad16")
               Else
                  If MsgBox("申請人【" & pa(25 + ii) & " " & Label2(2 + ii) & " 】之減免身分為中小企業但尚未勾選減免資格，是否確定要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                     Exit For
                  Else
                     Exit Function
                  End If
               End If
            End If
         Next
         If ii > 5 Then
            m_UseNewForm = True
         End If
      End If
   End If
   'end 2013/3/26
   
   'Added by Morgan 2013/6/25
   m_str414CP09 = ""
   If pa(9) = "000" Then
      If PUB_ChkCPExist(pa, "414", 1, m_str414CP09, strExc(1)) = True Then
         intI = MsgBox("系統將同時發文回復原狀？", vbOKCancel + vbDefaultButton2 + vbInformation)
         If intI = vbOK Then
            If strExc(1) = "" Then
               MsgBox "回復原狀尚未分案！", vbExclamation
               Exit Function
            End If
         Else
            Exit Function
         End If
      End If
   End If
   'end 2013/6/25
   
   'Added by Lydia 2021/05/25 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
   If Pub_ChkACS112isNull(pa(1), pa(2), pa(3), pa(4), txtCP113) = True Then
         txtCP113.SetFocus
         txtCP113_GotFocus
         Exit Function
   End If
   'end 2021/05/25
   
End If

   TxtValidate = True
End Function

'Add By Cheng 2002/12/31
'計算相關費用
'Modify By Sindy 2020/3/26
'+ Optional bolShowMsg As Boolean = True
'Private Function ChkPatentYearFee
Public Function ChkPatentYearFee( _
    strYF01 As String, strYF02 As String, strYF03 As String, _
    strYF04 As String, strYF05From As String, strYF05To As String, _
    blnDouble As Boolean, Optional bolShowMsg As Boolean = True) As Boolean
'strYF01  申請國家
'strYF02  專利種類
'strYF03  代理人
'strYF04  案件性質
'strYF05From  起始年度
'strYF05To  終止年度
'blnDouble  規費是否雙倍

Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim ii As Integer
'Add By Cheng 2003/05/07
Dim strOfficalFee1 As String
Dim strOfficalFee2 As String

'Add by Morgan 2004/6/21   繳費年度
Dim iYear As Integer
'Add by Morgan 2004/7/22   被異議控制
Dim lngDDate As Long '是否減免控制日期
Dim bol802 As Boolean   '是否有異議答辯
'Add by Morgan 2005/2/1 新型申請日
Dim stUtiAppDate As String

   bol802 = PUB_ChkCPExist(pa, "802")

    ChkPatentYearFee = False
    m_strOfficalFee = 0
    m_strServiceFee = 0
    m_strPoints = 0
    strOfficalFee1 = 0
    strOfficalFee2 = 0
    'Add by Morgan 2004/7/9
    m_strOfficalFee1 = 0
    m_lngFee1 = 0
    m_lngFee2 = 0
    m_lngDisc = 0
    m_lngDisc1Year = 0 '第一年減免金額 Add by Sindy 2020/4/8
    m_lngSub = 0
    
    ii = 1
    '取得案件性質為領證及繳年費的相關費用
   StrSQLa = "Select * From PatentYearFee Where YF01='" & strYF01 & "' AND YF02='" & strYF02 & "' AND YF03='" & strYF03 & "' AND YF04='" & strYF04 & "' AND YF05>=" & Val(strYF05From) & " AND YF05<=" & Val(strYF05From) & " Order By YF05 "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   While Not rsA.EOF
       '領證的規費不用Double
       strOfficalFee1 = Val(strOfficalFee1) + Val(rsA.Fields("YF07").Value)
       m_strServiceFee = Val(m_strServiceFee) + Val(rsA.Fields("YF06").Value)
       rsA.MoveNext
       'Add By Cheng 2003/01/02
       ii = ii + 1
   Wend
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
    
    ii = 1
    '取得案件性質為年費的相關費用
    StrSQLa = "Select * From PatentYearFee Where YF01='" & strYF01 & "' AND YF02='" & strYF02 & "' AND YF03='" & strYF03 & "' AND YF04='" & 年費 & "' AND YF05>=" & Val(strYF05From) & " AND YF05<=" & Val(strYF05To) & " Order By YF05 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    While Not rsA.EOF
      iYear = Val("" & rsA.Fields("YF05").Value)
      
      'Add by Morgan 2004/6/23
      strOfficalFee2 = Val(strOfficalFee2) + Val(rsA.Fields("YF07").Value)
      'If ii = 1 Then
      If iYear = 1 Then
         m_lngOfficalFee1Year = Val("" & rsA.Fields("YF07").Value) '第一年年費 Add By Sindy 2020/3/27
         m_lngFee1 = Val(strOfficalFee1) + Val(strOfficalFee2)
      Else
         m_lngFee2 = m_lngFee2 + Val(rsA.Fields("YF07").Value)
      End If
      
      '年費減免
      If cp(81) = "Y" Then
          '沒有公告日或公告日+繳費年度>930701的才能減免
          'Modify by Morgan 2004/7/22
          '若有被異議時使用過的年費都不可減免,否則當年度(使用中)的可減免
          'If (Val(pa(14)) = 0 Or Val(pa(14)) + (iYear) * 10000# > 930701) Then
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
          'Modify by Morgan 2005/1/31 只要年費期限930701以後的都要減免 -- 郭
          'If (lngDDate = 0 Or lngDDate > Val(strSrvDate(2))) Then
          If (lngDDate = 0 Or lngDDate > 930701) Then
          'End 2004/7/22
            If iYear >= 1 And iYear <= 3 Then
               'Add By Sindy 2020/4/8
               If iYear = 1 Then
                  m_lngDisc1Year = 800 '第一年減免金額
               End If
               '2020/4/8 END
               strOfficalFee2 = strOfficalFee2 - 800
               m_lngDisc = m_lngDisc + 800
            ElseIf iYear >= 4 And iYear <= 6 Then
               strOfficalFee2 = strOfficalFee2 - 1200
               m_lngDisc = m_lngDisc + 1200
            End If
         End If
         'Add by Morgan 2004/8/13
         '加倍補繳
         If blnDouble = True And ii = 1 Then
            strOfficalFee2 = strOfficalFee2 * 2#
            m_lngDisc = m_lngDisc * 2#
         End If
      Else
         '起始那年年費是否雙倍
         'Modify by Morgan 2004/6/21
         'strOfficalFee2 = Val(strOfficalFee2) + Val(rsA.Fields("YF07").Value) * IIf(blnDouble = True And ii = 1, 2, 1)
         If blnDouble = True And ii = 1 Then strOfficalFee2 = Val(rsA.Fields("YF07").Value) * 2
         'end
      End If
        
        m_strServiceFee = Val(m_strServiceFee) + Val(rsA.Fields("YF06").Value)
        rsA.MoveNext
        'Add By Cheng 2003/01/02
        ii = ii + 1
    Wend
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    m_strPoints = Val(m_strServiceFee) / 1000
    m_strOfficalFee = Val(strOfficalFee1) + Val(strOfficalFee2)
    'Add by  Morgan 2004/6/25
    m_strOfficalFee1 = strOfficalFee1
    'Add by Morgan 2004/7/8
    '台灣新型93.7.1以前申請,93.7.1(含)以後核准的規費可減免1500
   'Modify by Morgan 2011/5/18 日期全部都要轉西元年比較才不用考慮來源格式問題
   If (pa(8) = "2" And pa(9) = "000" And DBDATE(pa(20)) >= DBDATE(930701)) Then
      stUtiAppDate = PUB_UtiAppDate(pa, pa(10))
      If DBDATE(stUtiAppDate) < DBDATE(930701) Then
         m_strOfficalFee = Val(m_strOfficalFee) - 1500
         m_lngSub = 1500
      End If
   End If
'2005/2/1 end
   
   '若不等
   m_lngFinalFee = Val(m_strOfficalFee) + Val(txtFee.Text)
   If "" & cp(17) <> m_lngFinalFee Then
      'Modify By Sindy 2020/3/26
      If bolShowMsg = False Then
         ChkPatentYearFee = True
      Else
      '2020/3/26 END
         If MsgBox("計算出的規費( " & Format(m_lngFinalFee, "#,##0") & " )與目前進度檔的規費( " & Format(cp(17), "#,##0") & " )不同，是否要繼續作業???", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
             ChkPatentYearFee = True
         Else
             ChkPatentYearFee = False
         End If
      End If
   '若相等
   Else
       ChkPatentYearFee = True
   End If
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

'Add by Morgan 2004/6/24
Private Sub txtCP71_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtCP71.IMEMode = 2
   CloseIme
   TextInverse txtCP71
End Sub

'Add by Morgan 2004/6/24
Private Sub txtCP71_Validate(Cancel As Boolean)
   If txtCP71 = "" Then
      Exit Sub
   ElseIf Val(txtCP71) <> Val(m_str412CP71) Then
      MsgBox "延緩月數/日期必須與分案時相同！", vbCritical
      Cancel = True
   'Add by Morgan 2009/9/21
   '台灣領證催審期限=預定公告日+30天
   ElseIf pa(9) = "000" Then
      'Add by Morgan 2009/10/19 若指定日期時判斷是否超過預定公告日+3個月
      If Len(txtCP71) > 1 Then
         'Modified by Morgan 2014/11/20 +系統別參數
         'Modified by Morgan 2016/3/11 105/3/9日起延緩公告最長改6個月(原3個月)
         PUB_Get605NP pa(1), Text9.Text, "1", strExc, "6"
         If Val(DBDATE(txtCP71)) > Val(strExc(3)) Then
            MsgBox "延緩公告日期不可超過預定公告日+6個月(" & TransDate(strExc(3), 1) & ")！"
            Cancel = True
         End If
         'end 2016/3/11
      End If
      'end 2009/10/19
      'Modified by Morgan 2014/11/20 +系統別參數
      PUB_Get605NP pa(1), Text9.Text, Text7(0).Text, strExc, txtCP71
      txtChkRltDate = CompDate(2, 30, strExc(3))
   End If
End Sub

'Add by Morgan 2004/7/22
Private Sub txtAD_GotFocus(Index As Integer)
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtAD(Index).IMEMode = 2
   CloseIme
   TextInverse txtAD(Index)
End Sub

'Add by Morgan 2004/7/22
'只有公司可輸入 2,3,N
Private Sub txtAD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2014/7/15 學校改預設且不可改
   'If Not (KeyAscii = 8 Or KeyAscii = 50 Or KeyAscii = 51 Or KeyAscii = 78) Then
   If Not (KeyAscii = 8 Or KeyAscii = 51 Or KeyAscii = 78) Then
      KeyAscii = 0
   End If
End Sub

'Add by Morgan 2004/8/17 延緩公告
Private Function StartLetter1(ByVal strReceiveNo As String, ByVal ET01 As String, ByVal ET03 As String) As Boolean

   Dim strTxt(1 To 20) As String, strTmp As String, strTmp1 As String
   Dim iAppCnt As Integer
   Dim stAppData(1 To 1, 0 To 3) As String
   Dim ii As Integer, iLen As Integer, i As Integer
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   '1 發文日
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','發文日','" & ChangeTStringToTDateString(Text9.Text) & "')"
      
   '2~4 專利種類
   For i = 1 To 3
      If pa(8) = Format(i) Then
         strTmp = "■ "
      Else
         strTmp = "□ "
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','勾選" & Format(i) & "','" & strTmp & "')"

   Next
   
   '5 延緩期限
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','延緩期限','" & IIf(Len(txtCP71) = 1, txtCP71 & "個月。", "至民國 " & PUB_DBYEAR(txtCP71) - 1911 & " 年 " & Val(PUB_DBMONTH(txtCP71)) & " 月 " & Val(PUB_DBDAY(txtCP71)) & " 日。") & "')"
   
   
   '6 申請人數
   iAppCnt = 1
   For i = 27 To 30
      If pa(i) <> "" Then
         iAppCnt = iAppCnt + 1
      End If
   Next
   
   strTmp = Format(iAppCnt) '申請人數
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','申請人數','" & strTmp & "')"
    
   '3~12 國籍
   For i = 1 To iAppCnt
      Erase stAppData
      Call PUB_GetAppData(pa(25 + i), stAppData, 1)
      strTmp = IIf(Val(stAppData(1, 2)) < 10, "中華民國", stAppData(1, 3))
      strTmp1 = Label2(2 + i) & "　ID：" & stAppData(1, 0)
      
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','申請人" & Format(i) & "的國籍','" & strTmp & "')"
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','申請人" & Format(i) & "的名稱','" & strTmp1 & "')"
   Next
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(ii, strTxt) Then
   If Not ClsLawExecSQL(ii, strTxt) Then
       MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Function

'Add by Amy 2014/09/09 通知函
Private Sub StartLetter2(ByVal ET01 As String, ByVal ET03 As String)
    Dim strTxt(1 To 2) As String, strTmp As String
    Dim ii As Integer
    
    ii = 0
    EndLetter ET01, strReceiveNo, ET03, strUserNum
    If Text7(0).Text = Text7(1).Text Then
        strTmp = "第 " & Text7(0) & " 年"
    Else
        strTmp = "第 " & Text7(0) & " 至 " & Text7(1) & " 年"
    End If
    
    ii = ii + 1
    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
        "','第幾年至幾年費','" & strTmp & "年費')"
          
    If m_bol412 = True Then
        ii = ii + 1
        strTmp = "另，本案同時提出延緩公告至 " & PUB_DBYEAR(m_strPA14) - 1911 & " 年 " & Val(PUB_DBMONTH(m_strPA14)) & " 月 " & Val(PUB_DBDAY(m_strPA14)) & " 日。"

        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','延緩公告','" & strTmp & "')"
       
    End If
    If Not ClsLawExecSQL(ii, strTxt) Then
        MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
    End If
    
End Sub
'end 2014/09/09

Private Sub txtFee_GotFocus()
   TextInverse txtFee
End Sub

Private Sub txtFee_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

'Add by Morgan 2005/7/14
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   cp(110) = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         'Modified by Morgan 2021/12/14f Forms2.0 改用模組
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
   If bolCheck = True Then
      Text8(1) = ""
   Else
      Text8(1) = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

Private Function GetAgentName() As String
   Dim ii As Integer, strName As String
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         strName = strName & "、" & lstNameAgent.List(ii)
      End If
   Next
   If strName <> "" Then
      strName = Mid(strName, 2)
   End If
   GetAgentName = strName
End Function

'Add by Morgan 2009/8/13
Private Sub lblCaseFee_Click()
   frm12040102_2.txtCF(1) = cp(1)
   frm12040102_2.txtCF(2) = pa(9)
   frm12040102_2.txtCF(3) = cp(10)
   frm12040102_2.Show vbModal
   If Val(Text9) > 0 Then
      PUB_SetChkResultDate pa(1), pa(9), cp(10), Text9, txtChkRltDate, cp, pa(8)
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

'Added by Lydia 2020/04/07
Private Sub txtCP118_GotFocus()
   TextInverse txtCP118
End Sub

Private Sub txtCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP118_Change()
    'Added by Lydia 2020/04/08 電子送件要在發文前，先產生申請書；所以發文不用印
    If txtCP118 = "Y" Then
        Text8(0) = "N"
    End If
End Sub

Private Sub txtPayToday_GotFocus()
   TextInverse txtPayToday
End Sub

Private Sub txtPayToday_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub
