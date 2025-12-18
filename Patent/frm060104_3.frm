VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060104_3 
   Appearance      =   0  '平面
   BorderStyle     =   1  '單線固定
   Caption         =   "外專發文"
   ClientHeight    =   6708
   ClientLeft      =   396
   ClientTop       =   972
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6708
   ScaleWidth      =   8952
   Begin VB.TextBox txtDecreaseItemFee 
      Height          =   288
      Left            =   7800
      TabIndex        =   119
      Top             =   5460
      Width           =   1110
   End
   Begin VB.TextBox txtDecreasePageFee 
      Height          =   288
      Left            =   7800
      TabIndex        =   118
      Top             =   5190
      Width           =   1110
   End
   Begin VB.TextBox txtCP167 
      Height          =   280
      Left            =   8505
      TabIndex        =   115
      Top             =   4620
      Width           =   420
   End
   Begin VB.TextBox txtCP168 
      Height          =   280
      Left            =   8505
      TabIndex        =   114
      Top             =   4890
      Width           =   420
   End
   Begin VB.TextBox txtEmail 
      Height          =   270
      Left            =   3540
      MaxLength       =   1
      TabIndex        =   37
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox txtRecDate 
      Height          =   270
      Left            =   7740
      MaxLength       =   1
      TabIndex        =   22
      Top             =   2910
      Width           =   375
   End
   Begin VB.TextBox txtPAID 
      Height          =   270
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   21
      Top             =   2910
      Width           =   435
   End
   Begin VB.Frame fraTrans04 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame3"
      Height          =   660
      Left            =   60
      TabIndex        =   108
      Top             =   5955
      Width           =   8835
      Begin VB.TextBox txtTF37 
         Height          =   270
         Left            =   1170
         TabIndex        =   41
         Top             =   30
         Width           =   6000
      End
      Begin VB.ComboBox Combo6 
         Height          =   300
         Left            =   1170
         TabIndex        =   42
         Text            =   "Combo6"
         Top             =   330
         Width           =   3000
      End
      Begin VB.Label lblTF37 
         Caption         =   "翻譯瑕疵備註:"
         Height          =   165
         Left            =   0
         TabIndex        =   109
         Top             =   75
         Width           =   1245
      End
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   300
      Left            =   5100
      TabIndex        =   20
      Text            =   "cboPrinter"
      Top             =   810
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtPayToday 
      Height          =   264
      Left            =   8055
      MaxLength       =   1
      TabIndex        =   105
      Top             =   4290
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "案件進度(&C)"
      Height          =   255
      Left            =   6090
      TabIndex        =   7
      Top             =   1410
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.TextBox txtCP43 
      Height          =   270
      Left            =   4980
      MaxLength       =   9
      TabIndex        =   6
      Top             =   1410
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtCP138 
      Height          =   280
      Left            =   8505
      TabIndex        =   32
      Top             =   4035
      Width           =   420
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   4815
      Style           =   2  '單純下拉式
      TabIndex        =   36
      Top             =   4275
      Width           =   1095
   End
   Begin VB.TextBox txtCP118 
      Height          =   270
      Left            =   7590
      MaxLength       =   1
      TabIndex        =   10
      Top             =   1530
      Width           =   375
   End
   Begin VB.ComboBox cboAddCP64 
      Height          =   300
      ItemData        =   "frm060104_3.frx":0000
      Left            =   2145
      List            =   "frm060104_3.frx":0002
      Style           =   2  '單純下拉式
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5640
      Width           =   6360
   End
   Begin VB.CheckBox chkCP86 
      Caption         =   "否"
      Height          =   195
      Index           =   1
      Left            =   1950
      TabIndex        =   34
      Top             =   4335
      Width           =   510
   End
   Begin VB.CheckBox chkCP86 
      Caption         =   "是"
      Height          =   195
      Index           =   0
      Left            =   1350
      TabIndex        =   33
      Top             =   4335
      Width           =   555
   End
   Begin VB.TextBox txtCP135 
      Enabled         =   0   'False
      Height          =   280
      Left            =   8505
      TabIndex        =   29
      Top             =   3220
      Width           =   420
   End
   Begin VB.TextBox txtCP136 
      Height          =   280
      Left            =   8505
      TabIndex        =   30
      Top             =   3492
      Width           =   420
   End
   Begin VB.TextBox txtCP137 
      Height          =   280
      Left            =   8505
      TabIndex        =   31
      Top             =   3764
      Width           =   420
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "第三人提實審(&T)"
      Height          =   405
      Index           =   6
      Left            =   240
      TabIndex        =   95
      Top             =   15
      Width           =   1500
   End
   Begin VB.TextBox txtCP114 
      Height          =   270
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   25
      Top             =   3210
      Width           =   600
   End
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   3735
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1020
      Width           =   420
   End
   Begin VB.TextBox txtCP84 
      Height          =   288
      Left            =   7560
      TabIndex        =   3
      Top             =   1005
      Width           =   1110
   End
   Begin VB.TextBox txtEP04 
      Enabled         =   0   'False
      Height          =   270
      Left            =   3390
      MaxLength       =   6
      TabIndex        =   24
      Top             =   3210
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Index           =   5
      Left            =   4500
      TabIndex        =   45
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "修正/變更事項(&R)"
      Height          =   400
      Index           =   4
      Left            =   2976
      TabIndex        =   44
      Top             =   15
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   3
      Left            =   1752
      TabIndex        =   43
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6552
      TabIndex        =   47
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5724
      TabIndex        =   46
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7776
      TabIndex        =   48
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "補文件內容"
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   3210
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2580
      MaxLength       =   2
      TabIndex        =   55
      Top             =   690
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2340
      MaxLength       =   1
      TabIndex        =   54
      Top             =   690
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1500
      MaxLength       =   6
      TabIndex        =   53
      Top             =   690
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1020
      MaxLength       =   3
      TabIndex        =   52
      Top             =   690
      Width           =   495
   End
   Begin VB.Label lblCP167 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "刪未審頁:"
      Height          =   180
      Left            =   7710
      TabIndex        =   117
      Top             =   4650
      Width           =   765
   End
   Begin VB.Label lblCP168 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "刪已審頁:"
      Height          =   180
      Left            =   7710
      TabIndex        =   116
      Top             =   4920
      Width           =   765
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   20
      Left            =   3960
      TabIndex        =   35
      Top             =   4290
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   855
      Index           =   18
      Left            =   3870
      TabIndex        =   39
      Top             =   4740
      Width           =   3795
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "6694;1508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   19
      Left            =   1320
      TabIndex        =   19
      Top             =   2910
      Width           =   2265
      VariousPropertyBits=   671105051
      Size            =   "3995;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   855
      Index           =   17
      Left            =   60
      TabIndex        =   38
      Top             =   4740
      Width           =   3765
      VariousPropertyBits=   -1466941413
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "6641;1508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   16
      Left            =   1320
      TabIndex        =   28
      Top             =   4020
      Width           =   6150
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "10848;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   15
      Left            =   1320
      TabIndex        =   27
      Top             =   3750
      Width           =   6150
      VariousPropertyBits=   671105051
      MaxLength       =   250
      Size            =   "10848;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   14
      Left            =   1320
      TabIndex        =   26
      Top             =   3480
      Width           =   6150
      VariousPropertyBits=   671105051
      MaxLength       =   160
      Size            =   "10848;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   13
      Left            =   4980
      TabIndex        =   9
      Top             =   1560
      Width           =   375
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "661;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   12
      Left            =   4980
      TabIndex        =   5
      Top             =   1290
      Width           =   1260
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "2222;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   11
      Left            =   5205
      TabIndex        =   2
      Top             =   1020
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   10
      Left            =   4980
      TabIndex        =   12
      Top             =   1830
      Width           =   1365
      VariousPropertyBits=   671105051
      MaxLength       =   50
      Size            =   "2408;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   9
      Left            =   4980
      TabIndex        =   18
      Top             =   2640
      Width           =   1875
      VariousPropertyBits=   671105051
      Size            =   "3307;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   8
      Left            =   4980
      TabIndex        =   16
      Top             =   2370
      Width           =   1875
      VariousPropertyBits=   671105051
      Size            =   "3307;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   7
      Left            =   4980
      TabIndex        =   14
      Top             =   2100
      Width           =   1875
      VariousPropertyBits=   671105051
      Size            =   "3307;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   17
      Top             =   2640
      Width           =   2265
      VariousPropertyBits=   671105051
      Size            =   "3995;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   15
      Top             =   2370
      Width           =   2265
      VariousPropertyBits=   671105051
      Size            =   "3995;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   13
      Top             =   2100
      Width           =   2265
      VariousPropertyBits=   671105051
      Size            =   "3995;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   11
      Top             =   1830
      Width           =   1320
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "2328;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   270
      Index           =   2
      Left            =   1320
      TabIndex        =   8
      Top             =   1560
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   4
      Size            =   "1508;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   1290
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text7 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   1020
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   315
      Left            =   7350
      TabIndex        =   113
      Top             =   1800
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
      Caption         =   "Email維護:             (Y:是 / N:不寄)"
      Height          =   180
      Left            =   2640
      TabIndex        =   112
      Top             =   4485
      Width           =   2520
   End
   Begin VB.Label lblRecDate 
      AutoSize        =   -1  'True
      Caption         =   "當天報告:             (Y:是)"
      Height          =   180
      Left            =   6870
      TabIndex        =   111
      Top             =   2955
      Width           =   1815
   End
   Begin VB.Label lblPAID 
      Caption         =   "已收款:                (1-不寄D/N, 2-寄D/N)"
      Height          =   180
      Left            =   3660
      TabIndex        =   110
      Top             =   2955
      Width           =   3165
   End
   Begin VB.Label lblPrint 
      AutoSize        =   -1  'True
      Caption         =   "原文本印表機:"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3780
      TabIndex        =   107
      Top             =   840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblPayToday 
      AutoSize        =   -1  'True
      Caption         =   "電子送件是否當日扣款:         (Y/N)"
      Height          =   180
      Left            =   6120
      TabIndex        =   106
      Top             =   4335
      Width           =   2655
   End
   Begin VB.Label lblCP43 
      AutoSize        =   -1  'True
      Caption         =   "相關總收文號"
      Height          =   180
      Left            =   3660
      TabIndex        =   104
      Top             =   1410
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "下次管制收文日:"
      Height          =   180
      Left            =   2625
      TabIndex        =   103
      Top             =   4335
      Width           =   1305
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:          (Y: 是)"
      Height          =   180
      Left            =   6405
      TabIndex        =   102
      Top             =   1590
      Width           =   2085
   End
   Begin VB.Label lblAddCP64 
      AutoSize        =   -1  'True
      Caption         =   "新增補件內容至進度備註"
      Height          =   180
      Left            =   60
      TabIndex        =   101
      Top             =   5700
      Width           =   1980
   End
   Begin VB.Label Label5 
      Caption         =   "是否為複委任:"
      Height          =   225
      Left            =   60
      TabIndex        =   100
      Top             =   4320
      Width           =   1185
   End
   Begin VB.Label lblCP138 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "刪已審項:"
      Height          =   180
      Left            =   7710
      TabIndex        =   99
      Top             =   4035
      Width           =   765
   End
   Begin VB.Label lblCP135 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "增加頁數:"
      Height          =   180
      Left            =   7710
      TabIndex        =   98
      Top             =   3225
      Width           =   765
   End
   Begin VB.Label lblCP136 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "增加項數:"
      Height          =   180
      Left            =   7710
      TabIndex        =   97
      Top             =   3495
      Width           =   765
   End
   Begin VB.Label lblCP137 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "刪未審項:"
      Height          =   180
      Left            =   7710
      TabIndex        =   96
      Top             =   3765
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "核稿時數:"
      Height          =   180
      Index           =   1
      Left            =   5310
      TabIndex        =   94
      Top             =   3225
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工作時數:"
      Height          =   180
      Index           =   12
      Left            =   2925
      TabIndex        =   93
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人部門(日):"
      Height          =   180
      Left            =   60
      TabIndex        =   92
      Top             =   2940
      Width           =   1245
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6420
      TabIndex        =   91
      Top             =   1875
      Width           =   900
   End
   Begin VB.Label lblCP84 
      AutoSize        =   -1  'True
      Caption         =   "發文規費:"
      Height          =   180
      Left            =   6750
      TabIndex        =   90
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label lblEP04N 
      AutoSize        =   -1  'True
      Caption         =   "lblEP04N"
      Height          =   180
      Left            =   4365
      TabIndex        =   89
      Top             =   3225
      Width           =   675
   End
   Begin VB.Label lblEP04 
      AutoSize        =   -1  'True
      Caption         =   "核稿人:"
      Height          =   180
      Left            =   2790
      TabIndex        =   88
      Top             =   3225
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   30
      X2              =   8715
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   30
      X2              =   8710
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   1
      Left            =   5940
      TabIndex        =   87
      Top             =   690
      Width           =   765
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   7
      Left            =   6960
      TabIndex        =   86
      Top             =   690
      Width           =   1560
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2752;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   6
      Left            =   6300
      TabIndex        =   85
      Top             =   1335
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2328;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   5
      Left            =   2250
      TabIndex        =   84
      Top             =   1605
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2328;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   4
      Left            =   2250
      TabIndex        =   83
      Top             =   1335
      Width           =   1320
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2328;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   3
      Left            =   6960
      TabIndex        =   82
      Top             =   450
      Width           =   1560
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2752;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   2
      Left            =   4260
      TabIndex        =   81
      Top             =   690
      Width           =   1560
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2752;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   1
      Left            =   4260
      TabIndex        =   80
      Top             =   450
      Width           =   1560
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2752;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "案件備註:"
      Height          =   180
      Left            =   3870
      TabIndex        =   79
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "進度備註:"
      Height          =   180
      Left            =   60
      TabIndex        =   78
      Top             =   4560
      Width           =   765
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱(日):"
      Height          =   180
      Left            =   60
      TabIndex        =   77
      Top             =   4065
      Width           =   1065
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱(英):"
      Height          =   180
      Left            =   60
      TabIndex        =   76
      Top             =   3795
      Width           =   1065
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱(中):"
      Height          =   180
      Left            =   60
      TabIndex        =   75
      Top             =   3525
      Width           =   1065
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "補文件期限:"
      Height          =   180
      Left            =   60
      TabIndex        =   74
      Top             =   3225
      Width           =   945
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人(日)2:"
      Height          =   180
      Left            =   3660
      TabIndex        =   73
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人(英)2:"
      Height          =   180
      Left            =   3660
      TabIndex        =   72
      Top             =   2415
      Width           =   975
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人(中)2:"
      Height          =   180
      Index           =   0
      Left            =   3660
      TabIndex        =   71
      Top             =   2145
      Width           =   975
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人(日)1:"
      Height          =   180
      Left            =   60
      TabIndex        =   70
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人(英)1:"
      Height          =   180
      Left            =   60
      TabIndex        =   69
      Top             =   2415
      Width           =   975
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "聯絡人(中)1:"
      Height          =   180
      Left            =   60
      TabIndex        =   68
      Top             =   2145
      Width           =   975
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "(N: 不收)"
      Height          =   180
      Left            =   5400
      TabIndex        =   67
      Top             =   1590
      Width           =   690
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "是否向客戶收款:"
      Height          =   180
      Left            =   3660
      TabIndex        =   66
      Top             =   1605
      Width           =   1305
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號:"
      Height          =   180
      Left            =   3660
      TabIndex        =   65
      Top             =   1875
      Width           =   765
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "客戶案件案號:"
      Height          =   180
      Left            =   60
      TabIndex        =   64
      Top             =   1875
      Width           =   1125
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   60
      TabIndex        =   63
      Top             =   1605
      Width           =   765
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "鑑定報告代理人:"
      Height          =   180
      Left            =   3660
      TabIndex        =   62
      Top             =   1335
      Width           =   1305
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   60
      TabIndex        =   61
      Top             =   1335
      Width           =   585
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "催審期限:"
      Height          =   180
      Left            =   4410
      TabIndex        =   60
      Top             =   1065
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "發文日:"
      Height          =   180
      Left            =   30
      TabIndex        =   59
      Top             =   1065
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "收款後辦案:"
      Height          =   180
      Left            =   5940
      TabIndex        =   58
      Top             =   450
      Width           =   945
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   0
      Left            =   3420
      TabIndex        =   57
      Top             =   690
      Width           =   765
   End
   Begin VB.Label Label6 
      Caption         =   "本所案號:"
      Height          =   255
      Left            =   60
      TabIndex        =   56
      Top             =   690
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3420
      TabIndex        =   51
      Top             =   450
      Width           =   765
   End
   Begin MSForms.Label Label3 
      Height          =   285
      Index           =   0
      Left            =   1020
      TabIndex        =   50
      Top             =   450
      Width           =   1560
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "2752;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   49
      Top             =   450
      Width           =   585
   End
End
Attribute VB_Name = "frm060104_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo By Sindy 2021/11/15 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strReceiveNo As String
'Add By Sindy 2023/4/12
Dim strPD01 As String
Dim m_str938RecvNo As String '回傳退費的超頁費文號
Dim m_str939RecvNo As String '回傳退費的超項費文號
'2023/4/12 END
'Modify by Morgan 2005/8/4 改用動態陣列
'Dim pa(1 To T_pa) As String, cp(1 To T_CP) As String
Dim pa() As String, cp() As String
Dim pageD() As String 'Add By Sindy 2023/3/15
Dim intWhere As Integer
' 案件性質
Dim m_CP10 As String
'Add By Cheng 2003/10/06
Dim m_blnClkChgEvnBtn As Boolean '是否按下變更事項按鈕
'Add by Morgan 2004/3/23
'實審通知日本所期限,法定期限,母案收文號
Dim m_stVar(0 To 3) As String
Dim m_stCP09 As String
Dim bolDelay As Boolean 'Add by Morgan 2004/9/8 是否延期過
Dim m_strDelayCP09 As String 'Added by Morgan 2011/11/11 延期收文號
'Removed by Morgan 2015/9/1 取消--靜芳
''Add by Morgan 2005/9/8
'Dim bolXCase As Boolean '是否為特殊代理人的申請案
'Dim bolXCtrl As Boolean '是否管制核稿期限
'end 2015/9/1
Dim m_strMailCP09 As String
Dim stDivInfo(1 To 5) As String '分割案資料 1~4=母案號, 5=申請國家
'Add by Morgan 2006/4/28
Dim m_bol108 As Boolean '是否有收文申請寄存
'Add by Morgan 2006/6/8
Public m_bolSaveChgEvent As Boolean '更新變更資料
'2007/8/6 ADD BY SONIA
Public m_CP50 As String    '第三人提實審中文名稱
Public m_CP51 As String    '第三人提實審英文名稱
Public m_CP52 As String    '第三人提實審日文名稱
'Add by Morgan 2007/9/14
Dim m_RefCP10 As String '相關收文號案件性質
Dim m_PA143 As String 'Add by Morgan 2008/3/18
Dim m_CP09s As String, m_CP123s As String 'Add by Morgan 2009/3/20 收文號,是否算發文室案件
Dim m_CP130 As String 'Add by Morgan 2009/4/28 發文-主管機關
'Add by Morgan 2010/1/6
Dim m_bol99Case As Boolean '是否99年後申請案件
Dim m_bolChkFee As Boolean '是否需檢查規費
Dim m_bolChkPageItem As Boolean '是否要輸頁數與項數
'Dim strNewCaseCP118 As String 'Add By Sindy 2018/4/25
Dim m_lngOverPageFee As Long, m_lngOverItemFee As Long '超頁費,超項費
Dim m_lngOverPageFeeDiff As Long, m_lngOverItemFeeDiff As Long '超頁費,超項費差額
Dim m_lngRecOverPageFee As Long, m_lngRecOverItemFee As Long '已收文超頁費,超項費 Add by Morgan 2011/6/29
Dim m_FeeMemo As String '規費備註
Dim m_lngOfficialFee As Long '原始規費
'end 2010/1/6
Dim m_bolChkItem As Boolean '是否要檢查增刪項數 Add by Morgan 2010/9/27
Dim m_Div416OfficialFee As Long  '分割案實審規費  2010/12/8 add by sonia
Dim m_bolWebApp As Boolean 'Add by Morgan 2011/1/17 是否電子送件案
Dim m_PA162 As String 'Added by Morgan 2012/12/3
Dim m_PA163 As String 'Added by Morgan 2012/12/26

'Added by Morgan 2013/1/9
Dim m_bol107NewFee As Boolean '台灣再審是否用102年新規費計算
Dim m_bolFixNewFee As Boolean '台灣修正是否用102年新規費計算
Dim m_strReExamCP27 As String '台灣再審發文日(若再審延期發文日)
Dim m_strNewAppIpoNo As String 'Added by Morgan 2013/8/2 新申請案智慧局收文文號

'Add by Amy 2013/08/22 發明申請案發文時是否實審要掛 交承辦收文告代 (for 代理人Y49456)-靜芳
Dim bol416Msg As Boolean

Dim bol416Control As Boolean '是否管制實審期限 Added by Morgan 2013/12/11
Dim bolCreditNote As Boolean '是否抵帳 Added by Morgan 2014/2/7

'Add by Amy 2014/05/27 For 印簡易聯絡單
Dim iPage As Integer
Dim dblLineHeight As Double '行高
Dim m_dblTitleHeight As Double '抬頭
Dim m_dblTop As Double '上邊界
Dim m_dblLeft As Double '左邊界
Dim m_TBWidth As Double '表格寬
Dim intLine As Integer
Dim dblPrtX As Double
Dim dblPrtY As Double
Dim intFieldWidth
Dim m_PA60 As String 'Added by Morgan 2015/7/9 一案兩請是否放棄新型
Dim m_bol435 As Boolean 'Added by Morgan 2015/9/10
'Added by Lydia 2015/12/31 會稿924的相關總收文號資料
Dim m_CP43cpm As String
Dim m_CP43date1 As String
Dim bolAddSC As Boolean 'Added by Lydia 2016/01/21 是否產生行事曆管制
Dim bolDefWebApp As Boolean 'Added by Lydia 2017/12/12 客戶或代理人有設定為電子送件
'Added by Lydia 2018/03/27
Dim strPrinter As String  '系統預設印表機
Dim m_AttchPath As String  '說明書暫存區
Private Const cPrintORI As String = "101,102,103,125" 'Added by Lydia 2018/05/17 需要印ORI.PDF的新案性質(+衍生設計案125)
Dim m_allPage As String, m_allItem As String '總頁數,總項數
Dim bolChkSave As Boolean 'Added by Lydia 2018/10/31 本次是否已發文
Dim m_EP09 As String 'Added by Lydia 2018/12/08 翻譯完稿日
'Added by Lydia 2019/01/07
Dim m_strCP148 As String '申復205及再審107是否有一併修正 'Added by Lydia 2019/01/28 主動修正是否已併入中說送件
Dim m_GrpMan As String '(副本)工程師主管 'Memo by Lydia 2019/08/19 日文組副本:除各組主任(99034,94012)給主管,其餘人給審核主管
'end 2019/01/07
Dim mDate209210 As String 'Added by Lydia 2019/01/17 客戶提供中說期限
Dim mDate201CP158 As String 'Added by Lydia 2019/01/28 中說發文日
Dim mFA10 As String 'Added by Lydia 2019/07/03 代理人國籍
Dim m_TCTchk As String 'Added by Lydia 2019/07/30 命名記錄是否分組/工程師
Dim mDateTF30 As String 'Added by Lydia 2019/12/11 客戶提供英文翻譯本(行事曆日期)
Dim m_307CP64 As String 'Added by Morgan 2020/2/26 分割備註
Dim m_EdDivSugInform As Boolean 'Added by Morgan 2020/2/6 修改分割加註通知
'Added by Lydia 2021/01/21 FCP實審發文承辦單不出紙本改發email
Dim m_416Type As String '實審(實體審查)發文是否出帳單
Dim m_AddMcRecord As String '人工Email維護(語法)
Dim m_eFlag As String '是否e/E化
'end 2021/01/21
Dim m_pAgreeOnDate As String 'Modify By Sindy 2021/4/27
Dim m_AgentName As String 'Add By Sindy 2021/5/10
'Added by Lydia 2022/11/11
Dim m_LosCP84 As String '法律所案源之規費
Dim m_LOS15 As String '法律所案源單號
Dim m_LOS02 As String '法律所案源類別
Dim m_LosMemo As String 'EMail說明
Dim m_bolIns908 As Boolean 'Add By Sindy 2023/3/27 是否要同時產生內部收文代辦退費


'Add by Morgan 2005/9/8檢查是否有收文翻譯,檢視中說,製作中說
Private Function chkCPExist() As Boolean
     
On Error GoTo ErrHnd
   
   CheckOC3
   With AdoRecordSet3
      'Modified by Morgan 2013/11/6 +235核對中說格式
      strSql = "select 1 from caseprogress a WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 IN('201','209','235','210') AND CP27 IS NULL AND CP57 IS NULL"
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         chkCPExist = True
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub cboAddCP64_Click()
   Text7(17) = Text7(17) & IIf(Text7(17) = "", "", ", ") & cboAddCP64
   Text7(17).SelStart = Len(Text7(17))
End Sub

Private Sub chkCP86_Click(Index As Integer)
   If chkCP86(Index).Value = 1 Then
      chkCP86(Abs(Index - 1)).Value = 0
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bolEmail As Boolean 'Add by Morgan 2009/10/14
Dim bolPlusPaper As Boolean, iCopy As Integer 'Add by Morgan 2009/10/20
Dim strET03 As String 'Added by Morgan 2014/2/7
Dim strTo As String
'Dim strSubject As String, strContent As String 'Add By Sindy 2015/12/14
'Dim strTemp As String 'Add By Sindy 2015/12/14
'Added by Lydia 2018/03/26
Dim strFilePath As String '記錄智慧局收文文號
Dim strNewName As String '上傳到卷宗區自動更名
Dim strAutoPrinter As String '自動列印的印表機
Dim intAuto As Integer '列印說明書份數
Dim strAutoList As String '自動列印的檔案
'end 2018/03/26
Dim strNewCP64 As String 'Added by Lydia 2018/04/16 保留進度備註
Dim nFrm As Form 'Added by Lydia 2021/01/21

   'Added by Lydia 2017/12/13
   Dim fso As New FileSystemObject
   Dim fdr As Folder
   Dim fl
   Dim strDefDir As String, strNewFileName As String
   Dim strTmp As String
   'end 2017/12/13
   
   Select Case Index
      'Modify by Morgan 2009/3/26 將同時發文併入
      'Case 0 '確定
      Case 0, 3
         ' 90.08.29 modify by louis (先檢查變更事項檔是否存在)
         If Text7(2) = 變更 And IsChangeEventExist(strReceiveNo) = False Then
            MsgBox "請先輸入變更事項的資料", vbOKOnly + vbCritical, "檢核資料"
            Exit Sub
         End If
         'Add By Cheng 2002/05/21
         If CheckDataValid = False Then
            Exit Sub
         End If
         
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         If pa(1) = "FCP" Then 'Add by Morgan 2006/11/2
         
'Removed by Morgan 2015/9/1 取消--靜芳
'            'Add by Morgan 2005/9/8 提示更新核稿期限
'            bolXCtrl = False: bolXCase = False
'            'Modified by Morgan 2012/12/20 +衍生設計125
'            If pa(75) = "Y34232" And InStr("101,102,103,105,125", Text7(2).Text) > 0 Then
'               If chkCPExist = True Then
'                  bolXCase = True
'                  If PUB_ChkCPExist(cp, "924", 1) = False And PUB_ChkCPExist(cp, "949", 1) = False Then 'Added by Morgan 2015/6/9 若有收文會稿924或寄中說949則不管制核稿期限 --靜芳
'
'                     If MsgBox("是否管制核稿期限", vbYesNo + vbDefaultButton1) = vbYes Then
'                        bolXCtrl = True
'                     End If
'
'                  End If 'Added by Morgan 2015/6/9
'               End If
'            End If
'end 2015/9/1
          
            'Added by Lydia 2018/03/27 抓取預設印表機
            strAutoPrinter = "": strAutoList = ""
            intAuto = 0
            'Modified by Lydia 2018/05/17 改從畫面選取印表機
'            If InStr("101,102,103", Text7(2).Text) > 0 And TransDate(cp(5), 2) >= FCP案件命名啟用日 Then
'                strExc(2) = Pub_GetSpecMan("FCP程序機密列印")
'                If cboPrinter.ListCount > 0 And strExc(2) <> "" Then
'                    For intI = 0 To cboPrinter.ListCount - 1
'                         strExc(1) = cboPrinter.List(intI)
'                         '用系統特殊設定
'                         If strExc(1) = strExc(2) Then
'                              strAutoPrinter = strExc(1)
'                              Exit For
'                         End If
'                    Next intI
'                End If
'                If strAutoPrinter = "" Then
'                    If Pub_StrUserSt03 <> "M51" Then
'                        MsgBox "需要9F機密列印的印表機，請通知電腦中心！"
'                        Exit Sub
'                    Else
'                        MsgBox "列印說明書需要印表機" & vbCrLf & strExc(2) & vbCrLf & "，本次發文只會印承辦單！"
'                    End If
'                Else
'                    intAuto = 1
'                End If
'            End If
            'Mark by Lydia 2023/03/03 與Sharon,Phoebe 確認不用列印說明書
            'If cboPrinter.Visible = True Then
            '     strAutoPrinter = cboPrinter.Text
            '     intAuto = 1
            'End If
            'end 2018/03/27
            
            'Add by Amy 2013/08/22 +發明申請案是示是否實審要掛 交承辦收文告代 的備註
            bol416Msg = False
            If pa(75) = "Y49456" And Text7(2).Text = "101" Then
                If Check416Exist = False Then
                    If MsgBox("是否實審要掛 交承辦收文告代 ?", vbYesNo + vbDefaultButton1) = vbYes Then
                     bol416Msg = True
                  End If
                End If
            End If
            'end 2013/08/22
            
            'Added by  Lydia 2021/01/29 若實審發文日＝新案發明101、分割307發文日，即實審與發明或分割同一天發文，不出定稿及帳單，Email維護=N不發通知Email。
            If Text7(2) = "416" Then
                strExc(0) = "Select Max(Cp158||cpm03) Cp158 From Caseprogress,Casepropertymap  where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' " & _
                                 "and cp10 in ('101','307') and cp158 > 0 and cp159 = 0 and cp01=cpm01(+) and cp10=cpm02(+) "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                    If Mid("" & RsTemp.Fields("cp158"), 1, 8) = DBDATE(Text7(0)) Then
                        txtEmail = "N"
                        m_416Type = "A"
                    End If
                End If
                'Added by Lydia 2021/03/16 若該案之新案翻譯201、核對中說格式209、檢視中說210、製作中說235尚未發文，則預設實審發文不出定稿及帳單，Email維護=N不發通知Email。
                strExc(0) = "select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' " & _
                                 "and cp10 in ('201','209','210','235') and cp158 = 0 and cp159 = 0 "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                    txtEmail = "N"
                    m_416Type = "A"
                End If
                'end 2021/03/16
            End If
            'end 2021/01/29
            
            'Added by Lydia 2019/01/07 發文申復205及再審查申請107時，多一道詢問"是否有一併修正"(記錄CP148)，若"是"則和第1點一樣流程，若"否"則不作動作繼續發文。
            'Memo by Lydia 2023/05/18 因為收再審申請107是不會收主動修正,所以由人工判斷---from Phoebe ; ex.FCP-60457的再審AB2015671誤選擇"是"
            If (Text7(2).Text = "205" Or Text7(2).Text = "107") And m_strCP148 = "" Then
               If MsgBox("是否有一併修正？", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
                   m_strCP148 = "Y"
               End If
            'Added by Lydia 2019/01/28 主動修正發文時，若新案翻譯(含檢視中說、核對中說格式)為已發文且同日發文，則詢問「主動修正是否已併入中說送件」
            ElseIf Text7(2).Text = "203" And DBDATE(Text7(0).Text) = mDate201CP158 And m_strCP148 = "" Then
               If MsgBox("主動修正是否已併入中說送件？", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
                   m_strCP148 = "Y"
               End If
            'end 2019/01/28
            'Add By Sindy 2019/10/17
            '新案發文時檢查是否有收201.新案翻譯(不含檢視中說、核對中說格式)
            '及924.會稿且會稿有掛本所和法定期限
            '且新案翻譯的交稿日期和只交Claims期限均無日期時,才彈訊息"請管控交稿期限"
            ElseIf InStr(NewCasePtyList, Text7(2).Text) > 0 Then
               If PUB_ChkCPExist(cp, "201", 1) = True Then '有新案翻譯
                  If PUB_ChkCPExist(cp, "924", 1) = True Then '有會稿
                     '新案翻譯的交稿日期和只交Claims期限均無日期
                     strExc(0) = "SELECT cp09 FROM CASEPROGRESS,TransFee WHERE cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                                 " and cp10='201' and cp09=TF01 and TF26 is null and TF32 is null and cp27 is null"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        '會稿收文有掛本所和法定期限
                        strExc(0) = "SELECT cp09 FROM CASEPROGRESS WHERE cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                                    " and cp10='924' and cp06>0 and cp07>0 and cp27 is null"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           MsgBox "請管控交稿期限！"
                        End If
                     End If
                  End If
               End If
            '2019/10/17 END
            End If
            'end 2019/01/07
            
            'Modify By Sindy 2023/4/27 應該不用檢查此問題 ex:FCP-68521
'            'Add By Sindy 2023/3/21
'            '主動修正發文時，檢查
'            If Text7(2).Text = "203" And DBDATE(Text7(0).Text) = mDate201CP158 Then
'               If m_strCP148 = "" Then
'                  strExc(0) = "SELECT * FROM PageDetail WHERE pd01='" & cp(9) & "'"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     If "" & RsTemp.Fields("pd20") <> "" Then
'                        MsgBox "主動修正未於中說一併送件，但增刪頁數有掛中說一併修正之文號，請確認！", vbOKOnly + vbCritical, "檢核資料"
'                        Exit Sub
'                     End If
'                  End If
'               End If
'            End If
'            '2023/3/21 END
            
            '2006/3/10 ADD BY SONIA 未輸入補件內容時提示委任狀是否管制期限
            'Modif by Morgan 2006/4/20 加行政再審504,抗告509
            If m_CP10 = "501" Or m_CP10 = "503" Or m_CP10 = "507" Or m_CP10 = "504" Or m_CP10 = "509" Then
               strExc(0) = "SELECT NP08 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP06 IS NULL AND NP07=" & 補文件 & ""
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI <> 1 Then
                  If MsgBox("委任狀是否管制期限 ?", vbYesNo + vbDefaultButton1) = vbYes Then
                     Command1_Click
                     Exit Sub
                  End If
               End If
            End If
            'Add by Morgan 2006/6/8
            If Text7(2) = 變更 And m_bolSaveChgEvent = False Then
               MsgBox "案件性質為變更時須點【修正/變更事項】並存檔才可發文！"
               Exit Sub
            End If
            'end 2006/6/8
         
            '2007/8/6 ADD BY SONIA 代理人Y45148提醒是否為第三人提實審
            '2010/5/28 MODIFY BY SONIA 改判斷任一申請人為X45149才提醒
            'If m_CP10 = "416" And (m_CP50 & m_CP51 & m_CP52 = "") And pa(75) = "Y45148" Then
            If m_CP10 = "416" And (m_CP50 & m_CP51 & m_CP52 = "") And (pa(26) = "X45149" Or pa(27) = "X45149" Or pa(28) = "X45149" Or pa(29) = "X45149" Or pa(30) = "X45149") Then
               If MsgBox("申請人為NIKON, 是否為第三人提實審 ?", vbYesNo + vbDefaultButton1) = vbYes Then
                  cmdok_Click (6)
                  Exit Sub
               End If
            End If
            '2007/8/6 END
   
            'Add by Morgan 2008/3/28 主動修正,更正,專利權延長,舉發答辯
            If InStr("203,402,415,804", Text7(2)) > 0 Then
               '若未辦或不辦重新委任時不可發文
               If PUB_Check928NotOk(pa) = True Then
                  MsgBox "本案下一程序有重新委任之補文件未辦理，不可發文！"
                  Exit Sub
               End If
            End If
            
            '若基本檔年費申請人是否出名為N時提醒存檔將取消
            m_PA143 = pa(143)
            'Modified by Morgan 2012/3/27 +937(更換FC代理人) Ex.FCP-022297
            If pa(143) = "N" And InStr("901,902,903,904,905,906,912,937", Text7(2)) = 0 Then
               MsgBox "年費申請人是否出名現為【N】，存檔時將自動取消！"
               m_PA143 = ""
            End If
            'end 2008/3/28
            
'Removed by Morgan 2012/6/15 101.6.1 改與其他組相同預設，不必另外提醒
'            '2011/11/30 add by sonia電子電機組新案出名代理人若非林特助則提醒
'            If pa(150) = "1" And InStr(NewCasePtyList & ",803", m_CP10) > 0 And cp(110) <> "94007" Then
'               If MsgBox("電子電機組新案出名代理人非林景郁, 是否確定 ?", vbYesNo + vbDefaultButton1) = vbNo Then
'                  lstNameAgent.SetFocus
'                  Exit Sub
'               End If
'            End If
'            '2011/11/30 END
'end 2012/6/15

            'Added by Morgan 2013/7/10 台灣一案兩請提醒檢查申請書
            If (Text7(2) = "101" Or Text7(2) = "102") Then
               strExc(0) = "select cm05||'-'||cm06||decode(cm07||cm08,'000','','-'||cm07||'-'||cm08) p2 from casemap" & _
                  " where cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "' and cm10='3'" & _
                  " union select cm01||'-'||cm02||decode(cm03||cm04,'000','','-'||cm03||'-'||cm04) p2 from casemap" & _
                  " where cm05='" & pa(1) & "' and cm06='" & pa(2) & "' and cm07='" & pa(3) & "' and cm08='" & pa(4) & "' and cm10='3'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If MsgBox("本案與 " & RsTemp(0) & " 案為一案兩請，請檢查申請書上是否已載明本案同時申請發明及新型！" & vbCrLf & "是否要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                     Exit Sub
                  End If
               End If
            End If
            'end 2013/7/10
   
         End If
         '2006/3/10 END
         
         'Added by Lydia 2017/12/12 客戶或代理人有設定為電子送件
         If bolDefWebApp = True And InStr(NewCasePtyList, cp(10)) > 0 And txtCP118 <> "Y" Then
             If MsgBox("客戶或代理人有設定為FCP電子送件，是否繼續發文？", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                 If txtCP118.Enabled = True Then
                     txtCP118.SetFocus
                 End If
                 Exit Sub
             End If
         End If
         'end 2017/12/12
         
         'Added by Lydia 2021/01/21 FCP實審發文承辦單不出紙本改發email：Email維護
         m_AddMcRecord = ""
         If cp(10) = "416" And txtEmail.Visible = True And txtEmail.Text = "Y" Then
            strExc(5) = Pub_FcpSetPayToday("2", Text7(0).Text, txtPayToday.Text) '扣款日
            '開啟Email畫面
            Call PUB_GetFCPEmpMail416("2", strReceiveNo, m_eFlag, m_416Type, txtPAID, txtRecDate, DBDATE(Text7(0).Text), strExc(5), strExc(1), strExc(2), strExc(3), strExc(4))
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
               If m_AddMcRecord = "" Or m_AddMcRecord = "Y" Then
                   MsgBox "Email維護未確認，請重新確認Email !", vbCritical, "檢核資料"
                   Exit Sub
               End If
            End If
         End If
         'end 2021/01/21
         
         strNewCP64 = Text7(17).Text 'Added by Lydia 2018/04/18
         m_307CP64 = "" 'Added by Morgan 2020/2/26
         'Add by Morgan 2011/1/17
         If txtCP118 = "Y" Then
            m_CP123s = ""
            'Added by Morgan 2016/5/16 電子送件也要記錄主管機關
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text7(0), , True) = False Then
               Exit Sub
            End If
            'end 2016/5/16
            
            'Added by Morgan 2012/2/8
            'Modified by Morgan 2012/6/4
            '改要輸入收文文號
            'If MsgBox("本案要影印智慧局之【收件成功通知】!!", vbOKCancel + vbDefaultButton2) <> vbOK Then
            strExc(0) = InputBox("請輸入智慧局收文文號!!", , m_strNewAppIpoNo)
            If strExc(0) = "" Then
               Exit Sub
            Else
               strFilePath = strExc(0)  'Added by Lydia 2018/03/26 記錄智慧局收文文號
               'Modified by Lydia 2018/04/16 先保留進度備註，等檢查完後更新欄位
               'Text7(17) = "智慧局收文文號:" & strExc(0) & ";" & Text7(17)
               strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & Text7(17)
               m_307CP64 = "智慧局收文文號:" & strExc(0) & ";"  'Added by Morgan 2020/2/26
            End If
         
         Else
         'end 2011/1/17
         
            'Add by Morgan 2009/4/28
            If ModifyDispatchCp130(strReceiveNo, m_CP09s, m_CP123s, m_CP130, Text7(0)) = False Then
               Exit Sub
            End If
            
            If m_CP123s = "Y" Then
            'end 2009/4/28
               'Add by Morgan 2009/3/20 設定是否算發文室案件
               'modify by sonia 2014/6/23 加傳發文規費, P-108903
               If ModifyDispatch(strReceiveNo, m_CP09s, m_CP123s, txtCP84, Text7(0)) = False Then
                   Exit Sub
               End If
               'end 2009/3/20
            End If
         End If 'Add by Morgan 2011/1/17
         
         'Add By Sindy 2023/6/27 當發文規費為300且是否向客戶收款N，彈視窗詢問失誤人員，Key失誤人員姓名後確定→失誤人員: XXX
         '回寫進度備註→則雖有發文規費且是否向客戶收款N則發文過。
         If Text7(2) = "403" And Val(txtCP84) = 300 And Text7(13) = "N" Then
            strExc(0) = InputBox("請輸入失誤人員姓名!!", , m_strNewAppIpoNo)
            If strExc(0) = "" Then
               Exit Sub
            Else
               strExc(10) = GetPrjSalesNM(strExc(0)) '員編換姓名
               If strExc(10) <> "" Then strExc(0) = strExc(10)
               strNewCP64 = IIf(strNewCP64 <> "", strNewCP64 & ";", "") & "失誤人員:" & strExc(0)
            End If
         End If
         '2023/6/27 END
         
         'Added by Morgan 2012/12/26
         '發明案分割發文要設定是否初審階段提分割
         'Modified by Morgan 2015/9/30
         'If Text7(2) = "307" And pa(8) = "1" Then
         If Text7(2) = "307" And pa(8) = "1" And m_PA163 = "" Then
         'end 2015/9/30
            m_PA163 = PUB_GetDivCaseState(pa, DBDATE(Text7(0)))
            If m_PA163 = "" Then
               Exit Sub
            End If
         End If
         'end 2012/12/26
         
         'Added by Morgan 2015/7/9
         m_PA60 = ""
         If Text7(2) = "239" Then
            intI = MsgBox("是否放棄新型案？", vbYesNoCancel + vbQuestion + vbDefaultButton3, "一案兩請擇一選擇")
            If intI = vbYes Then
               m_PA60 = "Y"
            
            ElseIf intI = vbNo Then
               If MsgBox("是否確定放棄發明案？", vbYesNo + vbExclamation + vbDefaultButton2, "一案兩請擇一選擇") = vbYes Then
                  m_PA60 = "N"
               'Modified by Morgan 2016/7/28 可能會都不放棄 Ex.FCP-48486 --敏莉
               'Else
               '   Exit Sub
               'end 2016/7/28
               End If
            Else
               Exit Sub
            End If
         End If
         'end 2015/7/9
         
         'Added by Lydia 2018/03/26 不限新案，依據輸入的智慧局收文號(受理號,ex: 1073066637-0)，將本機C:\E-SET\RdcDocDir\(收文號ex: 1073066637-0)的pdf檔自動搬移到卷宗區(by Phoebe);
         'Memo by Lydia 2018/03/30 重新發文不做搬檔
         'Modified by Lydia 2018/08/10 重新發文要詢問(比照FCT發文自動上傳檔案)
         'If txtCP118.Text = "Y" And strFilePath <> "" And Val(cp(82)) = 0 Then
         If txtCP118.Text = "Y" And strFilePath <> "" Then
             strExc(1) = cp(82)
             If Val(cp(82)) > 0 Then
                 If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                      strExc(1) = ""
                 End If
             End If
        'end 2018/08/10
             'Modified by Lydia 2018/06/20 改成共用模組
'             strFilePath = "C:\E-SET\RdcDocDir\" & strFilePath
'             strExc(3) = Dir(strFilePath & "\*.pdf")
'             If strExc(3) = "" Then
'                 'Modified by Lydia 2018/04/16 因為實審和主動修正可併在一起，所以E-Set資料夾查無檔案則彈訊息問是否上傳檔案(by Phoebe)
'                 'MsgBox "尚未從智慧局下載檔案，不可發文", vbCritical
'                 If MsgBox("電子送件是否要上傳檔案到卷宗區？" & vbCrLf & "(Yes=要上傳；No=不要上傳，繼續發文)", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
'                     Exit Sub
'                 End If
'                 'end 2018/04/16
'             Else
'                 Do While strExc(3) <> ""
'                     strNewName = PUB_CaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
'                    '檢查檔案是否正在使用中
'                     If PUB_ChkFileOpening(strFilePath & "\" & strExc(3)) = True Then
'                         MsgBox strFilePath & "\" & strExc(3) & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
'                         Exit Sub
'                     End If
'                     'Added by Lydia 2018/04/12 電子送件自動匯入卷宗區請自動排除檔名中有.FIX_U(劃線本)和.COR_U的檔案by Phoebe
'                     If InStr(UCase(strExc(3)), ".FIX_U") > 0 Or InStr(UCase(strExc(3)), ".COR_U") > 0 Then
'                        Kill strFilePath & "\" & strExc(3)
'                     Else
'                     'end 2018/04/12
'                        If InStr(strExc(3), "中文本") > 0 Then
'                              'Modified by Lydia 2018/04/16 判斷改為專利種類，掛在中說進度的卷宗區
'                              'If Text7(2) = "101" Then
'                              If pa(8) = "1" Then
'                                   strNewName = strNewName & "." & Text7(2) & ".inv.pdf"
'                              'Modified by Lydia 2018/04/16
'                              'ElseIf Text7(2) = "102" Then
'                              ElseIf pa(8) = "2" Then
'                                   strNewName = strNewName & "." & Text7(2) & ".utl.pdf"
'                              'Modified by Lydia 2018/04/16
'                              'ElseIf Text7(2) = "103" Then
'                              ElseIf pa(8) = "3" Then
'                                   strNewName = strNewName & "." & Text7(2) & ".des.pdf"
'                              Else
'                                   MsgBox "下列檔案請手動上傳卷宗區後，再做發文!" & vbCrLf & strFilePath & "\" & strExc(3)
'                                   Exit Sub
'                              End If
'                        ElseIf InStr(strExc(3), "修正說明書") > 0 Then
'                              strNewName = strNewName & "." & Text7(2) & ".ori.fix.pdf"
'                        'Added by Lydia 2018/04/10
'                        ElseIf InStr(strExc(3), "申請書") > 0 Then
'                              strNewName = strNewName & "." & Text7(2) & ".data.pdf"
'                        'end 2018/04/10
'                        '其他
'                        Else
'                              If Left(UCase(strExc(3)), Len(strNewName)) = strNewName Then
'                                    strNewName = strNewName & "." & Text7(2) & "." & PUB_GetSimpleName(Mid(strExc(3), Len(strNewName) + 1))
'                              'FCPXXXXX(6碼) 開頭
'                              ElseIf Left(UCase(strExc(3)), Len(pa(1) & pa(2))) = pa(1) & pa(2) Then
'                                    strNewName = strNewName & "." & Text7(2) & "." & PUB_GetSimpleName(Mid(strExc(3), Len(pa(1) & pa(2)) + 1))
'                              '+案號
'                              Else
'                                    strNewName = strNewName & "." & Text7(2) & "." & PUB_GetSimpleName(strExc(3))
'                              End If
'                        End If
'                        strNewName = Replace(strNewName, "..", ".") 'Added by Lydia 2018/04/10
'                        If SaveAttFile_PDF(Label3(0), strFilePath & "\" & strExc(3), strNewName, Val(strSrvDate(1)), Val(Left(Format(ServerTime, "000000"), 4)), False) Then
'                              Kill strFilePath & "\" & strExc(3)
'                        Else
'                              Exit Sub
'                        End If
'                     End If 'end 2018/04/12
'                     strExc(3) = Dir(strFilePath & "\*.pdf")
'                 Loop
'                 '無檔案，刪除資料夾
'                 If Dir(strFilePath & "\*.*") = "" Then
'                      RmDir strFilePath
'                 End If
'             End If
             If Val(strExc(1)) = 0 Then 'Added by Lydia 2018/08/10 重新發文要詢問(比照FCT發文自動上傳檔案)
                'Modified by Lydia 2019/03/22 +傳入發文日
                If Pub_AutoEsetToCpp(True, pa(1), pa(2), pa(3), pa(4), pa(8), Label3(0).Caption, Text7(2).Text, strFilePath, Text7(0).Text) = False Then
                      Exit Sub
                End If
                'end 2018/06/20
             End If
         End If
         'end 2018/03/26
         
         'Added by Lydia 2018/03/27 下載說明書 (抓案件的最新一道ori.pdf、ori.repX.pdf或ori.fix.pdf檔案)
            '新案發文(101，102，103)時，自動抓卷宗區說明書(ori.pdf、ori.repX.pdf或ori.fix.pdf)列印，印表機設定機密列印(ApeosPort-IV C6685(9F影印機2.可彩色.機密列印))；
            '原新案發文產生之簡易連絡單改由機密列印，順序先印簡易連絡單，後印說明書
            '非電子送件的說明書：發明，新型有收文新案翻譯，列印2份，無收文新案翻譯列印1份；設計案列印1份。有序列表(.SEQ.)一併列印
            '電子送件的說明書：發明，新型有收文新案翻譯，列印1份，無收文新案翻譯不列印；設計案不列印。有序列表(.SEQ.)一併列印
         'Modified by Lydia 2018/05/17 改成常數
         'If InStr("101,102,103", Text7(2).Text) > 0 And TransDate(cp(5), 2) >= FCP案件命名啟用日 And strAutoPrinter <> "" Then
         'Mark by Lydia 2023/03/03 與Sharon,Phoebe 確認不用列印說明書
         'If InStr(cPrintORI, Text7(2).Text) > 0 And TransDate(cp(5), 2) >= FCP案件命名啟用日 And strAutoPrinter <> "" Then
         '   If txtCP118 = "Y" Then
         '       intAuto = IIf(Text7(2) = "103", 0, 1)
         '       If PUB_ChkCPExist(pa, "201") = False Then
         '           intAuto = 0
         '       End If
         '   Else
         '       intAuto = IIf(Text7(2) = "103", 1, 2)
         '       If PUB_ChkCPExist(pa, "201") = False Then
         '           intAuto = 1
         '       End If
         '   End If
         '   'Modified by Lydia 2018/08/10 電子送件已不列印ORI,
         '   'If Val(cp(82)) > 0 And intAuto > 0 Then '有發文時間=重新發文
         '   If txtCP118 <> "Y" And Val(cp(82)) > 0 And intAuto > 0 Then
         '       If MsgBox("重新發文是否重新列印說明書？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
         '             intAuto = 0
         '       End If
         '   End If
         '   'Added by Lydia 2018/08/10 翻譯分案上線後,可少印新案翻譯一份 (紙本送件印2份或電子送件印1份)
         '   'Memo by Lydia 2018/08/15 從下方移上來
         '   If intAuto > 1 Or (txtCP118 = "Y" And intAuto > 0) Then intAuto = intAuto - 1
         '
         '   '抓案件的最新一道ori.pdf、ori.repX.pdf或ori.fix.pdf檔案,有序列表(.SEQ.)一併列印
         '   If intAuto > 0 Then
         '       'Modified by Lydia 2018/10/05 ORI.FIX=>改成FIX.ORI
         '       strExc(0) = "SELECT CPP01,CPP02,CPP14 FROM CASEPROGRESS A,CASEPAPERPDF B " & _
         '                         "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP159=0 AND CP09=CPP01(+) " & _
         '                         "AND NVL(CPP10,'N') <> 'D' AND (UPPER(CPP02) LIKE '%.ORI.PDF' OR UPPER(CPP02) LIKE '%.ORI.REP%.PDF' OR UPPER(CPP02) LIKE '%.FIX%.ORI.PDF' OR UPPER(CPP02) LIKE '%.SEQ.%') " & _
         '                         "ORDER BY CPP06 DESC, CPP07 DESC "
         '       intI = 1
         '       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '       If intI = 0 Then
         '            'Modified by Lydia 2018/08/06紙本新案送件，若是僅以中說提申，因有管制無ORI則無法發文，則發文會有問題，故若是發文時無ORI，則彈訊息: 是否僅以中說提申，若"是"則讓他過， 若"否"則會回前畫面，不可發文。 by敏莉
         '            'MsgBox "卷宗區無說明書可供列印！", vbCritical
         '            If MsgBox("卷宗區無原文說明書，本案是否僅以中說提申？", vbYesNo + vbDefaultButton2) = vbYes Then
         '                 intAuto = 0
         '            Else
         '                 Exit Sub
         '            End If 'end 2018/08/06
         '       Else
         '            RsTemp.MoveFirst
         '            Do While Not RsTemp.EOF
         '                 If "" & RsTemp.Fields("CPP01") <> "" And "" & RsTemp.Fields("CPP02") <> "" And "" & RsTemp.Fields("CPP14") <> "" Then
         '                     '說明書
         '                     strExc(1) = m_AttchPath & "\" & RsTemp.Fields("CPP02")
         '                     'Modified by Lydia 2018/10/05 ORI.FIX=>改成FIX ; ORI.REP => 改成REP
         '                     If InStr(UCase("" & RsTemp.Fields("CPP02")), ".ORI.") > 0 And InStr(UCase(strAutoList), ".ORI.PDF") = 0 And _
         '                                  InStr(UCase(strAutoList), ".REP") = 0 And InStr(UCase(strAutoList), ".FIX") = 0 Then
         '                            'Added by Lydia 2018/08/09 翻譯分案上線後,可少印新案翻譯一份
         '                            'Remark by Lydia 2018/08/15 改從份數控制
         '                            'If intAuto = 1 Then
         '                            '       strAutoList = strAutoList & RsTemp.Fields("CPP02") & "&"
         '                            'Else
         '                            'end 2018/08/09
         '                               If PUB_GetFtpFile("" & RsTemp.Fields("CPP14"), m_AttchPath & "\" & RsTemp.Fields("CPP02")) = True Then
         '                                   strAutoList = strAutoList & RsTemp.Fields("CPP02") & "&"
         '                               End If
         '                            'End If 'end 2018/08/09
         '                     End If
         '                     '序列表
         '                     If InStr(UCase("" & RsTemp.Fields("CPP02")), ".SEQ.") > 0 And InStr(UCase(strAutoList), ".SEQ.") = 0 Then
         '                            'Added by Lydia 2018/08/09 翻譯分案上線後,可少印新案翻譯一份
         '                            'Remark by Lydia 2018/08/15 改從份數控制
         '                            'If intAuto = 1 Then
         '                            '       strAutoList = strAutoList & RsTemp.Fields("CPP02") & "&"
         '                           ' Else
         '                            'end 2018/08/09
         '                                   If PUB_GetFtpFile("" & RsTemp.Fields("CPP14"), m_AttchPath & "\" & RsTemp.Fields("CPP02")) = True Then
         '                                       strAutoList = strAutoList & RsTemp.Fields("CPP02") & "&"
         '                                   End If
         '                            'End If 'end 2018/08/09
         '                     End If
         '                 End If
         '                 RsTemp.MoveNext
         '            Loop
         '            If InStr(UCase(strAutoList), ".ORI.") = 0 Then
         '                  'Modified by Lydia 2018/08/06紙本新案送件，若是僅以中說提申，因有管制無ORI則無法發文，則發文會有問題，故若是發文時無ORI，則彈訊息: 是否僅以中說提申，若"是"則讓他過， 若"否"則會回前畫面，不可發文。 by敏莉
         '                  'MsgBox "卷宗區無說明書可供列印！", vbCritical
         '                  If MsgBox("卷宗區無原文說明書，本案是否僅以中說提申？", vbYesNo + vbDefaultButton2) = vbYes Then
         '                        intAuto = 0
         '                  Else
         '                        Exit Sub
         '                  End If
         '                  'end 2018/08/06
         '                  '刪除暫存檔
         '                   If Dir(m_AttchPath & "\" & pa(1) & "*" & Val(pa(2)) & "*.*") <> "" Then
         '                         Kill m_AttchPath & "\" & pa(1) & "*" & Val(pa(2)) & "*.*"
         '                   End If
         '                  Exit Sub
         '            End If
         '       End If
         '   End If
         'End If
         ''end 2018/03/27
         'end 2023/03/03
         
         'Added by Lydia 2018/04/16 檢查完畢，更新備註欄位
         Text7(17).Text = strNewCP64

         mDate209210 = "" 'Added by Lydia 2019/01/17
         mDateTF30 = "" 'Added by Lydia 2019/12/11
         
      '*************************
      'Modify By Sindy 2021/7/21 從FormSave移出來外層詢問,及改新增下一程序不存放行事曆
      '*************************
         'Added by Lydia 2019/01/02 新案101,102發文時若檢視中說209 or核對中說格式235 未發文，請設行事曆
         If Text7(2) = "101" Or Text7(2) = "102" Then
            strExc(0) = "select count(*) cnt from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' " & _
                              "and cp10 in ('209','235') and cp158=0 and cp159=0 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Val("" & RsTemp.Fields("cnt")) > 0 Then
                   If MsgBox("是否管制客戶提供中說期限？" & vbCrLf & "選是：產生下一程序" & vbCrLf & "選否：不產生下一程序，繼續發文", vbYesNo + vbInformation + vbDefaultButton1) = vbYes Then
JumpReInput:
                      'Modified by Lydia 2019/01/09 +自訂日期
                      strExc(2) = UCase(InputBox("管制期限為提申後(新案發文日)＋１～３個月，" & vbCrLf & "請在下方輸入1~3或自訂日期(民國年月日)：", "管制客戶提供中說期限", "2"))
                      If strExc(2) = "" Then
                          'Modified by Lydia 2019/01/09 +自訂日期
                          If MsgBox("未輸入1~3或自訂日期，是否要管制客戶提供中說期限？", vbYesNo + vbDefaultButton1) = vbYes Then
                               GoTo JumpReInput
                          End If
                      Else
                          'Modified by Lydia 2019/01/09 +自訂日期
                          'If InStr("1,2,3", strExc(2)) = 0 And Val(strExc(2)) = 0 Then
                          '    MsgBox "請輸入管制期限為提申後(新案發文日)＋１～３個月！", vbCritical
                          If (Len(strExc(2)) = 1 And InStr("1,2,3", strExc(2)) = 0) Or Val(strExc(2)) = 0 Or (Len(strExc(2)) > 1 And Len(strExc(2)) <> 7) Then
                              MsgBox "請輸入管制期限為提申後(新案發文日)＋１～３個月或自訂日期(民國年月日)！", vbCritical
                              GoTo JumpReInput
                          End If
                          'Added by Lydia 2019/01/09 檢查日期
                          If Len(strExc(2)) = 7 Then
                              If ChkDate(strExc(2)) = False Then
                                  GoTo JumpReInput
                              End If
                          End If
                          'end 2019/01/09
                      End If
                      If strExc(2) <> "" Then
                          '新案發文日起算X個月,日期若碰到放假則往前抓1個工作天。
                          'Added by Lydia 2019/01/09 自訂日期
                          If Len(strExc(2)) > 1 Then
                              strExc(1) = CompWorkDay(1, DBDATE(strExc(2)), 1)
                          Else
                          'end 2019/01/09
                              strExc(1) = CompWorkDay(1, CompDate(1, strExc(2), DBDATE(Text7(0))), 1)
                          End If
                          'end 2019/01/09
                          mDate209210 = strExc(1) 'Added by Lydia 2019/01/17
                          'Modify By Sindy 2021/7/23 皆不再產生相關之行事曆
'                          strExc(4) = "催客戶提供中說期限"
'                          strExc(3) = PUB_GetFCPHandler(cp(1), cp(2), cp(3), cp(4))
'                          strExc(5) = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
'                          If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(3) & "," & strExc(5), strExc(4), strExc(3) & "," & strExc(5), "1", cp(1), cp(2), cp(3), cp(4)) Then
'                          End If
                      End If
                   End If
               End If
            End If
         End If
         'end 2019/01/02
         'Added by Lydia 2019/12/11  FCP新案發文時檢查有新案翻譯未發文並且尚"待英文本翻譯"TF30='Y'，
         '則彈訊息"是否管制催客戶提供英文翻譯本之行事曆"並比照催客戶提供英文翻譯本的行事曆方式，其內容為" 催客戶提供英文翻譯本"。
         If InStr(NewCasePtyList, Text7(2)) > 0 Then
            strExc(0) = "select cp09,tf30 from caseprogress,transfee where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' " & _
                              "and cp10='201' and cp158=0 and cp159=0 and cp09=tf01(+)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If "" & RsTemp.Fields("tf30") = "Y" Then
                   If MsgBox("是否管制催客戶提供英文翻譯本之下一程序？" & vbCrLf & "選是：產生下一程序" & vbCrLf & "選否：不產生下一程序，繼續發文", vbYesNo + vbInformation + vbDefaultButton1) = vbYes Then
JumpReInput2:
                      strExc(2) = UCase(InputBox("管制期限為提申後(新案發文日)＋１～３個月，" & vbCrLf & "請在下方輸入1~3或自訂日期(民國年月日)：", "催客戶提供英文翻譯本", "2"))
                      If strExc(2) = "" Then
                          If MsgBox("未輸入1~3或自訂日期，是否要管制催客戶提供英文翻譯本期限？", vbYesNo + vbDefaultButton1) = vbYes Then
                               GoTo JumpReInput2
                          End If
                      Else
                          If (Len(strExc(2)) = 1 And InStr("1,2,3", strExc(2)) = 0) Or Val(strExc(2)) = 0 Or (Len(strExc(2)) > 1 And Len(strExc(2)) <> 7) Then
                              MsgBox "請輸入管制期限為提申後(新案發文日)＋１～３個月或自訂日期(民國年月日)！", vbCritical
                              GoTo JumpReInput2
                          End If
                          If Len(strExc(2)) = 7 Then
                              If ChkDate(strExc(2)) = False Then
                                  GoTo JumpReInput2
                              End If
                          End If
                      End If
                      If strExc(2) <> "" Then
                         '新案發文日起算X個月,日期若碰到放假則往前抓1個工作天。
                          If Len(strExc(2)) > 1 Then  '自訂日期
                              strExc(1) = CompWorkDay(1, DBDATE(strExc(2)), 1)
                          Else
                              strExc(1) = CompWorkDay(1, CompDate(1, strExc(2), DBDATE(Text7(0))), 1)
                          End If
                          mDateTF30 = strExc(1)
                          'Modify By Sindy 2021/7/23 皆不再產生相關之行事曆
'                          strExc(4) = "催客戶提供英文翻譯本"
'                          strExc(3) = PUB_GetFCPHandler(cp(1), cp(2), cp(3), cp(4))
'                          strExc(5) = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
'                          If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(3) & "," & strExc(5), strExc(4), strExc(3) & "," & strExc(5), "1", cp(1), cp(2), cp(3), cp(4)) Then
'                          End If
                      End If
                   End If
               End If
            End If
         End If
      '2021/7/21 END
      '*************************
         
         'Add By Sindy 2023/3/27 所有設定皆不變，只有修改判斷若有退費先詢問
         m_bolIns908 = False
         'Modify By Sindy 2025/7/22 增加判斷進度備註中有存在"本案提申時未附英文說明書，但所檢附之申請書中發明名稱、申請人姓名或名稱..."此段文字時,詢問"是否內部收文代辦退費？"
         If Val(txtDecreasePageFee.Text) > 0 Or Val(txtDecreaseItemFee.Text) > 0 _
            Or InStr(Text7(17), "本案提申時未附英文說明書，但所檢附之申請書中發明名稱、申請人姓名或名稱、發明人姓名及摘要同時附有英文翻譯，故可減收申請規費800") > 0 Then
            If MsgBox("是否內部收文代辦退費？" & vbCrLf & "選是：則收文代辦退費並上發文日" & vbCrLf & "選否：則不收文代辦退費", vbYesNo + vbInformation + vbDefaultButton1) = vbYes Then
               m_bolIns908 = True
            End If
         End If
         '2023/3/27 END
         
         'Add by Sindy 2021/11/15 檢查畫面上的物件是否含有Unicode文字
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         
         ' 設定滑鼠游標為等待狀態
         Screen.MousePointer = vbHourglass

         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
         
'cancel by sonia 2018/3/6 因要放正確版本,程式先不上傳由程序人工做
'        'Added by Lydia 2017/12/13 FCP案件命名電子化：搬移外文原文本PDF檔到卷宗區
'          If strSrvDate(1) >= FCP案件命名啟用日 And cp(1) = "FCP" And InStr(NewCasePtyList, Text7(2)) > 0 Then
'              'Added by Lydia 2018/03/05 排除無權限的人員
'              If Pub_StrUserSt03 <> "M51" And Left(Pub_StrUserSt03, 1) <> "F" Then
'                  MsgBox "非國外部人員無權限進入\\English_Vers，請手動上傳下列資料夾中的*" & FcpTcnFKey02 & "檔" & vbCrLf & "路徑: " & Pub_GetFCPcaseFilePath(pa(2))
'              Else
'              'end 2018/03/05
'                  'Modified by Lydia 2017/12/28
'                  'strDefDir = FCP命名追蹤收文區 & "\" & Mid(Val(cp(2)), 1, 3) & "\" & Val(cp(2)) '預設檔案路徑
'                  strDefDir = Pub_GetFCPcaseFilePath(cp(2))
'                  strExc(0) = Dir(strDefDir & "\*" & FcpTcnFKey02)
'                  If strExc(0) = "" Then
'                     '若缺檔案,發mail通知承辦業務和主管
'                     If PUB_GetTCTmail(True, 5, cp(1), cp(2), cp(3), cp(4), cp(9)) Then
'                     End If
'                  Else
'                     '只上傳一次
'                           strTmp = Format(ServerTime, "000000") '時間
'                           strNewFileName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & Trim(Text7(2)) & ".ForeignSpec" & FcpTcnFKey02
'                           If SaveAttFile_PDF(cp(9), strDefDir & "\" & strExc(0), strNewFileName, Val(strSrvDate(1)), Val(strTmp), False) Then
'                                If PUB_ChkFileOpening(strDefDir & "\" & strExc(0)) = True Then
'                                     MsgBox strDefDir & "\" & strExc(0) & vbCrLf & "檔案正在使用中，請關閉才可刪除！", vbExclamation
'                                Else
'                                    If Pub_FtpDelTyping2(Me.Name, strDefDir, strExc(0)) = True Then
'                                    End If
'                                End If
'                           Else
'                               MsgBox "上傳失敗:" & strDefDir & "\" & strExc(0)
'                           End If
'                  End If
'              End If  'end 2018/03/05
'          End If
'          'end 2017/12/13
'end 2018/3/6
          
          'Added by Lydia 2018/03/01 中說發文刪除english_vers案號資料夾的*.msg
'Remove by Lydia 2018/03/07 改到批次
'          If strSrvDate(1) >= FCP案件命名啟用日 And cp(1) = "FCP" And InStr("201,209,210,235", Text7(2)) > 0 Then
'              'Added by Lydia 2018/03/05 排除無權限的人員
'              If Pub_StrUserSt03 <> "M51" And Left(Pub_StrUserSt03, 1) <> "F" Then
'                  MsgBox "非國外部人員無權限進入\\English_Vers，請手動刪除下列資料夾中的*" & FcpTcnFKey01 & "檔" & vbCrLf & "路徑: " & Pub_GetFCPcaseFilePath(pa(2))
'              Else
'              'end 2018/03/05
'                strDefDir = Pub_GetFCPcaseFilePath(cp(2))
'                strExc(0) = Dir(strDefDir & "\*" & FcpTcnFKey01)
'                Do While strExc(0) <> ""
'                    If PUB_ChkFileOpening(strDefDir & "\" & strExc(0)) = True Then
'                         MsgBox strDefDir & "\" & strExc(0) & vbCrLf & "檔案正在使用中，請關閉才可刪除！", vbExclamation
'                         Exit Do
'                    Else
'                         If Pub_FtpDelTyping2(Me.Name, strDefDir, strExc(0)) = False Then
'                             Exit Do
'                         End If
'                    End If
'                    strExc(0) = Dir(strDefDir & "\*" & FcpTcnFKey01)
'                Loop
'              End If 'end 2018/03/05
'          End If
          'end 2018/03/01
'end 2018/03/07
         
         'Added by Lydia 2018/03/27 有自動列印說明書改用機密列印
         'Mark by Lydia 2023/03/03 與Sharon,Phoebe 確認不用列印說明書
         'If intAuto > 0 And strAutoList <> "" And strAutoPrinter <> "" Then
         '     PUB_RestorePrinter strAutoPrinter
         '     PUB_SetOsDefaultPrinter strAutoPrinter
         'End If
         ''end 2018/03/27
         
         'Add by Amy 2014/05/27 案件性質為101,102,103,125 +印簡易聯絡單
         'Modify by Amy 2016/04/29 +傳案件性質
         If InStr("101,102,103,125", Text7(2)) > 0 Then
            'Modified by Lydia 2023/05/19 改模組
            'm_strContactSheetA4 = PrintContactSheetA4(Label3(0), Text1, Text2, Text3, Text4, Text7(2))
            m_strContactSheetA4 = PUB_FCPPrintContactSheetA4(True, Label3(0), Text1, Text2, Text3, Text4, Text7(2), , mDate209210, mDateTF30)
         End If
         'end 2014/05/27
         
         'Added by Lydia 2018/03/27 自動列印說明書
         'Mark by Lydia 2023/03/03 與Sharon,Phoebe 確認不用列印說明書
         'If intAuto > 0 And strAutoList <> "" And strAutoPrinter <> "" Then
         '     Call PrintFileList(strAutoList, intAuto)
         '     PUB_RestorePrinter strPrinter '還原預設印表機
         '     PUB_SetOsDefaultPrinter strPrinter
         'End If
         ''end 2018/03/27
         
         If Text7(2) = 加註追加 Or Text7(2) = 加註聯合 Or Me.Text7(2).Text = 加註專用權延長 Then
            EndLetter "01", strReceiveNo, "00", strUserNum
            NowPrint strReceiveNo, "01", "00", False, strUserNum, 0
         End If
         
         'Add by Morgan 2009/10/6
         '退審查費要出通知信函
         If Text7(2) = 退費 And cp(43) <> "" Then
            'Modified by Morgan 2013/10/21 +435
            strExc(0) = "select 1 from caseprogress where cp09='" & cp(43) & "' and cp10 in ('416','435','107') and cp27>0"
            'Added by Morgan 2013/6/6
            strExc(0) = strExc(0) & " union select 2 from  caseprogress,nextprogress where cp09='" & cp(43) & "' and cp10='404' and np01(+)=cp43 and np07='107'"
            strExc(0) = strExc(0) & " union select 3 from  caseprogress a,caseprogress b where a.cp09='" & cp(43) & "' and a.cp10='404' and b.cp09(+)=a.cp43 and b.cp10='107'"
         
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Modified by Morgan 2014/2/7 +日文定稿;加選擇是否以CreditNote退費出不同定稿
               bolCreditNote = False
               If PUB_GetLanguage(cp(1), cp(2), cp(3), cp(4)) = "3" Then
                  strET03 = "02"
               Else
                  strET03 = "01"
               End If
               If MsgBox("是否以Credit Note方式退費", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                  bolCreditNote = True
               End If
               'end 2014/2/7
               
               StartLetter "02", strReceiveNo, strET03
               bolEmail = PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4), , , bolPlusPaper)
               'Add by Morgan 2009/10/20 +判斷是否EMail同時寄紙本
               If bolPlusPaper Then
                  iCopy = 0
               Else
                  iCopy = 1
               End If
               'end 2009/10/20
               If bolEmail Then
                  NowPrint strReceiveNo, "02", strET03, False, strUserNum, , , , , iCopy, , True, True
                  MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(pa(1)) & " ]！"
               Else
                  NowPrint strReceiveNo, "02", strET03, False, strUserNum
                  MsgBox "報告函已產生!!"
               End If
            End If
         End If
         
         'Add by Morgan 2010/11/12
         If Text7(2) = "202" Then
            strUserNum = strFMPNum
            StartLetter2 "02", pa(1) & pa(2) & pa(3) & pa(4) & "&202", "01"
            NowPrint pa(1) & pa(2) & pa(3) & pa(4) & "&202", "02", "01", False, strUserNum, 0
            strUserNum = strUser1Num
         End If
         'end 2010/11/12
   
         'Add by Morgan 2008/8/21
         '主動修正發文時若申請程序已請款時提醒
         If Text7(2) = "203" Then
            'Modify by Morgan 2010/4/12 +工程師提申 940,105
            'Modified by Morgan 2012/12/20 +衍生設計125
            strExc(0) = "select cp60 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in ('101','102','103','105','125','940') and cp60 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "此道主動修正程序將由工程師報告請款！", vbInformation
            End If
         End If
         
         If Text7(2) = "201" And cp(60) <> "" Then
            '2008/12/1 modify by sonia FCP-038136 加實審已發文條件
            'MsgBox "若有超頁請收文 938 超頁！", vbInformation
            strExc(0) = "select cp60 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='416' and cp27 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Modified by Morgan 2015/8/14
               'MsgBox "若有超頁請收文 938 超頁！", vbInformation
               MsgBox "本程序已請款若有超頁超項費請確認請款金額是否正確！", vbInformation
            End If
            '2008/12/1 end
         End If
         '2009/4/29 ADD BY SONIA 檢視中說發文若已請款則提醒
         '2011/6/27 modify by sonia 加入pa08條件 FCP-043622
         If pa(8) = "1" And Text7(2) = "209" And cp(60) <> "" Then
            'Modified by Morgan 2013/3/25 +實審已發文條件 FCP-046773
            'MsgBox "若有超頁請收文 938 超頁！", vbInformation
            strExc(0) = "select cp60 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='416' and cp27 is not null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Modified by Morgan 2015/8/14
               'MsgBox "若有超頁請收文 938 超頁！", vbInformation
               MsgBox "本程序已請款若有超頁超項費請確認請款金額是否正確！", vbInformation
            End If
            'end 2013/3/25
         End If
         '2009/4/29 END
         'end 2008/8/21
         
         If pa(1) = "FCP" Then 'Added by Morgan 2015/7/6
'Modified by Morgan 2020/3/3 改呼叫共用
'            'Add By Sindy 2016/7/7 + 代理人為Y4829203Hewlett-Packard Company Intellectual Property Administration
'            '承辦人為工程師(ST03 IN ('F21','F51','F52))時,於存檔後彈訊息
'            'If Text7(2) <> "926" Then  'Modify By Sindy 2016/7/11 926.核對已准專利除外
'               If ChangeCustomerL(pa(75)) = "Y48292030" And _
'                  (PUB_GetST03(Text7(1).Text) = "F21" Or PUB_GetST03(Text7(1).Text) = "F51" Or PUB_GetST03(Text7(1).Text) = "F52") Then
'                  'Add By Sindy 2016/7/18
'                  strExc(0) = "select cp130 from caseprogress where cp09='" & strReceiveNo & "'"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     If "" & RsTemp.Fields(0) <> "" Then
'                  '2016/7/18 END
'                        MsgBox "請優先請款並且在提申當天上傳報告!!"
'                     End If
'                  End If
'               End If
'            'End If
'            '2016/7/7 END
'
'            'Add By Sindy 2016/10/17 凡代理人Y33844   KLARQUIST SPARKMAN, LLP的案件，
'            '若是工程師中間程序(例: 申復、再審、訴願、補充說明、...)發文時，
'            '彈訊息"請在送件後3天內並且要當月優先請款"，請排除901.告代、902.回代、1202.審查意見、1002.核駁.....。
'            If (PUB_GetST03(Text7(1).Text) = "F21" Or PUB_GetST03(Text7(1).Text) = "F51" Or PUB_GetST03(Text7(1).Text) = "F52") And _
'               Not (m_CP10 = "901" And m_CP10 = "902" And m_CP10 = "1202" And m_CP10 = "1002") Then
'               'Memo by Morgan 2019/9/17 有發文主管機關的才算中間程序送件(如核對926就不算)--敏莉
'               strExc(0) = "select cp130 from caseprogress where cp09='" & strReceiveNo & "'"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  If "" & RsTemp.Fields(0) <> "" Then
'                     If ChangeCustomerL(pa(75)) = "Y33844000" Then
'                        MsgBox "請在送件後3天內並且要當月優先請款!!"
'                     'Add By Sindy 2016/10/20
'                     ElseIf ChangeCustomerL(pa(75)) = "Y51982000" Then
'                        MsgBox "申復/再審/修正等送智慧局的案件，於收到指示後7天內送程序送件同時請款(由Wilson指示備註)!!"
'                     'Add By Sindy 2016/10/24
'                     ElseIf ChangeCustomerL(pa(75)) = "Y20272000" Then
'                        MsgBox "中間程序送件當日簡單報告!!"
'                     'Add By Sindy 2016/11/16
'                     ElseIf ChangeCustomerL(pa(75)) = "Y34440B30" Then
'                        MsgBox "請當日優先請款報告!!"
'                     'Add By Sindy 2017/3/20 Y30053010 HARNESS, DICKEY & PIERCE, PLC的案件
'                     'Add By Sindy 2017/9/12 + Y54171 Boston Biomedical, Inc.
'                     'Add By Sindy 2018/5/17 + Y54682 Snyder, Clark, Lesch & Chung,LLP
'                     ElseIf InStr("Y30053010,Y54171000,Y54682000", ChangeCustomerL(pa(75))) > 0 Then
'                        MsgBox "請當天簡單報告!!"
'
'                     'Added by Morgan 2019/9/17 --Bobbie
'                     ElseIf ChangeCustomerL(pa(75)) = "Y20088000" Then
'                        MsgBox "中間程序送件後、法限前報告+請款。如無法，請先退承辦當天簡單報告!!"
'
'                     ElseIf ChangeCustomerL(pa(75)) = "Y21042000" Then
'                        MsgBox "所有程序送件後退承辦簡單報告!!"
'                     'end 2019/9/17
'                     End If
'                  End If
'               End If
'            End If
'            '2016/10/17 END
'
'            'Added by Morgan 2019/9/17 --Bobbie
'            If m_CP10 = "416" Or m_CP10 = "202" Then
'               If ChangeCustomerL(pa(75)) = "Y20088000" Then
'                  MsgBox "中間程序送件後、法限前報告+請款。如無法，請先退承辦當天簡單報告!!"
'
'               ElseIf ChangeCustomerL(pa(75)) = "Y21042000" Then
'                  MsgBox "所有程序送件後退承辦簡單報告!!"
'               End If
'            End If
'            'end 2019/9/17
            PUB_FCPAlert strReceiveNo
'end 2020/3/3
            
            'Added by Lydia 2015/06/04 檢查非日本案在翻譯中說(含核對中說格式、檢視中說、製作中說)發文時，若有收文主動修正203,但工程師沒有輸入修正內容,則請彈訊息
            'Remove by Lydia 2021/05/13 因流程已有改變，故請刪除彈提醒
'             strExc(10) = GetPrjNationNumber(ChangeCustomerL(pa(75)))
'            If InStr("201,209,210,235", Text7(2)) > 0 And Left(strExc(10), 3) <> "011" Then
'               strExc(0) = "select cp05,cp09,cp27 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' " & _
'                           "and cp10='203' and cp57 is null order by 3,1 "
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                   strExc(0) = "" & RsTemp.Fields("CP27") '最後發文的203之發文日
'                   strExc(1) = "select amd05 from amendedtext where amd01='" & pa(1) & "' and amd02='" & pa(2) & "' and amd03='" & pa(3) & "' and amd04='" & pa(4) & "' "
'                   intI = 1
'                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
'                   If intI = 1 Then
'                      strExc(1) = "" & RsTemp.Fields(0)
'                   Else
'                      strExc(1) = ""
'                   End If
'                   '美國區101,美洲區1xx,亞洲區0xx(韓國012除外)時,若最後發文203主動修正之發文日<=申請日則不彈訊息
'                   If Left(strExc(10), 3) = "101" Or Left(strExc(10), 1) = "1" Or (Left(strExc(10), 1) = "0" And Left(strExc(10), 3) <> "012") Then
'                      If strExc(0) <= DBDATE(pa(10)) Then
'                         strExc(3) = "0"
'                      Else
'                         strExc(3) = "1"
'                      End If
'                   Else
'                         strExc(3) = "1"
'                   End If
'                   '沒有輸入修正內容
'                   If strExc(1) = "" And strExc(3) = "1" Then MsgBox "請退工程師輸入修正內容!", vbExclamation
'               End If
'            End If
'            'end 2015/06/04
            'end 2021/05/12
         End If 'Added by Morgan 2015/7/6
         
         'Add by Morgan 2008/2/20 檢查代理人Email
         If pa(1) = "FCP" Then
            PUB_CheckEMail pa(75), pa(144)
            If pa(145) <> "" Then
               PUB_CheckEMail pa(75), pa(145)
            End If
         Else
            PUB_CheckEMail pa(26), pa(76)
            If pa(77) <> "" Then
               PUB_CheckEMail pa(26), pa(77)
            End If
         End If
         'end 2008/2/20
            
         'Add by Morgan 2005/12/6
         '發E-Mail給承辦人
         If m_strMailCP09 <> "" Then
             MailToPromoter m_strMailCP09
         End If
         
         'Added by Lydia 2018/10/24 FCP設計案在發文(210)製作中說時，若工程師無上傳DES至專利案件則發mail給工程師。
         If pa(8) = "3" And cp(10) = "210" Then
             strExc(1) = ""
             'Modified by Lydia 2020/09/08 檔案改放在原始檔區
             'strExc(2) = Dir("\\Typing2\專利案件\" & Left(Val(pa(2)), 3) & "\" & pa(1) & "*" & Val(pa(2)) & ".des.*")
             'Do While strExc(2) <> ""
             '    If InStr(".DES.PDF;.DES.DOC", Right(UCase(strExc(2)), 8)) > 0 Or InStr(".DES.DOCX", Right(UCase(strExc(2)), 9)) > 0 Then
             '        strExc(1) = strExc(1) & "," & strExc(2)
             '    End If
             '    strExc(2) = Dir()
             'Loop
             'Modify By Sindy 2025/11/10 + ) OR CP09='" & strReceiveNo & "')
             strExc(0) = "SELECT CP01,CP02,CP03,CP04,CP09,NVL(A02,0) CNT1,NVL(B02,0) CNT2 " & _
                         "FROM CASEPROGRESS,(SELECT CPF01 A01,COUNT(*) A02 FROM CASEPAPERFILE WHERE UPPER(CPF02) LIKE '%.DES.DOC%' GROUP BY CPF01) VT01 " & _
                         ",(SELECT CPF01 B01,COUNT(*) B02 FROM CASEPAPERFILE WHERE UPPER(CPF02) LIKE '%.DES.PDF' GROUP BY CPF01) VT02 " & _
                         "WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' " & _
                         "AND ((SUBSTR(CP09,1,1)='D' AND CP10='" & cnt專利案件 & "') OR CP09='" & strReceiveNo & "') AND CP159=0 AND CP09=A01(+) AND CP09=B01(+) "
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
             If intI = 1 Then
                If Val("" & RsTemp.Fields("cnt1")) + Val("" & RsTemp.Fields("cnt2")) > 0 Then
                    strExc(1) = "Y"
                End If
             End If
             'end 2020/09/08
             If strExc(1) = "" Then
                  '發mail給承辦工程師，CC: 工程師主管; 承辦管制人員; 程序管制人員
                  'Modified by Lydia 2019/01/09
                  'Select Case pa(150)
                  '     Case "1": strExc(3) = Pub_GetSpecMan("T")
                  '     Case "2": strExc(3) = Pub_GetSpecMan("R")
                  '     Case "3": strExc(3) = Pub_GetSpecMan("S")
                  '     Case "4": strExc(3) = Pub_GetSpecMan("T1")
                  '     Case Else: strExc(3) = ""
                  'End Select
                  'Memo by Lydia 2019/08/19 日文組副本:除各組主任(99034,94012)給主管,其餘人給審核主管
                  strExc(3) = m_GrpMan
                  
                  strExc(4) = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
                  strExc(3) = strExc(3) & IIf(strExc(3) <> "", ";", "") & strExc(4)
                  strExc(4) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
                  strExc(3) = strExc(3) & IIf(strExc(3) <> "", ";", "") & strExc(4)
                  strExc(0) = pa(1) & "-" & pa(2) & "請工程師自行上傳設計案之說明書及圖式至專利案件，以供承辦請款時寄給代理人，謝謝。"
                  PUB_SendMail strUserNum, Text7(1).Text, "", strExc(0), "同主旨", "", "", , , , strExc(3)
             End If
         End If
         'end 2018/10/24
         
         'Add By Sindy 2017/3/20 實審發文
         'Modify By Sindy 2017/3/27 檢查同發文日是否有其他道發文,若無,才詢問是否要出帳單
         'Modified by Lydia 2021/01/12 另外分成模組
'         If cp(10) = "416" Then
'            strExc(1) = "select cp09 from caseprogress" & _
'                        " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
'                        " and cp27=" & DBDATE(Text7(0)) & " and cp57 is null" & _
'                        " and cp09<>'" & Label3(0) & "' and cp43<>'" & Label3(0) & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
'            If intI = 0 Then
'               If MsgBox("實審是否要產生帳單？", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
'                  'Add By Sindy 2017/3/20 實審發文直接產生承辦單、請款信及帳單
'                  '列印FCP承辦單
'                  'Modified by Morgan 2017/10/27
'                  'Call PUB_PrintFCPEmpBill(pa(1), pa(2), pa(3), pa(4), "08", , 實體審查, "實體審查請款")
'                  'Modified by Lydia 2019/01/04 傳入實審收文號
'                  'Call PUB_PrintFCPEmpBill(pa(1), pa(2), pa(3), pa(4), "09", , 實體審查, "實體審查請款")
'                  'Modified by Lydia 2019/03/04 更換類別代號;09=>05
'                  'Modified by Lydia 2020/07/09 傳入備註
'                  'Call PUB_PrintFCPEmpBill(pa(1), pa(2), pa(3), pa(4), "05", Label3(0).Caption, 實體審查, "實體審查請款")
'                  strExc(1) = ""
'                  If txtPAID = "1" Then strExc(1) = "D/N已會財務處。"  '已收款不寄D/N
'                  If txtPAID = "2" Then strExc(1) = "D/N已加蓋PAID章，D/N已會財務處。"    '已收款寄D/N
'                  Call PUB_PrintFCPEmpBill(pa(1), pa(2), pa(3), pa(4), "05", Label3(0).Caption, 實體審查, "實體審查請款", , strExc(1))
'                  'end 2020/07/09
'
'                  Dim nFrm As Form
'                  '檢查表單是否已開啟，若是，則關閉
'                  For Each nFrm In Forms
'                     If StrComp(nFrm.Name, "frm060306_7", vbTextCompare) = 0 Then
'                        Unload frm060306_7
'                        'Exit For
'                     End If
'                     If StrComp(nFrm.Name, "frm060306", vbTextCompare) = 0 Then
'                        Unload frm060306
'                        Exit For
'                     End If
'                  Next
'                  frm060306.Show
'                  frm060306.Text1.Text = pa(1)
'                  frm060306.Text2.Text = pa(2)
'                  frm060306.Text3.Text = pa(3)
'                  frm060306.Text4.Text = pa(4)
'                  frm060306.m_quy416 = True
'                  frm060306.Command1_Click
'                  If frm060306.MSHFlexGrid1.Rows >= 2 Then
'                     If frm060306.MSHFlexGrid1.TextMatrix(1, 2) <> "" Then
'                        frm060306.MSHFlexGrid1.TextMatrix(1, 0) = "v"
'                        Call frm060306.cmdok_Click(1)
'                        frm060306_7.Show
'                        If cp(60) = "" Then frm060306_7.Text1(1) = "Y" '要產生請款單
'                        'Added by Lydia 2020/08/17
'                        frm060306_7.txtPAID.Text = Me.txtPAID.Text ' 已收款
'                        frm060306_7.m_CallName = Me.Name  '呼叫的表單名稱
'                        'end 2020/08/17
'                        Call frm060306_7.cmdok_Click(0)
'                        Unload frm060306
'                     End If
'                  End If
'                  '2017/3/20 END
'               End If
'            End If
'         End If
'         '2017/3/27 END
         Call ProcDNfor416
         'end 2021/01/21
         
         ProcDNfor447 'Added by Morgan 2024/11/21
         
         strExc(10) = Pub_GetCP31toCP27(pa(1), pa(2), pa(3), pa(4)) 'Added by Lydia 2019/01/10 新申請案發文日

         'Added by Lydia 2019/01/07 主動修正203、修正204、誤譯訂正433和申復、再審發文(有一併修正)時，若工程師沒有上傳中說word檔最終版本，系統會自動發email給工程師提醒
         '若到了請款階段(203,204)還沒上傳，會再發一次通知提醒
         'Modified by Lydia 2019/01/28 + 主動修正發文時，若新案翻譯(含檢視中說、核對中說格式)為已發文且同日發文，則詢問「主動修正是否已併入中說送件」
         '若選擇「是」則不用檢查中說最終版是否存在，若選擇「否」否: 則依舊在主動修正發文或請款時檢查中說最終版。
         'If (InStr("203,204,433", cp(10)) > 0) Or (InStr("107,205", cp(10)) > 0 And m_strCP148 = "Y") Then
         If (InStr("107,205", cp(10)) > 0 And m_strCP148 = "Y") Or _
              (InStr("203,204,433", cp(10)) > 0 And m_strCP148 = "") Then
             If strExc(10) <> "" And ((cp(10) = "203" And DBDATE(Text7(0)) > strExc(10)) Or (cp(10) <> "203" And DBDATE(Text7(0)) >= strExc(10))) Then  'Added by Lydia 2019/01/10 判斷提申後才檢查
                'Modified by Lydia 2020/03/03 改成模組
                ''中說word檔最終版本=>案號-送件日.FIX_U
                'strExc(0) = "\\Typing2\專利案件\" & Left(Val(pa(2)), 3) & "\" & pa(1) & pa(2) & "-" & Text7(0) & ".fix_u.doc*"
                'strExc(1) = Dir(strExc(0))
                ''Modified by Lydia 2019/01/10 +圖式
                ''If strExc(1) = "" And Text7(1) <> "" Then
                'strExc(2) = "\\Typing2\專利案件\" & Left(Val(pa(2)), 3) & "\" & pa(1) & pa(2) & "-" & Text7(0) & ".fig.pdf"
                'strExc(3) = Dir(strExc(2))
                'If strExc(1) & strExc(3) = "" And Text7(1) <> "" Then
                ''end 2019/01/10
                'Modify By Sindy 2025/10/22 +, cp(9)
                If Pub_ChkFixUExists(pa(1), pa(2), pa(3), pa(4), Text7(0).Text, cp(9)) = False And Text7(1) <> "" Then
                'end 2020/03/03
                    'Memo by Lydia 2019/08/19 日文組副本:除各組主任(99034,94012)給主管,其餘人給審核主管
                    'Modified by Morgan 2024/3/5 機械組案件主旨都加【機械設計組】--Sharon
                    PUB_SendMail strUserNum, Text7(1), "", IIf(pa(150) = "4", "【機械設計組】", "") & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "未上傳中說Word檔最終版本，請儘速上傳！", vbCrLf & "同主旨", , , , , , IIf(Text7(1) <> m_GrpMan, m_GrpMan, "")
                End If
             End If 'end 2019/01/10 判斷提申後才檢查
         End If
         'end 2019/01/07
         
'         'Add By Sindy 2015/12/14 相關總收文號為機關來函且未發文時,發mail通知承辦工程師及主管
'         If cp(43) <> "" Then
'            If Left(cp(43), 1) = "C" Then
'               strExc(0) = "SELECT cp09,cp27,st15,st52,cp14,cp10,decode('" & pa(9) & "','000',cpm03,cpm04) cp10nm FROM CasePROGRESS,staff,casepropertymap" & _
'                           " WHERE CP09 = '" & cp(43) & "' and cp27 is null" & _
'                           " And cp14=st01(+)" & _
'                           " And cp01=cpm01(+) And cp10=cpm02(+)"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               strTo = ""
'               If intI = 1 Then
'                  '工程師
'                  If "" & RsTemp.Fields("st15") = "F21" Then
'                     strTo = RsTemp.Fields("cp14") & ";"
'                     If pa(150) = "" Then
'                        If InStr(strTo, Pub_GetSpecMan("N")) = 0 Then strTo = strTo & Pub_GetSpecMan("N")
'                     Else
'                        strTemp = IIf(pa(150) = "1", Pub_GetSpecMan("T"), IIf(pa(150) = "2", Pub_GetSpecMan("R"), IIf(pa(150) = "3", Pub_GetSpecMan("S"), Pub_GetSpecMan("T1"))))
'                        If InStr(strTo, strTemp) = 0 Then strTo = strTo & strTemp
'                     End If
'                  '程序人員
'                  Else
'                     strTo = RsTemp.Fields("cp14") & ";"
'                     If "" & RsTemp.Fields("st52") <> "" Then
'                        If InStr(strTo, RsTemp.Fields("st52")) = 0 Then strTo = strTo & RsTemp.Fields("st52")
'                     End If
'                  End If
'                  strSubject = "已屆期限，但OA尚未發文"
'                  strContent = "本所案號：" + pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) + vbCrLf + _
'                               "案件性質：" + RsTemp.Fields("cp10nm") + vbCrLf + vbCrLf + _
'                               "*本程序請儘速通知代理人/客戶" + vbCrLf
'                  PUB_SendMail strUserNum, strTo, "", strSubject, strContent
'               End If
'            End If
'         End If
'         '2015/12/14 END
         
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
         
         'Added by Lydia 2019/08/23 提申前主動修正發文彈提醒
         If Me.Text7(2) = "203" And (pa(8) = "1" Or pa(8) = "2") And Val(pa(10)) > 0 And DBDATE(pa(10)) = DBDATE(Text7(0)) Then
            ''主動修正發文與發明 , 新型申請發文日同一日則彈提醒:
            MsgBox "請更正修正後原文字數！", vbInformation, "提申前主動修正"
         End If
         If InStr("101,102", Me.Text7(2)) > 0 Then
              '發明 , 新型申請發文時, 若有主動修正已發文, 則彈提醒: 請更正修正後原文字數
              If PUB_ChkCPExist(cp, "203", 2) = True Then
                  MsgBox "請更正修正後原文字數！", vbInformation, "提申前主動修正"
              End If
         End If
         'end 2019/08/23
         
         'Added by Lydia 2022/04/28 核對已准專利發文確定後，檢查核對已准專利是否有請款單號或已上不請款，若有，請彈訊息
         If pa(1) = "FCP" And Me.Text7(2) = "926" Then
            If cp(60) > "X" Or Text7(13) = "N" Then '有請款單號 或 CP20=N
               MsgBox "核對已准專利已有請款單號或已上不請款，請直接寄二核報告信！", vbInformation
            'Add By Sindy 2022/5/9 若無彈訊息，則發Mail通知智權人員
            Else
               strTo = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
               strExc(0) = "SELECT st01,st52 FROM staff" & _
                           " WHERE st01='" & strTo & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(3) = "" & RsTemp.Fields("st52")
               End If
               strExc(3) = strExc(3) & IIf(strExc(3) <> "", ";", "") & strUserNum & ";backup"
               'Modified by Lydia 2022/09/23 請將(Murgitroyd案優先)移至主旨前面，以便承辦可以快速辨認(by Bobbie); 判斷改用模組
               'strExc(0) = "【已發文_核對已准專利】可進行請款 Our Ref: " & pa(1) & "-" & pa(2) & " [INCOM.926]" & IIf(Left(ChangeCustomerL(pa(75)), 8) = "Y2099001", "(Murgitroyd案)", "")
               strExc(0) = PUB_GetSetMailSubF2(pa(75)) & "【已發文_核對已准專利】可進行請款 Our Ref: " & pa(1) & "-" & pa(2) & " [INCOM.926]"
               strExc(10) = "1.程序已上發文【核對已准專利】" & vbCrLf & _
                            "2.請承辦同仁處理請款，謝謝您！"
               PUB_SendMail strUserNum, strTo, "", strExc(0), strExc(10), "", "", , , , strExc(3)
            End If
         End If
         'end 2022/04/28
         
         'Added by Lydia 2024/03/06 外專機械設計組人員異動調整程式：內專協辦工程師完成送件之後，需通知外專工程師進行請款
         'Move by Lydia 2024/03/12 改使用Outlook草稿，從FormSave移出
         'Mark by Lydia 2024/04/18 FCP案直接併入frm060104_k的Outlook，所以也不用---Sharon
         'If cp(1) = "FCP" And Mid(Text7(1), 4, 1) = "9" Then
         '   Call Pub_SetEngMail(cp(9))
         'End If
         ''end 2024/03/06
         'end 2024/04/18
         
         bolChkSave = True 'Added by Lydia 2018/10/31
         'Added by Lydia 2019/06/25 FCP特定案件性質的電子送件，發文確定後直接跳到"申請案號輸入"畫面，直接key 申請案號。
         'Modified by Lydia 2020/07/13 +衍生設計125
         'Modified by Lydia 2024/11/07 +改請獨立306  --- Gill
         If InStr("101,102,103,125,301,302,303,306,307,308", Me.Text7(2)) > 0 And Me.txtCP118 = "Y" Then
            If Index = 0 Then
               '若有未發文資料顯示警告
               PUB_GetCPunIssueDatas "" & Me.Text1.Text & "-" & Me.Text2.Text & "-" & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & "-" & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)
            End If
            'Added by Lydia 2019/06/27 101發明申請,102新型申請,103設計申請有客戶提供文件未處理,自動跳到客戶提供文件畫面
                                      'ex.優先權文件:必須新案發文時勾補文件-優先權才會自動產生下一程序,等到客戶提供文件做補文件收文才有下一程序可勾選
            strExc(1) = ""
            If InStr("101,102,103", Me.Text7(2)) > 0 Then
                Call PUB_ChkCPExist(pa, "1920", 1, strExc(1), , "D")
            End If
            If strExc(1) <> "" Then
                '檢查表單是否已開啟，若是，則關閉
                For Each nFrm In Forms
                   If StrComp(nFrm.Name, "frm060104_1", vbTextCompare) = 0 Then
                      Unload frm060104_1
                      Exit For
                   End If
                Next
                Call frm060121.SetParent(Me, Me.Text1 & Me.Text2 & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text))
                frm060121.Show
                Call frm060121.Command1_Click
            Else
            'end 2019/06/27
                 frm060105_2.PubOtherCall = Index & "frm060104_1;" & Me.Text1 & Me.Text2 & IIf(Len("" & Me.Text3.Text) <= 0, "0", Me.Text3.Text) & IIf(Len("" & Me.Text4.Text) <= 0, "00", Me.Text4.Text)
                 frm060105_2.Show
            End If
            
         Else
         'end 2019/06/25
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
            
            'Add By Sindy 2022/5/12
            If txtEmail = "Y" And cp(10) <> "416" Then
               frm060104_k.m_CP09 = cp(9)
               frm060104_k.m_strRecDate = txtRecDate
               frm060104_k.Hide
               frm060104_k.cmdOK(0) = 1
               Unload frm060104_k
            End If
            '2022/5/12 END
         End If
         Unload Me
      Case 1
         frm060104_1.Show
         Unload Me
      Case 2
         Unload frm060104_1
         Unload Me

      Case 4
         Me.Hide
         'Modify by Morgan 2006/7/4
         'frm060104_5.LoadMe strReceiveNo, pA(1), pA(2), pA(3), pA(4), 3
         'frm060104_5.Caption = "外專發文-變更事項"
         Set frm06010303_1.oParent = Me 'Add by Morgan 2011/10/5
         frm06010303_1.LoadMe strReceiveNo, pa(1), pa(2), pa(3), pa(4), 64
        m_blnClkChgEvnBtn = True
      Case 5
         Where1103ComeFrom Me, pa(1), pa(2), pa(3), pa(4)
      '2007/8/6 ADD BY SONIA
      Case 6
         frm060104_f.Text7(0) = m_CP50
         frm060104_f.Text7(1) = m_CP51
         frm060104_f.Text7(2) = m_CP52
         frm060104_f.Show
      '2007/8/6 END
   End Select
End Sub

Private Function FormSave() As Boolean
   Dim i As Integer, intStep As Integer, strTmp As String, intMax As Long, strTxt(1 To 20) As String
   Dim stCtrlDate(0 To 2) As String
   '2005/9/13 ADD BY SONIA
   Dim rsA As New ADODB.Recordset
   Dim rsB As New ADODB.Recordset
   Dim StrSQLa As String
   Dim StrSqlB As String
   Dim m_203CP48 As String    '2009/10/6 add by sonia
   'Add by Morgan 2010/1/6
   Dim stCP12 As String, stCP13 As String '最新收文智權人員,業務區
   Dim strCP137 As String, strCP167 As String 'Add By Sindy 2023/4/12
   Dim stCP17 As String
   'end 2010/1/6
   Dim stUpdate As String 'Add by Morgan 2010/5/13
   Dim strMemo416 As String   '2011/10/12 add by sonia
   Dim lngFee As Long 'Added by Morgan 2011/11/11
   Dim stCP118 As String, stCP152 As String 'Added by Morgan 2017/9/25
   Dim strCP06 As String, strCP48 As String 'Add By Sindy 2021/6/24
   Dim strPA64 As String, strPA65 As String, strPA67 As String, strPA68 As String 'Add By Sindy 2023/3/20
   
On Error GoTo CheckingErr

   cnnConnection.BeginTrans
  
   stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
   stCP12 = GetSalesArea(stCP13)
   
   intMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了  intMax = objPublicData.GetNextProgressNo
   
   strExc(0) = ""
   Select Case Text7(2)
      '2011/1/24 MODIFY BY SONIA 加積體電路117 (FCP-042803)
      'Modified by Morgan 2012/12/20 +衍生設計125
      Case "101", "102", "103", "104", "105", "117", "125"
         If pa(1) = "FCP" Then
            strExc(0) = strExc(0) & "pa10=" & TransDate(Text7(0), 2) & ","
         ElseIf pa(1) = "FG" Then
            strExc(0) = strExc(0) & "SP10=" & TransDate(Text7(0), 2) & ","
         End If
      '2008/8/21 add by sonia 414回復原狀發文時專用期仍有效者才更新專利權是否存在pa17
      Case "414"
         'Modify by Morgan 2010/8/12 百年蟲
         'If pa(1) = "FCP" And pa(25) >= strSrvDate(2) Then
         If pa(1) = "FCP" And Val(pa(25)) >= Val(strSrvDate(2)) Then
            strExc(0) = strExc(0) & "pa17='Y',"
         End If
      '2008/8/21 end
   End Select
   
   'Added by Morgan 2012/12/3
   If m_PA162 <> "" Then
      strExc(0) = strExc(0) & "pa162='" & m_PA162 & "',"
   End If
   'end 2012/12/3
   
   'Added by Morgan 2015/7/9
   If m_PA60 <> "" Then
      strExc(0) = strExc(0) & "pa60='" & m_PA60 & "',"
      If m_PA60 = "Y" Then
         Text7(17) = "放棄新型;" & Text7(17)
      ElseIf m_PA60 = "N" Then
         Text7(17) = "放棄發明;" & Text7(17)
      End If
   End If
   'end 2015/7/9
   
   'Added by Morgan 2012/12/27
   'Modified by Morgan 2023/10/16 +設計案的分割/改請設計/改請衍生設計
   If (Text7(2) = "307" And pa(8) = "1") Or (pa(8) = "3" And (Text7(2) = "307" Or Text7(2) = "303" Or Text7(2) = "308")) Then
      strExc(0) = strExc(0) & "pa163='" & m_PA163 & "',"
   End If
   
   Select Case pa(1)
      Case "FCP"
         'Modify by Morgan 2006/10/19 加 pa139
         'Modify by Morgan 2008/3/18 +pa143
         strTxt(1) = "UPDATE patent SET " & strExc(0) & "pa48=" & CNULL(ChgSQL(Text7(3))) & _
            ",pa77=" & CNULL(ChgSQL(Text7(10))) & ",pa05=" & CNULL(ChgSQL(Text7(14))) & _
            ",pa06=" & CNULL(ChgSQL(Text7(15))) & ",pa07=" & CNULL(ChgSQL(Text7(16))) & _
            ",pa91=" & CNULL(ChgSQL(Text7(18))) & ",pa51=" & CNULL(ChgSQL(Text7(4))) & _
            ",pa52=" & CNULL(ChgSQL(Text7(5))) & ",pa53=" & CNULL(ChgSQL(Text7(6))) & _
            ",pa54=" & CNULL(ChgSQL(Text7(7))) & ",pa55=" & CNULL(ChgSQL(Text7(8))) & _
            ",pa56=" & CNULL(ChgSQL(Text7(9))) & ",pa139=" & CNULL(ChgSQL(Text7(19))) & _
            ",pa143='" & m_PA143 & "'" & _
            " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      Case "FG"
         'Modify by Morgan 2006/10/19 加 sp71
         'Modify by Morgan 2007/5/4 SP30=" & CNULL(ChgSQL(Text7(4)))-->,SP30=" & CNULL(ChgSQL(Text7(5)))
         strTxt(1) = "UPDATE SERVICEPRACTICE SET " & strExc(0) & "SP29=" & CNULL(Text7(3)) & _
            ",SP27=" & CNULL(Text7(10)) & ",SP30=" & CNULL(ChgSQL(Text7(5))) & _
            ",SP05=" & CNULL(ChgSQL(Text7(14))) & ",SP06=" & CNULL(ChgSQL(Text7(15))) & _
            ",SP07=" & CNULL(ChgSQL(Text7(16))) & ",SP18=" & CNULL(ChgSQL(Text7(18))) & _
            ",SP71=" & CNULL(ChgSQL(Text7(19))) & _
            " WHERE " & ChgService(pa(1) & pa(2) & pa(3) & pa(4))
   End Select
   '911105 nick transation
   cnnConnection.Execute strTxt(1)
   
   'Add by Morgan 2010/1/7 有超頁超項費時收文規費只要放原始規費
   If m_bolChkFee Then
      stCP17 = m_lngOfficialFee
   Else
      stCP17 = txtCP84
   End If
   
   '2010/12/8 ADD BY SONIA 分割案之實審發文,若母案有已收未取消的再審程序且申請日在2010/1/1以前者,分割案實審規費應為8000元
   If m_Div416OfficialFee <> 0 Then
      stCP17 = txtCP84
   End If
   
   'Add by Morgan 2010/5/13
   stUpdate = ""
   If Text7(2) = "202" Then
      If chkCP86(0).Value = 1 Then
         stUpdate = ",CP86='Y'"
      ElseIf chkCP86(1).Value = 1 Then
         stUpdate = ",CP86=NULL"
      End If
   End If
   'end 2010/5/13
   'Added by Lydia 2015/12/31 +CP43 (會稿924)
   'Modify By Sindy 2019/1/18 更改也會輸相關總收文號
   If Text7(2) = "924" And txtCP43.Visible = True Then
   'If txtCP43.Visible = True Then
   '2019/1/18 END
      stUpdate = stUpdate & ",CP43=" & CNULL(txtCP43.Text)
   End If
   'end 2015/12/31
   
   'Added by Morgan 2017/9/25
   '電子送件有規費的一律設自動扣款(同內專) --敏莉
   stCP118 = txtCP118
   stCP152 = ""
   If txtCP118 = "Y" And Val(txtCP84) > 0 Then
      stCP118 = "A"
      'Modifed by Lydia 2018/09/11 改成模組
'      If txtPayToday <> "" Then
'         If txtPayToday = "Y" Then
'            stCP152 = CompWorkDay(2, DBDATE(Text7(0)))
'         Else
'            stCP152 = CompWorkDay(3, DBDATE(Text7(0)))
'         End If
'      End If
      stCP152 = Pub_FcpSetPayToday("2", Text7(0).Text, txtPayToday.Text)
      'end 2018/09/11
   End If
   'end 2017/9/25
   
   'Add By Sindy 2018/5/22 發文時,清除前面的總頁數,總項數;後面會更新在此筆文號
'   If txtCP118 = "" And _
'      (cp(10) = 實體審查 Or _
'       cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Or _
'       cp(10) = "307") Then
   'Modify By Sindy 2018/6/11 ex:FCP-050420:再審申請 + ,cp137=null,cp138=null
   'Modify By Sindy 2018/6/26 敏莉說不控管210.製作中說(取消Or cp(10) = "210")
'   If m_bolChkPageItem = True Or _
'      (txtCP118 = "" And (cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "307")) Then
   'Add By Sindy 2019/3/22 敏莉:新案翻譯,檢視中說,核對中說格式,製作中說;都要執行此動作清除此案號的頁項數,只留該筆後面存檔
   '                       ex:FCP-059999(送中說+主動修正時尚未收實審)
   If m_bolChkPageItem = True Or _
      cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Then
      'Modify By Sindy 2023/3/16 +,cp167=null,cp168=null
      strSql = "UPDATE CASEPROGRESS SET cp135=null,cp136=null,cp137=null,cp138=null,cp167=null,cp168=null" & _
               " WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'"
      cnnConnection.Execute strSql
   End If
   '2018/5/22 END
   
   'Modify by morgan 2005/8/4 加 cp110
   'Modify by Morgan 2007/7/19 加 cp113,cp114
   '2007/8/6 MODIFY BY SONIA 加 CP50,CP51,CP52
   'Modify by Morgan 2010/1/6 +CP135,CP136,CP137,CP138,m_FeeMemo
   'Modify by Morgan 2010/4/20 940也不要回寫cp16,cp17
   'Modify by Morgan 2011/1/17 +CP118
   'Modify by Sindy 2016/8/18 + ,cp32=" & CNULL(Text7(13)) & ","
   'Modified by Morgan 2017/9/25 +CP152
   'Modified by Lydia 2019/01/07 +CP148 是否有一併修正
   strTxt(2) = "UPDATE CASEPROGRESS SET cp27=" & CNULL(TransDate(Text7(0), 2)) & "," & _
      "cp14=" & CNULL(Text7(1)) & ",cp44=" & CNULL(ChangeCustomerL(Text7(12))) & "," & _
      "cp10=" & CNULL(Text7(2)) & ",cp20=" & CNULL(Text7(13)) & ",cp32=" & CNULL(Text7(13)) & "," & _
      "CP50=" & CNULL(m_CP50) & ", CP51=" & CNULL(m_CP51) & ", CP52=" & CNULL(m_CP52) & "," & _
      "cp64=" & CNULL(ChgSQL(m_FeeMemo & Text7(17))) & ",cp84=" & Format(Val(txtCP84.Text)) & IIf(bolDelay Or Text7(2) = "940", "", ", CP16=NVL(CP16,0)-NVL(CP17,0)+" & Format(Val(stCP17)) & _
      ", CP17=" & Format(Val(stCP17))) & ", CP18=NVL(CP18,0)" & _
      ", CP113=" & CNULL(txtCP113.Text, True) & ", CP114=" & CNULL(txtCP114.Text, True) & ",cp135=" & CNULL(txtCP135, True) & ",cp136=" & CNULL(txtCP136, True) & _
      ",cp137=" & CNULL(txtCP137, True) & ",cp138=" & CNULL(txtCP138, True) & _
      ",cp110=" & CNULL(cp(110)) & ",CP22=NULL" & stUpdate & ",CP118='" & stCP118 & "',CP152='" & stCP152 & "',CP148=" & CNULL(m_strCP148) & _
      " WHERE CP09='" & strReceiveNo & "'"
   '911105 nick transation
   cnnConnection.Execute strTxt(2), intI
   intStep = 3
   
   'Add By Sindy 2023/3/20 有專利說明書頁數明細要回寫基本檔
   strExc(0) = "Select * From PageDetail where pd01='" & strReceiveNo & "' and pd20 is null" & _
               " union all " & _
               "Select * From PageDetail where pd20='" & strReceiveNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strExc(0) = ""
      If Val("" & RsTemp.Fields("pd02")) > 0 Or Val("" & RsTemp.Fields("pd06")) > 0 Or Val("" & RsTemp.Fields("pd10")) > 0 Then
         strPA64 = Val(pa(64)) + Val("" & RsTemp.Fields("pd02")) - Val("" & RsTemp.Fields("pd06")) - Val("" & RsTemp.Fields("pd10"))
         If Val(strPA64) <> pa(64) Then
            strExc(0) = strExc(0) & ",pa64=" & strPA64
         End If
      End If
      If Val("" & RsTemp.Fields("pd03")) > 0 Or Val("" & RsTemp.Fields("pd07")) > 0 Or Val("" & RsTemp.Fields("pd11")) > 0 Then
         strPA65 = Val(pa(65)) + Val("" & RsTemp.Fields("pd03")) - Val("" & RsTemp.Fields("pd07")) - Val("" & RsTemp.Fields("pd11"))
         If Val(strPA65) <> pa(65) Then
            strExc(0) = strExc(0) & ",pa65=" & strPA65
         End If
      End If
      If Val("" & RsTemp.Fields("pd04")) > 0 Or Val("" & RsTemp.Fields("pd08")) > 0 Or Val("" & RsTemp.Fields("pd12")) > 0 Then
         strPA67 = Val(pa(67)) + Val("" & RsTemp.Fields("pd04")) - Val("" & RsTemp.Fields("pd08")) - Val("" & RsTemp.Fields("pd12"))
         If Val(strPA67) <> pa(67) Then
            strExc(0) = strExc(0) & ",pa67=" & strPA67
         End If
      End If
      If Val("" & RsTemp.Fields("pd05")) > 0 Or Val("" & RsTemp.Fields("pd09")) > 0 Or Val("" & RsTemp.Fields("pd13")) > 0 Then
         strPA68 = Val(pa(68)) + Val("" & RsTemp.Fields("pd05")) - Val("" & RsTemp.Fields("pd09")) - Val("" & RsTemp.Fields("pd13"))
         If Val(strPA68) <> pa(68) Then
            strExc(0) = strExc(0) & ",pa68=" & strPA68
         End If
      End If
      '圖式圖數
      If Val("" & RsTemp.Fields("pd21")) <> Val(pa(173)) Then
         strExc(0) = strExc(0) & ",pa173=" & Val("" & RsTemp.Fields("pd21"))
      End If
      '更新項數
      If "" & RsTemp.Fields("pd20") = strReceiveNo Then '中說一併修正
         If Val(txtCP136) <> Val(pa(172)) And Val(txtCP136) > 0 Then
            strExc(0) = strExc(0) & ",pa172=" & txtCP136
         End If
      Else
         If Val(m_allItem) <> Val(pa(172)) And Val(m_allItem) > 0 Then
            strExc(0) = strExc(0) & ",pa172=" & m_allItem
         End If
      End If
      If strExc(0) <> "" Then
         strExc(0) = Mid(strExc(0), 2)
         strSql = "update patent set " & strExc(0) & " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
         Pub_SeekTbLog strSql, , , , Me.Caption & "(" & Me.Name & ")"
         cnnConnection.Execute strSql
      End If
   End If
   '2023/3/20 END
   
   'Add By Sindy 2023/3/27 內部收文"908.代辦退費"
   'If Val(txtDecreasePageFee.Text) > 0 Or Val(txtDecreaseItemFee.Text) > 0 Then
      'Add By Sindy 2023/10/26
      If m_bolIns908 = True Then
      '2023/10/26 END
         stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
         stCP12 = GetSalesArea(stCP13)
         If strPD01 <> "" Then
            strExc(0) = " SELECT CP09,CP167,CP137" & _
                        " FROM CaseProgress" & _
                        " WHERE CP09='" & strPD01 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strCP167 = Val("" & RsTemp.Fields("CP167"))
               strCP137 = Val("" & RsTemp.Fields("CP137"))
            End If
         Else
            strCP167 = Val(txtCP167.Text)
            strCP137 = Val(txtCP137.Text)
         End If
         strExc(1) = AutoNo("B", 6) 'B類總收文號
         strExc(9) = Val(txtDecreasePageFee.Text) + Val(txtDecreaseItemFee.Text) '退費金額
         strExc(10) = "" '進度備註
         strExc(8) = ""
         If Val(txtDecreasePageFee.Text) > 0 Then
            strExc(8) = m_str938RecvNo
            strExc(10) = "此次修正刪除" & Val(strCP167) & "頁，一併辦理退費" & Format(txtDecreasePageFee, "###,###,##0") & "元"
         End If
         If Val(txtDecreaseItemFee.Text) > 0 Then
            If strExc(8) = "" Then strExc(8) = m_str939RecvNo
            strExc(10) = IIf(strExc(10) <> "", strExc(10) & ";", "") & _
               "此次修正刪除" & Val(strCP137) & "項，一併辦理退費" & Format(txtDecreaseItemFee, "###,###,##0") & "元"
         End If
         'Modify By Sindy 2024/7/30 敏莉:自動內部收文"908代辦退費請預設不請款（案例：FCP-070120）
         strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp16,cp17,cp18,cp19,cp20,cp26,cp43,cp84,cp64,cp27) values " & _
            " ('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & _
            ",'" & strExc(1) & "','908','" & stCP12 & "','" & stCP13 & "','" & strUserNum & "',0,0,0," & strExc(9) & ",'N','N'" & _
            ",'" & strExc(8) & "',0,'" & strExc(10) & "'," & strSrvDate(1) & ")"
         cnnConnection.Execute strSql, intI
         
         'Added by Lydia 2025/04/11 FCP若有自動產生代辦退費且同時上發文日時，下一程序都要掛代辦退費的【催審】411; 例：FCP-066290
         strExc(0) = GetUrgeDate(pa(1), pa(9), "908", strSrvDate(1))
         strExc(0) = CompWorkDay(1, strExc(0), 1)
         intMax = GetNextProgressNo
         strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) VALUES ('" & strExc(1) & "','" & pa(1) & _
                  "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 催審 & "," & strExc(0) & "," & strExc(0) & ",'" & strUserNum & "'," & intMax & ")"
         cnnConnection.Execute strSql, intI
         'end 2025/04/11
      End If
   'End If
   '2023/3/27 END
   
   'Add By Sindy 2025/4/15 "901"告代發文時若相關總收文號掛"945"電話連絡單時,
   '  若有輸入管制下一程序期限者,自動產生下一程序"204"修正
   If Text7(2) >= 告知代理人 And cp(43) <> "" Then
      strExc(0) = "select cp09,cp10 from caseprogress where cp09='" & cp(43) & "' and cp10='945'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(0) = "select * from EMPELECTRONDATA where eed01='" & strReceiveNo & "' and eed14 is not null and eed15 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            intMax = GetNextProgressNo
            strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP23,NP08,NP10,NP15,NP22)" & _
                     " VALUES ('" & strReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'" & _
                     ",'" & 修正 & "'," & RsTemp.Fields("eed14") & "," & RsTemp.Fields("eed15") & ",'" & cp(13) & "','電話通知修正'," & intMax & ")"
            cnnConnection.Execute strSql, intI
         End If
      End If
   End If
   '2025/4/15 END
   
   'Added by Morgan 2011/11/11
   '更新收文規費為延期發文規費(含補收款) Ex.FCP-030936
   If bolDelay = True And m_strDelayCP09 <> "" Then
      lngFee = PUB_GetDelayPayFee(m_strDelayCP09)
      If lngFee > 0 Then
         strSql = "update caseprogress set cp16=NVL(cp16,0)-NVL(cp17,0)+" & lngFee & ",cp17=" & lngFee & " where cp09='" & strReceiveNo & "'"
         cnnConnection.Execute strSql, intI
      End If
   End If
   
   'Added by Morgan 2020/2/25
   '改請衍生設計發文若有分割未發文時
   '1.307(分割)自動上發文，並帶入改請衍生設計之智慧局收文文號（稽核用）
   '2.分割之發文規費0
   If Text7(2) = "308" Then
      strSql = "update caseprogress set cp27=" & CNULL(TransDate(Text7(0), 2), True) & ", cp64='" & m_307CP64 & "'||cp64, cp84=0" & _
         " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'" & _
         " and cp10='307' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql, intI
   End If
   'end 2020/2/25
   
   intMax = GetNextProgressNo  'edit by nickc 2007/02/02 不用 dll 了  intMax = objPublicData.GetNextProgressNo
   '1
   If Text7(11) <> "" Then
      'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
      strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
         "NP07,NP08,NP09,NP10,NP22) VALUES ('" & strReceiveNo & "','" & pa(1) & _
         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 催審 & "," & _
         PUB_GetWorkDay1(TransDate(Text7(11), 2), True) & "," & TransDate(Text7(11), 2) & ",'" & strUserNum & "'," & intMax & ")"
      cnnConnection.Execute strTxt(intStep)
         
      intMax = intMax + 1
      intStep = intStep + 1
   End If
   
   If Text7(2) = 修正 Or Text7(2) = 申復 Then
      strExc(0) = "SELECT NVL(MIN(NP22),0) FROM NEXTPROGRESS " & _
                  "WHERE NP02 = '" & pa(1) & "' AND " & _
                        "NP03 = '" & pa(2) & "' AND " & _
                        "NP04 = '" & pa(3) & "' AND " & _
                        "NP05 = '" & pa(4) & "' AND " & _
                        "NP07=" & 催審 & " AND " & _
                        "( NP06 IS NULL OR NP06 = '' ) "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.Fields(0) <> "0" Then
            strTmp = CompDate(1, 6, TransDate(Text7(0), 2))
            'Modify by Morgan 2006/1/24 加本所案號,因NP22有可能會重複
            'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
            strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP08=" & PUB_GetWorkDay1(strTmp, True) & "," & _
               "NP09=" & strTmp & " WHERE NP22=" & RsTemp.Fields(0) & _
               " AND NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "'"
               
            '911105 nick transation
            cnnConnection.Execute strTxt(intStep)
               
            intStep = intStep + 1
         End If
      End If
   End If
   
   '3
   If Text7(2) <> cp(10) Then
      If Left(Text7(2), 1) = "3" Then
         strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP09," & _
            "CP10,CP20,CP26,CP32,CP43,CP27) VALUES ('" & pa(1) & "','" & pa(2) & _
            "','" & pa(3) & "','" & pa(4) & "','" & AutoNo("B", 6) & "','" & Text7(2) & _
            "','N','N','N','" & strReceiveNo & "'," & TransDate(Text7(0), 2) & ")"
            
        '911105 nick transation
        cnnConnection.Execute strTxt(intStep)
            
         intStep = intStep + 1
         
         If Text7(2) = "301" Or Text7(2) = "302" Or Text7(2) = "303" Then
            strTxt(intStep) = "UPDATE paTENT SET pa08='" & Right(Text7(2), 1) & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
            
            '911105 nick transation
            cnnConnection.Execute strTxt(intStep)
                
            intStep = intStep + 1
         End If
      ElseIf Text7(2) = 舉發 Then
         strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP09," & _
            "CP10,CP20,CP26,CP32,CP43,CP27) VALUES ('" & pa(1) & "','" & pa(2) & _
            "','" & pa(3) & "','" & pa(4) & "','" & AutoNo("B", 6) & "','" & 舉發 & _
            "','N','N','N','" & strReceiveNo & "'," & TransDate(Text7(0), 2) & ")"
            
        '911105 nick transation
        cnnConnection.Execute strTxt(intStep)
            
         intStep = intStep + 1
         strTxt(intStep) = "UPDATE paTENT SET pa23='3' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
         
        '911105 nick transation
        cnnConnection.Execute strTxt(intStep)
         
         intStep = intStep + 1
      End If
   End If
  
   'Add by Morgan 2004/3/23
   '更新核稿人資料
   If txtEP04.Enabled = True And (txtEP04.Tag <> txtEP04.Text) Then
      strTxt(intStep) = "Update EngineerProgress Set EP04='" & txtEP04.Text & "' Where EP02='" & cp(9) & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   'FCP及P案件，若專利種類為'發明'且案件性質為'分割'時，實審期限=母案之申請日＋3年<=分割案之收文日＋1月
   '若有收文未取消收文之'實體審查'，則更新該筆'實體審查'之期限，若無則新增下一程序'實體審查'期限，並顯示'此分割案尚未收文實體審查，期限為XXXXXX，請提醒智權人員 !!'
   'Modify by Morgan 2004/8/6
   '改請發明案也要控制實體審查
   'If pa(8) = "1" And Text7(2).Text = "307" Then
   If (pa(8) = "1" And Text7(2).Text = "307") Or (Text7(2).Text = 改請發明) Then
      
      'Modified by Morgan 2015/9/10 +續行母案再審
      'If bol416Control = True Then 'Added byMorgan 2013/12/11
      If bol416Control Or m_bol435 Then
         If m_bol435 Then
            strExc(1) = "435"
         Else
            strExc(1) = "416"
         End If
               
         '有收'實體審查'
         If m_stCP09 <> "" Then
            strTxt(intStep) = " UPDATE CASEPROGRESS SET CP06=" & m_stVar(0) & ",CP07=" & m_stVar(3) & " WHERE CP09='" & m_stCP09 & "'"
         '沒有收'實體審查'
         Else
            'Modify By Sindy 2021/4/27 + ,NP23=" & CNULL(DBDATE(m_pAgreeOnDate)):約定期限
            strTxt(intStep) = _
               " DECLARE" & _
                  " V_NP22 NUMERIC(10,0);" & _
                  " V_NP02 VARCHAR2(9);" & _
               " BEGIN" & _
                     " SELECT MAX(NP02) INTO V_NP02 FROM NEXTPROGRESS WHERE NP01='" & cp(9) & "' AND NP07='" & strExc(1) & "';" & _
                     " IF V_NP02 IS NULL THEN" & _
                        " SELECT NVL(MAX(NP22),0)+1 INTO V_NP22 FROM NEXTPROGRESS;" & _
                        " INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22,NP23)" & _
                        " VALUES ('" & cp(9) & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'" & _
                        " ,'" & strExc(1) & "'," & m_stVar(0) & "," & m_stVar(3) & ",'" & cp(13) & "',V_NP22," & CNULL(DBDATE(m_pAgreeOnDate)) & ");" & _
                        " ELSE" & _
                           " UPDATE NEXTPROGRESS SET NP08=" & m_stVar(0) & ",NP09=" & m_stVar(3) & ",NP23=" & CNULL(DBDATE(m_pAgreeOnDate)) & " WHERE NP01='" & cp(9) & "' AND NP07='" & strExc(1) & "';" & _
                     " END IF;" & _
               " END;"
         End If
         cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
         
         'Add By Sindy 2021/8/11
         '分割發文時若有主動修正未發文，則Update主動修正本所期限和實審本所期限一致(前提是實審期限要先update)
         '主動修正本所期限=實審本所期限
         strCP06 = m_stVar(0)
         '承辦期限=本所期限往前-5個工作天
         strCP48 = CompWorkDay(5, CompDate(2, -1, DBDATE(strCP06)), 1)
         strSql = "Update caseprogress set CP48=" & strCP48 & ",CP06='" & strCP06 & "' WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                  " and cp10 ='203' and CP27 IS NULL AND CP57 IS NULL "
         cnnConnection.Execute strSql
      End If
      
      If bol416Control = True Then
      'end 2015/9/10
      
      '2008/10/13 add by sonia
      '2009/2/6 MODIFY BY SONIA 加入X20438000,X21775000
      'If ChangeCustomerL(pa(26)) = "X55560000" Or ChangeCustomerL(pa(26)) = "X55560010" Or ChangeCustomerL(pa(27)) = "X55560000" Or ChangeCustomerL(pa(27)) = "X55560010" Or ChangeCustomerL(pa(28)) = "X55560000" Or ChangeCustomerL(pa(28)) = "X55560010" Or ChangeCustomerL(pa(29)) = "X55560000" Or ChangeCustomerL(pa(29)) = "X55560010" Or ChangeCustomerL(pa(30)) = "X55560000" Or ChangeCustomerL(pa(30)) = "X55560010" Then
'Modify by Morgan 2011/6/13 改抓共用函數
'      If ChangeCustomerL(pa(26)) = "X55560000" Or ChangeCustomerL(pa(27)) = "X55560000" Or ChangeCustomerL(pa(28)) = "X55560000" Or ChangeCustomerL(pa(29)) = "X55560000" Or ChangeCustomerL(pa(30)) = "X55560000" Or _
'         ChangeCustomerL(pa(26)) = "X55560010" Or ChangeCustomerL(pa(27)) = "X55560010" Or ChangeCustomerL(pa(28)) = "X55560010" Or ChangeCustomerL(pa(29)) = "X55560010" Or ChangeCustomerL(pa(30)) = "X55560010" Or _
'         ChangeCustomerL(pa(26)) = "X20438000" Or ChangeCustomerL(pa(27)) = "X20438000" Or ChangeCustomerL(pa(28)) = "X20438000" Or ChangeCustomerL(pa(29)) = "X20438000" Or ChangeCustomerL(pa(30)) = "X20438000" Or _
'         ChangeCustomerL(pa(26)) = "X21775000" Or ChangeCustomerL(pa(27)) = "X21775000" Or ChangeCustomerL(pa(28)) = "X21775000" Or ChangeCustomerL(pa(29)) = "X21775000" Or ChangeCustomerL(pa(30)) = "X21775000" Then
'         strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP15='不寄函逕收文;'||NP15 WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07='416' AND NP06 IS NULL"
'         cnnConnection.Execute strTxt(intStep)
'         intStep = intStep + 1
'      End If
         strExc(1) = ""
         'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
         'For intI = 0 To 4
         '   If pa(26 + intI) <> "" Then
         '      'Modified by Morgan 2012/2/2 改與發明申請一致 +pa75
         '      'Modified by Morgan 2013/9/11 改抓設定檔
         '      'strExc(1) = PUB_Get416Memo(ChangeCustomerL(pa(26 + intI)), ChangeCustomerL(pa(75)))
         '      strExc(1) = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26 + intI)))
         '      If strExc(1) <> "" Then Exit For
         '   End If
         'Next
         strExc(1) = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL(pa(75)), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30))
         'end 2022/08/02
         If strExc(1) <> "" Then
            'Modified by Lydia 2022/08/02
            'strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP15='" & strExc(1) & "'||NP15 WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07='416' AND NP06 IS NULL"
            strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP15='" & ChgSQL(strExc(1)) & "'||';'||NP15 " & _
                   "WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07='416' AND NP06 IS NULL and instr(np15,'" & strExc(1) & "') = 0 "
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
         End If
'end 2011/6/13

      '2008/10/13 end
      
      End If 'Added byMorgan 2013/12/11
      
      'Add by Morgan 2006/5/1
      '申請寄存108之存活證明221期限管制
      If m_bol108 = True Then
        ' Added by Lydia 2022/09/15 抓約定期限
         If Trim(m_stVar(3)) = "" Then m_stVar(3) = cp(7)
         If m_pAgreeOnDate = "" Then
           Call PUB_GetFCPOurDeadline(m_stVar(3), 4, , m_pAgreeOnDate)
         End If
         'end 2022/09/15
         'Modify By Sindy 2021/4/27 + , DBDATE(m_pAgreeOnDate)
         strSql = PUB_Get221SQL(cp, m_stVar(0), m_stVar(3), cp(13), cp(9), DBDATE(m_pAgreeOnDate))
         cnnConnection.Execute strSql
      End If
   End If
   'Add end
   
   '2005/9/13 ADD BY SONIA 發明案發文日>=91/10/26
   If pa(8) = "1" And Val(Text7(0)) >= 911026 Then
      '若案件性質為"發明申請"(101)
      If cp(10) = "101" Then
         Dim strTmp1A(0 To 4) As String, strTmpA(1 To 3) As String
         strTmp1A(0) = cp(9)
         For i = 1 To 4
            strTmp1A(i) = pa(i)
         Next
         If GetMoneyDate(Val(pa(8) + 3), pa(9), strTmp1A, strTmpA(1), strTmpA(2), strTmpA(3)) = True Then
            '法定期限
            If strTmpA(3) <> "" Then
               '法定期限應為止日加一天
               strTmpA(3) = CompDate(2, 1, strTmpA(3))
               'Modified by Morgan 2014/11/20 外專改回舊規則
               ''Added by Morgan 2014/10/29
               'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
               '   strTmpA(2) = PUB_GetOurDeadline(strTmpA(3))
               'Else
               ''end 2014/10/29
               
               'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
               If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
                  'Modify By Sindy 2021/4/27 + m_pAgreeOnDate
                  strTmpA(2) = PUB_GetFCPOurDeadline(strTmpA(3), 4, , m_pAgreeOnDate)
               Else
               'end 2019/7/11
      
                  '本所期限 = 法定期限 - 4天
                  strTmpA(2) = CompDate(2, -4, strTmpA(3))
                  
               End If 'Added by Morgan 2019/7/11
               'End If 'Added by Morgan 2014/10/29
               'end 2014/11/20
               
                If rsB.State <> adStateClosed Then rsB.Close
                Set rsB = Nothing
                StrSqlB = "Select * From CaseProgress Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10='" & 實體審查 & "' And CP05 IS NOT NULL And CP57 IS NULL "
                rsB.CursorLocation = adUseClient
                rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
                '若案件進度檔無實體審查的資料(即未收文)
                If rsB.RecordCount <= 0 Then
                    '2011/10/12 add by sonia改用PUB_GetMemo
                     strMemo416 = ""
                     'Modified by Lydia 2022/08/02 整合模組：修改為複數新規則
                     'For intI = 26 To 30
                     '   If Not IsNull(pa(intI)) Then
                     '      'Modified by Morgan 2013/9/11 改抓設定檔
                     '      'strMemo416 = PUB_Get416Memo(ChangeCustomerL(pa(intI)), ChangeCustomerL(pa(75)))
                     '      strMemo416 = PUB_GetNpMemo(pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL(pa(75)), ChangeCustomerL(pa(intI)))
                     '      If strMemo416 <> "" Then Exit For
                     '   End If
                     'Next
                     strMemo416 = PUB_GetNpMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), "416", ChangeCustomerL(pa(75)), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30))
                     'end 2022/08/02
                    '2011/10/12 end
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                    StrSQLa = "Select NP01,NP07,NP22,NP15 From Nextprogress Where NP22= (SELECT MAX(NP22) FROM NEXTPROGRESS WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP07='" & 實體審查 & "' AND NP06 IS NULL ) "
                    rsA.CursorLocation = adUseClient
                    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                    If rsA.RecordCount > 0 Then
'2011/10/12 modify by sonia 改用PUB_GetMemo
'                       '2008/10/13 MODIFY by sonia
'                       'strTxt(intStep) = "Update NEXTPROGRESS SET NP08=" & Val(strTmpA(2)) & ",NP09=" & Val(strTmpA(3)) & " WHERE NP01='" & rsA.Fields(0).Value & "' AND NP07='" & rsA.Fields(1).Value & "' AND NP22=" & rsA.Fields(2).Value
'                       '2009/2/6 MODIFY BY SONIA 加入X20438000,X21775000
'                       '2009/5/4 modify by sonia 加入Y51304020
'                       'If ChangeCustomerL(pa(26)) = "X55560000" Or ChangeCustomerL(pa(26)) = "X55560010" Or ChangeCustomerL(pa(27)) = "X55560000" Or ChangeCustomerL(pa(27)) = "X55560010" Or ChangeCustomerL(pa(28)) = "X55560000" Or ChangeCustomerL(pa(28)) = "X55560010" Or ChangeCustomerL(pa(29)) = "X55560000" Or ChangeCustomerL(pa(29)) = "X55560010" Or ChangeCustomerL(pa(30)) = "X55560000" Or ChangeCustomerL(pa(30)) = "X55560010" Then
'                       If ChangeCustomerL(pa(26)) = "X55560000" Or ChangeCustomerL(pa(27)) = "X55560000" Or ChangeCustomerL(pa(28)) = "X55560000" Or ChangeCustomerL(pa(29)) = "X55560000" Or ChangeCustomerL(pa(30)) = "X55560000" Or _
'                          ChangeCustomerL(pa(26)) = "X55560010" Or ChangeCustomerL(pa(27)) = "X55560010" Or ChangeCustomerL(pa(28)) = "X55560010" Or ChangeCustomerL(pa(29)) = "X55560010" Or ChangeCustomerL(pa(30)) = "X55560010" Or _
'                          ChangeCustomerL(pa(26)) = "X20438000" Or ChangeCustomerL(pa(27)) = "X20438000" Or ChangeCustomerL(pa(28)) = "X20438000" Or ChangeCustomerL(pa(29)) = "X20438000" Or ChangeCustomerL(pa(30)) = "X20438000" Or _
'                          ChangeCustomerL(pa(26)) = "X21775000" Or ChangeCustomerL(pa(27)) = "X21775000" Or ChangeCustomerL(pa(28)) = "X21775000" Or ChangeCustomerL(pa(29)) = "X21775000" Or ChangeCustomerL(pa(30)) = "X21775000" Or _
'                          ChangeCustomerL(pa(75)) = "Y51304020" Then
'                          strTxt(intStep) = "Update NEXTPROGRESS SET NP08=" & Val(strTmpA(2)) & ",NP09=" & Val(strTmpA(3)) & ",NP15='不寄函逕收文;'||NP15 WHERE NP01='" & rsA.Fields(0).Value & "' AND NP07='" & rsA.Fields(1).Value & "' AND NP22=" & rsA.Fields(2).Value
'                       Else
'                          strTxt(intStep) = "Update NEXTPROGRESS SET NP08=" & Val(strTmpA(2)) & ",NP09=" & Val(strTmpA(3)) & " WHERE NP01='" & rsA.Fields(0).Value & "' AND NP07='" & rsA.Fields(1).Value & "' AND NP22=" & rsA.Fields(2).Value
'                       End If
'                       '2008/10/13 END
                       '2011/10/12 同時更新智權人員
                       'Modify By Sindy 2021/4/27 + ,NP23=" & CNULL(DBDATE(m_pAgreeOnDate)):約定期限
                       'Modified by Lydia 2022/08/02 判斷備註不存在才加註
                       'strTxt(intStep) = "Update NEXTPROGRESS SET NP08=" & Val(strTmpA(2)) & ",NP09=" & Val(strTmpA(3)) & ",NP23=" & CNULL(DBDATE(m_pAgreeOnDate)) & ",NP10='" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "',np15=decode(NP15,null,'" & strMemo416 & "','" & strMemo416 & "'||NP15) WHERE NP01='" & rsA.Fields(0).Value & "' AND NP07='" & rsA.Fields(1).Value & "' AND NP22=" & rsA.Fields(2).Value
                       StrSQLa = ""
                       If InStr("" & rsA.Fields("np15") & ";", strMemo416) = 0 And strMemo416 <> "" Then
                           StrSQLa = ", NP15='" & ChgSQL(strMemo416) & IIf("" & rsA.Fields("np15") <> "", ";", "") & "'||NP15 "
                       End If
                       strTxt(intStep) = "Update NEXTPROGRESS SET NP08=" & Val(strTmpA(2)) & ",NP09=" & Val(strTmpA(3)) & ",NP23=" & CNULL(DBDATE(m_pAgreeOnDate)) & ",NP10='" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "' " & StrSQLa & " WHERE  NP01='" & rsA.Fields(0).Value & "' AND NP07='" & rsA.Fields(1).Value & "' AND NP22=" & rsA.Fields(2).Value
                       'end 2022/08/02
'2011/10/12 end
                       cnnConnection.Execute strTxt(intStep)
                       intStep = intStep + 1
                    Else
                        strTxt(intStep) = "declare intMax number;begin   select max(np22)+1 into intMax from nextprogress;IF intMax IS NULL THEN intMax:=1; END IF;"
'2011/10/12 modify by sonia 改用PUB_GetMemo
'                        '2008/10/13 MODIFY by sonia
'                        'strTxt(intStep) = strTxt(intStep) & "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'                           "NP09,NP10,NP22) VALUES ('" & cp(9) & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & _
'                           pa(4) & "'," & 實體審查 & "," & Val(strTmpA(2)) & "," & Val(strTmpA(3)) & ",'" & cp(13) & "',intMax);"
'                        '2009/2/6 MODIFY BY SONIA 加入X20438000,X21775000
'                        '2011/10/7 modify by sonia 加入X45814同時Y45814才要
'                        'If ChangeCustomerL(pa(26)) = "X55560000" Or ChangeCustomerL(pa(26)) = "X55560010" Or ChangeCustomerL(pa(27)) = "X55560000" Or ChangeCustomerL(pa(27)) = "X55560010" Or ChangeCustomerL(pa(28)) = "X55560000" Or ChangeCustomerL(pa(28)) = "X55560010" Or ChangeCustomerL(pa(29)) = "X55560000" Or ChangeCustomerL(pa(29)) = "X55560010" Or ChangeCustomerL(pa(30)) = "X55560000" Or ChangeCustomerL(pa(30)) = "X55560010" Then
'                        If ChangeCustomerL(pa(26)) = "X55560000" Or ChangeCustomerL(pa(27)) = "X55560000" Or ChangeCustomerL(pa(28)) = "X55560000" Or ChangeCustomerL(pa(29)) = "X55560000" Or ChangeCustomerL(pa(30)) = "X55560000" Or _
'                           ChangeCustomerL(pa(26)) = "X55560010" Or ChangeCustomerL(pa(27)) = "X55560010" Or ChangeCustomerL(pa(28)) = "X55560010" Or ChangeCustomerL(pa(29)) = "X55560010" Or ChangeCustomerL(pa(30)) = "X55560010" Or _
'                           ChangeCustomerL(pa(26)) = "X20438000" Or ChangeCustomerL(pa(27)) = "X20438000" Or ChangeCustomerL(pa(28)) = "X20438000" Or ChangeCustomerL(pa(29)) = "X20438000" Or ChangeCustomerL(pa(30)) = "X20438000" Or _
'                           ChangeCustomerL(pa(26)) = "X21775000" Or ChangeCustomerL(pa(27)) = "X21775000" Or ChangeCustomerL(pa(28)) = "X21775000" Or ChangeCustomerL(pa(29)) = "X21775000" Or ChangeCustomerL(pa(30)) = "X21775000" Or _
'                           ChangeCustomerL(pa(75)) = "Y51304020" Then
'                           strTxt(intStep) = strTxt(intStep) & "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'                              "NP09,NP10,NP15,NP22) VALUES ('" & cp(9) & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & _
'                              pa(4) & "'," & 實體審查 & "," & Val(strTmpA(2)) & "," & Val(strTmpA(3)) & ",'" & cp(13) & "','不寄函逕收文;',intMax);"
'                        Else
'                           strTxt(intStep) = strTxt(intStep) & "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
'                              "NP09,NP10,NP22) VALUES ('" & cp(9) & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & _
'                              pa(4) & "'," & 實體審查 & "," & Val(strTmpA(2)) & "," & Val(strTmpA(3)) & ",'" & cp(13) & "',intMax);"
'                        End If
'                        '2008/10/13 END
                        '2011/10/12 改掛PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
                        'Modify By Sindy 2021/4/27 + ,NP23=" & CNULL(DBDATE(m_pAgreeOnDate)):約定期限
                        strTxt(intStep) = strTxt(intStep) & "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
                           "NP09,NP10,NP15,NP22,NP23) VALUES ('" & cp(9) & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & _
                           pa(4) & "'," & 實體審查 & "," & Val(strTmpA(2)) & "," & Val(strTmpA(3)) & ",'" & _
                           PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "','" & strMemo416 & "',intMax," & CNULL(DBDATE(m_pAgreeOnDate)) & ");"
'2011/10/12 end
                        strTxt(intStep) = strTxt(intStep) & " end;"
                        cnnConnection.Execute strTxt(intStep)
                        intStep = intStep + 1
                    End If
                '若案件進度檔有未發文的實體審查資料
                ElseIf "" & rsB("CP27") = "" Then
                    strTxt(intStep) = "Update CaseProgress Set CP06=" & Val(strTmpA(2)) & ",CP07=" & Val(strTmpA(3)) & " Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10='" & 實體審查 & "' And CP27 IS NULL "
                    cnnConnection.Execute strTxt(intStep)
                    intStep = intStep + 1
                End If
                If rsB.State <> adStateClosed Then rsB.Close
                Set rsB = Nothing
                
               'Add by Morgan 2006/5/1
               '發明案發文時若有收申請寄存108則須控管存活證明221期限
               If m_bol108 = True Then
                  'Added by Lydia 2022/09/15 抓約定期限; ex.FCP-67719的申請寄存108沒有先抓到約定期限(比照分案frm060101_1)
                  If Trim(strTmpA(2)) = "" Then strTmpA(2) = cp(6)
                  If Trim(strTmpA(3)) = "" Then strTmpA(3) = cp(7)
                  If m_pAgreeOnDate = "" Then
                    Call PUB_GetFCPOurDeadline(strTmpA(3), 4, , m_pAgreeOnDate)
                  End If
                  'end 2022/09/15
                  'Modify By Sindy 2021/4/27 + , DBDATE(m_pAgreeOnDate)
                  strSql = PUB_Get221SQL(cp, Val(strTmpA(2)), Val(strTmpA(3)), cp(13), cp(9), DBDATE(m_pAgreeOnDate))
                  cnnConnection.Execute strSql
               End If
               
            End If
         End If
      End If
   End If
   '2005/9/13 END
   
   'Add by Morgan 2004/10/21 用新案的發文日計算並更新 201,209,210 的期限
   'Modified by Morgan 2012/12/20 +衍生設計125
   'Modified by Morgan 2013/11/6 +235核對中說格式也要被更新
   If InStr("101,102,103,105,125", Text7(2).Text) > 0 Then
      'Modify by Morgan 2007/1/10 新型申請時期限改掛發文日次日起2個月
      'strExc(3) = CompDate(1, 4, TransDate(Text7(0), 2))
      If Text7(2).Text = "102" Then
         '2007/8/3 MODIFY BY SONIA 改掛發文日當日起2個月
         '2008/1/22 MODIFY BY SONIA 新型若無優先權證明文件才掛發文當日起2個月
         'Modified by Morgan 2013/8/22 102新法改固定發文日起2個月 --譚文容
         'If rsA.State <> adStateClosed Then rsA.Close
         'Set rsA = Nothing
         'StrSQLa = "Select NP01 From nextprogress Where np01='" & cp(9) & "' and np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07='202' and instr(np15,'優先權證明文件')>0 "
         'rsA.CursorLocation = adUseClient
         'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         'If rsA.RecordCount = 0 Then
         '   strExc(3) = CompDate(1, 2, TransDate(Text7(0), 2))
         'Else
         '   strExc(3) = CompDate(1, 4, TransDate(Text7(0), 2))
         'End If
         strExc(3) = CompDate(1, 2, TransDate(Text7(0), 2))
         'end 2013/8/22
         '2008/1/22 END
      Else
         strExc(3) = CompDate(1, 4, TransDate(Text7(0), 2))
      End If
      
      'Modified by Morgan 2014/11/20 外專改回舊規則
      ''Added by Morgan 2014/10/29
      'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
      '   strExc(0) = PUB_GetOurDeadline(strExc(3))
      'Else
      ''end 2014/10/29
      
      'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
      If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
         'Modify By Sindy 2021/4/27 + m_pAgreeOnDate
         strExc(0) = PUB_GetFCPOurDeadline(strExc(3), 4, , m_pAgreeOnDate)
      Else
      'end 2019/7/11
      
         strExc(0) = CompDate(2, -4, strExc(3))
         
      End If 'Added by Morgan 2019/7/11
      
      'End If 'Added by Morgan 2014/10/29
      'end 2014/11/20
      
'Removed by Morgan 2015/9/1 取消--靜芳
'      'Modify by Morgan 2005/9/8 加更新承辦/核稿期限/是否會稿
'      'strSQL = "Update Caseprogress set CP06=" & strExc(0) & ",CP07=" & strExc(3) & " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 IN('201','209','210') AND CP27 IS NULL AND CP57 IS NULL"
'      If bolXCase = True Then
'         stCtrlDate(0) = TransDate(Text7(0), 2)
'         stCtrlDate(1) = CompDate(2, 45, stCtrlDate(0))
'         stCtrlDate(1) = PUB_GetWorkDay1(stCtrlDate(1), True) '承辦期限
'         stCtrlDate(2) = CompDate(2, 58, stCtrlDate(0))
'         stCtrlDate(2) = PUB_GetWorkDay1(stCtrlDate(2), True) '核稿期限
'         If bolXCtrl = True Then
'            '2012/5/15 MODIFY BY SONIA 會稿改為存Y
'            strSql = " Begin" & _
'               " Update ENGINEERPROGRESS Set EP34='Y', EP08=" & stCtrlDate(2) & " Where EP02 IN (SELECT CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 IN('201','209','210') AND CP27 IS NULL AND CP57 IS NULL);" & _
'               " Update Caseprogress set CP06=" & strExc(0) & ",CP07=" & strExc(3) & ",CP48=" & stCtrlDate(1) & " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 IN('201','209','235','210') AND CP27 IS NULL AND CP57 IS NULL;" & _
'               " End;"
'         Else
'            strSql = " Begin" & _
'               " Update ENGINEERPROGRESS Set EP34='N' Where EP02 IN (SELECT CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 IN('201','209','210') AND CP27 IS NULL AND CP57 IS NULL);" & _
'               " Update Caseprogress set CP06=" & strExc(0) & ",CP07=" & strExc(3) & ",CP48=" & stCtrlDate(1) & " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 IN('201','209','235','210') AND CP27 IS NULL AND CP57 IS NULL;" & _
'               " End;"
'         End If
'      Else
         strSql = "Update Caseprogress set CP06=" & strExc(0) & ",CP07=" & strExc(3) & " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 IN('201','209','235','210') AND CP27 IS NULL AND CP57 IS NULL"
'      End If
'end 2015/9/1
      cnnConnection.Execute strSql, intI
      
      'Add By Sindy 2021/5/5【968回復說明書校閱】承辦期限：201新案翻譯本所期限前10個工作天
      'Add By Sindy 2021/7/14 【968回復說明書校閱】新案翻譯法定期限前15個工作天，本所期限再加5個工作天
      'Modify By Sindy 2022/2/16 + 核對中說格式209、檢視中說210、製作中說235(至於「210製作中說」屬設計案不會有968)
      If PUB_ChkCPExist(cp, "201", 1) Or PUB_ChkCPExist(cp, "209", 1) Or PUB_ChkCPExist(cp, "210", 1) Then
         'strExc(10) = CompWorkDay(10, CompDate(2, -1, strExc(0)), 1)
         strExc(10) = CompWorkDay(15, CompDate(2, -1, strExc(3)), 1)
         'strCP06 = ""
         'If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
            strCP06 = PUB_GetFCPOurDeadline(DBDATE(strExc(10)), , , , "N")
         'End If
         'Modify By Sindy 2025/7/22 + 請設定<Y5473200>此代理人，一併在進度備註加上:「原文說明書若有實質內容上的誤記或待確認事項，需於補呈中說期限前1~2週另函通知客戶，不可併入會稿結果報告信(Check Sheet相關回函)，若工程師核稿後認定無需報告，再請工程師通知承辦銷該收文。」
         '                          + ,cp64='" & strExc(9) & "'||cp64
         If ChangeCustomerL(pa(75)) = "Y54732000" Then
            strExc(9) = "「原文說明書若有實質內容上的誤記或待確認事項，需於補呈中說期限前1~2週另函通知客戶，不可併入會稿結果報告信(Check Sheet相關回函)，若工程師核稿後認定無需報告，再請工程師通知承辦銷該收文。」"
         Else
            strExc(9) = ""
         End If
         '2025/7/22 END
         strSql = "Update Caseprogress set CP48=" & strExc(10) & ",CP06='" & strCP06 & "',cp64='" & strExc(9) & "'||cp64" & _
                  " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='968' AND CP27 IS NULL AND CP57 IS NULL"
         cnnConnection.Execute strSql, intI
      End If
      '2021/5/5 END
      
      'Added by Lydia 2018/04/20 提申後告代/主動修正自動掛承辦期限
      '更新提申後告代：新案發文日起算6個工作天
       'Memo by Lydia 2021/08/23 規則：與frm090902_2
       '主動修正:
       '提申前:  本所期限 = 新案收文日起算15個工作天(本所期限若超過提申所限, 則以提申所限當本所期限)，承辦期限=本所期限往前-5個工作天
       '提申後: 1. 新案翻譯已發文並且與申請日同一天，承辦期限 = 新案發文日起算15個工作天，本所期限 = 承辦期限 + 再加5個工作天
       '              2. 新案翻譯未發文，本所期限 = 新案翻譯的本所期限，承辦期限=本所期限往前-5個工作天
       '告代:
       '提申前: 承辦期限 = 新案收文日起算(依案件國家收費表天數計算,FCP案5天,FG案6天,P案7天)個工作天，本所期限 = 承辦期限 + 再加5個工作天(本所期限若超過提申所限, 則以提申所限當本所期限)
       '提申後: 承辦期限 = 新案發文日起算(依案件國家收費表天數計算,FCP案5天,FG案6天,P案7天)個工作天，本所期限 = 承辦期限 + 再加5個工作天
       'end 2021/08/23
       strExc(1) = Pub_GetHandleDay(pa(1), pa(9), "901", TransDate(Text7(0), 2))
       If strExc(1) <> "" Then
            'Add By Sindy 2021/6/24
            strCP06 = ""
            If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
               strCP06 = PUB_GetFCPOurDeadline(DBDATE(strExc(1)), , , , "N")
            End If
            '2021/6/24 END
            'Modify By Sindy 2021/9/8 取消 and substr(cp09,1,1)='B' 因為A,B類都要更新
            strSql = "Update caseprogress  set CP48=" & strExc(1) & ",CP06='" & strCP06 & "' WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                       " and cp10 ='901' and instr(cp64,'提申後') > 0  and CP27 IS NULL AND CP57 IS NULL "
            cnnConnection.Execute strSql
       End If
      '更新主動修正：(提申後+新案翻譯未發文)=新案翻譯的本所期限
      'Add By Sindy 2021/6/24
      strCP06 = ""
      If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
         '本所期限=承辦期限(新案翻譯的本所期限)
         strCP06 = strExc(0)
         'Modify By Sindy 2021/8/10 承辦期限=本所期限往前-5個工作天
         strCP48 = CompWorkDay(5, CompDate(2, -1, DBDATE(strCP06)), 1)
      End If
      '2021/6/24 END
'新申請案101、102、103發文時，若有主動修正未發文：
'舊程式為:
'主動修正的承辦期限=新案翻譯本所期限，本所期限空白
'(參FCP65000,65222)
'修改為:
'主動修正的本所期限 = 新案翻譯的本所期限
'承辦期限=本所期限往前-5個工作天
'      strSql = "Update caseprogress set CP48=" & strExc(0) & ",CP06='" & strCP06 & "' WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                  " and cp10 ='203' and substr(cp09,1,1)='B' and instr(cp64,'提申後') > 0  and CP27 IS NULL AND CP57 IS NULL "
      'Modify By Sindy 2021/9/8 取消 and substr(cp09,1,1)='B' 因為A,B類都要更新
      strSql = "Update caseprogress set CP48=" & strCP48 & ",CP06='" & strCP06 & "' WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                  " and cp10 ='203' and instr(cp64,'提申後') > 0  and CP27 IS NULL AND CP57 IS NULL "
      cnnConnection.Execute strSql
      'end 2018/04/20
      
'Removed by Morgan 2015/9/1 取消--靜芳
'      'Add by Morgan 2010/4/29 Y34232申請案發文若核稿期限大於會稿之本所期限時，更新核稿期限為會稿之本所期限
'      'Modified by Morgan 2013/11/6 +235核對中說格式
'      If bolXCase = True Then
'         strSql = "update engineerprogress a set ep08=(select c.cp06 from caseprogress b,caseprogress c" & _
'            " where b.cp09=ep02 and c.cp01=b.cp01 and c.cp02=b.cp02 and c.cp03=b.cp03 and c.cp04=b.cp04 and c.cp10='924' and c.cp06<ep08)" & _
'            " where ep02 in (select b.cp09 from caseprogress b,caseprogress c,engineerprogress d" & _
'            " where b.cp01='" & cp(1) & "' and b.cp02='" & cp(2) & "' and b.cp03='" & cp(3) & "'" & _
'            " and b.cp04='" & cp(4) & "' and b.cp10 in('201','209','235','210') and c.cp01(+)=b.cp01" & _
'            " and c.cp02(+)=b.cp02 and c.cp03(+)=b.cp03 and c.cp04(+)=b.cp04 and c.cp10='924'" & _
'            " and d.ep02(+)=b.cp09 and c.cp06<d.ep08)"
'         cnnConnection.Execute strSql, intI
'      End If
'      'add by sonia 2014/6/23 申請案發文若核稿期限大於寄中說949之本所期限時，更新核稿期限為寄中說之本所期限
'      If bolXCase = True Then
'         strSql = "update engineerprogress a set ep08=(select c.cp06 from caseprogress b,caseprogress c" & _
'            " where b.cp09=ep02 and c.cp01=b.cp01 and c.cp02=b.cp02 and c.cp03=b.cp03 and c.cp04=b.cp04 and c.cp10='949' and c.cp06<ep08)" & _
'            " where ep02 in (select b.cp09 from caseprogress b,caseprogress c,engineerprogress d" & _
'            " where b.cp01='" & cp(1) & "' and b.cp02='" & cp(2) & "' and b.cp03='" & cp(3) & "'" & _
'            " and b.cp04='" & cp(4) & "' and b.cp10 in('201','209','235','210') and c.cp01(+)=b.cp01" & _
'            " and c.cp02(+)=b.cp02 and c.cp03(+)=b.cp03 and c.cp04(+)=b.cp04 and c.cp10='949'" & _
'            " and d.ep02(+)=b.cp09 and c.cp06<d.ep08)"
'         cnnConnection.Execute strSql, intI
'      End If
'end 2015/9/1

   End If
   '2004/10/21 END

   
   'Add by Morgan 2005/12/6 翻譯發文時檢查是否有關聯國外案以判斷是否提醒並發Mail
   m_strMailCP09 = ""
   
'Removed by Morgan 2013/1/3 取消
'
'   '2009/10/6 add by sonia 檢查若有未發文的主動修正,補充說明,則以此發文日+10工作天更新承辦期限,但不可大於該筆本所期限
'   'Modified by Morgan 2012/4/6 +檢視中說,製作中說
'   'If Text7(2) = "201" Then
'   If Text7(2) = "201" Or Text7(2) = "209" Or Text7(2) = "210" Then
'   'End 2012//4/6
'      m_203CP48 = CompWorkDay(10, DBDATE(Text7(0)), 0)
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      StrSQLa = "Select cp09,cp06 From caseprogress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 IN('203','206') AND CP27 IS NULL AND CP57 IS NULL "
'      rsA.CursorLocation = adUseClient
'      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'         rsA.MoveFirst
'         While Not rsA.EOF
'            If Not IsNull(rsA.Fields("cp06")) Then
'               If Val(m_203CP48) < Val(rsA.Fields("cp06")) Then
'                  strSql = "Update caseprogress SET cp48=" & DBDATE(m_203CP48) & " WHERE cp09='" & rsA.Fields("cp09").Value & "'"
'                  cnnConnection.Execute strSql, intI
'               End If
'            Else
'               strSql = "Update caseprogress SET cp48=" & DBDATE(m_203CP48) & " WHERE cp09='" & rsA.Fields("cp09").Value & "'"
'               cnnConnection.Execute strSql, intI
'            End If
'            rsA.MoveNext
'         Wend
'      End If
'   End If
'   '2009/10/6 end
'
'end 2013/1/3
   
   If Text7(2) = "201" Then

      strSql = "SELECT CM01,CM02,CM03,CM04 FROM CASEMAP,paTENT WHERE " & ChgCaseMap(pa(1) & pa(2) & pa(3) & pa(4), 0, 1) & " AND pa01(+)=CM01 AND pa02(+)=CM02 AND pa03(+)=CM03 AND pa04(+)=CM04 AND CM10='0' AND pa57 IS NULL"
      CheckOC
      With adoRecordset
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
         If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
               '更新國外案(未發文未取消收文未齊備)的承辦期限(CP48)
               strSql = "Select CP09,CP06 From CaseProgress,EngineerProgress WHERE " & ChgCaseprogress("" & .Fields(0) & .Fields(1) & .Fields(2) & .Fields(3)) & _
                            " AND To_Number(CP10)>=101 AND To_Number(CP10)<=104 And CP27 Is Null AND  CP57 IS NULL AND EP02=CP09 AND (EP06 IS NULL OR EP06=0)"
               CheckOC2
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
               If adoRecordset1.RecordCount > 0 Then
                  adoRecordset1.MoveFirst
                  While Not adoRecordset1.EOF
                     '更新文件齊備日
                     strExc(0) = "Update EngineerProgress Set EP06=" & strSrvDate(1) & " Where EP02='" & adoRecordset1.Fields("CP09").Value & "' AND EP06 IS NULL"
                     cnnConnection.Execute strExc(0), intI
                     
                     'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
                     '抓承辦天數
                     'strSQL = "Select NVL(CF04,0) From CaseProgress, patent, Casefee Where CP09='" & adoRecordset1.Fields("CP09").Value & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  and pa01=cf01(+) and pa09=cf02(+) and cp10=cf03 "
                     strSql = "Select cp01,cp10,pa09 From CaseProgress, patent Where CP09='" & adoRecordset1.Fields("CP09").Value & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
                     CheckOC3
                     AdoRecordSet3.CursorLocation = adUseClient
                     AdoRecordSet3.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
                     If AdoRecordSet3.RecordCount > 0 Then
                        'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
                        'If Val("" & AdoRecordSet3.Fields(0).Value) > 0 Then
                        '   '計算承辦期限
                        '    strExc(1) = CompWorkDay(AdoRecordSet3.Fields(0).Value, strSrvDate(1), 0)
                        '    '有本所期限
                        '    If Not IsNull(adoRecordset1.Fields(1).Value) Then
                        '        '若承辦期限大於本所期限
                        '        If strExc(1) > AdoRecordSet1.Fields(1).Value Then
                        '            strExc(1) = AdoRecordSet1.Fields(1).Value
                        '        End If
                        '    End If
                        strExc(1) = Pub_GetHandleDay(AdoRecordSet3.Fields("cp01"), AdoRecordSet3.Fields("pa09"), AdoRecordSet3.Fields("cp10"), , "" & adoRecordset1.Fields(1), adoRecordset1.Fields("CP09"))
                        If strExc(1) <> "" Then
                        'end 2007/10/11
                            '更新承辦期限
                            strExc(0) = "Update CaseProgress Set CP48=" & CNULL(strExc(1)) & " Where CP09='" & adoRecordset1.Fields("CP09").Value & "' "
                            cnnConnection.Execute strExc(0)
                            m_strMailCP09 = m_strMailCP09 & adoRecordset1.Fields("CP09").Value & ";"
                        End If
                     End If
                     adoRecordset1.MoveNext
                  Wend
               End If
               .MoveNext
            Loop
         End If
      End With
   End If
   '2005/12/6 END
   
   'Added by Lydia 2018/04/20 提申後主動修正自動掛承辦期限,(提申後+新案翻譯已發文=新案發文日)新案發文日起算15個工作天
   If InStr("201,209,235,210", Text7(2)) > 0 And Val(pa(10)) > 0 And Val(TransDate(pa(10), 2)) = Val(TransDate(Text7(0).Text, 2)) Then       '
       strExc(1) = Pub_GetHandleDay(pa(1), pa(9), "203", TransDate(pa(10), 2))
       strCP06 = PUB_GetFCPOurDeadline(DBDATE(strExc(1)), , , , "N") 'Add By Sindy 2021/8/11
       If strExc(1) <> "" Then
            'Add By Sindy 2021/8/11 + ,CP06='" & strCP06 & "'
            'Modify By Sindy 2021/9/8 取消 and substr(cp09,1,1)='B' 因為A,B類都要更新
            strSql = "Update caseprogress  set CP48=" & strExc(1) & ",CP06='" & strCP06 & "' WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                       " and cp10 ='203' and instr(cp64,'提申後') > 0  and CP27 IS NULL AND CP57 IS NULL "
            cnnConnection.Execute strSql, intI
       End If
   End If
   'Added by Lydia 2018/04/20 工程師提申後從命名系統修改專利名稱則有欄位註記，待程序從各式申請書產生申請書時，電子送件申請書在【備註】欄位自動加註"一併修改專利名稱"，待中說或補文件發文後自動將註記的欄位清空。
   If InStr("201,209,235,210,202", Text7(2)) > 0 Then
         strSql = "update transcasetitle set tct15=null where tct15='Y' and tct01 in (" & _
                    "select cp09 from caseprogress where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp31='Y') "
         cnnConnection.Execute strSql, intI
   End If
   'end 2018/04/20
   'Added by Lydia 2019/07/30 衍生設計新案發文時檢查命名記錄尚未分組,才刪除命名記錄
   If Text7(2) = "125" And m_TCTchk = "" Then
         strSql = "delete from transcasetitle where tct01='" & Label3(0) & "' "
         cnnConnection.Execute strSql, intI
   End If
   
   '2006/5/16 ADD BY SONIA 改請程序發文時, 原程序之催審期限要上Y
   If Text7(2) >= 改請發明 And Text7(2) <= 改請獨立 Then
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      StrSQLa = "Select CP43 From CASEprogress Where CP09='" & cp(43) & "'"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         '2010/2/4 MODIFY BY SONIA 改為 NP06='N'
         strTxt(intStep) = "Update NEXTPROGRESS SET NP06='N' WHERE NP01='" & rsA.Fields(0).Value & "' AND NP02='" & pa(1) & "' AND NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "'"
         cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
      End If
   End If
   '2006/5/16 END
   
   
'Removed by Morgan 2013/1/7 102 新法取消
'
'   'Add by Morgan 2007/8/28 若有主動修正未發文時更新期限 法限=申請日(最早優先權日)+15個月(新型2個月);所限=法限-4天
'   'Modified by Morgan 2012/3/5  +206 補充說明,並控制法限遇假日順延(與分案相同)
'   If cp(10) = "101" Or cp(10) = "102" Then
'      strExc(0) = "select cp09,pd05 from caseprogress,pridate" & _
'         " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
'         " and cp10 in ('203','206') and cp27 is null and cp57 is null" & _
'         " and pd01(+)=cp01 and pd02(+)=cp02 and pd03(+)=cp03 and pd04(+)=cp04" & _
'         " order by pd05"
'
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         '主動修正收文號
'         strExc(1) = RsTemp.Fields(0)
'         '發明
'         If cp(10) = "101" Then
'            '最早優先權日
'            If Not IsNull(RsTemp.Fields(1)) Then
'               strExc(2) = RsTemp.Fields(1)
'            '發文日
'            Else
'               strExc(2) = DBDATE(Text7(0))
'            End If
'         '新型
'         Else
'            strExc(2) = DBDATE(Text7(0))
'         End If
'         If strExc(2) <> "" Then
'            '發明
'            If cp(10) = "101" Then
'               strExc(3) = CompDate(1, 15, strExc(2))
'            '新型
'            Else
'               strExc(3) = CompDate(1, 2, strExc(2))
'            End If
'            strExc(3) = PUB_GetWorkDay1(strExc(3), False) 'Added by Morgan 2012/3/5
'            strExc(4) = CompDate(2, -4, strExc(3))
'            strSql = "update caseprogress set cp06=" & strExc(4) & ",cp07=" & strExc(3) & " where cp09='" & strExc(1) & "'"
'            cnnConnection.Execute strSql, intI
'         End If
'      End If
'   End If
'   'End 2007/8/28
'
'end 2013/1/7
   
   'Add by Morgan 2007/9/13 自撤發文時若相關總收文為重新委任時要將該收文號的未收文的補文件上不續辦
   If Text7(2) = "413" And m_RefCP10 = "928" And cp(43) <> "" Then
      strSql = "update nextprogress set np06='N' where np01='" & cp(43) & "' and np06 is null and np07='202' and exists(select * from caseprogress where cp09=np01 and cp10='928' and cp27>0)"
      cnnConnection.Execute strSql, intI
   End If
   'end 2007/9/13
   
   '2007/10/8 add by sonia 自撤發文時更新其相關總收文下一程序的催審為不續辦FCP-030162
   If Text7(2) = "413" And m_RefCP10 <> "928" And cp(43) <> "" Then
      strSql = "update nextprogress set np06='N' where np01='" & cp(43) & "' and np06 is null and np07='411' and np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
      cnnConnection.Execute strSql, intI
   End If
   '2007/10/8 END
   
   'Add by Morgan 2008/10/14 重新委任之補文件發文，若下一程序有催審不續辦且備註為"重新委任補文件未收文"時，新增一筆催審
   If Text7(2) = "202" And cp(43) <> "" Then
      strExc(0) = "select np01,np09 from caseprogress,nextprogress where cp09='" & cp(43) & "' and cp10='928' and np02(+)=cp01 and np03(+)=cp02 and np04(+)=cp03 and np05(+)=cp04 and np06='N' and np15='重新委任補文件未收文'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         intMax = GetNextProgressNo
         With RsTemp
            '未過期帶原期限
            If .Fields("np09") > strSrvDate(1) Then
               strExc(1) = .Fields("np09")
            '已過期改發文日+3月
            Else
               strExc(1) = CompDate(1, 3, Text7(0))
            End If
            'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
            strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
               "NP07,NP08,NP09,NP10,NP22) VALUES ('" & .Fields("np01") & "','" & pa(1) & _
               "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 催審 & "," & _
               PUB_GetWorkDay1(strExc(1), True) & "," & strExc(1) & ",'" & strUserNum & "'," & intMax & ")"
            cnnConnection.Execute strSql, intI
         End With
      End If
   End If
   
   'Add by Morgan 2010/8/10
   '申復,補充,修正+14個月更新催審期限
   If Text7(2) = "204" Or Text7(2) = "205" Or Text7(2) = "206" And cp(43) > "C" Then
      strExc(1) = CompDate(1, 14, Text7(0))
      'Modified by Morgan 2012/12/20 +衍生設計125,改請衍生設計308
      'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
      strSql = "Update NextProgress Set NP08=" & PUB_GetWorkDay1(strExc(1), True) & ",NP09=" & strExc(1) & " WHERE NP02='" & cp(1) & "' and NP03='" & cp(2) & "'" & _
         " and NP04='" & cp(3) & "' and NP05='" & cp(4) & "' AND NP07='" & 催審 & "' AND NP06 IS NULL" & _
         " and exists(select * from caseprogress where cp09=np01 and instr('101,102,103,104,105,125,301,302,303,304,305,306,307,308,107',cp10)>0)"
      cnnConnection.Execute strSql, intI
   End If
   
   'Add by Morgan 2010/1/5 補收文超頁超項費
   'Add By Sindy 2019/4/24 增加檢查若無規費,就不需要產生超頁超項費進度 ex:FCP-60163中說發文
   If Val(txtCP84) > 0 Then
   '2019/4/24 END
      If m_lngOverPageFee + m_lngOverItemFee > 0 Then
         If m_lngOverPageFee > 0 Then
            'Add by Morgan 2011/6/29
            '已收文
            If m_lngRecOverPageFee > 0 Then
               '更新發文日及規費
               strSql = "Update caseprogress set cp16=cp16-cp17+" & m_lngOverPageFee & ",cp17=" & m_lngOverPageFee & ",cp27=" & DBDATE(Text7(0)) & ",cp84=0" & _
                  " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='938' and cp27||cp57 is null and rownum<2"
               cnnConnection.Execute strSql, intI
            '未收文
            Else
            'end 2011/6/29
               strExc(1) = AutoNo("B", 6) 'B類總收文號
               strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp16,cp17,cp18,cp26,cp27,cp43,cp84) values " & _
                  " ('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & _
                  ",'" & strExc(1) & "','938','" & stCP12 & "','" & stCP13 & "','" & strUserNum & "'," & m_lngOverPageFee & "," & m_lngOverPageFee & ",0,'N'" & _
                  "," & DBDATE(Text7(0)) & ",'" & cp(9) & "',0)"
               cnnConnection.Execute strSql, intI
            End If 'Add by Morgan 2011/6/29
         End If
         If m_lngOverItemFee > 0 Then
            'Add by Morgan 2011/6/29
            '已收文
            If m_lngRecOverItemFee > 0 Then
               '更新發文日及規費
               strSql = "Update caseprogress set cp16=cp16-cp17+" & m_lngOverItemFee & ",cp17=" & m_lngOverItemFee & ",cp27=" & DBDATE(Text7(0)) & ",cp84=0" & _
                  " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='939' and cp27||cp57 is null and rownum<2"
               cnnConnection.Execute strSql, intI
            '未收文
            Else
            'end 2011/6/29
               strExc(1) = AutoNo("B", 6) 'B類總收文號
               strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp16,cp17,cp18,cp26,cp27,cp43,cp84) values " & _
                  " ('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & _
                  ",'" & strExc(1) & "','939','" & stCP12 & "','" & stCP13 & "','" & strUserNum & "'," & m_lngOverItemFee & "," & m_lngOverItemFee & ",0,'N'" & _
                  "," & DBDATE(Text7(0)) & ",'" & cp(9) & "',0)"
               cnnConnection.Execute strSql, intI
            End If 'Add by Morgan 2011/6/29
         End If
      End If
   End If
   
   PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130 'Add by Morgan 2009/3/20
   
   'Add by Morgan 2010/4/12
   '工程師提申發文時更新新申請案程序為不請款
   If Text7(2) = "940" Then
      'Modified by Morgan 2012/12/20 +衍生設計125
      If Val(cp(16)) = 0 Then
         strSql = "update caseprogress c1 set (cp16,cp17,cp18)=(select c2.cp16,c2.cp17,c2.cp18 from caseprogress c2 where c2.cp01=c1.cp01" & _
            " and c2.cp02=c1.cp02 and c2.cp03=c1.cp03 and c2.cp04=c1.cp04 and c2.cp10 in ('101','102','103','105','125') and c2.cp16>0 and c2.cp60||c2.cp57 is null)" & _
            " where cp09='" & cp(9) & "' and exists(select * from caseprogress c2 where c2.cp01=c1.cp01" & _
            " and c2.cp02=c1.cp02 and c2.cp03=c1.cp03 and c2.cp04=c1.cp04 and c2.cp10 in ('101','102','103','105','125') and c2.cp16>0 and c2.cp60||c2.cp57 is null)"
         cnnConnection.Execute strSql, intI
      End If
      
      strSql = "update caseprogress set cp16=0,cp17=0,cp18=0,cp20='N' where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('101','102','103','105','125') and cp16>0 and cp60||cp57 is null"
      cnnConnection.Execute strSql, intI
   End If
   
   'Added by Morgan 2012/8/1
   If Text7(20).Visible = True Then
      If Text7(20) <> "" Then
         strExc(1) = DBDATE(Text7(20))
         
         'Modified by Morgan 2014/11/20 外專改回舊規則
         ''Added by Morgan 2014/10/29
         'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
         '   strExc(2) = PUB_GetOurDeadline(strExc(1))
         'Else
         ''end 2014/10/29
         
         'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
         If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
            'Modify By Sindy 2021/4/27 + m_pAgreeOnDate
            strExc(2) = PUB_GetFCPOurDeadline(strExc(1), 2, , m_pAgreeOnDate)
         Else
         'end 2019/7/11
      
            strExc(2) = CompDate(2, -2, strExc(1))
            
         End If 'Added by Morgan 2019/7/11
         
         'End If 'Added by Morgan 2014/10/29
         'end 2014/11/20
         
         'Modify By Sindy 2021/4/27 + ,NP23=" & CNULL(DBDATE(m_pAgreeOnDate)):約定期限
         strSql = "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22,NP23) " & _
            " Values ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & cp(10) & "," & strExc(2) & "," & strExc(1) & _
            ",'" & PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "','例行;',GETNP22," & CNULL(DBDATE(m_pAgreeOnDate)) & ")"
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2012/8/1
   
   'Add by Amy 2013/08/22 代理人Y49456 發明申請案發文時提示是否實審要 交承辦收文告代，是則寫入下一程序備註
   If bol416Msg Then
         strSql = "Update NextProgress set NP15='交承辦收文告代'||NP15 Where " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07='416' "
         cnnConnection.Execute strSql, intI
   End If
   'end 2013/08/22
   
   'Add by Lydia 2014/12/24 代辦退費發文時,進度檔自動產生一道"自請撤回"(413,B類單),收文日為系統日;發文日同代辦退費之發文日;自請撤回的相關收文號掛"申請"那一道
   If Text7(2) = "908" Then
        'Modified by Morgan 2013/6/6 +檢查再審延期
        'strExc(0) = "select 1 from caseprogress a,caseprogress b where a.cp09='" & strCP09 & "' and b.cp09(+)=a.cp43 and b.cp10 in ('416','107')"
        'Modified by Morgan 2022/10/12 +435續行母案再審
        strExc(0) = "select 1 from caseprogress a,caseprogress b where a.cp09='" & cp(9) & "' and b.cp09(+)=a.cp43 and b.cp10 in ('416','107','435')" & _
           " union select 2 from  caseprogress a,caseprogress b,nextprogress where a.cp09='" & cp(9) & "' and b.cp09(+)=a.cp43 and b.cp10='404' and np01(+)=b.cp43 and np07='107'" & _
           " union select 3 from  caseprogress a,caseprogress b,caseprogress c where a.cp09='" & cp(9) & "' and b.cp09(+)=a.cp43 and b.cp10='404' and c.cp09(+)=b.cp43 and c.cp10='107'"
          
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            strExc(1) = AutoNo("B", 6) 'B類總收文號
            strExc(9) = PUB_GetFCPHandler(cp(1), cp(2), cp(3), cp(4))
            
            strExc(0) = "select cp09 from caseprogress " & _
             " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'" & _
             " and instr('" & NewCasePtyList & "',cp10)>0 order by cp05"
            
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
             strExc(0) = RsTemp.Fields(0) '抓新案的總收文號
            End If
            
            strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp26,cp27,cp43) values " & _
               " ('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & _
               ",'" & strExc(1) & "','413','" & stCP12 & "','" & stCP13 & "','" & strExc(9) & "','N'" & _
               "," & DBDATE(Text7(0)) & ",'" & strExc(0) & "')"
            cnnConnection.Execute strSql, intI
        End If
   End If
   
   'Added by Lydia 2015/02/26 若已開請款單則換承辦人或核稿人時發Mail通知靜芳
   If cp(60) > "X" Then
      'Modified by Lydia 2019/10/17 本所案號+"-"
      'PUB_PointReAssignInform Text1 & Text2 & Text3 & Text4, cp(60), Text7(1).Tag, Text7(1).Text, txtEP04.Tag, txtEP04.Text
      PUB_PointReAssignInform pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)), cp(60), Text7(1).Tag, Text7(1).Text, txtEP04.Tag, txtEP04.Text
   End If
   
   'Added by Lydia 2015/12/31 FCP會稿發文存檔時，自動新增國外部行事曆資料
   'Modified by Lydia 2016/01/21 +詢問是否產生(bolAddSC)
   If Text7(2) = "924" And txtCP43.Visible = True And m_CP43cpm <> "" And bolAddSC = True Then
      '若為新案翻譯201、檢視中說209、製作中說210，管制日期=相關總收文號之本所期限前二週
      'Modified by Lydia 2019/07/03 只有中說之會稿要管制
      'If InStr("201,209,210", m_CP43cpm) > 0 Then
      '   strExc(1) = CompDate(2, -14, m_CP43date1)
      ''非新案翻譯，管制日期=相關總收文號之本所期限前一週
      'Else
      '   strExc(1) = CompDate(2, -7, m_CP43date1)
      'End If
      strExc(1) = CompDate(2, -14, m_CP43date1)
      'end 2019/07/03
      '若結果小於發文日時則放發文日
      If strExc(1) < DBDATE(Text7(0)) Then strExc(1) = DBDATE(Text7(0))
      strExc(1) = PUB_GetWorkDay1(strExc(1), True) 'Added by Lydia 2025/11/12 改抓最近工作天
      
      '提醒人員1預設為該本所案號之FCP承辦業務員，提醒人員2預設為該案之最後工程師；可解除人員預設為提醒人員1；事由='追蹤會稿結果'；本所案號也要存；
      strExc(3) = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
      strExc(5) = ""
      'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
      strExc(0) = "select cp14,st04,st02 from caseprogress,staff where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' " & _
                  "and st01(+)=cp14 and st03='F21' and cp14<>'F4102' and cp57 is null order by cp05 desc,cp09 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RTrim("" & RsTemp("st04")) = "1" Then
            strExc(5) = "" & RsTemp.Fields("cp14")
         End If
      End If
      If strExc(3) <> "" Then
          strExc(4) = "追蹤會稿結果"
          strExc(2) = strExc(3) & IIf(strExc(5) <> "", ",", "") & strExc(5)
          '可解除人員預設為提醒人員1
          If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(2), strExc(4), strExc(3), "1", cp(1), cp(2), cp(3), cp(4)) Then
          End If
      End If
   End If
   'end 2015/12/31
   
   'Added by Lydia 2017/10/11 FCP電子送件若發文時若有規費，則自動產生行事曆。
   If stCP118 = "A" Then '電子送件cp118=Y,規費cp84>0 =>自動扣款CP118=A
      'stCP152:若是當天扣款=Y則抓隔一工作天，若當天扣款"N"則抓隔二工作天
      'Modified by Lydia 2018/09/11 改成模組
'      strExc(3) = PUB_GetFCPHandler(cp(1), cp(2), cp(3), cp(4), cp(10))
'      strExc(5) = ""
'      strExc(5) = GetABS001_17(strExc(3))
'      If strExc(3) <> "" Then
'          strExc(4) = "請當日下午下載電子送件之收據"
'          '解除人員請抓"管制人"及其"案件第一職代"
'          strExc(2) = strExc(3) & IIf(strExc(5) <> "", ",", "") & strExc(5)
'          'Added by Lydia 2018/06/15 發文人員不是期限管制人請增加"發文人員"到提醒人員及解除人員
'          If strUserNum <> strExc(3) Then
'               strExc(3) = strExc(3) & "," & strUserNum
'               If InStr(strExc(2), strUserNum) = 0 Then
'                   strExc(2) = strExc(2) & "," & strUserNum
'               End If
'          End If
'          'end 2018/06/15
'          '提醒人員請抓"管制人"
'          If PUB_AddFCPStaffCalendar(stCP152, "1", strExc(3), strExc(4), strExc(2), "1", cp(1), cp(2), cp(3), cp(4)) Then
'          End If
'      End If
      If Pub_AddReceiptCalendar1(cp(1), cp(2), cp(3), cp(4), cp(10), stCP152) = True Then
      End If
      'end 2018/09/11
   End If
   'end 2017/10/11
   
   'Modify By Sindy 2021/7/21 Move到FormSave外層詢問,及改新增下一程序不存放行事曆
   If mDate209210 <> "" Then
      intMax = GetNextProgressNo
      strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
         "NP10,NP15,NP22,NP23) VALUES ('" & strReceiveNo & "','" & pa(1) & _
         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','202'" & _
         "," & DBDATE(mDate209210) & ",'" & _
         PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "','客戶提供中說（簡/繁）'" & _
         "," & intMax & "," & DBDATE(mDate209210) & ")"
      cnnConnection.Execute strSql
   End If
   If mDateTF30 <> "" Then
      intMax = GetNextProgressNo
      strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08," & _
         "NP10,NP15,NP22,NP23) VALUES ('" & strReceiveNo & "','" & pa(1) & _
         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','202'" & _
         "," & DBDATE(mDateTF30) & ",'" & _
         PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) & "','英文參考本'" & _
         "," & intMax & "," & DBDATE(mDateTF30) & ")"
      cnnConnection.Execute strSql
   End If
   '2021/7/21 END
'   'Added by Lydia 2019/01/02 新案101,102發文時若檢視中說209 or核對中說格式235 未發文，請設行事曆
'   If Text7(2) = "101" Or Text7(2) = "102" Then
'      strExc(0) = "select count(*) cnt from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' " & _
'                        "and cp10 in ('209','235') and cp158=0 and cp159=0 "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If Val("" & RsTemp.Fields("cnt")) > 0 Then
'             If MsgBox("是否管制客戶提供中說期限？" & vbCrLf & "選是：產生行事曆" & vbCrLf & "選否：不產生行事曆，繼續發文", vbYesNo + vbInformation + vbDefaultButton1) = vbYes Then
'JumpReInput:
'                'Modified by Lydia 2019/01/09 +自訂日期
'                strExc(2) = UCase(InputBox("管制期限為提申後(新案發文日)＋１～３個月，" & vbCrLf & "請在下方輸入1~3或自訂日期(民國年月日)：", "管制客戶提供中說期限", "2"))
'                If strExc(2) = "" Then
'                    'Modified by Lydia 2019/01/09 +自訂日期
'                    If MsgBox("未輸入1~3或自訂日期，是否要管制客戶提供中說期限？", vbYesNo + vbDefaultButton1) = vbYes Then
'                         GoTo JumpReInput
'                    End If
'                Else
'                    'Modified by Lydia 2019/01/09 +自訂日期
'                    'If InStr("1,2,3", strExc(2)) = 0 And Val(strExc(2)) = 0 Then
'                    '    MsgBox "請輸入管制期限為提申後(新案發文日)＋１～３個月！", vbCritical
'                    If (Len(strExc(2)) = 1 And InStr("1,2,3", strExc(2)) = 0) Or Val(strExc(2)) = 0 Or (Len(strExc(2)) > 1 And Len(strExc(2)) <> 7) Then
'                        MsgBox "請輸入管制期限為提申後(新案發文日)＋１～３個月或自訂日期(民國年月日)！", vbCritical
'                        GoTo JumpReInput
'                    End If
'                    'Added by Lydia 2019/01/09 檢查日期
'                    If Len(strExc(2)) = 7 Then
'                        If ChkDate(strExc(2)) = False Then
'                            GoTo JumpReInput
'                        End If
'                    End If
'                    'end 2019/01/09
'                End If
'                If strExc(2) <> "" Then
'                   '新案發文日起算X個月,日期若碰到放假則往前抓1個工作天。
'                    'Added by Lydia 2019/01/09 自訂日期
'                    If Len(strExc(2)) > 1 Then
'                        strExc(1) = CompWorkDay(1, DBDATE(strExc(2)), 1)
'                    Else
'                    'end 2019/01/09
'                        strExc(1) = CompWorkDay(1, CompDate(1, strExc(2), DBDATE(Text7(0))), 1)
'                    End If
'                    'end 2019/01/09
'                    mDate209210 = strExc(1) 'Added by Lydia 2019/01/17
'                    strExc(4) = "催客戶提供中說期限"
'                    strExc(3) = PUB_GetFCPHandler(cp(1), cp(2), cp(3), cp(4))
'                    strExc(5) = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
'                    If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(3) & "," & strExc(5), strExc(4), strExc(3) & "," & strExc(5), "1", cp(1), cp(2), cp(3), cp(4)) Then
'                    End If
'                End If
'             End If
'         End If
'      End If
'   End If
'   'end 2019/01/02
'
'   'Added by Lydia 2019/12/11  FCP新案發文時檢查有新案翻譯未發文並且尚"待英文本翻譯"TF30='Y'，
'   '則彈訊息"是否管制催客戶提供英文翻譯本之行事曆"並比照催客戶提供英文翻譯本的行事曆方式，其內容為" 催客戶提供英文翻譯本"。
'   If InStr(NewCasePtyList, Text7(2)) > 0 Then
'      strExc(0) = "select cp09,tf30 from caseprogress,transfee where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' " & _
'                        "and cp10='201' and cp158=0 and cp159=0 and cp09=tf01(+)"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If "" & RsTemp.Fields("tf30") = "Y" Then
'             If MsgBox("是否管制催客戶提供英文翻譯本之行事曆？" & vbCrLf & "選是：產生行事曆" & vbCrLf & "選否：不產生行事曆，繼續發文", vbYesNo + vbInformation + vbDefaultButton1) = vbYes Then
'JumpReInput2:
'                strExc(2) = UCase(InputBox("管制期限為提申後(新案發文日)＋１～３個月，" & vbCrLf & "請在下方輸入1~3或自訂日期(民國年月日)：", "催客戶提供英文翻譯本", "2"))
'                If strExc(2) = "" Then
'                    If MsgBox("未輸入1~3或自訂日期，是否要管制催客戶提供英文翻譯本期限？", vbYesNo + vbDefaultButton1) = vbYes Then
'                         GoTo JumpReInput2
'                    End If
'                Else
'                    If (Len(strExc(2)) = 1 And InStr("1,2,3", strExc(2)) = 0) Or Val(strExc(2)) = 0 Or (Len(strExc(2)) > 1 And Len(strExc(2)) <> 7) Then
'                        MsgBox "請輸入管制期限為提申後(新案發文日)＋１～３個月或自訂日期(民國年月日)！", vbCritical
'                        GoTo JumpReInput2
'                    End If
'                    If Len(strExc(2)) = 7 Then
'                        If ChkDate(strExc(2)) = False Then
'                            GoTo JumpReInput2
'                        End If
'                    End If
'                End If
'                If strExc(2) <> "" Then
'                   '新案發文日起算X個月,日期若碰到放假則往前抓1個工作天。
'                    If Len(strExc(2)) > 1 Then  '自訂日期
'                        strExc(1) = CompWorkDay(1, DBDATE(strExc(2)), 1)
'                    Else
'                        strExc(1) = CompWorkDay(1, CompDate(1, strExc(2), DBDATE(Text7(0))), 1)
'                    End If
'                    mDateTF30 = strExc(1)
'                    strExc(4) = "催客戶提供英文翻譯本"
'                    strExc(3) = PUB_GetFCPHandler(cp(1), cp(2), cp(3), cp(4))
'                    strExc(5) = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4))
'                    If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(3) & "," & strExc(5), strExc(4), strExc(3) & "," & strExc(5), "1", cp(1), cp(2), cp(3), cp(4)) Then
'                    End If
'                End If
'             End If
'         End If
'      End If
   If InStr(NewCasePtyList, Text7(2)) > 0 Then
      'Added by Lydia 2020/10/14 Murgitroyd呈送期限設定: 翻譯中說－ 提申日起算3月內完成中說並報告
      strExc(0) = Pub_GetSpecMan("外專MURGITROYD設定")
      If strExc(0) <> "" And InStr(strExc(0), ChangeCustomerL(pa(75))) > 0 And pa(75) <> "" Then
          'Modified by Lydia 2021/01/06 抓翻譯中說的所限
          'If PUB_ChkCPExist(cp, "201", 1, strExc(9)) = True Then 'Added by Lydia 2020/12/09 有新案翻譯才做
          strSql = "select cp09,cp06 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='201' and cp158=0 and cp159=0 "
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strSql)
          If intI = 1 Then
                strExc(6) = "" & RsTemp.Fields("cp06") '中說的所限
                strExc(9) = "" & RsTemp.Fields("cp09") '中說收文號
          'end 2021/01/06
                'Modified by Lydia 2020/10/19 改為「新案發文日＋2.5個月再往前推3個工作天」(發生中說翻譯不滿3個月，MURGITROYD即來函催促的情形)
                'strExc(1) = CompWorkDay(4, CompDate(1, 3, DBDATE(Text7(0))), 1)
                strExc(1) = CompWorkDay(4, CompDate(2, 15, CompDate(1, 2, DBDATE(Text7(0)))), 1)
                '自動帶翻譯進度備註：為Murgitroyd案需xx月xx日（計算結果）完成中說並報告
                strExc(4) = "為Murgitroyd案需" & ChangeWStringToTDateString(strExc(1)) & "完成中說並報告"
                'Added by Lydia 2021/01/06 新案發文時一併設定中說核稿期限
                'Remove by Lydia 2021/11/03 經過確認會影響國外部期限的「未交稿,已完稿無核稿管制」並且也不符現況，所以拿掉這項更新 --- Sharon
                'If PUB_FCPsetEP08M(cp(1), cp(2), cp(3), cp(4), strExc(6), cp(10), DBDATE(Text7(0)), strExc(8), False) = True Then
                '     strSql = "update Engineerprogress set ep08='" & strExc(8) & "' where ep02='" & strExc(9) & "' and ep08 is null"
                '     cnnConnection.Execute strSql
                'End If
                ''end 2021/01/06
                strExc(3) = PUB_GetFCPHandler(cp(1), cp(2), cp(3), cp(4))
                If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(3), strExc(4), strExc(3), "1", cp(1), cp(2), cp(3), cp(4)) = True Then
                    '翻譯交稿期限：提申日起1個月期限
                    'If PUB_ChkCPExist(cp, "201", 1, strExc(9)) = True Then 'Remove by Lydia 2020/12/09 debug
                        strSql = "Update CaseProgress set cp64=" & CNULL(strExc(4)) & "||';'||cp64 where cp09=" & CNULL(strExc(9))
                        cnnConnection.Execute strSql
                        'Modified by Lydia 2020/10/22 新增行事曆
                        'strSql = "Update TransFee set TF26=" & CompWorkDay(1, CompDate(1, 1, DBDATE(Text7(0))), 1) & " where TF01=" & CNULL(strExc(9))
                        'cnnConnection.Execute strSql
                        strExc(1) = CompWorkDay(1, CompDate(1, 1, DBDATE(Text7(0))), 1)
                        strSql = "Update TransFee set TF26=" & strExc(1) & " where TF01=" & CNULL(strExc(9))
                        cnnConnection.Execute strSql
                        strExc(2) = Pub_GetSpecMan("M") 'Added by Lydia 2020/10/26 國外部專利處翻譯未交稿管制人
                        'Modified by Lydia 2020/10/26 增加提醒人員: 翻譯管制人
                        'If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(3), "譯者翻譯交稿期限", strExc(3), "1", cp(1), cp(2), cp(3), cp(4)) = True Then
                        If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(3) & IIf(strExc(3) <> strExc(2), "," & strExc(2), ""), "譯者翻譯交稿期限", strExc(3), "1", cp(1), cp(2), cp(3), cp(4)) = True Then
                        End If
                        'end 2020/10/22
                    'End If 'Remove by Lydia 2020/12/09 debug
                End If

          End If 'Added by Lydia 2020/12/09
      End If
      'end 2020/10/14
   End If
   'end 2019/12/11
   
   'Add By Sindy 2017/5/12 實審請款若當初不繳規費，之後才補繳規費，會收文"補收款"，來繳納實審的申請費7000元，
   '實審請款時會漏請"補收款"此道程序
   '若是以上狀況，發文"補收款(相關收文號為實體審查之收文號)"時，請先檢查"實體審查"的發文規費為0及是否還未請款，
   '若符合以上二個條件，則把"補收款(相關收文號為實體審查之收文號)"的收文規費回寫到實體審查的收文規費，
   '且自動將"補收款"上不請款。
   'Modify By Sindy 取消鎖定實體審查案件性質( and b.cp10 in ('416') )
   If Text7(2) = "911" Then '補收款
        strExc(0) = "select b.cp09,a.cp17 from caseprogress a,caseprogress b" & _
                    " where a.cp09='" & cp(9) & "' and b.cp09(+)=a.cp43" & _
                    " and b.cp158>0 and nvl(b.cp84,0)=0" & _
                    " and b.cp60 is null"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            'Modified by Morgan 2017/11/8 費用要加上點數
            'strSql = "Update caseprogress set cp17=" & RsTemp.Fields(1) & ",cp16=" & RsTemp.Fields(1) & ",cp18=0 where cp09='" & RsTemp.Fields(0) & "'"
            strSql = "Update caseprogress set cp17=" & RsTemp.Fields(1) & ",cp16=" & RsTemp.Fields(1) & "+nvl(cp18,0)*1000 where cp09='" & RsTemp.Fields(0) & "'"
            cnnConnection.Execute strSql, intI
            'modify by sonia 2021/3/25 既然回寫相關總收文號，補收款本身應取消規費。另配合財務處有發文規費不可列不請款，並於請款時一併自動點入請款單
            'strSql = "Update caseprogress set cp20='N' where cp09='" & cp(9) & "'"
            strSql = "Update caseprogress set cp16=cp16-" & RsTemp.Fields(1) & ",cp17=cp17-" & RsTemp.Fields(1) & " where cp09='" & cp(9) & "'"
            cnnConnection.Execute strSql, intI
        End If
   End If
   '2017/5/12 END
   
   'Add By Sindy 2018/4/19 非變更案,但有變更檔則更新個案地址
   If InStr(翻譯 & "," & 檢視中說 & "," & 製作中說 & ",235," & 補文件, Text7(2)) > 0 And _
      IsChangeEventExist(strReceiveNo) = True Then
      
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      StrSQLa = "SELECT * FROM CHANGEEVENT WHERE CE01='" & strReceiveNo & "'"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      strExc(9) = ""
      If rsA.RecordCount > 0 Then
         If "" & rsA.Fields("CE23") <> "" Then strExc(9) = strExc(9) & ",pa31='" & ChgSQL(rsA.Fields("CE23")) & "'"
         If "" & rsA.Fields("CE26") <> "" Then strExc(9) = strExc(9) & ",pa32='" & ChgSQL(rsA.Fields("CE26")) & "'"
         If "" & rsA.Fields("CE29") <> "" Then strExc(9) = strExc(9) & ",pa33='" & ChgSQL(rsA.Fields("CE29")) & "'"
         If "" & rsA.Fields("CE32") <> "" Then strExc(9) = strExc(9) & ",pa34='" & ChgSQL(rsA.Fields("CE32")) & "'"
         If "" & rsA.Fields("CE35") <> "" Then strExc(9) = strExc(9) & ",pa35='" & ChgSQL(rsA.Fields("CE35")) & "'"
         If "" & rsA.Fields("CE24") <> "" Then strExc(9) = strExc(9) & ",pa36='" & ChgSQL(rsA.Fields("CE24")) & "'"
         If "" & rsA.Fields("CE27") <> "" Then strExc(9) = strExc(9) & ",pa37='" & ChgSQL(rsA.Fields("CE27")) & "'"
         If "" & rsA.Fields("CE30") <> "" Then strExc(9) = strExc(9) & ",pa38='" & ChgSQL(rsA.Fields("CE30")) & "'"
         If "" & rsA.Fields("CE33") <> "" Then strExc(9) = strExc(9) & ",pa39='" & ChgSQL(rsA.Fields("CE33")) & "'"
         If "" & rsA.Fields("CE36") <> "" Then strExc(9) = strExc(9) & ",pa40='" & ChgSQL(rsA.Fields("CE36")) & "'"
         If strExc(9) <> "" Then
            strExc(9) = Mid(strExc(9), 2)
            strSql = "Update Patent set " & strExc(9) & " where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & cp(3) & "' and pa04='" & cp(4) & "'"
            cnnConnection.Execute strSql, intI
         End If
      End If
   End If
   '2018/4/19 END
   
   'Added by Lydia 2019/10/25 FCP新案翻譯增加"翻譯瑕疵"
   If fraTrans04.Visible = True And lblTF37.Tag <> "" Then
       strSql = "Update TransFee set TF37=" & CNULL(ChgSQL(txtTF37.Text)) & " Where TF01=" & CNULL(lblTF37.Tag)
       cnnConnection.Execute strSql, intI
   End If
   'end 2019/10/25
   
   'Added by Morgan 2019/11/12 --何淑華
   '代辦退費-相關收文號為領證601,發文時需將下一程序-年費期限刪除
   '需整筆刪除,不能上N,因系統有2道不同期限以為6個月逾繳,故通知年費沒通知,例FCP54168
   If Text7(2) = "908" And cp(43) <> "" Then
      strExc(0) = "select np22 from caseprogress,nextprogress where cp09='" & cp(43) & "' and cp10='601' and np01(+)=cp09 and np07='605'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "delete from nextprogress where np01='" & cp(43) & "' and np07='605' and np22=" & RsTemp("np22")
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2019/11/12
   
   'Added by Morgan 2020/2/26
   '修改分割加註通知
   If m_EdDivSugInform = True Then
      strExc(0) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
      strExc(1) = strExc(0) & "(" & cp(9) & ")請修改分割建議內容!"
      strExc(2) = strExc(0) & "前次修正已寫建議分割內容，" & ChangeTStringToTDateString(Text7(0)) & "發文之" & Label3(5) & _
         "尚未修改內容，請修改分割建議內容!'"
      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
         "values('" & strUserNum & "','" & Text7(1) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(2)) & "')"
      cnnConnection.Execute strSql, intI
   End If
   'end 2020/2/26
   
   'Add By Sindy 2023/4/17 設計新案發文後自動發mail通知分案(210)製作中說
   'Modify By Sindy 2023/4/28 +敏莉說增加案件性質125.衍生設計申請
   If Text7(2) = "103" Or Text7(2) = "125" Then
      strSql = "select cp06,cp07,cp48,cp10 from caseprogress" & _
              " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
              " and cp10 in('210') and cp27||cp57 is null and cp14 is null"
      CheckOC
      With adoRecordset
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
         If .RecordCount > 0 Then
            'Modify By Sindy 2023/4/28
            '案件性質名稱
            strExc(10) = ""
            Call ClsPDGetCasePropertyL(1, pa(1), .Fields("cp10"), strExc(10))
            '2023/4/28 END
            '"承辦期限：" & ChangeWStringToTDateString("" & .Fields("cp48")) & vbCrLf
            strExc(9) = "To " & GetPrjSalesNM(Pub_GetSpecMan("C")) & "," & vbCrLf & vbCrLf & _
                        pa(1) & "-" & pa(2) & "今日已提申完畢，請分案" & strExc(10) & vbCrLf & _
                        "本所期限：" & ChangeWStringToTDateString("" & .Fields("cp06")) & vbCrLf & _
                        "法定期限：" & ChangeWStringToTDateString("" & .Fields("cp07")) & vbCrLf & _
                        strExc(10) & "工程師：" & GetPrjSalesNM(m_TCTchk) & vbCrLf & _
                        "謝謝" & vbCrLf
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
               "values('" & strUserNum & "','" & Pub_GetSpecMan("C") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'【請分案" & strExc(10) & "】Our Ref: " & pa(1) & "-" & pa(2) & " [INCOM.210]'" & _
               ",'" & ChgSQL(strExc(9)) & "','" & strUserNum & "')"
            cnnConnection.Execute strSql, intI
         End If
      End With
   End If
   '2023/4/17 END
      
   'Added by Lydia 2022/11/11 法律所案源(限制B2案源)：法務案無發文日則更新發文日為系統日，但法務案不更新CP84。另加發EMAIL給法務案承辦人，提供案號及案件性質、總收文號，提醒他去案件進度檔補輸工作時數及工作點數分配。
   If m_LOS15 <> "" And m_LOS02 = "B2" Then
       Call PUB_UpdateLosCP27(m_LOS15)
   End If
   'end 2022/11/11
     
   'Added by Lydia 2023/03/03 外專新案認領：急件(新案)發文時，重新進入認領流程
   'Modifed by Lydia 2023/05/23 限第一次發文cp82
   If strSrvDate(1) >= 外專新案認領啟用日 And pa(1) = "FCP" And Val(cp(82)) = 0 Then
      If PUB_UpdateReTCN(pa, cp) = False Then
         GoTo CheckingErr
      End If
   End If
   'end 2023/03/03
   
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
End Function

Private Sub Combo3_Click()
   If Combo3.ListIndex >= 0 Then
      If Text7(0) <> "" Then
         'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
         Text7(20) = TransDate(PUB_GetWorkDay1(CompDate(1, Combo3.ItemData(Combo3.ListIndex), Text7(0)), True), 1)
      End If
   End If
End Sub

'Modify by Morgan 2006/5/1 改寫
Private Sub Command1_Click()

   strExc(0) = Name
   strExc(1) = pa(1)
   strExc(2) = pa(2)
   strExc(3) = pa(3)
   strExc(4) = pa(4)
   strExc(5) = Text7(2).Text '案件性質
   strExc(6) = strReceiveNo '收文號
   strExc(7) = Text7(0) '發文日
   
   Me.Hide
   frm060104_4.Show
End Sub

Private Sub Form_Activate()
    'Add By Cheng 2003/10/06
    '若有按下變更事項按鈕, 則重新讀取資料
    If m_blnClkChgEvnBtn = True Then
        ReadPatent
        Label3(0) = strReceiveNo
        m_blnClkChgEvnBtn = False
    End If
End Sub

Private Sub Form_Load()
Dim m_notREC As String   'add by sonia 2017/11/21
Dim m_CP27 As String
   
   'MoveFormToCenter Me 'Mark by Lydia 2019/10/25 移到下面
   intWhere = 國外_FC
   With frm060104_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   
   'Add By Sindy 2023/4/7
   If Pub_StrUserSt03 <> "M51" Then
      txtDecreasePageFee.Visible = False
      txtDecreaseItemFee.Visible = False
   End If
   '2023/4/7 END
   
   'Add by Morgan 2005/8/4
   lblEP04N = ""
   ReDim pa(TF_PA)
   ReDim cp(TF_CP)
   ReadPatent
   'Add by Morgan 2005/8/4
   '加出名代理人清單供勾選
   lstNameAgent.Clear
   'Modify by Morgan 2005/8/29 控制'FG'不選出名代理人
   'PUB_SetOurAgent lstNameAgent, pa(), cp(110)
   If cp(1) = "FCP" Then
'2010/5/7 MODIFY BY SONIA 改在PUB_SetOurAgent預設
'      'FCP 下列性質不預設,PS.員工號"X"不存在所以不會有預設值
'      If (Text7(2) = "901" Or Text7(2) = "902" Or Text7(2) = "903" Or Text7(2) = "904" Or Text7(2) = "906" Or Text7(2) = "912") Then
'         PUB_SetOurAgent lstNameAgent, pa(), "X"
'      'Add by Morgan 2006/1/2 101-103預設"桂齊恆"and"閻啟泰"
'      ElseIf cp(110) = "" And (Text7(2) = "101" Or Text7(2) = "102" Or Text7(2) = "103") Then
'         'PUB_SetOurAgent lstNameAgent, pa(), "76012,81040"
'      Else
'         PUB_SetOurAgent lstNameAgent, pa(), cp(110), cp(10)
'      End If
      PUB_SetOurAgent lstNameAgent, pa(), cp(110), cp(10), True
      '2010/5/7 END
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   'Added by Sindy 2021/5/10 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 1000
   lstNameAgent.Width = 1300

   'Added by Lydia 2019/10/25 FCP新案翻譯增加"翻譯瑕疵"
   fraTrans04.Visible = False
   fraTrans04.Top = cboAddCP64.Top
   fraTrans04.BackColor = &H8000000F
   If cp(1) = "FCP" And cp(10) = "201" Then
       lblAddCP64.Visible = False: cboAddCP64.Visible = False
       fraTrans04.Visible = True
       Me.Height = 6650 'Modify By Sindy 2023/3/20 6550
       Combo6.Clear
       Combo6.AddItem "漏譯"
       Combo6.AddItem "誤譯"
       Combo6.AddItem "贅字太多"
       Combo6.AddItem "語句不通順"
       Combo6.AddItem "其他(自行輸入內容)"
   Else
       Me.Height = 6300 'Modify By Sindy 2023/3/20 6120
   End If
   MoveFormToCenter Me  '先調整畫面大小,再置中
   'end 2019/10/25
         
   '2007/8/6 ADD BY SONIA 加第三人提實審按鈕
   If m_CP10 = "416" Then
      cmdOK(6).Visible = True
      cmdOK(6).Enabled = True
   Else
      cmdOK(6).Visible = False
      cmdOK(6).Enabled = False
   End If
   '2007/8/6 END
   
   Label3(0) = strReceiveNo
   ' 90.12.19 modify by louis
   'If Left(strReceiveNo, 1) = "A" Then Text7(13) = "Y"
   '93.1.20 CANCEL BY SONIA
   'If Left(strReceiveNo, 1) = "A" Then Text7(13) = Empty
   '93.1.20 END
    m_blnClkChgEvnBtn = False
   
   'Add by Morgan 2010/1/6 增加頁數,項數欄位
   m_bolChkFee = False
   txtCP135.Enabled = False
   txtCP135.BackColor = Me.BackColor
   txtCP136.Enabled = False
   txtCP136.BackColor = Me.BackColor
   txtCP137.Enabled = False
   txtCP137.BackColor = Me.BackColor
   txtCP138.Enabled = False
   txtCP138.BackColor = Me.BackColor
   'Add By Sindy 2023/3/16
   txtCP167.Enabled = False
   txtCP167.BackColor = Me.BackColor
   txtCP168.Enabled = False
   txtCP168.BackColor = Me.BackColor
   '2023/3/16 END
   txtCP84.Enabled = True
   
   'Add By Sindy 2018/4/23
   '讀取總頁數和總項數(統計已發文)
   m_allPage = 0: m_allItem = 0
   'Modify By Sindy 2023/3/9 改成共用函數
   '取得總頁數/總項數
   Call PUB_GetAllPageItem("", cp, pa, m_allPage, m_allItem)
'   '總頁數:最近一筆進度的頁數
'   'Modify By Sindy 2018/5/21 取消 and cp158>0,改為 and cp159=0
'   strExc(0) = "select cp09,cp10,nvl(cp135,0) from caseprogress" & _
'               " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
'               " and cp159=0" & _
'               " and nvl(cp135,0)>0" & _
'               " ORDER BY CP69 DESC,CP70 DESC"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      m_allPage = Val("" & RsTemp.Fields(2))
'   End If
'   '總項數:增加項數-刪除未審項數-刪除已審項數
'   strExc(0) = "select sum(nvl(cp136,0)),sum(nvl(cp137,0)),sum(nvl(cp138,0)) from caseprogress" & _
'               " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
'               " and cp159=0" & _
'               " and (nvl(cp136,0)>0 or nvl(cp137,0)>0 or nvl(cp138,0)>0)" & _
'               " ORDER BY CP69 DESC,CP70 DESC"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      m_allItem = Val("" & RsTemp.Fields(0)) - Val("" & RsTemp.Fields(1)) - Val("" & RsTemp.Fields(2))
'   End If
   '2018/4/23 END
   
   'Add By Sindy 2018/6/22 工程師有輸入時,要帶出欄位值
   If Val(cp(135)) > 0 And Val(txtCP135.Text) = 0 Then
      txtCP135.Text = cp(135)
   End If
   If Val(cp(136)) > 0 And Val(txtCP136.Text) = 0 Then
      txtCP136.Text = cp(136)
   End If
   If Val(cp(137)) > 0 And Val(txtCP137.Text) = 0 Then
      txtCP137.Text = cp(137)
   End If
   If Val(cp(138)) > 0 And Val(txtCP138.Text) = 0 Then
      txtCP138.Text = cp(138)
   End If
   'Add By Sindy 2023/3/16
   If Val(cp(167)) > 0 And Val(txtCP167.Text) = 0 Then
      txtCP167.Text = cp(167)
   End If
   If Val(cp(168)) > 0 And Val(txtCP168.Text) = 0 Then
      txtCP168.Text = cp(168)
   End If
   '2023/3/16 END
   '2018/6/22 END
   'Add By Sindy 2023/3/21 檢查是否需要重新計算總頁數(中說一併修正)
   '讀取專利說明書頁數明細
   ReDim pageD(1 To 21) As String
   Dim strChgPA64 As String, strChgPA65 As String, strChgPA67 As String, strChgPA68 As String
   Dim strTotPage As String
   If PUB_ReadPageDetail(strPD01, pageD, , , , strChgPA64, strChgPA65, strChgPA67, strChgPA68, strReceiveNo) = True Then
      If Val(strChgPA64) <> 0 Or Val(strChgPA65) <> 0 Or Val(strChgPA67) <> 0 Or Val(strChgPA68) <> 0 Then
         strTotPage = Val(pa(64)) + Val(pa(65)) + Val(pa(67)) + Val(pa(68)) + _
                      Val(strChgPA64) + Val(strChgPA65) + Val(strChgPA67) + Val(strChgPA68)
         If txtCP135.Text <> strTotPage And Val(strTotPage) > 0 Then
            txtCP135.Text = strTotPage
         End If
      End If
   End If
   '2023/3/21 END
   
   m_strReExamCP27 = "" 'Added by Morgan 2013/1/10
   m_bolFixNewFee = False 'Added by Morgan 2013/1/10
   'Modified by Morgan 2013/11/6 +235核對中說格式
   If (cp(10) = "416" Or cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Or cp(10) = "203" Or cp(10) = "204" Or cp(10) = "205" Or cp(10) = "206") Then
      m_Div416OfficialFee = 0 '2010/12/8 add by sonia
      
      'Added by Morgan 2013/1/10
      If (cp(10) = "210" Or cp(10) = "203" Or cp(10) = "204" Or cp(10) = "205" Or cp(10) = "206") Then
         m_strReExamCP27 = PUB_GetReExamDate(cp)
         If m_strReExamCP27 > "20130000" Then
            m_bolFixNewFee = True
         End If
      End If
      'end 2013/1/10
      
      If m_strReExamCP27 = "" Then 'Added by Morgan 2013/1/10
      
         m_bol99Case = Chk99NewCase(cp(1), cp(2), cp(3), cp(4))
         If m_bol99Case Then
            '實審
            If cp(10) = "416" Then
               '新案翻譯已發文
               'Modify by Morgan 2010/4/28 +307
               'Modified by Morgan 2013/11/6 +235核對中說格式
               If PUB_ChkCPExist(cp, "201", 2) Or PUB_ChkCPExist(cp, "209", 2) Or _
                  PUB_ChkCPExist(cp, "235", 2) Or PUB_ChkCPExist(cp, "210", 2) Or _
                  PUB_ChkCPExist(cp, "307", 2) Then
                  m_bolChkFee = True
                  m_bolChkPageItem = True
                  txtCP84.Enabled = False
                  lblCP135.Caption = "總頁數:" 'Add By Sindy 2023/3/16
'                  txtCP135.Enabled = True
'                  txtCP135.BackColor = vbWhite
                  lblCP136.Caption = "總項數:"
                  txtCP136.Enabled = True
                  txtCP136.BackColor = vbWhite
               Else
                  'Modify By Sindy 2018/9/5 ex:FCP-52686
                  'Modify By Sindy 2018/8/14 +203主動修正(同日發文) => (Val(m_CP27) > 0 And Val(Text7(0)) = Val(m_CP27) - 19110000)
                  Call PUB_ChkCPExist(cp, "203", 2, , , , m_CP27) 'Add By Sindy 2018/8/14 ex:FCP-55336
                  If (Val(m_CP27) > 0 And Val(Text7(0)) = Val(m_CP27) - 19110000) Then
                     '中說未收文
                     If PUB_ChkCPExist(cp, "201", 1) = False And PUB_ChkCPExist(cp, "209", 1) = False And _
                        PUB_ChkCPExist(cp, "235", 1) = False And PUB_ChkCPExist(cp, "210", 1) = False Then
                        m_bolChkFee = True
                        m_bolChkPageItem = True
                        txtCP84.Enabled = False
                        lblCP135.Caption = "總頁數:" 'Add By Sindy 2023/3/16
'                        txtCP135.Enabled = True
'                        txtCP135.BackColor = vbWhite
                        lblCP136.Caption = "總項數:"
                        txtCP136.Enabled = True
                        txtCP136.BackColor = vbWhite
                     End If
                  End If
               End If
               'Modify By Sindy 2018/11/20 Mark,不控管; Ex:FCP-059934實審要先收超頁超項費用請款,中說假發文
'               'Add By Sindy 2018/4/25
'               Call GetCP31isY_CP05(cp(1), cp(2), cp(3), cp(4), "cp118", strNewCaseCP118)
'               If strNewCaseCP118 = "Y" Then '新案電子送件時,走新規則:頁數,項數,規費在填申請書時處理
'                  m_bolChkFee = False
'               End If
'               '2018/4/25 END
            '新案翻譯,檢視中說,製作中說
            'Modified by Morgan 2013/11/6 +235核對中說格式
            ElseIf cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Then
               '實審已發文
               If PUB_ChkCPExist(cp, "416", 2) Then
                  m_bolChkFee = True
                  m_bolChkPageItem = True
                  txtCP84.Enabled = False
                  lblCP135.Caption = "總頁數:" 'Add By Sindy 2023/3/16
'                  txtCP135.Enabled = True
'                  txtCP135.BackColor = vbWhite
                  lblCP136.Caption = "總項數:"
                  txtCP136.Enabled = True
                  txtCP136.BackColor = vbWhite
               End If
            '修正
            ElseIf cp(10) = "203" Or cp(10) = "204" Or cp(10) = "205" Or cp(10) = "206" Then
               '實審已發文
               'MODIFY BY SONIA 2014/6/20 加入435續行母案再審 FCP-048155
               If PUB_ChkCPExist(cp, "416", 2) Or PUB_ChkCPExist(cp, "435", 2) Then
                  m_bolChkItem = True 'Add by Morgan 2010/9/27
                  m_bolChkFee = True
                  txtCP84.Enabled = False
                  lblCP136.Caption = "增加項數:"
                  txtCP136.Enabled = True
                  txtCP136.BackColor = vbWhite
                  txtCP137.Enabled = True
                  txtCP137.BackColor = vbWhite
                  '已輸審查意見通知
                  If PUB_ChkCPExist(cp, "1202") = True Then
                     txtCP138.Enabled = True
                     txtCP138.BackColor = vbWhite
                  End If
                  'Add By Sindy 2023/3/16
                  lblCP135.Caption = "增加頁數:"
'                  txtCP135.Enabled = True
'                  txtCP135.BackColor = vbWhite
'                  txtCP167.Enabled = True
'                  txtCP167.BackColor = vbWhite
                  '已輸審查意見通知
                  If PUB_ChkCPExist(cp, "1202") = True Then
'                     txtCP168.Enabled = True
'                     txtCP168.BackColor = vbWhite
                  End If
                  '2023/3/16 END
               End If
            End If
            If m_bolChkFee Then
               'Modify By Sindy 2019/3/26
               'SetOfficialFee
               'Modify By Sindy 2023/3/16 +, , , txtCP167
               Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
                                         txtCP135, txtCP136, txtCP137, txtCP84, , txtDecreaseItemFee, _
                                         m_lngOverPageFee, m_lngOverItemFee, , , m_lngOfficialFee, _
                                         Text7(0).Text, , txtDecreasePageFee, txtCP167, , m_str938RecvNo, m_str939RecvNo)
               '2019/3/26 END
            End If
         '2010/12/8 ADD BY SONIA 分割案之實審發文,若母案有已收未取消的再審程序且申請日在2010/1/1以前者,分割案實審規費應為8000元
         '2011/3/24 MODIFY BY SONIA FCP-034512申復發文誤帶規費8000
         'ElseIf PUB_ChkCPExist(cp, "307") Then
         ElseIf cp(10) = "416" And PUB_ChkCPExist(cp, "307") Then
            txtCP84 = "8000"
            'Modify by Morgan 2011/7/26 會有超頁費 Ex.FCP-044051
            'txtCP84.Enabled = False
            MsgBox "本案請依舊法規則計算規費！"
            'end 2011/7/26
            cp(17) = txtCP84      '同時改收文規費
            m_Div416OfficialFee = txtCP84
         '2010/12/8 END
         End If
         
      'Added by Morgan 2013/1/10
      '有再審102年後發文,修正要收超項費
      ElseIf m_bolFixNewFee = True Then
         m_bolChkItem = True
         m_bolChkFee = True
         txtCP84.Enabled = False
         lblCP136.Caption = "增加項數"
         txtCP136.Enabled = True
         txtCP136.BackColor = vbWhite
         txtCP137.Enabled = True
         txtCP137.BackColor = vbWhite
         txtCP138.Enabled = True
         txtCP138.BackColor = vbWhite
         'Add By Sindy 2023/3/16
         lblCP135.Caption = "增加頁數"
'         txtCP135.Enabled = True
'         txtCP135.BackColor = vbWhite
'         txtCP167.Enabled = True
'         txtCP167.BackColor = vbWhite
'         txtCP168.Enabled = True
'         txtCP168.BackColor = vbWhite
         '2023/3/16 END
      End If
      'end 2013/1/10
      
   'Added by Morgan 2013/1/9 102/1/1 起再審也要算超頁超項費(若有延期則以延期發文日判斷)
   'Modified by Morgan 2013/6/10 只有發明的再審
   'ElseIf cp(10) = "107" And m_bol107NewFee = True Then
   'Modified by Morgan 2013/10/18 +435
   ElseIf pa(8) = "1" And (cp(10) = "107" Or cp(10) = "435") And m_bol107NewFee = True Then
      m_bolChkFee = True
      m_bolChkPageItem = True
      txtCP84.Enabled = False
      lblCP135.Caption = "總頁數:" 'Add By Sindy 2023/3/16
'      txtCP135.Enabled = True
'      txtCP135.BackColor = vbWhite
      lblCP136.Caption = "總項數:"
      txtCP136.Enabled = True
      txtCP136.BackColor = vbWhite
         
   'Added by Morgan 2013/1/3
   '申請技術報告要輸項數以便計算規費
   ElseIf cp(10) = "421" Or cp(10) = "807" Then
      'Add By Sindy 2024/8/6
      m_bolChkFee = True
      m_bolChkPageItem = True
      '2024/8/6 END
      lblCP136.Caption = "項數:"
      txtCP136.Enabled = True
      txtCP136.BackColor = vbWhite
      txtCP84.Enabled = False
   'end 2013/1/3
   End If
   
   'Add by Morgan 2012/8/1
   'Modified by Morgan 2015/11/24 +230 --何淑華
   If cp(10) = "903" Or cp(10) = "904" Or cp(10) = "230" Then
      Label10.Visible = True
      Text7(20).Visible = True
      Combo3.Visible = True
      SetCombo3
   Else
      Label10.Visible = False
      Text7(20).Visible = False
      Combo3.Visible = False
   End If
   'end 2012/8/1
   'Added by Lydia 2015/12/31 FCP案會稿924發文時,檢查一定要有相關總收文號
   'Add By Sindy 2019/1/18 更改也要顯示相關總收文號
   'If text1 = "FCP" And (Text7(2) = "924" Or Text7(2) = "403") Then
   If Text1 = "FCP" And (Text7(2) = "924") Then
      Label14.Visible = False
      Text7(12).Visible = False
      Label3(6).Visible = False
      txtCP43.Visible = True
      lblCP43.Visible = True
      lblCP43.Top = Label14.Top: lblCP43.Left = Label14.Left
      txtCP43.Top = Text7(12).Top: txtCP43.Left = Text7(12).Left
      Command2.Visible = True
      Command2.Top = txtCP43.Top
   End If
   'end 2015/12/31
   
   'add by sonia 2017/11/21 中說發文時若其他進度尚未請款時要提醒(不管CP16是否有值)
   If cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Then
      m_notREC = ""
      strSql = "select cp09,cp10,cpm03 from caseprogress,casepropertymap " & _
              " where cp01='" & Text1 & "' and cp02='" & Text2 & "' and cp03='" & Text3 & "' and cp04='" & Text4 & "'" & _
              " and cp20 is null AND CP159=0 and cp60 is null and cp01=cpm01(+) and cp10=cpm02(+) " & _
              " order by cp09"
      CheckOC
      With adoRecordset
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
         If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
               m_notREC = m_notREC & .Fields("cpm03") & ";"
               .MoveNext
            Loop
            m_notREC = "請注意！尚有未請款之進度：" & m_notREC
            MsgBox m_notREC, vbExclamation
         Else
            MsgBox "此案皆已請款，可直接寄中說！", vbExclamation  '2017/11/22 add by sonia
         End If
      End With
   End If
   'end 2017/11/21
   txtCP118_Validate False 'Add By Sindy 2018/1/31
   
'   'Add By Sindy 2018/5/22
'   If m_bolChkPageItem = True Then
'      If Val(txtCP135) = 0 Or Val(txtCP136) = 0 Then
'         strExc(0) = "select cp09,cp10,nvl(cp135,0) as cp135,nvl(cp136,0) as cp136 from caseprogress" & _
'                     " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
'                     " and cp159=0" & _
'                     " and nvl(cp135,0)>0" & _
'                     " ORDER BY CP69 DESC,CP70 DESC"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            txtCP135 = Val(RsTemp.Fields("cp135")) '總頁數
'            txtCP136 = Val(RsTemp.Fields("cp136")) '總項數
'         End If
'      End If
'   End If
'   '2018/5/22 END
   
   'Added by Lydia 2018/03/27 建User暫存區
   m_AttchPath = App.path & "\" & strUserNum
   If Dir(m_AttchPath, vbDirectory) = "" Then
        MkDir m_AttchPath
   End If
   'end 2018/03/27
   
   'Added by Lydia 2018/03/27 載入印表機選單
   PUB_SetPrinter Me.Name, Me.cboPrinter, strPrinter
   
   'Added by Lydia 2018/05/17 顯示印表機選單
   'Mark by Lydia 2023/03/03 與Sharon,Phoebe 確認不用列印說明書
   'If InStr(cPrintORI, cp(10)) > 0 And TransDate(cp(5), 2) >= FCP案件命名啟用日 Then
   '     lblPrint.Visible = True: cboPrinter.Visible = True
   '     'Added by Lydia 2018/05/18 敏莉要求預設為機密列印,而非上次列印的印表機
   '     strExc(2) = Pub_GetSpecMan("FCP程序機密列印")
   '     If cboPrinter.ListCount > 0 And strExc(2) <> "" Then
   '         For intI = 0 To cboPrinter.ListCount - 1
   '              strExc(1) = cboPrinter.List(intI)
   '              '用系統特殊設定
   '              If strExc(1) = strExc(2) Then
   '                   cboPrinter.ListIndex = intI
   '                   Exit For
   '              End If
   '         Next intI
   '         'end 2018/05/18
   '     End If
   'Else
        lblPrint.Visible = False: cboPrinter.Visible = False
   'End If 'end 2023/03/03
   'Added by Lydia 2021/01/21 FCP實審發文承辦單不出紙本改發email
   lblEmail.Top = 4335:   txtEmail.Top = 4290
   lblPrint.Top = 2940:   cboPrinter.Top = 2910
   'Modify By Sindy 2022/5/18 Mark
'   lblEmail.Visible = False: txtEmail.Visible = False  'Email維護
'   lblRecDate.Visible = False: txtRecDate.Visible = False  '當天報告
   
   'Added by Lydia 2020/08/17 實審: 增加已收款
   lblPAID.Visible = False: txtPAID.Visible = False
   If cp(1) = "FCP" And cp(10) = "416" Then
       lblPAID.Visible = True: txtPAID.Visible = True
       'Added by Lydia 2021/01/21 FCP實審發文承辦單不出紙本改發email：增加「當天報告」、「Email維護」欄位設定
       lblRecDate.Visible = True: txtRecDate.Visible = True
       If Val(cp(82)) = 0 Then  '第一次發文
           lblEmail.Visible = True: txtEmail.Visible = True
       End If
       m_416Type = "0"
       '下列兩種情況不用問是否「出帳單」
       If PUB_ChkCPExist(cp, "203", 1) = True Then
           '1.為主動修正+實審(由工程師出請款信及承辦出帳單)：判斷案件尚有未發文之主動修正
           m_416Type = "1"
       ElseIf cp(148) = "Y" Then
           '2.特殊請款(由承辦出帳單)
           m_416Type = "2"
       End If
       If m_416Type <> "0" Then
           txtEmail = "Y"
       End If
       'end 2021/01/21
       
       strExc(0) = "select GetEmailFlag('" & cp(1) & cp(2) & cp(3) & cp(4) & "') as eFlag from dual "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
           m_eFlag = "" & RsTemp.Fields("eFlag")
       End If
       'end 2021/01/21
       
   'Modify By Sindy 2022/5/18
   Else
      'Added by Morgan 2024/11/18 447再審查加速審查預設空且鎖住--敏莉
      If cp(10) = "447" Then
         txtEmail = ""
         txtEmail.Enabled = False
         lblEmail.Enabled = False
      'end 2024/11/18
      ElseIf InStr("202,101,102,103,125,307,308,447", cp(10)) = 0 Then '排除補文件...不用預設Y
         txtEmail = "Y"
      End If
      lblEmail.Caption = "Email維護:             (Y:是)"
      '2022/5/18 END
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2017/12/18
   
   DelTempFile 'Added by Lydia 2018/10/19 另外處理暫存檔 (FCP-59630因為有開啟卷宗區檔,造成無法刪檔)
   'Added by Lydia 2018/05/17 若印表機變動, 則更新列印設定
   If Me.cboPrinter.Text <> Me.cboPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cboPrinter.Name, "0", "0", Me.cboPrinter.Text
   End If
   'end 2018/05/17
   
   Set frm060104_3 = Nothing
End Sub
'Added by Lydia 2018/10/19 刪除說明書暫存檔
Private Sub DelTempFile()
'Modified by Lydia 2018/11/01 如果電子送件暫存區無法刪除,彈訊息
'On Error Resume Next
Dim stRtn As String

On Error GoTo ErrMsg
'end 2018/11/01

   'Added by Lydia 2018/03/27 刪除說明書暫存檔
   'Move by Lydia 2018/10/19 從FormUnload移出來
   If Dir(m_AttchPath & "\" & pa(1) & "*" & Val(pa(2)) & "*.*") <> "" Then
         Kill m_AttchPath & "\" & pa(1) & "*" & Val(pa(2)) & "*.*"
   End If
   'end 2018/03/27
   
    'Added by Lydia 2018/10/29  FCP發文A類非"924會稿"時，則自動去電子送件暫存區將同一案件的資料匣刪除。
    'Modified by Lydia 2018/10/31 判斷有發文
    'If Left(cp(9), 1) = "A" And Text7(2) <> "924" Then
    'Remove by Lydia 2018/12/03 改在程序人員在"A"類請款時
'    If bolChkSave = True And Left(cp(9), 1) = "A" And Text7(2) <> "924" Then
'        If Dir("\\Typing2\電子送件暫存區\" & pa(1) & pa(2), vbDirectory) <> "" Then
'            stRtn = "無法刪除\\Typing2\電子送件暫存區\" & pa(1) & pa(2) & vbCrLf & "，請手動刪除資料夾！"  'Added by Lydia 2018/11/01
'            If Dir("\\Typing2\電子送件暫存區\" & pa(1) & pa(2) & "\*.*") <> "" Then
'                 Kill "\\Typing2\電子送件暫存區\" & pa(1) & pa(2) & "\*.*"
'            End If
'            If stRtn <> "" Then 'Added by Lydia 2018/11/01
'                 RmDir "\\Typing2\電子送件暫存區\" & pa(1) & pa(2)
'            End If
'        End If
'    End If
'    'end 2018/10/29
    
'Added by Lydia 2018/11/01
    Exit Sub

ErrMsg:
If Err.Number <> 0 Then
    If stRtn <> "" Then
        MsgBox stRtn, vbCritical
        stRtn = ""
    End If
    Resume Next
End If
'end 2018/11/01
End Sub
Private Sub ReadPatent()
Dim Lbl As Object, txt As Object, i As Integer, iPos1 As Integer, iPos2 As Integer
   
   m_RefCP10 = ""
   m_strNewAppIpoNo = "" 'Added by Morgan 2013/8/2
   
   For Each Lbl In Label3
      Lbl = ""
   Next
   For Each txt In Text7
      txt = ""
   Next
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   bolDefWebApp = False 'Added by Lydia 2017/12/12
   m_TCTchk = "" 'Added by Lydia 2019/07/30
   
   Select Case pa(1)
      Case "FCP"
         'Modify by Morgan 2006/10/19 不再用dll的函數讀取基本資料
         If PUB_ReadPatentDatabase(pa(), intWhere) Then
            Text7(3) = pa(48)
            Text7(10) = pa(77)
            For i = 51 To 56
               Text7(i - 47) = pa(i)
            Next
            For i = 5 To 7
               Text7(i + 9) = pa(i)
            Next
            Text7(18) = pa(91)
            Text7(19) = pa(139) 'Add by Morgan 2006/10/19
            If CU72FA39(pa(26), pa(75)) Then Label3(3) = "Y"
            
            m_PA163 = pa(163) 'Added by Morgan 2015/9/30
            
            'Added by Lydia 2019/01/09 工程師主管
            'Memo by Lydia 2019/08/19 日文組副本:除各組主任(99034,94012)給主管,其餘人給審核主管
            'Mark by Lydia 2020/02/10 改到承辦人設定
            'm_GrpMan = Pub_GetFCPGrpMan(pa(150))
         End If
      Case "FG"
         'Modify by Morgan 2006/10/19 不再用dll的函數讀取基本資料
         If PUB_ReadServicePracticeDatabase(pa(), intWhere) Then
            Text7(3) = pa(29)
            Text7(10) = pa(27)
            Text7(5) = pa(30)
            For i = 5 To 7
               Text7(i + 9) = pa(i)
            Next
            Text7(18) = pa(18)
            Text7(19) = pa(71) 'Add by Morgan 2006/10/19
            'Modified by Lydia 2017/12/12 改成模組
'            If pa(26) = "" Then
'               'pa(8)
'               strExc(0) = "SELECT CU72 FROM CUSTOMER WHERE " & ChgCustomer(pa(8))
'            Else
'               strExc(0) = "SELECT FA39 FROM FAGENT WHERE " & ChgFagent(pa(26))
'            End If
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               If Not IsNull(RsTemp.Fields(0)) Then Label3(3) = RsTemp.Fields(0)
'            End If
            If CU72FA39(pa(8), pa(26)) Then Label3(3) = "Y"
         End If
   End Select
   
   'Added by Lydia 2017/12/12  判斷客戶或代理人有設定為電子送件
   strExc(0) = ""
   If pa(1) = "FCP" Then
       strExc(0) = "select nvl(cu174,'N') from customer WHERE " & ChgCustomer(pa(26))
       strExc(0) = strExc(0) & " union select nvl(fa104,'N') from fagent WHERE " & ChgFagent(pa(75))
       If pa(75) <> "" Then mFA10 = GetPrjNationNumber(ChangeCustomerL(pa(75))) 'Added by Lydia 2019/07/03
   ElseIf pa(1) = "FG" Then
       strExc(0) = "select nvl(cu174,'N') from customer WHERE " & ChgCustomer(pa(8))
       strExc(0) = strExc(0) & " union select nvl(fa104,'N') from fagent WHERE " & ChgFagent(pa(26))
   End If
   If strExc(0) <> "" Then
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
           strExc(1) = RsTemp.GetString(adClipString, , , ",")
           If InStr(strExc(1), "Y") > 0 Then
               bolDefWebApp = True
           End If
       End If
   End If
   'end 2017/12/12
      
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      ' 90.06.27 modify by louis, 暫存案件性質
      m_CP10 = cp(10)
      m_CP50 = cp(50): m_CP51 = cp(51): m_CP52 = cp(52) '2007/8/6 ADD BY SONIA
      If cp(13) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(13), strExc(0)) Then Label3(1) = strExc(0)
         If ClsPDGetStaff(cp(13), strExc(0)) Then Label3(1) = strExc(0)
      End If
      Label3(2) = cp(6)
      Label3(7) = cp(7)
      If cp(27) <> "" Then
         Text7(0) = cp(27)
      Else
         Text7(0) = strSrvDate(2)
      End If
      If cp(14) <> "" Then
         Text7(1) = cp(14)
         ChgType 1
         m_GrpMan = PUB_GetFCPEngSup(cp(14)) 'Added by Lydia 2020/02/10 抓外專工程師主管, 日文組要抓第二、三級主管
      End If
      If cp(44) <> "" Then
         Text7(12) = cp(44)
         ChgType 12
      End If
      If cp(10) <> "" Then
         Text7(2) = cp(10)
         ChgType 2
      End If
      
      m_strCP148 = cp(148) 'Added by Lydia 2019/01/07 暫存 'Memo by Lydia 2019/01/28 +主動修正是否已併入中說送件
            
      Text7(13) = cp(20)
      Text7(17) = cp(64)
      'Add by Morgan 2007/7/20
      If cp(113) = "" Then
         '請求面詢407,申請技術報告421,實體審查416(承辦人為工程師)預設0.5小時
         'Modify by Morgan 2007/8/31 加807第三申請技術報告
         'If cp(10) = "407" Or cp(10) = "421" Then
         If cp(10) = "407" Or cp(10) = "421" Or cp(10) = "807" Then
            txtCP113 = 0.5
         ElseIf cp(10) = "416" Then
            If PUB_GetStaffST15(cp(14), 1) = "F21" Or PUB_GetStaffST15(cp(14), 1) = "F81" Then  '2008/4/8 MODIFY BY SONIA 加 F81
               txtCP113 = 0.5
            End If
         End If
      Else
         txtCP113 = cp(113)
      End If
      txtCP114 = cp(114)
      
      'Modify By Sindy 2019/2/21 Mark,不能加此if判斷,申請書規費是總額,但發文時新案一個規費,實審又一個規費要在此計算
'      'Add By Sindy 2019/2/18
'      If Val(cp(84)) > 0 Then
'         txtCP84.Tag = cp(84)
'         txtCP84.Text = txtCP84.Tag
'      Else
      'Modify By Sindy 2023/8/8 收文規費為0且發文規費有"值",則發文介面的發文規費抓進度檔的發文規費的"值"
      If Val(cp(17)) = 0 And Val(cp(84)) > 0 Then
         txtCP84.Tag = cp(84)
         txtCP84.Text = txtCP84.Tag
      Else
      '2023/8/8 END
         'Add by Morgan 2004/8/12
         txtCP84.Tag = cp(17)
         txtCP84.Text = txtCP84.Tag
      End If
      
      'Added by Lydia 2022/11/11 法律所案源：取得案源類別、發文規費、email加註
      m_LOS15 = cp(162)
      '限制B2類發文規費; 只需考慮B2類, A類不會有PT案,B1類PT案不會繳規費, C類已取消(有也是照原來規則不必改)
      '模組回傳m_LOS02
      m_LosCP84 = PUB_GetLosCP84(m_LOS15, pa(1), pa(2), pa(3), pa(4), "B2", m_LOS02, m_LosMemo)
      'end 2022/11/11
   
      'add by sonia 2021/3/26 面詢408時再檢查本次請求面詢及其補收款是否已繳規費
      If cp(10) = "408" Then
         strExc(0) = "select sum(nvl(c1.cp17,0)+nvl(c2.cp17,0)) cp17,sum(nvl(c1.cp84,0)+nvl(c2.cp84,0)) cp84 FROM caseprogress c1,caseprogress c2, " & _
                     "(SELECT MAX(cp05||cp09) cp09 FROM caseprogress WHERE cp01='" & pa(1) & "' AND cp02='" & pa(2) & "' AND cp03='" & pa(3) & "' AND cp04='" & pa(4) & "' and cp10='407' and (cp158>0 or cp159=0)) c3 " & _
                     "where substr(c3.cp09,9)=c1.cp09(+) and c1.cp09=c2.cp43(+) AND '911'=c2.cp10(+)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            txtCP84.Tag = cp(17) - Val(RsTemp.Fields("CP84"))
            txtCP84.Text = cp(17) - Val(RsTemp.Fields("CP84"))
         End If
      End If
      'end 2021/3/26
      
      'end 2007/7/20
      'txtCP118 = cp(118) 'Modify by Amy 2013/05/16 不顯示 'Add by Morgan 2011/1/17
      'Added by Lydia 2017/12/18 配合FCP案增加客戶檔和代理人檔增加"FCP是否電子送件"的控制
      '                                         過去是預設不代CP118由人工重複確認,現在預設代入電子送件的控制
      'Modified by Morgan 2018/1/29 若重新發文cp118可能是"A"
      'txtCP118 = cp(118)
      If cp(118) <> "" Then txtCP118 = "Y"
      'end 2018/1/29
      'end 2017/12/18
      txtCP43 = cp(43) 'Added by Lydia 2015/12/31
      txtCP135 = cp(135) 'Add By Sindy 2018/2/8
      txtCP136 = cp(136) 'Add By Sindy 2018/2/8
   End If
   
   'Add by Morgan 2007/9/13 重新委任自撤不必掛催審期限
   If Text7(2) = "413" And cp(43) <> "" Then
      strExc(0) = "select cp10 from caseprogress where cp09='" & cp(43) & "' and cp10='928'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_RefCP10 = "" & RsTemp(0)
      End If
   End If
   If Not (Text7(2) = "413" And m_RefCP10 = "928") Then
   'end 2007/9/13
      strExc(0) = "SELECT CF05 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND " & _
         "CF02='" & pa(9) & "' AND CF03='" & Text7(2) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Not IsNull(RsTemp.Fields(0)) And RsTemp.Fields(0) <> 0 Then
            Text7(11) = TransDate(CompDate(2, Val(RsTemp.Fields(0)), TransDate(Text7(0), 2)), 1)
            '2006/5/16 ADD BY SONIA 發明案之分割發文時不掛催審, 於通知實審日才掛
            If pa(8) = "1" And m_CP10 = 分割 Then
               Text7(11) = ""
            End If
            '2006/5/16 END
         End If
      End If
   End If
   
   'Added by Lydia 2019/01/28 中說發文日
   mDate201CP158 = ""
   If m_CP10 = "203" Then
      strExc(0) = "select cp158 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' " & _
                        " and cp10 in (201,209,210,235) and cp159=0 and cp158> 0 "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         mDate201CP158 = "" & RsTemp(0)
      End If
   End If
   'end 2019/01/28
   
   'Added by Lydia 2018/12/06
   m_EP09 = ""
   Text7(1).Locked = False
   'end 2018/12/06
   'Add by Morgan 2004/3/23
   '只有發明,新型的翻譯，檢視中說，製作中說會有核稿人
   'Modified by Morgan 2013/9/13 +927其他翻譯
   'Modified by Morgan 2013/11/6 +235核對中說格式
   If pa(8) <> "3" And (m_CP10 = "201" Or m_CP10 = "209" Or m_CP10 = "235" Or m_CP10 = "210" Or m_CP10 = "927") Then
      Dim stSQL As String, rsQuery As New ADODB.Recordset
      '取得核稿人
      'Modified by Lydia 2018/12/06 +EP09
      stSQL = "Select EP04,ST02,EP09 From EngineerProgress, Staff Where EP04=ST01(+) And EP02='" & cp(9) & "'"
      rsQuery.CursorLocation = adUseClient
      rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsQuery.RecordCount > 0 Then
          txtEP04.Text = "" & rsQuery.Fields(0).Value
          lblEP04N.Caption = "" & rsQuery.Fields(1).Value
      Else
          txtEP04.Text = ""
          lblEP04N.Caption = ""
      End If
      'Added by Lydia 2018/12/06 新案翻譯在輸入翻譯完稿日後,分案及發文之承辦人,不能被修改
      m_EP09 = "" & rsQuery.Fields("EP09")
      If m_EP09 <> "" Then
           Text7(1).Locked = True
      End If
      'end 2018/12/06
      
      Set rsQuery = Nothing
      '未輸入核稿人時才開放使用者輸入
      'Modify by Morgan 2006/3/31 發文有時需修改,改都開放--靜芳
      'If txtEP04.Text = "" Then txtEP04.Enabled = True
      txtEP04.Enabled = True
   Else
      txtEP04.Text = ""
      lblEP04N.Caption = ""
   End If
   '紀錄承辦人, 核稿人
   Text7(1).Tag = Text7(1).Text
   txtEP04.Tag = txtEP04.Text
   
   'Added by Lydia 2019/10/25 FCP新案翻譯增加"翻譯瑕疵"
   lblTF37.Tag = ""
   txtTF37.Text = ""
   If cp(1) = "FCP" And cp(10) = "201" Then
      strExc(0) = "select tf01,tf37 from transfee where tf01='" & cp(9) & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
          lblTF37.Tag = "" & RsTemp.Fields("tf01")
          txtTF37.Text = "" & RsTemp.Fields("tf37")
      End If
   End If
   'end 2019/10/25
   
   'Added by Morgan 2013/1/8
   m_bol107NewFee = True
   bolDelay = False
   'end 2013/1/8
   
   'Modify by Morgan 2006/8/18 加判斷107(再審),803(舉發),301,302,303,305(改請)才要
   'Modified by Morgan 2013/8/26 +507 -- FCP032929
   If InStr("107,803,301,302,303,305,507", cp(10)) > 0 Then
      'Add by Morgan 2004/9/8 檢查是否有延期，若有則規費預設0
      bolDelay = PUB_ChkDelay(strReceiveNo, m_strDelayCP09, strExc(1))
      If bolDelay = True Then
         If strExc(1) < "20130101" Then m_bol107NewFee = False 'Added by Morgan 2013/1/9
         cp(17) = "0"
      End If
   End If
  
   'Add by Morgan 2010/5/13
   '補文件且進度備註有委任書時要點選是否為複委任
   If Text7(2) = "202" Then
      chkCP86(0).Enabled = True
      chkCP86(1).Enabled = True
   Else
      chkCP86(0).Value = 0
      chkCP86(0).Enabled = False
      chkCP86(1).Value = 0
      chkCP86(1).Enabled = False
   End If
   
   'Add by Morgan 2010/11/22
   If pa(1) = "FCP" And Text7(2) = "202" Then
      cboAddCP64.Enabled = True
      strExc(1) = ""
      If cp(43) <> "" Then
         strExc(0) = "select cp10 from caseprogress where cp09='" & cp(43) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = RsTemp(0)
         End If
      End If
      PUB_SetCombo202 cboAddCP64, strExc(1)
   Else
      cboAddCP64.Enabled = True
   End If
   
   'Add by Morgan 2011/1/17 檢查是否為電子送件案
   'Modified by Lydia 2017/12/12 改成新案申請
   'strExc(0) = "select cp27,cp64 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('101','102','103') and cp118='Y' and cp27>0 and cp57 is null"
   'Modified by Morgan 2018/1/29 電子送件要考慮自動扣款 cp118='Y' => cp118 is not null
   strExc(0) = "select cp27,cp64 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in (" & NewCasePtyList & ") and cp118 is not null and cp27>0 and cp57 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_bolWebApp = True
      'Added by Morgan 2013/8/2
      '實審發文時檢查若申請案同日發文且為電子送件時自動設定為電子送件並預設智慧局收文文號 -- 靜芳 Ex.FCP-47964
      If Text7(2) = "416" Then
         If RsTemp("cp27") = strSrvDate(1) Then
            MsgBox "本案申請程序為電子送件，" & Label3(5) & "也將自動設定為電子送件!!", vbExclamation
            txtCP118 = "Y"
            If Not IsNull(RsTemp("cp64")) Then
               iPos1 = InStr(RsTemp("cp64"), "智慧局收文文號:")
               If iPos1 > 0 Then
                  iPos2 = InStr(iPos1, RsTemp("cp64"), ";")
                  If iPos2 > 0 Then
                     m_strNewAppIpoNo = Mid(Left(RsTemp("cp64"), iPos2 - 1), iPos1 + 8)
                  Else
                     m_strNewAppIpoNo = Mid(RsTemp("cp64"), iPos1 + 8)
                  End If
               End If
            End If
         End If
      'Added by Lydia 2018/08/22 若新案為電子送件,則主動修正也設為電子送件; ex.FC-58733命名作業產生之主動修正不經過外專分案,未能預設為電子送件
      ElseIf Text7(2) = "203" And txtCP118 = "" Then
            MsgBox "本案申請程序為電子送件，" & Label3(5) & "也將自動設定為電子送件!!", vbExclamation
            txtCP118 = "Y"
      End If
      'end 2013/8/2
   Else
      m_bolWebApp = False
   End If
   'Modified by Morgan 2013/11/6 +235核對中說格式
   If m_bolWebApp = True And (Text7(2) = "201" Or Text7(2) = "209" Or Text7(2) = "235" Or Text7(2) = "210") And txtCP118 <> "Y" Then
      MsgBox "本案應以電子送件方式呈送!!", vbExclamation
   End If
   'end 2011/1/17
   
    'Added by Lydia 2017/12/12 參考內專發文frm040104_3,增加下列檢查
    '新案發文時若為電子送件,當天發文其他案件性質也應為電子送件--玲玲
    If cp(118) = "" And pa(9) = 台灣國家代號 And InStr(NewCasePtyList, cp(10)) = 0 Then
       '排除 501 訴願/503 行政訴訟/803 舉發/804 舉發答辯,不設電子送件--雅娟
       If cp(10) = "501" Or cp(10) = "503" Or cp(10) = "803" Or cp(10) = "804" Then
          txtCP118 = ""
          txtCP118.Locked = True: txtCP118.Enabled = False
       Else
          strExc(0) = "select cp118 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'" & _
             " and cp27='" & strSrvDate(1) & "' and cp10 in (" & NewCasePtyList & ") and cp118 is not null"
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
          If intI = 1 Then
               txtCP118 = "Y"
          End If
       End If
    End If
    'end 2017/12/12
    'Added by Lydia 2018/05/17 排除對象非智慧局(告知代理人901,會稿924,回覆代理人902,其他翻譯927) 和消管制的C類發文
    'Modified by Lydia 2019/12/13 +核對已准專利926
    'Modified by Lydia 2021/02/02 排除藥品連結相關(告代)：959藥品專利連結告代,960專利連結通知,962登錄分析,963資訊登錄,964資訊變更,965回覆第三人通知
    If InStr("901,924,902,927,926,959,960,962,963,964,965", cp(10)) > 0 Or Left(cp(9), 1) = "C" Then
        txtCP118 = ""
    End If
    
    'Added by Lydia 2018/05/17 命名作業有勾選彩圖提申,增加提醒
    'Modified by Lydia 2023/05/24 排除電子送件txtCP118 <> "Y" ---- 'Mark by Lydia 2023/03/03 與Sharon,Phoebe 確認不用列印說明書
    'Modify By Sindy 2023/5/29 拿掉 And txtCP118 <> "Y"
    If InStr(cPrintORI, cp(10)) > 0 And TransDate(cp(5), 2) >= FCP案件命名啟用日 Then
            'Remove by Lydia 2019/07/30 改成衍生設計新案發文時檢查命名記錄尚未分組,才刪除
            'If Text1(1) = "125" Then
            '    strSql = "delete from transcasetitle where tct01='" & Label3(8) & "' and tct05||tct08||tct11 is null "
            '    cnnConnection.Execute strSql, intI
            ''End If
         'Modified by Lydia 2019/07/30 + nvl(tct10,tct04) as grpman
         strExc(0) = "select tct118,nvl(tct10,tct04) as grpman from transcasetitle where tct01='" & cp(9) & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
              'Modify By Sindy 2023/5/29 + And txtCP118 <> "Y"
              If "" & RsTemp.Fields("tct118") = "Y" And txtCP118 <> "Y" Then
                  MsgBox "本案須彩圖提申，請選擇已設定彩色列印的原文本印表機。", vbExclamation
              End If
              m_TCTchk = "" & RsTemp.Fields("grpman") 'Added by Lydia 2019/07/30 命名記錄是否分組/工程師
         End If
    End If
    'end 2018/05/17
    
'Removed by Morgan 2013/4/11 取消--靜芳,FCP案狀況不同無需提醒 Ex.FCP-46826
'   'Added by Morgan 2012/11/14
'   If Text7(2) = "413" Then
'      strExc(1) = PUB_GetFirstPriDate(pa)
'      If strExc(1) = "" Then strExc(1) = DBDATE(pa(10))
'      strExc(2) = CompDate(1, 15, strExc(1))
'      If strSrvDate(1) > strExc(2) Then
'         MsgBox "已超過申請日(優先權日)起算15個月！", vbExclamation
'      End If
'   End If
'   'end 2012/11/14
'end 2013/4/11
End Sub

Private Function ChgType(i As Integer) As Boolean
Dim strTempName As String
   
   ChgType = False
   Select Case i
      Case 0 '發文日
         'Modify By Cheng 2002/07/03
'         If Not ChkDate(Text7(i)) Or Val(Text7(i)) > Val(strSrvDate(2)) Then
'            MsgBox "發文日期不正確或發文日大於系統日，請重新輸入 !", vbCritical
         If Not ChkDate(Text7(i)) Then
         ElseIf DBDATE(Val(Text7(i))) > DBDATE(PUB_GetWorkDayAfterSysDate(Val(strSrvDate(1)))) Then
            MsgBox "發文日大於系統日下一個工作日，請重新輸入 !", vbCritical
         Else
            ChgType = True
         End If
      Case 1
         'ADD BY SONIA 2015/9/21 承辦人為外專程序時,改為操作人員
         Text7(1) = GetFCPUser(Text7(1))
         'END 2015/9/21
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(Text7(i), strTempName) Then
         If ClsPDGetStaff(Text7(i), strTempName) Then
            Label3(4) = strTempName
            ChgType = True
         End If
      Case 2
         If pa(1) = "FCP" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetCaseProperty("FCP", Text7(i), strTempName, False) Then
            If ClsPDGetCaseProperty("FCP", Text7(i), strTempName, False) Then
               Label3(5) = strTempName
               ChgType = True
            End If
         ElseIf pa(1) = "FG" Then
            
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetCaseProperty("FG", Text7(i), strTempName, False) Then
            If ClsPDGetCaseProperty("FG", Text7(i), strTempName, False) Then
               Label3(5) = strTempName
               ChgType = True
            End If
         End If
      Case 12
         strExc(0) = Text7(i)
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetAgent(strExc(0), strTempName) Then
         If ClsPDGetAgent(strExc(0), strTempName) Then
            Text7(i).Text = strExc(0)
            Label3(6) = strTempName
            ChgType = True
         End If
   End Select
End Function

Private Sub Text7_GotFocus(Index As Integer)
  TextInverse Text7(Index)
End Sub

Private Sub Text7_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 13
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 1, 12
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

' 重新計算催審期限
Private Sub ReCaculateSpecDate()
   
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   
   'Add by Morgan 2007/9/14 重新委任自撤不必掛催審期限
   If (Text7(2) = "413" And m_RefCP10 = "928") Then
      Text7(11) = ""
   Else
   'end 2007/9/14
      strSql = "SELECT CF05 FROM CASEFEE " & _
               "WHERE CF01='" & pa(1) & "' AND " & _
                     "CF02='" & pa(9) & "' AND " & _
                     "CF03='" & Text7(2) & "'"
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If Not IsNull(rsTmp.Fields("CF05")) And rsTmp.Fields("CF05") Then
            Text7(11) = TransDate(CompDate(2, Val(rsTmp.Fields("CF05")), TransDate(Text7(0), 2)), 1)
            '2006/5/16 ADD BY SONIA 發明案之分割發文時不掛催審, 於通知實審日才掛
            If pa(8) = "1" And m_CP10 = 分割 Then
               Text7(11) = ""
            End If
            '2006/5/16 END
         End If
      End If
      rsTmp.Close
   End If
   
   Set rsTmp = Nothing
End Sub

Private Sub Text7_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0 '發文日
         If Text7(Index) <> "" Then
            ' 90.08.29 modify by louis
            'If ChgType(Index) = False Then Cancel = True
            If ChgType(Index) = False Then
               Cancel = True
            Else
               'Added by Morgan 2017/9/25
               '當發文日有改時
               If Text7(0).Tag <> Text7(0) Then
                  Text7(0).Tag = Text7(0)
                  SetPayToday
               End If
               'end 2017/9/25
               
               ' 重新計算催審期限
               ReCaculateSpecDate
            End If
         Else
            MsgBox "發文日不可空白 !", vbCritical
            Cancel = True
         End If
      Case 1
         If Text7(Index) <> "" Then
            If ChgType(Index) = False Then Cancel = True
         Else
            MsgBox "承辦人不可空白 !", vbCritical
            Cancel = True
         End If
         'Added by Lydia 2019/08/19 日文組副本:除各組主任(99034,94012)給主管,其餘人給審核主管
         If Text7(Index).Tag <> Text7(Index).Text Then
            'Modified by Lydia 2020/02/10 抓外專工程師主管, 日文組要抓第二、三級主管
            'strExc(5) = PUB_GetStaffST16(Text7(1))
            'If pa(150) = "3" And strExc(5) = "3" Then
            '    strExc(3) = Replace(GetABS001_2(Text7(1)), m_GrpMan, "") '審核主管(replace排除工程師主管)
            '    If strExc(3) = "" Or strExc(3) = "," Then
            '       strExc(3) = m_GrpMan '99034,94012 審核主管只有日文組主管
            '    Else
            '       strExc(3) = Replace(strExc(3), ",", ";")
            '    End If
            '    m_GrpMan = strExc(3)
            'End If
            m_GrpMan = PUB_GetFCPEngSup(Text7(1))
            'end 2020/02/10
         End If
         'end 2019/08/19
         'Add by Morgan 2004/3/23
         '核稿人檢查，檢查失敗時 Cancel=True
         If Cancel = False And txtEP04.Enabled = True Then
            Cancel = Not CheckEP04
         End If
         If Cancel = False Then Text7(Index).Tag = Text7(Index).Text
      Case 2
         If Text7(Index) <> cp(10) Then
            Select Case Text7(Index)
               Case "301", "302", "303", "304", "305", "306", "307", "803"
                  ' 90.08.29 modify by louis
                  'If ChgType(Index) = False Then Cancel = True
                  If ChgType(Index) = False Then
                     Cancel = True
                  Else
                     ' 重新計算催審期限
                     ReCaculateSpecDate
                  End If
               Case Else
                  MsgBox "案件性質只可為改請程序之案件性質或舉發 !", vbCritical
                  Cancel = True
            End Select
         Else
            ' 重新計算催審期限
            ReCaculateSpecDate
         End If
      Case 11
         '鑑定報告
         If Text7(2) = 鑑定報告 And Text7(12) <> "" Then
            If Text7(Index) = "" Then
               MsgBox "案件性質為鑑定報告且有發文代理人時，此欄不可為空白 !", vbCritical
               Cancel = True
            Else
               If Not ChkDate(Text7(Index)) Then
                  MsgBox "日期不正確，請重新輸入 !", vbCritical
                  Cancel = True
               End If
            End If
         End If
      Case 12
         If Text7(Index) <> "" Then
            Text7(Index) = UCase(Text7(Index))
            If Text7(2) = 鑑定報告 Then
               If ChgType(Index) = False Then Cancel = True
            Else
               MsgBox "案件性質為鑑定報告時，才可輸入 !", vbCritical
               Text7(Index) = ""
            End If
         End If
         'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
         If Cancel = False Then
            If PUB_CheckStatus(Text7(Index).Text) = False Then Cancel = True
         End If
      Case 13
         'Add By Sindy 2016/8/18 若有發文規費時, 存檔更新進度檔時同時更新 CP20及CP32 為NULL(即要向客戶請款)
         'Modified by Morgan 2018/10/2 核對中說格式435除外--敏莉
         'Modified by Morgan 2019/11/12 補收款911除外--敏莉
         'modify by sonia 2021/3/25 再取消911
         'If Val(txtCP84.Text) > 0 And Text7(13) = "N" And Text7(2) <> "235" And Text7(2) <> "911" Then
         'Modify By Sindy 2023/6/27 + 排除403更改
         If Val(txtCP84.Text) > 0 And Text7(13) = "N" And Text7(2) <> "235" And Text7(2) <> "403" Then
            MsgBox "有發文規費時，是否向客戶收款不可為N !", vbCritical
            Text7(13).SetFocus
            Cancel = True
         End If
         '2016/8/18 END
         'Modify by Morgan 2004/11/11 不再檢查
'         If Left(strReceiveNo, 1) = "A" Then
'            ' 90.12.19 modify by louis
'            'If Text7(Index) <> "Y" Then
'            'Modify by Morgan 2004/9/29
'            'If Not IsEmptyText(Text7(Index)) Then
'            'Modify by Morgan 2004/11/5 退費908,主動修正203不檢查
'            'If Not IsEmptyText(Text7(Index)) And Val(txtCP84.Text) > 0 Then
'            If Not IsEmptyText(Text7(Index)) And Val(txtCP84.Text) > 0 And Text7(2).Text <> "908" And Text7(2).Text <> "203" Then
'               'MsgBox "A 類收文時，必須為 Y !", vbCritical
'               MsgBox "A 類收文且有發文規費時，必須為空的 !", vbCritical
'               Cancel = True
'            End If
'         End If
      Case 16
         'Modify By Cheng 2002/10/21
'         If Text7(14) = "" Or Text7(15) = "" Or Text7(16) = "" Then
         If Text7(14) = "" And Text7(15) = "" And Text7(16) = "" Then
            MsgBox "案件名稱不可同時空白 !", vbCritical
            Text7(14).SetFocus
            Cancel = True
         End If
      'Added by Lydia 2017/06/21
      Case 4, 7  '聯絡人中文
         If Not CheckLengthIsOK(Text7(Index), 30) Then
            Cancel = True
         End If
      Case 5, 8  '聯絡人英文
         If Not CheckLengthIsOK(Text7(Index), 35) Then
            Cancel = True
         End If
      Case 6, 9, 10 '聯絡人日文,部門
         If Not CheckLengthIsOK(Text7(Index), 60) Then
            Cancel = True
         End If
      'end 2017/06/21
      'Add by Morgan 2006/10/19
      'Modified by Lydia 2017/06/21
      'Case 4, 5, 6, 7, 8, 9, 19
      Case 19
         If Not CheckLengthIsOK(Text7(Index), Text7(Index).MaxLength) Then
            Cancel = True
         End If
      'Added by Morgan 2012/8/1
      Case 20
         If Text7(Index) <> "" Then
            If Not ChkDate(Text7(Index)) Then
               Cancel = True
            End If
         End If
         
   End Select
   If Cancel = True Then TextInverse Text7(Index)
End Sub

Private Function CheckDataValid() As Boolean
CheckDataValid = False
'檢查發文日
If Text7(0) <> "" Then
   ' 90.08.29 modify by louis
   'If ChgType(Index) = False Then Cancel = True
   If ChgType(0) = False Then
      Me.Text7(0).SetFocus
      Text7_GotFocus 0
      Exit Function
   Else
      ' 重新計算催審期限
      ReCaculateSpecDate
   End If
Else
   MsgBox "發文日不可空白 !", vbCritical
   Me.Text7(0).SetFocus
   Text7_GotFocus 0
   Exit Function
End If
'檢查承辦人
If Text7(1) <> "" Then
   If ChgType(1) = False Then
      Me.Text7(1).SetFocus
      Text7_GotFocus 1
      Exit Function
   End If
Else
   MsgBox "承辦人不可空白 !", vbCritical
   Me.Text7(1).SetFocus
   Text7_GotFocus 1
   Exit Function
End If

'Added by Lydia 2018/12/06 新案翻譯之發文,若承辦人與核稿人為同一人,彈訊息"承辦人與核稿人為同一人,不可發文",且不得發文(Sharon)
If Text7(2) = "201" Then
   If Left(Text7(1), 1) = "F" Then
        strExc(1) = PUB_GetMapID(Text7(1), 1)  '抓上班翻譯代號
   Else
        strExc(1) = PUB_GetMapID(Text7(1), 0)   '抓下班翻譯代號
   End If
   If Text7(1) = txtEP04 Or (txtEP04 <> "" And InStr(Text7(1) & "," & strExc(1), txtEP04) > 0) _
        Or (InStr(Label3(4).Caption, lblEP04N.Caption) > 0 And Label3(4).Caption <> "" And lblEP04N.Caption <> "") Then '因為有可能未建立翻譯編號關聯，所以額外判斷名稱
       MsgBox "承辦人與核稿人為同一人，不可發文！", vbCritical
       Exit Function
   End If
End If
'end 2018/12/06

'檢查案件性質
If Text7(2) <> cp(10) Then
   Select Case Text7(2)
      Case "301", "302", "303", "304", "305", "306", "307", "803"
         ' 90.08.29 modify by louis
         'If ChgType(Index) = False Then Cancel = True
         If ChgType(2) = False Then
            Me.Text7(2).SetFocus
            Text7_GotFocus 2
            Exit Function
         Else
            ' 重新計算催審期限
            ReCaculateSpecDate
         End If
      Case Else
         MsgBox "案件性質只可為改請程序之案件性質或舉發 !", vbCritical
         Me.Text7(2).SetFocus
         Text7_GotFocus 2
         Exit Function
   End Select
Else
   ' 重新計算催審期限
   ReCaculateSpecDate
End If
'檢查催審期限
If Text7(2) = 鑑定報告 And Text7(12) <> "" Then
   If Text7(11) = "" Then
      MsgBox "案件性質為鑑定報告且有發文代理人時，此欄不可為空白 !", vbCritical
      Me.Text7(11).SetFocus
      Text7_GotFocus 11
      Exit Function
   Else
      If Not ChkDate(Text7(11)) Then
         MsgBox "日期不正確，請重新輸入 !", vbCritical
         Me.Text7(11).SetFocus
         Text7_GotFocus 11
         Exit Function
      End If
   End If
End If
'檢查發文代理人
If Text7(12) <> "" Then
   Text7(12) = UCase(Text7(12))
   If Text7(2) = 鑑定報告 Then
      If ChgType(12) = False Then
         Me.Text7(12).SetFocus
         Text7_GotFocus 12
         Exit Function
      End If
   Else
      MsgBox "案件性質為鑑定報告時，才可輸入 !", vbCritical
      Text7(12) = ""
   End If
End If

'Remove by Morgan 2004/11/15 不再檢查是否向客戶收款欄位
'檢查是否向客戶收款
'If Left(strReceiveNo, 1) = "A" Then
'   ' 90.12.19 modify by louis
'   'If Text7(Index) <> "Y" Then
'   If Not IsEmptyText(Text7(13)) Then
'      'MsgBox "A 類收文時，必須為 Y !", vbCritical
'      MsgBox "A 類收文時，必須為空的 !", vbCritical
'      Me.Text7(13).SetFocus
'      Text7_GotFocus 13
'      Exit Function
'   End If
'End If
'2004/11/15 end

'檢查專利名稱
'Modify By Cheng 2002/10/21
'If Text7(14) = "" Or Text7(15) = "" Or Text7(16) = "" Then
If Text7(14) = "" And Text7(15) = "" And Text7(16) = "" Then
   MsgBox "案件名稱不可同時空白 !", vbCritical
   Text7(14).SetFocus
   Text7_GotFocus 14
   Exit Function
End If

CheckDataValid = True
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim Chk_CP118 As Boolean 'Add by Amy 2013/05/16

TxtValidate = False
For Each objTxt In Me.Text7
   If objTxt.Enabled = True Then
      Cancel = False
      Text7_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next

'Add by Morgan 2010/5/13
If Text7(2) = "202" Then
   'Modify by Morgan 2010/12/30
   'If InStr(Text7(17), "委任書") > 0 Then
   If InStr(Text7(17), "委任書") > 0 Or InStr(Text7(17), "委任狀") > 0 Then
      If chkCP86(0) + chkCP86(1) = 0 Then
         MsgBox "補文件為委任書時請勾選是否為複委任！"
         If chkCP86(0).Enabled = True Then chkCP86(0).SetFocus
         Exit Function
      End If
   Else
      chkCP86(0).Value = 0
      chkCP86(1).Value = 0
   End If
End If
'end 2010/5/13

'Add by Morgan 2004/4/30
'發明與新型的翻譯，核稿人不可空白！
If pa(8) <> "3" And m_CP10 = "201" And txtEP04.Text = "" Then
   MsgBox "發明與新型的翻譯，核稿人不可空白！", vbExclamation
   If txtEP04.Enabled Then txtEP04.SetFocus
   Exit Function
End If

'Add by Morgan 2004/3/23
'FCP及P案件,當專利種類為'發明'且案件性質為'分割'時,若無收文未取消收文之'實體審查'則顯示'此分割案尚未收文實體審查，期限為XXXXXX，請提醒智權人員 !!
'讀取母案申請日
Erase m_stVar(): m_stCP09 = ""
Erase stDivInfo
bol416Control = True 'Added by Morgan 2013/12/11
m_bol435 = False 'Added by Morgan 2015/9/9 是否管制續行母案再審

Dim strTmp1(0 To 4) As String, strTmp(1 To 3) As String, i As Integer, bolContinue As Boolean

If pa(8) = "1" And Text7(2).Text = "307" Then
   
   'Modified by Morgan 2025/3/13 +bolContinue 母案可能為非本所案件 Ex:FCP-073311--Winfrey
   If GetDivInfo(stDivInfo(), pa(), , bolContinue) = False Then
      If bolContinue = False Then
         Exit Function
      Else
         bol416Control = False
         MsgBox "無法計算 實體審查/續行母案再審 期限請自行管制！", vbExclamation
      End If
   'end 2025/3/13
   ElseIf stDivInfo(5) <> "000" Then
      MsgBox "發明分割案的母案申請國家必須為台灣！", vbCritical
      Exit Function
      
   Else
      strTmp1(0) = Label3(0).Caption
      For i = 1 To 4
         strTmp1(i) = stDivInfo(i)
      Next
      
      'Modified by Morgan 2013/12/11
      '若已收文435(續行母案再審)則不必再管制實審期限--靜芳
      'Modified by Morgan 2015/9/10 更新435(續行母案再審)期限為發文日+30天
      If PUB_ChkCPExist(pa(), "435", , m_stCP09) = True Then
         bol416Control = False
         m_bol435 = True
         
      Else
         'Modified by Morgan 2015/9/30 未設定才讓User確認
         If m_PA163 = "" Then
            m_PA163 = PUB_GetDivCaseState(pa(), strSrvDate(1), True)
            If m_PA163 = "" Then
               intI = MsgBox("資訊不足無法判斷!!請問本案是否為初審階段提分割??", vbYesNoCancel + vbQuestion + vbDefaultButton3)
               If intI = vbYes Then
                  m_PA163 = "Y"
               ElseIf intI = vbNo Then
                  m_PA163 = "N"
               Else
                  Exit Function
               End If
            End If
         End If
         'end 2015/9/30
         
         If m_PA163 = "N" Then
            MsgBox "此分割案尚未收文435(續行母案再審)，請提醒智權人員!!!", vbExclamation
            bol416Control = False
            m_bol435 = True
            
         'end 2015/9/10
         '讀取實體審查得法定期限
         ElseIf GetMoneyDate(4, stDivInfo(5), strTmp1, strTmp(1), strTmp(2), strTmp(3)) = True Then
            If strTmp(3) <> "" Then
               strTmp(3) = CompDate(2, 1, strTmp(3))
               '法定期限
               m_stVar(3) = PUB_Get416LawLimit(Text7(0), strTmp(3))
               
               'Modified by Morgan 2014/11/20 外專改回舊規則
               ''Added by Morgan 2014/10/29
               'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
               '   m_stVar(0) = PUB_GetOurDeadline(m_stVar(3))
               'Else
               ''end 2014/10/29
               
               'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
               If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
                  'Modify By Sindy 2021/4/27 + m_pAgreeOnDate
                  m_stVar(0) = PUB_GetFCPOurDeadline(m_stVar(3), 4, , m_pAgreeOnDate)
               Else
               'end 2019/7/11
      
                  '本所期限= 法定期限-4天,不必管假日 93/8/6
                  m_stVar(0) = CompDate(2, -4, m_stVar(3))
               
               End If 'Added by Morgan 2019/7/11
               'End If 'Added by Morgan 2014/10/29
               'end 2014/11/20
               
               '檢查有收'實體審查'否，有則抓收文號-->m_stCP09
               If PUB_Get416CP09(m_stCP09, ChangeWStringToTString(m_stVar(0)), pa()) = False Then
                  Exit Function
               End If
            Else
               MsgBox "無法讀取實體審查的法定期限！", vbCritical
               Exit Function
            End If
            
         Else
            MsgBox "無法讀取實體審查的法定期限！", vbCritical
            Exit Function
            
         End If 'Added by Morgan 2013/12/11
         
      End If
      
      'Added by Morgan 2015/9/10
      If m_bol435 Then
         'Modified by Morgan 2022/5/3
         ''收文日/發文日+30天
         'm_stVar(3) = PUB_Get416LawLimit(Text7(0), Text7(0))
         ''Modified by Morgan 2019/8/23
         ''m_stVar(0) = CompDate(2, -4, m_stVar(3))
         ''Modify By Sindy 2021/4/27 + m_pAgreeOnDate
         'm_stVar(0) = PUB_GetFCPOurDeadline(m_stVar(3), 4, , m_pAgreeOnDate)
         
         '分割案發文時，產生下一程序續行母案再審期限，為分割案發文日+4個月--陳亭妙
         m_stVar(3) = CompDate(1, 4, (Text7(0)))
         m_stVar(0) = PUB_GetFCPOurDeadline(m_stVar(3), 4, , m_pAgreeOnDate)
         'end 2022/5/3
      End If
      'end 2015/9/10
   End If
   
'Add by Morgan 2004/8/6
'改請發明案也要控制實體審查
ElseIf Text7(2).Text = 改請發明 Then

   '讀取實體審查得法定期限
   If GetMoneyDate(4, pa(9), pa, strTmp(1), strTmp(2), strTmp(3)) = True Then
      If strTmp(3) <> "" Then
         strTmp(3) = CompDate(2, 1, strTmp(3))
         '法定期限
         m_stVar(3) = PUB_Get416LawLimit(Text7(0), strTmp(3))
         '本所期限= 法定期限-4天
         'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
         m_stVar(0) = PUB_GetWorkDay1(CompDate(2, -4, m_stVar(3)), True)
         '檢查有收'實體審查'否，有則抓收文號-->m_stCP09
         If PUB_ChkCPExist(pa, 實體審查, 0, m_stCP09) = False Then
            MsgBox "此改請案尚未收文實體審查，期限為【" & ChangeWStringToTString(m_stVar(0)) & "】，請提醒智權人員!!!", vbExclamation
         End If
      Else
         MsgBox "無法讀取實體審查的法定期限！", vbCritical
         Exit Function
      End If
   Else
      MsgBox "無法讀取實體審查的法定期限！", vbCritical
      Exit Function
   End If

End If

   'Added by Morgan 2023/10/16
   '設計案的分割/改請設計/改請衍生設計要檢查是否初審階段提出
   '分割:母案1.已核駁則為再審階段,2.為申請案且未審定則為初審階段,3.其他則詢問User後設定
   '改請:原案1.已核駁則為再審階段,2.為申請案且未審定則為初審階段,3.其他則詢問User後設定
   If pa(8) = "3" And (Text7(2) = "307" Or Text7(2) = "303" Or Text7(2) = "308") And m_PA163 = "" Then
      '分割
      If Text7(2) = "307" Then
         strExc(0) = "select pa16,cp09,cp10,cp24 from divisioncase,patent,caseprogress" & _
            " where dc01='" & pa(1) & "' and dc02='" & pa(2) & "' and dc03='" & pa(3) & "' and dc04='" & pa(4) & "'" & _
            " and pa01(+)=dc05 and pa02(+)=dc06 and pa03(+)=dc07 and pa04(+)=dc08" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10(+)='103' and cp09(+)<'B' and cp24(+) is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '母案已核駁
            If RsTemp("pa16") = "2" Then
               m_PA163 = "N"
            '母案為設計申請案且未審定
            ElseIf IsNull(RsTemp("pa16")) And Not IsNull(RsTemp("cp09")) Then
               m_PA163 = "Y"
            End If
         End If
      '改請
      Else
         '原案已核駁
         If pa(16) = "2" Then
            m_PA163 = "N"
         '原案為申請案且未審定
         ElseIf pa(16) = "" Then
            strExc(0) = "select cp10 from caseprogress" & _
               " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
               " and cp10 in ('101','102','103') and cp09<'B' and cp24 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_PA163 = "Y"
            End If
         End If
      End If
      If m_PA163 = "" Then
         intI = MsgBox("無法自動判斷!!" & vbCrLf & "請問本案是否為初審階段提" & Label3(5) & "??", vbYesNoCancel + vbExclamation + vbDefaultButton3)
         If intI = vbYes Then
            m_PA163 = "Y"
         ElseIf intI = vbNo Then
            m_PA163 = "N"
         Else
            Exit Function
         End If
      End If
   End If
   'end 2023/10/16
         
   'Add by Morgan 2004/8/12
   If txtCP84.Enabled = True Then
      Cancel = False
      txtCP84_Validate Cancel
      If Cancel = True Then
         txtCP84.SetFocus
         txtCP84_GotFocus
         Exit Function
      End If
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
   
   'Add by Morgan 2006/5/1 實審發文之申請寄存及存活證明控制
   m_bol108 = False
   Dim strCP09 As String, bolDivCase As Boolean
   '判斷是否為分割案並讀取母案資料
   If Text7(2).Text = "416" Or (pa(8) = "1" And Text7(2).Text = "307") Then
      Erase stDivInfo
      If GetDivInfo(stDivInfo, cp, False) = True Then
         bolDivCase = True
      End If
   End If
   '發明,改請發明,實審,發明分割要檢查是否有收文申請寄存
   If Text7(2).Text = "101" Or Text7(2).Text = "301" Or Text7(2).Text = "416" Or (pa(8) = "1" And Text7(2).Text = "307") Then
      If bolDivCase = True Then
         m_bol108 = PUB_ChkCPExist(stDivInfo, "108", 0)
      Else
         m_bol108 = PUB_ChkCPExist(cp, "108", 0)
      End If
      '有收申請寄存108且為實審發文
      If m_bol108 = True And Text7(2).Text = "416" Then
         If bolDivCase = True Then
            If PUB_ChkCPExist(stDivInfo, "108", 1) = True Then
               MsgBox "本案為分割案，母案有【申請寄存】尚未發文！", vbExclamation
               Exit Function
            End If
         Else
            If PUB_ChkCPExist(cp, "108", 1) = True Then
               MsgBox "本案有【申請寄存】尚未發文！", vbExclamation
               Exit Function
            End If
         End If
         If PUB_ChkCPExist(cp, "221") = False Then
            If bolDivCase = True Then
               MsgBox "本案為分割案，母案有收文【申請寄存】但本案尚未收文【存活證明】！", vbExclamation
            Else
               MsgBox "本案有收文【申請寄存】但尚未收文【存活證明】！", vbExclamation
            End If
            Exit Function
         ElseIf PUB_ChkCPExist(cp, "221", 1) = True Then
            MsgBox "本案有【存活證明】尚未發文！", vbExclamation
            Exit Function
         End If
      End If
   End If
   '2006/5/1 end
   
   'Add by Morgan 2007/7/19
   txtCP113_Validate Cancel
   If Cancel = True Then
      txtCP113.SetFocus
      Exit Function
   End If
   txtCP114_Validate Cancel
   If Cancel = True Then
      txtCP114.SetFocus
      Exit Function
   End If
   'end 2007/7/19
   
   'Add By Sindy 2020/10/19 檢查是否有同一天發文的超頁超項費,若有請拿掉發文日,重新發文
   If Val(txtCP135) > 0 Or Val(txtCP136) > 0 Then
      strExc(0) = "select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "'" & _
                  " and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp158>0 and cp159=0" & _
                  " and cp10 in('938','939') and cp27=" & DBDATE(Text7(0).Text)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            MsgBox "本案有【超頁超項費】同時發文,請拿掉【超頁超項費】的發文日,一併重新發文！", vbExclamation
            Exit Function
         End If
      End If
   End If
   '2020/10/19 END
   
   'Add by Morgan 2010/1/6
   m_lngOverPageFee = 0
   m_lngOverPageFee = 0
   m_FeeMemo = ""
   If m_bolChkFee Then
      'Modify By Sindy 2019/3/26
      'SetOfficialFee
      'Modify By Sindy 2023/3/16 +, , , txtCP167
      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
                                txtCP135, txtCP136, txtCP137, txtCP84, , txtDecreaseItemFee, _
                                m_lngOverPageFee, m_lngOverItemFee, , , m_lngOfficialFee, _
                                Text7(0).Text, , txtDecreasePageFee, txtCP167, , m_str938RecvNo, m_str939RecvNo)
      '2019/3/26 END
      'Modify By Sindy 2023/3/24 Mark,改呼叫共用函數
'      If Not CheckOfficialFee Then
'         Exit Function
'      End If
      If Not PUB_CheckOfficialFee_P(cp(), m_bolChkPageItem, m_bolChkItem, _
                                    txtCP135, txtCP136, txtCP137, txtCP138, txtCP84, _
                                    m_lngRecOverPageFee, m_lngRecOverItemFee, m_FeeMemo, _
                                    m_lngOverPageFee, m_lngOverItemFee, _
                                    m_lngOverPageFeeDiff, m_lngOverItemFeeDiff, txtCP167, txtCP168, True) Then
         Exit Function
      End If
      '2023/3/24 END
   End If
   'end 2010/1/6
   
   'Modify By Sindy 2018/9/5 Mark ex:FCP-59301新案一併實體送件,但之前又先送主動修正了,導致會彈頁/項數不可空白,狀況之多暫不控制
'   'Add By Sindy 2018/5/25
'   'Modify By Sindy 2018/6/26 敏莉說不控管210.製作中說(取消Or cp(10) = "210")
''   If m_bolChkPageItem = True Or _
''      (txtCP118 = "" And (cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "307")) Then
'   If m_bolChkPageItem = True Then
'      If Val(txtCP135) = 0 And txtCP135.Enabled = True Then
'         MsgBox "頁數不可空白！", vbExclamation
'         txtCP135.SetFocus
'         Exit Function
'      ElseIf Val(txtCP136) = 0 And txtCP136.Enabled = True Then
'         MsgBox "項數不可空白！", vbExclamation
'         txtCP136.SetFocus
'         Exit Function
'      End If
'   End If
   
   'Add by Morgan 2011/1/17
   'Modified by Morgan 2013/11/6 +235核對中說格式
   If m_bolWebApp = True And (Text7(2) = "201" Or Text7(2) = "209" Or Text7(2) = "235" Or Text7(2) = "210") And txtCP118 <> "Y" Then
      MsgBox "本案應以電子送件方式呈送!!", vbExclamation
      txtCP118.SetFocus
      Exit Function
   End If
   'end 2011/1/17
   
   'Add by Amy 2013/05/16
   'Modified by Morgan 2018/1/29 電子送件要考慮自動扣款 cp118='Y' => cp118 is not null
   strExc(0) = "Select CP118 From CaseProgress Where CP09='" & cp(9) & "' And CP118 is not null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Chk_CP118 = intI Xor IIf(txtCP118 = "Y", 1, 0)
   If Text7(2) = "101" And Chk_CP118 Then
      MsgBox "新案建檔設定之是否電子送件與目前輸入不一致，不可發文!!", vbExclamation
      txtCP118.SetFocus
      Exit Function
   End If
   'end 2013/05/16
   
'   'Add by Morgan 2011/8/15
'   If Text7(2) = "422" Then
'      If Val(txtCP84) = 0 Then
'         If MsgBox("是否不為商業上之施行之必要？若選[ 是 ]將繼續發文，選[ 否 ]則請輸入發文規費！", vbYesNo + vbDefaultButton2) = vbNo Then
'            txtCP84.SetFocus
'            Exit Function
'         End If
'
'      ElseIf txtCP84 <> "4000" Then
'         MsgBox "規費金額錯誤，請重新輸入！"
'         txtCP84.SetFocus
'         Exit Function
'      End If
'   End If
'   'end 2011/8/15
   'Modify By Sindy 2023/5/29
   If Text7(2).Text = "422" Then '加速審查
      If Val(txtCP84) = 0 Then
         If MsgBox("請確認是否為 1.以商業上之實施所必要 或 2.為綠色技術相關案件(是：需繳納4000元)" & vbCrLf & _
                   "是: 則發文規費帶4000 並完成發文" & vbCrLf & _
                   "否: 則完成發文", vbYesNo + vbDefaultButton2) = vbYes Then
            txtCP84 = "4000"
         End If
      End If
   End If
   '2023/5/29 END
   
   'Added by Morgan 2012/4/11
   If Text7(2) = "401" And Me.Tag <> "True" Then
      MsgBox "案件性質為[ " & Label3(5) & " ]，需先點變更事項並按確定！"
      cmdok_Click 4
      Exit Function
   End If
   
   'Added by Morgan 2012/8/1
   If Text7(20).Visible = True Then
      If Text7(20) = "" Then
         If InStr(Text7(18), "例行") > 0 Then
            MsgBox "例行案件必須輸入下次管制日期！"
            Text7(20).SetFocus
            Exit Function
            
         ElseIf MsgBox("本案是否為例行案件?", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
            MsgBox "例行案件必須輸入下次管制日期！"
            Text7(20).SetFocus
            Exit Function
         End If
      ElseIf InStr(Text7(18), "例行") = 0 Then
         Text7(18) = "例行;" & Text7(18)
      End If
   End If
   'end 2012/8/1
   
   m_EdDivSugInform = False 'Added by Morgan 2020/2/26
   'Added by Morgan 2012/12/3
   m_PA162 = ""
   'Modified by Morgan 2019/10/5 +改發明/新型的再審107,主動修正203,申復204,修正205
   'If Text7(2) = "204" Or Text7(2) = "205" Then
   If (pa(8) = "1" Or pa(8) = "2") And (Text7(2) = "107" Or Text7(2) = "203" Or Text7(2) = "204" Or Text7(2) = "205") Then
   'end 2019/10/5
   
      '發明實審
      'Modified by Morgan 2013/10/7 +307--靜芳
      'Modified by Morgan 2019/10/5 +不必再限定發明初審階段,但發明的主動修正要判斷實審/再審已發文,新型的主動修正要判斷新型/改請新型已發文
      'strExc(0) = "select pa162,dst09 from caseprogress a,patent,divsugtext" & _
         " where cp09='" & Label3(0) & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa08='1'" & _
         " and dst01(+)=pa01 and dst02(+)=pa02 and dst03(+)=pa03 and dst04(+)=pa04" & _
         " and exists(select * from caseprogress b where b.cp01=pa01 and b.cp02=pa02 and b.cp03=pa03 and b.cp04=pa04 and b.cp10 in ('101','307') and b.cp27>0)" & _
         " and exists(select * from caseprogress b where b.cp01=pa01 and b.cp02=pa02 and b.cp03=pa03 and b.cp04=pa04 and b.cp10='416' and b.cp27<" & DBDATE(Text7(0)) & ")" & _
         " and not exists(select * from caseprogress b where b.cp01=pa01 and b.cp02=pa02 and b.cp03=pa03 and b.cp04=pa04 and b.cp10='107' and b.cp27>0)"
      strExc(0) = "select pa162,dst09 from caseprogress a,patent,divsugtext" & _
         " where cp09='" & Label3(0) & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and pa08 in ('1','2') and dst01(+)=pa01 and dst02(+)=pa02 and dst03(+)=pa03 and dst04(+)=pa04"
         
      If Text7(2) = "203" Then
         If pa(8) = "1" Then
            strExc(0) = strExc(0) & " and exists(select * from caseprogress b where b.cp01=pa01 and b.cp02=pa02 and b.cp03=pa03 and b.cp04=pa04 and b.cp10 in ('416','107') and b.cp27<" & DBDATE(Text7(0)) & ")"
         ElseIf pa(8) = "2" Then
            strExc(0) = strExc(0) & " and exists(select * from caseprogress b where b.cp01=pa01 and b.cp02=pa02 and b.cp03=pa03 and b.cp04=pa04 and b.cp10 in ('102','302') and b.cp27<" & DBDATE(Text7(0)) & ")"
         End If
      End If
      'end 2019/10/5
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp("pa162") = "Y" Then
            strExc(1) = "目前設定為""是"""
            
            'Added by Morgan 2019/10/8
            If Not IsNull(RsTemp("dst09")) Then '工程師有進維護確認
               If RsTemp("dst09") <> cp(9) Then
                  'Modified by Morgan 2020/2/26 改可發文另EMail通知工程師
                  'MsgBox "本次" & Label3(5) & "工程師尚未輸入分割建議定稿，請通知其維護該筆進度資料！", vbCritical, Label3(5) & "發文"
                  'Exit Function
                  m_EdDivSugInform = True
                  'end 2020/2/26
               End If
            End If
            'end 2019/10/8
            
         ElseIf RsTemp("pa162") = "N" Then
            strExc(1) = "目前設定為""否"""
         Else
            strExc(1) = "目前未設定"
            
            'Added by Morgan 2019/10/8 改未設定才問--Sharon
            intI = MsgBox("是否加註核准分割建議？" & vbCrLf & vbCrLf & "( " & strExc(1) & " )", vbYesNoCancel + vbDefaultButton3 + vbQuestion, Label3(5) & "發文")
            If intI = vbCancel Then
               Exit Function
            ElseIf intI = vbYes Then
               m_PA162 = "Y"
            ElseIf intI = vbNo Then
               m_PA162 = "N"
            End If
            'end 2019/10/8
         End If
         
         'Removed by Morgan 2019/10/8 改未設定才問
         'intI = MsgBox("是否要另函通知初審核准後分割？" & vbCrLf & vbCrLf & "( " & strExc(1) & " )", vbYesNoCancel + vbDefaultButton3 + vbQuestion, Label3(5) & "發文")
         'If intI = vbCancel Then
         '   Exit Function
         'Else
         '   If intI = vbYes Then
         '      m_PA162 = "Y"
         '   ElseIf intI = vbNo Then
         '      m_PA162 = "N"
         '   End If
         '   If Not IsNull(RsTemp("dst09")) Then '工程師有進維護確認
         '      If RsTemp("pa162") <> m_PA162 Then
         '         MsgBox "您輸入的選項與目前工程師設定不同，請再確認承辦單內容！", vbCritical, "申復修正發文"
         '         Exit Function
         '      ElseIf RsTemp("pa162") = "Y" And RsTemp("dst09") <> cp(9) Then
         '         MsgBox "本次" & Label3(5) & "工程師尚未輸入分割建議定稿，請通知其維護該筆進度資料！", vbCritical, Label3(5) & "發文"
         '         Exit Function
         '      End If
         '   End If
         'End If
         'end 2019/10/8
         
      End If
   End If
   'end 2012/12/3
   
   'Added by Morgan 2013/1/3
   If (cp(10) = "421" Or cp(10) = "807") And txtCP136.Enabled = True Then
      If txtCP136 = "" Then
         MsgBox "請輸入項數!!!"
         txtCP136.SetFocus
         Exit Function
      End If
   End If
   'end 2013/1/3
   
   'Added by Morgan 2015/10/13
   '若有主張優先權時檢查優先權證明的補件期限筆數是否與優先權數相符
   'Modify By Sindy 2021/9/8 + Or Text7(2) = "105"
   '發文的案件性質=101發明申請,102新型申請,103設計申請,105聯合申請,125衍生設計申請
   If Text7(2) = "101" Or Text7(2) = "102" Or Text7(2) = "103" Or Text7(2) = "105" Or Text7(2) = "125" Then
      strExc(0) = "select * from pridate where pd01='" & pa(1) & "' and pd02='" & pa(2) & "' and pd03='" & pa(3) & "' and pd04='" & pa(4) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(1) = RsTemp.RecordCount
         strExc(2) = 0
         strExc(0) = "select * from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07='202' and instr(np15,'優先權證明')>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(2) = RsTemp.RecordCount
         End If
         If Val(strExc(1)) <> Val(strExc(2)) Then
            If MsgBox("本案優先權資料有 " & strExc(1) & " 筆與優先權證明補文件期限 " & strExc(2) & " 筆不符！是否仍要繼續發文？", vbQuestion + vbYesNo + vbDefaultButton2, "優先權文件之管制補呈個數有誤") = vbNo Then
               Exit Function
            End If
         End If
      End If
      
      'Add By Sindy 2021/9/8 檢查是否有主動修正未發文未取消收文，並且進度備註是否有"分案-提申前"或沒有註明屬什麼提申字樣
      '則彈詢問訊息，因若為提申後需在申請案發文時，主動修正的案件備註裡需有"提申後"的字樣，才會更新主動修正的期限
      strExc(0) = "select cp64 from caseprogress WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                  " and cp10 ='203' and (instr(cp64,'分案-提申前') > 0 or instr(cp64,'提申')=0)" & _
                  " and CP27 IS NULL AND CP57 IS NULL "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If MsgBox("有【提申前主動修正】未發文，是否為提申前？" & vbCrLf & _
            "是：繼續發文作業" & vbCrLf & _
            "否：將此主動修正的備註自動改為【提申後主動修正】", vbQuestion + vbYesNo + vbDefaultButton2, "主動修正的提醒") = vbNo Then
            '否：
            strSql = "UPDATE CASEPROGRESS SET cp64='" & ChangeWStringToTDateString(strSrvDate(1)) & "改為提申後主動修正;'||cp64" & _
                     " WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                     " and cp10 ='203' and CP27 IS NULL AND CP57 IS NULL "
            cnnConnection.Execute strSql
         End If
      End If
      '2021/9/8 END
   End If
   'end 2015/10/13
   
   'Add By Sindy 2015/12/17 檢查是否有指定送件日期,若有不可小於指定日期送件
   If cp(142) <> "" Then
      'Modified by Morgan 2016/1/14
      'If cp(142) > strSrvDate(1) Then
       'Modify By Sindy 2021/12/3 淑華說之後可以含當天發文
      'If cp(142) >= DBDATE(Text7(0)) Then
      If cp(142) > DBDATE(Text7(0)) Then
      'Sindy 2021/12/3 END
         'Add By Sindy 2021/4/20
         'Modify By Sindy 2021/10/20 + 3.之後
         If ((cp(164) = "1" Or cp(164) = "") And cp(142) > DBDATE(Text7(0))) Or _
            cp(164) = "3" Then '1.當天 3.之後
         '2021/4/20 END
            MsgBox "有指定送件日期（" & ChangeWStringToTDateString(cp(142)) & "），不可提前送件!!!"
            Exit Function
         End If
      End If
   End If
   '2015/12/17 END
   
   bolAddSC = False 'Added by Lydia 2016/01/21
   'Added by Lydia 2015/12/31 FCP案會稿924發文時，檢查一定要有相關總收文號
   If Text7(2) = "924" And cp(1) = "FCP" And txtCP43.Visible = True Then
      If txtCP43 = "" Then
         MsgBox "案件性質為會稿時，一定要有相關總收文號!!", vbCritical
         txtCP43.SetFocus
         txtCP43_GotFocus
         Exit Function
      'Modified by Lydia 2019/07/03 改成只針對011A日本區之代理人
      'Else
      'Modified by Lydia 2019/09/17 +案件為Y52218000(PanKorea Patent & Law Firm)+X80582000(三星BIOEPIS)
      'ElseIf mFA10 = "011" Then
      'Modified by Lydia 2019/11/13 取消日本代理人011(A字母開頭)追蹤會稿之行事曆期限彈跳之控管
      'ElseIf mFA10 = "011" Or (ChangeCustomerL(pa(75)) = "Y52218000" And ChangeCustomerL(pa(26)) = "X80582000") Then
      ElseIf ChangeCustomerL(pa(75)) = "Y52218000" And ChangeCustomerL(pa(26)) = "X80582000" Then
         'Added by Lydia 2016/01/21 詢問是否產生行事曆管制 'Move by Lydia 2016/02/03
         'Mark by Lydia 2019/07/03 改成只針對011A日本區之代理人的的翻譯201、檢視中說209、製作中說210和核對中說格式235之會稿，彈跳提醒;申復等其它中間程序之會稿發文不用彈跳提醒
         'If MsgBox("是否產生管制會稿結果期限?", vbInformation + vbYesNo) = vbYes Then
         '    bolAddSC = True
         'End If
         ''end 2016/01/21
         'If bolAddSC = True Then 'Added by Lydia 2016/02/03
            strExc(0) = "select cp06,cp10,cp27,cp57 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp09='" & txtCP43 & "' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_CP43cpm = "" & RsTemp.Fields("cp10")
               m_CP43date1 = "" & RsTemp.Fields("cp06")
                'Move by Lydia 2016/02/03
               'Added by Lydia 2016/02/03
               'Mark by Lydia 2019/07/03
               'If m_CP43date1 = "" Then
               '   MsgBox "相關總收文號無所限!!", vbCritical
               '   txtCP43.SetFocus
               '   txtCP43_GotFocus
               '   Exit Function
               'End If
            Else
               MsgBox "相關總收文號不存在!!", vbCritical
               txtCP43.SetFocus
               txtCP43_GotFocus
               Exit Function
            End If
         'End If
         'Added by Lydia 2019/07/03 先判斷是否為中說之會稿,再彈提醒
         If InStr("201,209,210,235", m_CP43cpm) > 0 Then
            If m_CP43date1 = "" Then
               MsgBox "相關總收文號無所限!!", vbCritical
               txtCP43.SetFocus
               txtCP43_GotFocus
               Exit Function
            End If
            If MsgBox("是否產生管制會稿結果期限?", vbInformation + vbYesNo) = vbYes Then
                bolAddSC = True
            End If
         End If
         'end 2019/07/03
      End If
   End If
   'end 2015/12/31
   
   'Added by Morgan 2017/9/25
   If txtCP118 = "Y" And Val(txtCP84) > 0 Then
      If txtPayToday = "" Then
         MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
         txtPayToday.SetFocus
         Exit Function
      End If
   End If
   'end 2017/9/25
   
'   'Add By Sindy 2019/1/18 更改且有公告公報進度時
'   If text1 = "FCP" And Text7(2) = "403" And Trim(txtCP43) = "" Then
'      If PUB_ChkCPExist(cp, "1228", , strCP09) Then '公告公報
'         If MsgBox("是否為公告公報勘誤的更改？", vbInformation + vbYesNo) = vbYes Then
'            txtCP43 = strCP09
'         End If
'      End If
'   End If
   
   TxtValidate = True
End Function

Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

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
   Cancel = Not PUB_CheckCP113(txtCP113, pa(1), m_CP10, Text7(1))
   
End Sub

Private Sub txtCP114_GotFocus()
   TextInverse txtCP114
End Sub

Private Sub txtCP114_Validate(Cancel As Boolean)
   If txtCP114 <> "" Then
      If Not IsNumeric(txtCP114) Then
         MsgBox "請輸入數字！", vbExclamation
         Cancel = True
         Exit Sub
      'Add by Morgan 2008/3/12
      ElseIf Val(txtCP114) > 25 Then
         If txtCP114.Tag <> txtCP114 Then
            If MsgBox("核稿時數超過25小時，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
               Cancel = True
               Exit Sub
            End If
            txtCP114.Tag = txtCP114
         End If
      End If
   Else
      If Text7(2) = "201" Then
         'Modify by Morgan 2007/9/4 判斷核稿人非承辦人
         If Text7(1) <> txtEP04 Then
            MsgBox Label3(5) & "核稿人非承辦人時核稿時數不可空白！", vbExclamation
            txtCP114.SetFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub txtCP118_Change()
   SetPayToday
   setCP84 'Added by Morgan 2018/1/29 因改為會預設故從 Validate 事件移來
End Sub

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

Private Sub setCP84()
Dim strFee As String 'Added by Lydia 2017/12/11
   'Modified by Lydia 2017/12/12 新案申請用模組計算規費
'   If txtCP118 = "Y" Then
'      If Text7(2) = "101" Then
'         txtCP84 = "2900"
'      ElseIf Text7(2) = "102" Or Text7(2) = "103" Then
'         txtCP84 = "2400"
'      End If
'   End If
   
   'Modify By Sindy 2019/2/21 Mark,不能加此if判斷,申請書規費是總額,但發文時新案一個規費,實審又一個規費要在此計算
   'Add By Sindy 2019/2/18 發文規費無值才計算
'   If Val(txtCP84) = 0 Then
   '2019/2/18 END
      If InStr(NewCasePtyList, Text7(2)) > 0 Then
         strFee = GetPatentOfficialFee(cp(1), cp(10), cp(7), pa(8), pa(9), pa(16), pa(14), pa(2), pa(3), pa(4), Trim(txtCP118))
         If Val(strFee) > 0 Then
            txtCP84 = strFee
            'Add By Sindy 2019/2/21
            If InStr(cp(64), "減收申請規費800元") > 0 Then
               txtCP84 = Val(txtCP84) - 800
            End If
            '2019/2/21 END
         End If
      End If
'   End If
   'end 2017/12/12
End Sub

Private Sub txtCP118_Validate(Cancel As Boolean)
   'Add By Sindy 2018/5/25 ex:FCP-058321.中說
   'Modify By Sindy 2018/6/26 敏莉說不控管210.製作中說(取消Or cp(10) = "210")
   'Modify By Sindy 2018/7/27 敏莉說分割的設計案不用輸入頁,項數 ex:FCP-59278
   If m_bolChkPageItem = True Or _
      (txtCP118 = "" And (cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or (cp(10) = "307" And pa(8) <> "3"))) Then
      m_bolChkPageItem = True
      lblCP135.Caption = "總頁數:" 'Add By Sindy 2023/3/16
'      txtCP135.Enabled = True
'      txtCP135.BackColor = vbWhite
'      If Val(cp(135)) = 0 Then txtCP135 = m_allPage
      lblCP136.Caption = "總項數:"
      txtCP136.Enabled = True
      txtCP136.BackColor = vbWhite
'      'Modify By Sindy 2018/6/11 ex:FCP-050420:再審申請
'      'If Val(cp(136)) = 0 Then txtCP136 = m_allItem
'      'Modify By Sindy 2019/1/15 ex:FCP-051198:再審申請
'      'txtCP136 = m_allItem
'      If Val(cp(136)) = 0 Then
'         txtCP136 = m_allItem
'      Else
'         txtCP136 = cp(136)
'      End If
'      '2018/6/11 END
   End If
   '2018/5/25 END
   'Add By Sindy 2019/4/16 存檔時會以此條件清除該案號進度的頁,項數 ex:FCP-060726;FCP-060493
   If m_bolChkPageItem = True Or _
      cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Then
      If Val(cp(135)) = 0 Then
         txtCP135 = m_allPage
      'Modify By Sindy 2023/4/18
      Else
         txtCP136 = cp(135)
      End If
      '2023/4/18 END
      'Modify By Sindy 2018/6/11 ex:FCP-050420:再審申請
      'If Val(cp(136)) = 0 Then txtCP136 = m_allItem
      'Modify By Sindy 2019/1/15 ex:FCP-051198:再審申請
      'txtCP136 = m_allItem
      If Val(cp(136)) = 0 Then
         txtCP136 = m_allItem
      Else
         txtCP136 = cp(136)
      End If
      '2018/6/11 END
   End If
   '2019/4/16 END
   'Add By Sindy 2018/1/31 電子送件實審發文時,總頁數及總項數欄位鎖住不得由實審發文處更改,僅能重新產生申請書來更改。
   'Modify By Sindy 2018/5/22 + 因紙本實體審查申請書已開放輸入頁項數
   'Modify By Sindy 2018/6/11 + 紙本實體審查申請書雖有程序輸入,但若加主動修正是工程師印出紙本給程序,所以紙本送件均不鎖頁項數
   If txtCP118 = "Y" And cp(10) = 實體審查 Then
   'If cp(10) = 實體審查 Then
   '2018/5/22 END
      txtCP135.Enabled = False
      txtCP135.BackColor = Me.BackColor
      txtCP136.Enabled = False
      txtCP136.BackColor = Me.BackColor
   End If
   '2018/1/31 END
   If m_bolChkFee Then
      'Modify By Sindy 2019/3/26
      'SetOfficialFee
      'Modify By Sindy 2023/3/16 +, , , txtCP167
      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
                                txtCP135, txtCP136, txtCP137, txtCP84, , txtDecreaseItemFee, _
                                m_lngOverPageFee, m_lngOverItemFee, , , m_lngOfficialFee, _
                                Text7(0).Text, , txtDecreasePageFee, txtCP167, , m_str938RecvNo, m_str939RecvNo)
      '2019/3/26 END
   End If
End Sub

Private Sub txtCP135_GotFocus()
   TextInverse txtCP135
   CloseIme
End Sub

Private Sub txtCP135_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP135_Validate(Cancel As Boolean)
   'Modify By Sindy 2018/5/25
   'SetOfficialFee
   If m_bolChkFee = True Then
      'Modify By Sindy 2019/3/26
      'SetOfficialFee
      'Modify By Sindy 2023/3/16 +, , , txtCP167
      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
                                txtCP135, txtCP136, txtCP137, txtCP84, , txtDecreaseItemFee, _
                                m_lngOverPageFee, m_lngOverItemFee, , , m_lngOfficialFee, _
                                Text7(0).Text, , txtDecreasePageFee, txtCP167, , m_str938RecvNo, m_str939RecvNo)
      '2019/3/26 END
   End If
End Sub

Private Sub txtCP136_GotFocus()
   TextInverse txtCP136
   CloseIme
End Sub

Private Sub txtCP136_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP136_Validate(Cancel As Boolean)
   'Modify By Sindy 2018/5/25
   'SetOfficialFee
   If m_bolChkFee = True Then
      'Modify By Sindy 2019/3/26
      'SetOfficialFee
      'Modify By Sindy 2023/3/16 +, , , txtCP167
      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
                                txtCP135, txtCP136, txtCP137, txtCP84, , txtDecreaseItemFee, _
                                m_lngOverPageFee, m_lngOverItemFee, , , m_lngOfficialFee, _
                                Text7(0).Text, , txtDecreasePageFee, txtCP167, , m_str938RecvNo, m_str939RecvNo)
      '2019/3/26 END
   End If
End Sub

Private Sub txtCP137_GotFocus()
   TextInverse txtCP137
   CloseIme
End Sub

Private Sub txtCP137_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP137_Validate(Cancel As Boolean)
   'Modify By Sindy 2018/5/25
   'SetOfficialFee
   If m_bolChkFee = True Then
      'Modify By Sindy 2019/3/26
      'SetOfficialFee
      'Modify By Sindy 2023/3/16 +, , , txtCP167
      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
                                txtCP135, txtCP136, txtCP137, txtCP84, , txtDecreaseItemFee, _
                                m_lngOverPageFee, m_lngOverItemFee, , , m_lngOfficialFee, _
                                Text7(0).Text, , txtDecreasePageFee, txtCP167, , m_str938RecvNo, m_str939RecvNo)
      '2019/3/26 END
   End If
End Sub

'Add By Sindy 2023/3/16
Private Sub txtCP167_GotFocus()
   TextInverse txtCP167
   CloseIme
End Sub
Private Sub txtCP167_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub
Private Sub txtCP167_Validate(Cancel As Boolean)
   If m_bolChkFee = True Then
      Call PUB_SetOfficialFee_P(cp(), pa(), bolDelay, m_strDelayCP09, m_strReExamCP27, _
                                txtCP135, txtCP136, txtCP137, txtCP84, , txtDecreaseItemFee, _
                                m_lngOverPageFee, m_lngOverItemFee, , , m_lngOfficialFee, _
                                Text7(0).Text, , txtDecreasePageFee, txtCP167, , m_str938RecvNo, m_str939RecvNo)
   End If
End Sub
'2023/3/16 END

Private Sub txtCP138_GotFocus()
   TextInverse txtCP138
   CloseIme
End Sub

Private Sub txtCP138_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub

'Add By Sindy 2023/3/16
Private Sub txtCP168_GotFocus()
   TextInverse txtCP168
   CloseIme
End Sub
Private Sub txtCP168_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub
'2023/3/16 END

'Add by Morgan 2004/8/11
Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   '只能輸倒退及數字鍵
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      Beep
      KeyAscii = 0
   End If
End Sub
'Add by Morgan 2004/8/11
Private Sub txtCP84_Validate(Cancel As Boolean)
   '台灣
   If pa(9) = "000" Then
      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            txtCP84.Tag = txtCP84.Text
         Else
            txtCP84_GotFocus
            Cancel = True
         End If
      End If
   End If
End Sub

'Add by Morgan 2004/3/23
Private Sub txtEP04_GotFocus()
   TextInverse txtEP04
End Sub

Private Sub txtEP04_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add by Morgan 2004/3/23
Private Sub txtEP04_Validate(Cancel As Boolean)
   If txtEP04.Enabled = True Then
      If txtEP04 = "" Then
         lblEP04N.Caption = ""
      Else
         lblEP04N.Caption = GetStaffName(txtEP04.Text)
         If lblEP04N.Caption = "" Then
            MsgBox "核稿人輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            Call txtEP04_GotFocus
         'Add by Morgan 2008/2/27 控制只能為外專工程師   2008/4/8 加 F81
         ElseIf InStr("F21,F52,F81", GetStaffDepartment(txtEP04)) = 0 Then
            MsgBox "核稿人僅能輸外專工程師！"
            Cancel = True
            Call txtEP04_GotFocus
         End If
      End If
   End If
End Sub

'Add by Morgan 2004/3/23
'核稿人欄位控制
Private Function CheckEP04() As Boolean
   Dim stDept As String
   'Modified by Morgan 2013/11/6 +235核對中說格式
   If pa(8) <> "3" And (Text7(2).Text = "201" Or Text7(2).Text = "209" Or Text7(2).Text = "235" Or Text7(2).Text = "210") Then
      If Text7(1) <> Text7(1).Tag Then
         stDept = GetStaffDepartment(Text7(1))
         Select Case stDept
            Case "F22"
               MsgBox "該案件性質的承辦人不可為程序!!!"
               Exit Function
            Case "F51"
               txtEP04.Text = ""
               lblEP04N.Caption = ""
            Case Else
               txtEP04.Text = Text7(1).Text
               lblEP04N.Caption = GetStaffName(txtEP04.Text)
            End Select
      End If
   Else
      txtEP04.Text = ""
      lblEP04N.Caption = ""
   End If
   CheckEP04 = True
End Function
'Add by Morgan 2003/3/30
'檢查母案是否存在
Private Function GetDivInfo(ByRef stDivInfo() As String, ByRef stPA() As String, Optional ByRef p_bolMsg As Boolean = True, Optional p_Continue As Boolean = False) As Boolean

On Error GoTo flgErr

   Dim stSQL As String, rsQuery As New ADODB.Recordset
   
   stSQL = "select pa01,pa02,pa03,pa04,pa09 from patent, divisioncase where pa01=dc05 and pa02=dc06 and pa03=dc07 and pa04=dc08" & _
      " and dc01='" & stPA(1) & "' and dc02='" & stPA(2) & "' and  dc03='" & stPA(3) & "' and dc04='" & stPA(4) & "'"
   
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      stDivInfo(1) = "" & rsQuery("pa01").Value
      stDivInfo(2) = "" & rsQuery("pa02").Value
      stDivInfo(3) = "" & rsQuery("pa03").Value
      stDivInfo(4) = "" & rsQuery("pa04").Value
      stDivInfo(5) = "" & rsQuery("pa09").Value
      GetDivInfo = True
   ElseIf p_bolMsg = True Then
      'Modified by Morgan 2025/3/13 +bolContinue 母案可能為非本所案件 Ex:FCP-073311--Winfrey
      'MsgBox "分割母案本所案號不存在！", vbExclamation
      If MsgBox("分割母案本所案號不存在！是否仍要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
         p_Continue = True
      End If
      'end 2025/3/13
   End If
   
flgErr:
   Set rsQuery = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    
End Function

'Add by Morgan 2005/8/4
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer
   Cancel = True
   cp(110) = "": m_AgentName = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'cp(110) = cp(110) & "," & lstNameAgent.ItemData(ii)
         'Modify By Sindy 2021/5/10
         'cp(110) = cp(110) & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         cp(110) = cp(110) & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         '2021/5/10 END
         Cancel = False
      End If
   Next
   If Cancel = True Then
      'Modify by Morgan 2005/8/29
      'Modified by Lydia 2025/08/29 +242製作提申外文本
      If cp(1) = "FG" Or Text7(2) = "901" Or Text7(2) = "902" Or Text7(2) = "903" Or Text7(2) = "904" Or Text7(2) = "906" Or Text7(2) = "912" Or Text7(2) = "242" Then
         Cancel = False
      Else
         MsgBox "出名代理人不可空白！", vbExclamation
      End If
   Else
      If Left(cp(110), 1) = "," Then cp(110) = Mid(cp(110), 2)
      m_AgentName = Mid(m_AgentName, 2) 'Add By Sindy 2021/5/10
   End If
End Sub
'發E-Mail給承辦人
Private Sub MailToPromoter(strMailCp09 As String)

   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim arrMailCP09
   Dim ii As Integer
   'Dim strOfficeKind As String '所別
   Dim bolMsg As Boolean 'P案提醒

   bolMsg = False
   If strMailCp09 <> "" Then
      'strOfficeKind = PUB_GetST06(strUserNum)
      arrMailCP09 = Split(strMailCp09, ";")
      For ii = LBound(arrMailCP09) To UBound(arrMailCP09)
         If arrMailCP09(ii) <> "" Then
            strSql = "Select * From CaseProgress Where CP09='" & arrMailCP09(ii) & "' And CP14 Is Not Null "
            CheckOC
            With adoRecordset
               .CursorLocation = adUseClient
               .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
               If .RecordCount > 0 Then
                  'Modify by Morgan 2008/8/26 改呼叫共用函數
                  'Load frm880005
                  'Select Case "" & .Fields("CP14").Value
                  '   Case "F5416" '林新晞改發職務代理人張偉城
                  '      frm880005.txtEmail(0).Text = "89026"
                  '   Case "F5479" '洪　磊改發職務代理人蕭圓錚
                  '      frm880005.txtEmail(0).Text = "79075"
                  '   Case "F5484" '李皇　改發職務代理人粘竺儒
                  '      frm880005.txtEmail(0).Text = "84012"
                  '   Case Else
                  '      frm880005.txtEmail(0).Text = "" & .Fields("CP14").Value
                  'End Select
                  ''若使用者不為北所人員, 則E-Mail後面加@taie.com.tw
                  'If strOfficeKind <> "1" Then
                  '   frm880005.txtEmail(0).Text = frm880005.txtEmail(0).Text & "@taie.com.tw"
                  'End If
                  'frm880005.txtEmail(1).Text = .Fields("CP01").Value & "-" & .Fields("CP02").Value & "-" & .Fields("CP03").Value & "-" & .Fields("CP04").Value & "已上文件齊備日及承辦期限!(因為" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "已送件)"
                  'frm880005.txtEmail(2).Text = ""
                  'frm880005.Form_Activate: DoEvents
                  'frm880005.cmdok_Click 0: DoEvents
                  If .Fields("cp01") = "P" Then bolMsg = True
                  PUB_SendMail strUserNum, .Fields("CP14").Value, arrMailCP09(ii), .Fields("CP01").Value & "-" & .Fields("CP02").Value & "-" & .Fields("CP03").Value & "-" & .Fields("CP04").Value & "已上文件齊備日及承辦期限!(因為" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "已送件)", " "
                  'end 2008/8/26
               End If
            End With
         End If
      Next ii
      'Add by Morgan 2008/8/26
      If bolMsg = True Then
         MsgBox "有 P 的相關案，請影印一份中文圖式給專利處相關人員！"
      End If
   End If
End Sub
''Add by Morgan 2010/1/4
''計算實審,修正規費
'Private Sub SetOfficialFee()
'   Dim iItems As Integer, iItemsOld As Integer, iItemsAdd As Integer
'   m_lngOverPageFee = 0
'   m_lngOverItemFee = 0
'
'   'Added by Morgan 2013/1/3
'   If cp(10) = "421" Or cp(10) = "807" Then
'      '發文規費
'      txtCP84 = PUB_GetReportFee(pa(1), pa(9), cp(10), Val(txtCP136))
'   Else
'   'end 2013/1/3
'
'      'Added by Morgan 2013/1/10
'      If cp(10) = "107" Then
'         m_lngOfficialFee = GetPatentOfficialFee(cp(1), cp(10), cp(7), pa(8), pa(9), pa(16))
'         '有延期過則須扣除延期的發文規費
'         If bolDelay = True Then
'            strExc(0) = "select cp84 from caseprogress where cp09='" & m_strDelayCP09 & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               m_lngOfficialFee = m_lngOfficialFee - Val("" & RsTemp("cp84"))
'            End If
'         End If
'
'      Else
'      'end 2013/1/10
'
'         strExc(0) = "select cf08 from casefee where cf01='" & pa(1) & "' and cf02='" & pa(9) & "' and cf03='" & cp(10) & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            '原始規費
'            m_lngOfficialFee = Val("" & RsTemp.Fields(0))
'         End If
'
'      End If 'Added by Morgan 2013/1/9
'
'      '超頁費
'      'Modify By Sindy 2018/5/30 發明才會有超頁費
'      If Val(txtCP135) > 50 And pa(8) = "1" Then
'         m_lngOverPageFee = 500# * ((Val(txtCP135) - 1) \ 50)
'      End If
'
'      '超項費
'      'Modified by Morgan 2013/1/8 +再審107
'      'If cp(10) = "416" Or cp(10) = "201" Or cp(10) = "209" Or cp(10) = "210" Then
'      'Modified by Morgan 2013/10/18 +435
'      'Modified by Morgan 2013/11/6 +235核對中說格式
'      If cp(10) = "416" Or cp(10) = "435" Or cp(10) = "107" Or cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Then
'         iItems = Val(txtCP136)
'         'Modify By Sindy 2018/5/30 發明才會有超項費
'         If iItems > 10 And pa(8) = "1" Then
'            m_lngOverItemFee = 800# * (iItems - 10)
'         End If
'      Else
'         iItemsAdd = Val(txtCP136) - Val(txtCP137)
'
'         strExc(0) = "select sum(cp136),sum(cp137),sum(cp138) from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp27>0 and cp57 is null"
'
'         'Added by Morgan 2013/1/10
'         '若為再審後修正時只能抓再審以後發文的程序加總
'         If m_strReExamCP27 <> "" Then
'            strExc(0) = strExc(0) & " and cp27>=" & m_strReExamCP27
'         End If
'         'end 2013/1/10
'
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            iItemsOld = Val("" & RsTemp.Fields(0)) - Val("" & RsTemp.Fields(1))
'         End If
'         iItems = iItemsOld + iItemsAdd
'         '項數增加
'         If iItemsAdd > 0 Then
'            '超過10項
'            'Modify By Sindy 2018/5/30 發明才會有超項費
'            If iItemsOld >= 10 And pa(8) = "1" Then
'               m_lngOverItemFee = 800# * iItemsAdd
'            ElseIf iItems >= 10 And pa(8) = "1" Then
'               m_lngOverItemFee = 800# * (iItems - 10)
'            End If
'         '項數減少
'         ElseIf iItemsAdd < 0 Then
'            '超過10項
'            'Modify By Sindy 2018/5/30 發明才會有超項費
'            If iItems >= 10 And pa(8) = "1" Then
'               m_lngOverItemFee = 800# * iItemsAdd
'            '刪減後少於10項,但原來總項數>10,則可退原繳超項費=800*(原來總項數-10)
'            ElseIf iItemsOld > 10 And pa(8) = "1" Then
'               m_lngOverItemFee = -1 * 800# * (iItemsOld - 10)
'            End If
'         End If
'      End If
'      '發文規費
'      txtCP84 = m_lngOfficialFee + m_lngOverPageFee + m_lngOverItemFee
'   End If 'Added by Morgan 2013/1/3
'End Sub
'Add by Morgan 2010/1/5
'Modify by Morgan 2011/6/29 改和內專相同,要先扣除已收未發的超頁超項費
'Modify By Sindy 2023/3/27 Mark,改呼叫共用函數
''檢查規費
'Private Function CheckOfficialFee() As Boolean
'   Dim strMsg As String, bolBilled As Boolean
'   CheckOfficialFee = True
'   m_lngRecOverPageFee = 0
'   m_lngRecOverItemFee = 0
'
'   If m_bolChkPageItem = True Then
'      If txtCP135 = "" And txtCP135.Enabled = True Then 'Modify By Sindy 2018/5/22 + And txtCP135.Enabled = True
'         MsgBox "頁數不可空白！", vbExclamation
'         txtCP135.SetFocus
'         CheckOfficialFee = False
'         Exit Function
'      ElseIf txtCP136 = "" And txtCP136.Enabled = True Then 'Modify By Sindy 2018/5/22 + And txtCP136.Enabled = True
'         MsgBox "項數不可空白！", vbExclamation
'         txtCP136.SetFocus
'         CheckOfficialFee = False
'         Exit Function
'      End If
'   End If
'
'   'Add by Morgan 2010/9/27
'   If m_bolChkItem Then
'      'Modify By Sindy 2023/3/16 + or (txtCP135 = "" And txtCP167 = "" And txtCP168 = "")
'      If (txtCP136 = "" And txtCP137 = "" And txtCP138 = "") And _
'         (txtCP135 = "" And txtCP167 = "" And txtCP168 = "") Then
'         If MsgBox("(增加項數及刪除項數) 及 (增加頁數及刪除頁數)皆為空白，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
'            'Modify By Sindy 2023/3/16 + If txtCP135.Enabled Then txtCP135.SetFocus
'            If (txtCP136 = "" And txtCP137 = "" And txtCP138 = "") Then
'               If txtCP136.Enabled Then txtCP136.SetFocus
'            Else
'               If txtCP135.Enabled Then txtCP135.SetFocus
'            End If
'            CheckOfficialFee = False
'            Exit Function
'         End If
'      End If
'   End If
'   'end 2010/9/27
'
'   '退規費
'   If Val(txtCP84) < 0 Then
'      m_FeeMemo = "退規費 " & Format(-1 * Val(txtCP84), DAmount) & " 元;"
'      MsgBox "本次發文可退規費【" & Format(-1 * Val(txtCP84), DAmount) & "】元！", vbInformation
'      txtCP84 = 0
'
'   ElseIf m_lngOverPageFee + m_lngOverItemFee > 0 Then
'
'      m_lngOverPageFeeDiff = m_lngOverPageFee
'      m_lngOverItemFeeDiff = m_lngOverItemFee
'
'      'Add by Morgan 2011/6/29
'      strExc(0) = "select cp10,sum(nvl(cp17,0)-nvl(cp77,0)) Fee,max(cp60) BNo from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='938' and cp27||cp57 is null group by cp10" & _
'         " union select cp10,sum(nvl(cp17,0)-nvl(cp77,0)) Fee,max(cp60) BNo from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='939' and cp27||cp57 is null group by cp10"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         Do While Not RsTemp.EOF
'            If RsTemp.Fields("cp10") = "938" Then m_lngRecOverPageFee = m_lngRecOverPageFee + Val("" & RsTemp.Fields("Fee"))
'            If RsTemp.Fields("cp10") = "939" Then m_lngRecOverItemFee = m_lngRecOverItemFee + Val("" & RsTemp.Fields("Fee"))
'            If Not IsNull(RsTemp.Fields("BNo")) Then bolBilled = True
'            RsTemp.MoveNext
'         Loop
'         m_lngOverPageFeeDiff = m_lngOverPageFee - m_lngRecOverPageFee
'         m_lngOverItemFeeDiff = m_lngOverItemFee - m_lngRecOverItemFee
'      End If
'
'      If bolBilled And (m_lngOverPageFeeDiff <> 0 Or m_lngOverItemFeeDiff <> 0) Then
'         MsgBox "本案已收文超頁費或超項費且已請款但金額不符，請修正後再發文！"
'         CheckOfficialFee = False
'         Exit Function
'      End If
'      'end 2011/6/29
'
'      strMsg = ""
'      If m_lngOverPageFee > 0 Then
'         strMsg = "超頁費【" & Format(m_lngOverPageFee, DDollar) & "】元"
'      End If
'      If m_lngOverItemFee > 0 Then
'         strMsg = strMsg & IIf(strMsg <> "", "及", "") & "超項費【" & Format(m_lngOverItemFee, DDollar) & "】元"
'      End If
'      'Modify by Morgan 2011/6/29
'      'If MsgBox("本案須繳" & strMsg & "共【" & Format(m_lngOverPageFee + m_lngOverItemFee, DDollar) & "】元，存檔時會自動做內部收文並同時上發文日，是否確定要繼續？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
'      If MsgBox("本案須繳" & strMsg & "共【" & Format(m_lngOverPageFee + m_lngOverItemFee, DDollar) & "】元，存檔時會自動做內部收文並同時上發文日(若已收文將更新收文金額並同時上發文日)，是否確定要繼續？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
'      'end 2011/6/29
'         CheckOfficialFee = False
'      End If
'
'   End If
'
'End Function

Private Sub StartLetter(ByVal ET01 As String, ET02 As String, ByVal ET03 As String)
   Dim strTxt() As String, ii As Integer

   EndLetter ET01, ET02, ET03, strUserNum

   ii = 1
   ReDim Preserve strTxt(ii)
   If bolCreditNote = True Then
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','抵帳要印','♀')"
   Else
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','抵帳不印','♀')"
   End If
   'ii = ii + 1
   'ReDim Preserve strTxt(ii)
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Sub
'Add by Morgan 2010/11/12
Private Sub StartLetter2(ByVal ET01 As String, ET02 As String, ByVal ET03 As String)
   Dim strTxt() As String, i As Integer, j As Integer
   Dim strDoc As String, strList As String

   EndLetter ET01, ET02, ET03, strUserNum

   i = 1
   ReDim Preserve strTxt(i)
   strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','補文件發文日期','" & DBDATE(Text7(0)) & "')"
   
   strExc(0) = "select cp64 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
      " and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='202'" & _
      " and cp27=" & DBDATE(Text7(0))
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strDoc = ""
      strList = ""
      With RsTemp
      Do While Not .EOF
         j = i
         strExc(1) = "委任書"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
         strExc(1) = "申請權證明"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
            
         strExc(1) = "優先權證明"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
         
         strExc(1) = "切結書"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
         
         strExc(1) = "僱傭契約"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
         
         strExc(1) = "美國讓與"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               If InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X28186") > 0 Then
                  strExc(1) = "美國讓與免公證"
               End If
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
         
         strExc(1) = "國內寄存證明"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
         
         strExc(1) = "國外寄存證明"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
         
         strExc(1) = "死亡證明"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
         
         strExc(1) = "繼承證明"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
         
         strExc(1) = "法人地位證明"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
         
         strExc(1) = "國籍證明"
         If InStr(strList, strExc(1)) = 0 Then
            If InStr("" & .Fields(0), strExc(1)) > 0 Then
               i = i + 1
               ReDim Preserve strTxt(i)
               strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                  "','" & strExc(1) & "','♀')"
               strList = strList & strExc(1) & ";"
            End If
         End If
         
         'Remove by Morgan 2011/2/24 只帶出上列文件就好，如還有其他文件再另行通知新增 --David
         'If j = i Then
         '   If InStr("" & .Fields(0), "專利申請") = 0 Then
         '      strDoc = strDoc & vbCrLf & .Fields(0)
         '   End If
         'End If
         
         .MoveNext
      Loop
      End With
      If strDoc <> "" Then
         i = i + 1
         ReDim Preserve strTxt(i)
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','其他補文件','" & ChgSQL(strDoc) & "')"
      End If
   End If
   
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If

End Sub

Private Sub SetCombo3()
   Combo3.Clear
   Combo3.AddItem "3 個月", 0
   Combo3.ItemData(0) = 3
   Combo3.AddItem "2 個月", 0
   Combo3.ItemData(0) = 2
   Combo3.AddItem "1 個月", 0
   Combo3.ItemData(0) = 1
   Combo3.ListIndex = -1
End Sub

'Add by Amy 2013/08/22 代理人Y49456發明申請發文時提示是否實審要掛交承辦收文告代
Private Function Check416Exist() As Boolean
  On Error GoTo ErrHnd
   
   Check416Exist = True
   
   CheckOC3
   With AdoRecordSet3
      strSql = "Select * From CaseProgress Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10='" & 實體審查 & "'  And CP57 IS NULL "
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      '若案件進度檔無實體審查的資料(即未收文)
      If .RecordCount <= 0 Then
         Check416Exist = False
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical

End Function
'
''Add by Amy 2014/05/27 +簡易聯絡單
''Modify by Amy 2014/07/04 改共用讓frm060102使用
''Modify byAmy 2016/04/29 +cp10參數
''Modified by Lydia 2020/02/10 +外部呼叫bolCall
''Modify By Sindy 2022/5/11 回傳:簡易聯絡單內容
'Mark by Lydia 2023/05/019 改成basPublic.PUB_FCPPrintContactSheetA4
'Public Function PrintContactSheetA4(ByVal CP09 As String, ByVal cp01 As String, ByVal cp02 As String, ByVal cp03 As String, ByVal cp04 As String, _
'                      ByVal CP10 As String, Optional ByVal bolCall As Boolean = False) As String
'    Dim strTemp As String
'    Dim strTmpA As String '折行-剩的字串
'    Dim bolFirst As Boolean
'    Dim m_Line3 As String 'Added by Lydia 2018/01/05
'
'    PrintContactSheetA4 = "" 'Add By Sindy 2022/5/11
'    strExc(1) = 0
'    'Added by Lydia 2020/02/10 外部呼叫: 預設變數
'    If bolCall = True Then
'       mDate209210 = ""
'       mDateTF30 = ""
'    End If
'    'end 2020/02/10
'
'    intFieldWidth = Array(2000, 2500)
'    m_dblTop = 300: m_dblLeft = 600:   dblLineHeight = 200
'    iPage = 1
'
'    Printer.PaperSize = PUB_GetPaperSize(9) '設定紙張 A4
'    Printer.Orientation = 1 '直印
'    Printer.Font.Name = "標楷體"
'    m_TBWidth = Printer.ScaleWidth - 700
'
'    Call PrintStaticData(cp01, cp02, cp03, cp04)
'    PrintTableLine '畫表格
'    Call PrintField3Title(cp01, cp02) '第三欄抬頭-本所案號
'
'    intLine = intLine + 1
'    strExc(1) = Val(strExc(1)) + 1
'    strTemp = strExc(1) & "、告申請日、案號"
'    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'    Printer.CurrentX = dblPrtX
'    Printer.CurrentY = dblPrtY
'    Printer.Print strTemp
'    PrintContactSheetA4 = PrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
'
'    intLine = intLine + 1
'    'Modified by Morgan 2016/9/30 +申請書--何淑華
'    'Modified by Lydia 2019/12/09 修正頁->修正本 (by Phoebe)
'    strTemp = " ＊同時寄：收據、修正本、申請書"
'    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'    Printer.CurrentX = dblPrtX
'    Printer.CurrentY = dblPrtY
'    Printer.Print strTemp
'    PrintContactSheetA4 = PrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
'
'    '*** 第三欄第2點 ***
'    '取得202補文件及231寄存證明的本所期限,有期限印NP15並計算NP09筆數(列印份數),無期限印「文件已齊備」
'    'Modified by Morgan 2017/7/10 補文件的內容一份前的字數不一定是7，改剔除"一份"兩字就好--敏莉
'    'strExc(0) = "Select NP08,SubStr(NP15,1,7) as NP15,count(NP15) as CNP15 From NextProgress " & _
'                      "Where NP01='" & CP09 & "' And NP07 in ('202','231') And NP06 is null " & _
'                      "And InStr(NP15,'專利申請書')=0 Group by NP08,NP15 " & _
'            "Union Select NP08,NP15,0 as CNP15 From NextProgress " & _
'                      "Where NP01='" & CP09 & "' And NP07 in ('202','231') And NP06 is null " & _
'                      "And InStr(NP15,'專利申請書')>0 Group by NP08,NP15 "
'    'Modify By Sindy 2021/7/8 本所期限改抓約定期限 + ,np23
'    'Modify By Sindy 2021/7/23 排除 客戶提供中說、英文參考本 在下方獨立判斷, 因非真正智慧局的期限
'    strExc(0) = "Select NP08,np23,replace(NP15,'一份','') as NP15,count(NP15) as CNP15 From NextProgress " & _
'                      "Where NP01='" & CP09 & "' And NP07 in ('202','231') And NP06 is null " & _
'                      "And InStr(NP15,'專利申請書')=0 And InStr(NP15,'客戶提供中說')=0 And InStr(NP15,'英文參考本')=0 " & _
'                      "Group by NP08,np23,NP15 " & _
'            "Union Select NP08,np23,NP15,0 as CNP15 From NextProgress " & _
'                      "Where NP01='" & CP09 & "' And NP07 in ('202','231') And NP06 is null " & _
'                      "And InStr(NP15,'專利申請書')>0 And InStr(NP15,'客戶提供中說')=0 And InStr(NP15,'英文參考本')=0 " & _
'                      "Group by NP08,np23,NP15 "
'    'end 2017/7/10
'    intI = 1
'    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'    If intI = 1 Then
'        intLine = intLine + 1
'        strExc(1) = Val(strExc(1)) + 1
'        strTemp = strExc(1) & "、本案尚缺之文件："
'        dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'        Printer.CurrentX = dblPrtX
'        Printer.CurrentY = dblPrtY
'        Printer.Print strTemp
'        PrintContactSheetA4 = PrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
'
'        Dim i As Integer
'        With RsTemp
'            .MoveFirst
'            Do While Not .EOF
'                If intLine > 24 Then
'                    iPage = iPage + 1
'                    Printer.NewPage
'                    Call PrintStaticData(cp01, cp02, cp03, cp04)
'                    PrintTableLine '畫表格
'                    Call PrintField3Title(cp01, cp02)
'                End If
'
'                '專利申請書不需計算份數
'                If InStr(.Fields("NP15"), "專利申請書") > 0 Then
'                  'Modify By Sindy 2021/7/8 本所期限改抓約定期限
'                  If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
'                     strTemp = "(期限" & ChangeWStringToTDateString("" & .Fields("NP23")) & ")"
'                  Else
'                  '2021/7/8 END
'                     strTemp = "(期限" & ChangeWStringToTDateString(.Fields("NP08")) & ")"
'                  End If
'                Else
'                  'Modify By Sindy 2021/7/8 本所期限改抓約定期限
'                  If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
'                     strTemp = "(" & .Fields("CNP15") & "份)(期限" & ChangeWStringToTDateString("" & .Fields("NP23")) & ")"
'                  Else
'                  '2021/7/8 END
'                     strTemp = "(" & .Fields("CNP15") & "份)(期限" & ChangeWStringToTDateString(.Fields("NP08")) & ")"
'                  End If
'                End If
'                strTemp = " ＊" & .Fields("NP15") & strTemp
'                PrintContactSheetA4 = PrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
'                '超過20個字-換行
'                If strTemp <> StrToStr(strTemp, 19) Then
'                    bolFirst = True
'                    Do While strTemp <> ""
'                        If intLine > 24 Then
'                            iPage = iPage + 1
'                            Printer.NewPage
'                            Call PrintStaticData(cp01, cp02, cp03, cp04)
'                            PrintTableLine '畫表格
'                            Call PrintField3Title(cp01, cp02)
'                        End If
'
'                        If bolFirst = True Then
'                            strTmpA = StrToStr(strTemp, 19) '目前要取的字串
'                            bolFirst = False
'                            strTemp = Mid(strTemp, Len(strTmpA) + 1) '取完剩的字串
'                        Else
'                            strTmpA = Space(3) & StrToStr(strTemp, 19)
'                            strTemp = Mid(strTemp, Len(strTmpA) - 3 + 1) '取完剩的字串
'                        End If
'                        'strTemp = Mid(strTemp, Len(strTmpA) + 1) '取完剩的字串
'
'                        intLine = intLine + 1
'                        dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'                        Printer.CurrentX = dblPrtX
'                        Printer.CurrentY = dblPrtY
'                        Printer.Print strTmpA
'                    Loop
'                Else
'                    intLine = intLine + 1
'                    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'                    Printer.CurrentX = dblPrtX
'                    Printer.CurrentY = dblPrtY
'                    Printer.Print strTemp
'                End If
'                .MoveNext
'            Loop
'        End With
'    Else
'        intLine = intLine + 1
'        strExc(1) = Val(strExc(1)) + 1
'        strTemp = strExc(1) & "、文件已齊備"
'        dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'        Printer.CurrentX = dblPrtX
'        Printer.CurrentY = dblPrtY
'        Printer.Print strTemp
'        PrintContactSheetA4 = PrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
'    End If
'    '*** End 第三欄第2點 ***
'
'    'Added by Lydia 2018/01/05 新案發文時，有收文新案翻譯且原文字數為空白，簡易連聯絡單多加註記
'    If InStr(NewCasePtyList, CP10) > 0 Then
'        strExc(0) = "select cp09,cp10,tf01,tf23,tf19,tf20 from caseprogress,transfee " & _
'                          "where cp01='" & cp01 & "' and cp02='" & cp02 & "' and cp03='" & cp03 & "' and cp04='" & cp04 & "' " & _
'                          "and cp10='201' and cp09=tf01(+)"
'        intI = 1
'        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'        If intI = 1 Then
'             If Val("" & RsTemp.Fields("tf23")) = 0 Then
'                strExc(1) = Val(strExc(1)) + 1
'                m_Line3 = strExc(1) & "、原文字數為空白，請承辦填寫原文字數"
'                intLine = intLine + 1
'                dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'                Printer.CurrentX = dblPrtX
'                Printer.CurrentY = dblPrtY
'                Printer.Print m_Line3
'                PrintContactSheetA4 = PrintContactSheetA4 & m_Line3 & vbCrLf 'Add By Sindy 2022/5/11
'
'                m_Line3 = " 　，退程序輸入。"
'                intLine = intLine + 1
'                dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'                Printer.CurrentX = dblPrtX
'                Printer.CurrentY = dblPrtY
'                Printer.Print m_Line3
'                PrintContactSheetA4 = PrintContactSheetA4 & m_Line3 & vbCrLf 'Add By Sindy 2022/5/11
'             End If
'        End If
'    End If
''    If m_Line3 <> "" Then
''         strExc(1) = "4"
''    Else
''         strExc(1) = "3"
''    End If
'    'end 2018/01/05
'
'    strTemp = ""
'    'Modify by Amy 2016/04/29
'    If CP10 = "103" Or CP10 = "125" Then
'        'Modified by Lydia 2018/01/05
'        'strTemp = "3、退程序主管分案撰中文圖說"
'        strExc(1) = Val(strExc(1)) + 1
'        'Modify By Sindy 2022/12/29
'        'strTemp = strExc(1) & "、退程序主管分案撰中文圖說"
'        strTemp = strExc(1) & "、通知分案撰中文圖說"
'        If strTemp <> "" Then
'           intLine = intLine + 1
'           dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'           Printer.CurrentX = dblPrtX
'           Printer.CurrentY = dblPrtY
'           Printer.Print strTemp
'           PrintContactSheetA4 = PrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
'         End If
'        '2022/12/29 END
'    'Modify By Sindy 2022/12/29 設計案也要抓有告代
'    End If
'    'Else
'    strTemp = ""
'    '2022/12/29 END
'        'Modified by Lydia 2018/01/05
'        'strTemp = "3、退檔 or 退程序主管分案告代函"
'        'Modified by Lydia 2018/04/16
'        'strTemp = strExc(1) & "、退檔 or 退程序主管分案告代函"
'
'        'strTemp = strExc(1) & "、退檔 or 退程序(有告代函)"
'        'Modify By Sindy 2022/5/11
'        '請改成:若有收文告代(901)或主動修正(203)未發文
'        '且進度檔是帶提申後告代(ex:066746)or提申後主動修正，
'        '二者皆有則帶: 有告代函及主動修正--(請抓承辦人(工程師))，
'        '若只有一種則帶其一即可，例: 有告代函--(請抓承辦人)，若皆無則此欄可不帶。
'         strExc(0) = "SELECT * FROM caseprogress,staff" & _
'                     " WHERE cp01='" & cp01 & "' and cp02='" & cp02 & "' and cp03='" & cp03 & "' and cp04='" & cp04 & "'" & _
'                     " and ((cp10='901' and instr(cp64,'提申後告代')>0) or (cp10='203' and instr(cp64,'提申後主動修正')>0))" & _
'                     " and cp27||cp57 is null" & _
'                     " and cp14=st01(+)" & _
'                     " order by cp10 desc"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            RsTemp.MoveFirst
'            strExc(10) = 0
'            strExc(9) = "" & RsTemp.Fields("st02")
'            strExc(1) = Val(strExc(1)) + 1
'            strTemp = strExc(1) & "、有"
'            Do While Not RsTemp.EOF
'               strExc(10) = Val(strExc(10)) + 1
'               If Val(strExc(10)) > 1 Then
'                  strTemp = strTemp & "及"
'               End If
'               If RsTemp.Fields("cp10") = "901" Then
'                  strTemp = strTemp & "告代函"
'               ElseIf RsTemp.Fields("cp10") = "203" Then
'                  strTemp = strTemp & "主動修正"
'               End If
'               RsTemp.MoveNext
'            Loop
'            strTemp = strTemp & "--" & strExc(9)
'         End If
'         '2022/5/11 END
''    End If
'    'end 2016/04/29
'    If strTemp <> "" Then
'      intLine = intLine + 1
'      dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'      Printer.CurrentX = dblPrtX
'      Printer.CurrentY = dblPrtY
'      Printer.Print strTemp
'      PrintContactSheetA4 = PrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
'    End If
'
'    'Added by Lydia 2019/01/04 新案發文有扣款日期，則簡易連聯絡單多加註記
'    If InStr(NewCasePtyList, CP10) > 0 Then
'        'Modified by Lydia 2023/05/02 +cp84
'        strExc(0) = "select cp09,cp152,cp84 from caseprogress where cp09='" & CP09 & "' "
'        intI = 1
'        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'        If intI = 1 Then
'            If Val("" & RsTemp.Fields("cp152")) > 0 Then
'                intLine = intLine + 1
'                strExc(1) = Val(strExc(1)) + 1
'                'Modified by Lydia 2023/05/02 +規費金額
'                strTemp = strExc(1) & "、收據下載日期：" & ChangeTStringToTDateString(TransDate("" & RsTemp.Fields("cp152"), 1)) & IIf(Val("" & RsTemp.Fields("cp84")) > 0, "，規費金額：NTD " & RsTemp.Fields("cp84"), "")
'                dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'                Printer.CurrentX = dblPrtX
'                Printer.CurrentY = dblPrtY
'                Printer.Print strTemp
'                'Added by Lydia 2023/05/03 因為先發文新案，所以同時提醒該案亦有同日發文的規費金額 ---- frm060104_k使用
'                If PUB_ChkCPExist(cp, "416", 1) = True Then
'                    strTemp = strTemp & "+實體審查:"
'                End If
'                'end 2023/05/03
'                PrintContactSheetA4 = PrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
'            End If
'        End If
'        'Added by Lydia 2020/02/10 新案建檔有重新印「發文簡易聯絡單」之功能,所以要能抓取相關行事曆
'        If mDate209210 = "" And InStr("101,102", CP10) > 0 Then '客戶提供中說期限
'            'Add By Sindy 2021/7/23 不產生行事曆了,改新增下一程序
'            strExc(0) = "Select NP08,np23,NP15,0 as CNP15 From NextProgress " & _
'                      "Where NP01='" & CP09 & "' And NP07 in ('202','231') And NP06 is null " & _
'                      "And InStr(NP15,'客戶提供中說')>0 " & _
'                      "Group by NP08,np23,NP15 "
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               mDate209210 = "" & RsTemp.Fields("np23")
'            Else
'            '2021/7/23 END
'               strExc(0) = "select sc01 from staff_calendar where instr(sc04,'催客戶提供中說期限') > 0 and sc05='" & cp01 & "' and sc06='" & cp02 & "' and sc07='" & cp03 & "' and sc08='" & cp04 & "' and sc18 is null "
'               strExc(0) = strExc(0) & "order by sc01"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  mDate209210 = "" & RsTemp.Fields("sc01")
'               End If
'            End If
'        End If
'        If mDateTF30 = "" Then
'            'Add By Sindy 2021/7/23 不產生行事曆了,改新增下一程序
'            strExc(0) = "Select NP08,np23,NP15,0 as CNP15 From NextProgress " & _
'                      "Where NP01='" & CP09 & "' And NP07 in ('202','231') And NP06 is null " & _
'                      "And InStr(NP15,'英文參考本')>0 " & _
'                      "Group by NP08,np23,NP15 "
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               mDateTF30 = "" & RsTemp.Fields("np23")
'            Else
'            '2021/7/23 END
'               strExc(0) = "select sc01 from staff_calendar where instr(sc04,'催客戶提供英文翻譯本') > 0 and sc05='" & cp01 & "' and sc06='" & cp02 & "' and sc07='" & cp03 & "' and sc08='" & cp04 & "' and sc18 is null "
'               strExc(0) = strExc(0) & "order by sc01"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  mDateTF30 = "" & RsTemp.Fields("sc01")
'               End If
'            End If
'        End If
'    End If
'    'end 2019/01/04
'
'    'Added by Lydia 2019/01/17 FCP新案發文時若檢視中說or核對中說格式未發文，自動設行事曆；一併在簡易聯絡單增加第5點：客戶提供中說期限。
'    If mDate209210 <> "" Then
'        intLine = intLine + 1
'        strExc(1) = Val(strExc(1)) + 1
'        strTemp = strExc(1) & "、客戶提供中說期限：" & ChangeTStringToTDateString(TransDate(mDate209210, 1))
'        dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'        Printer.CurrentX = dblPrtX
'        Printer.CurrentY = dblPrtY
'        Printer.Print strTemp
'        PrintContactSheetA4 = PrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
'    End If
'    'end 2019/01/17
'
'    'Added by Lydia 2019/12/11 FCP新案發文時檢查有新案翻譯未發文並且尚"待英文本翻譯"，自動設行事曆；一併在簡易聯絡單增加第5點：客戶提供英文翻譯本。
'    If mDateTF30 <> "" Then
'        intLine = intLine + 1
'        strExc(1) = Val(strExc(1)) + 1
'        strTemp = strExc(1) & "、客戶提供英文翻譯本：" & ChangeTStringToTDateString(TransDate(mDateTF30, 1))
'        dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
'        Printer.CurrentX = dblPrtX
'        Printer.CurrentY = dblPrtY
'        Printer.Print strTemp
'        PrintContactSheetA4 = PrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
'    End If
'    'end 2019/12/11
'
'    Printer.EndDoc
'End Function

Private Sub PrintStaticData(cp01 As String, cp02 As String, cp03 As String, cp04 As String)
    intLine = 1
    
    'Removed by Morgan 2020/3/30
    'strExc(0) = "台一國際專利商標事務所"
    'm_dblTitleHeight = (intLine + 0.8) * 300
    'Printer.Font.Size = 22
    'dblPrtX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strExc(0)) / 2)
    'dblPrtY = m_dblTop + m_dblTitleHeight
    'Printer.CurrentX = dblPrtX
    'Printer.CurrentY = dblPrtY
    'Printer.Print strExc(0)
    'intLine = intLine + 1
    'end 2020/3/30
    
    strExc(0) = "簡易聯絡單"
    Printer.Font.Size = 20
    m_dblTitleHeight = (intLine + 1.5) * 300
    dblPrtX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strExc(0)) / 2)
    dblPrtY = m_dblTop + m_dblTitleHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)
    m_dblTitleHeight = m_dblTitleHeight + 400
    
    intLine = 1
    Printer.Font.Size = 18
    strExc(0) = "受 文 者"
    dblPrtX = m_dblLeft + intFieldWidth(0) / 2 - Printer.TextWidth(strExc(0)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 350
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)
    
    strExc(0) = "發 文 者"
    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 350
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)
 
    intLine = intLine + 1
    Printer.Font.Size = 16
    strExc(0) = GetStaffName(PUB_GetFCPSalesNo(cp01, cp02, cp03, cp04))
    dblPrtX = m_dblLeft + intFieldWidth(0) / 2 - Printer.TextWidth(strExc(0)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + dblLineHeight + intLine * 500
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)

    strExc(0) = GetStaffName(PUB_GetFCPHandler(cp01, cp02, cp03, cp04))
    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + dblLineHeight + intLine * 500
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)

    intLine = intLine + 2
    Printer.Font.Size = 18
    strExc(0) = "發文時間"
    dblPrtX = m_dblLeft + intFieldWidth(0) / 2 - Printer.TextWidth(strExc(0)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight * 1.2 + dblLineHeight + intLine * 500  'm_dblTitleHeight * 1.2 for 垂直靠下
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)
    
    Printer.Font.Size = 16
    strExc(0) = Year(Now) - 1911 & "年" & "  月" & "  日"
    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight * 1.2 + dblLineHeight + intLine * 500
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)
    
    intLine = intLine + 2
    Printer.Font.Size = 18
    strExc(0) = "發文地點"
    dblPrtX = m_dblLeft + intFieldWidth(0) / 2 - Printer.TextWidth(strExc(0)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + dblLineHeight + intLine * 500
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)
    
    Printer.Font.Size = 16
    strExc(0) = "國外部專利處"
    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + dblLineHeight + intLine * 500
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)
 
    'Added by Lydia 2017/10/20 新增承辦、判行、確認已報告+Email已回存之簽核欄位
    intLine = intLine + 15
    'Remove by Lydia 2019/03/21 取消承辦、判行(by A4011)
'    Printer.Font.Size = 18
'    strExc(0) = "承　辦"
'    dblPrtX = m_dblLeft + intFieldWidth(0) / 2 - Printer.TextWidth(strExc(0)) / 2
'    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 350
'    Printer.CurrentX = dblPrtX
'    Printer.CurrentY = dblPrtY
'    Printer.Print strExc(0)
'
'    strExc(0) = "判　行"
'    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) / 2 - Printer.TextWidth(strExc(0)) / 2
'    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 350
'    Printer.CurrentX = dblPrtX
'    Printer.CurrentY = dblPrtY
'    Printer.Print strExc(0)
    'end 2019/03/21
    '確認已報告+Email已回存
    intLine = intLine + 11
    Printer.Font.Size = 16
    strExc(0) = "確認已報告 + E-mail已回存"
    dblPrtX = m_dblLeft + (intFieldWidth(0) + intFieldWidth(1)) / 2 - Printer.TextWidth(strExc(0)) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 350
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)
    'end 2017/10/20
End Sub

Private Sub PrintTableLine()
    Dim i As Integer
    Dim startY As Integer 'Added by Lydia 2017/10/20
    '雙橫線-上
    dblPrtY = m_dblTop + m_dblTitleHeight + 50
    startY = dblPrtY 'Added by Lydia 2017/10/20
    Printer.Line (m_dblLeft, dblPrtY)-(m_TBWidth, dblPrtY)
    dblPrtY = m_dblTop + m_dblTitleHeight + 100
    Printer.Line (m_dblLeft + 50, dblPrtY)-(m_TBWidth - 50, dblPrtY)
    
    '雙橫線-下
    Printer.Line (m_dblLeft + 50, 14950)-(m_TBWidth - 50, 14950)
    Printer.Line (m_dblLeft, 15000)-(m_TBWidth, 15000)

    '雙直線-左
    Printer.Line (m_dblLeft, dblPrtY - 50)-(m_dblLeft, 15000)
    Printer.Line (m_dblLeft + 50, dblPrtY)-(m_dblLeft + 50, 14950)
    
    '雙直線-右
    Printer.Line (m_TBWidth - 50, dblPrtY)-(m_TBWidth - 50, 14950)
    Printer.Line (m_TBWidth, dblPrtY - 50)-(m_TBWidth, 15000)
      
    '欄位分隔線-橫線
    'Memo by Lydia 2017/10/20 ex.FCP-57679
    '受文者     |發文者
    '-------------------- i=1
    '羅XX       |蔡XX
    '-------------------- i=2
    '發文時間   |106年 月 日
    '-------------------- i=3
    '發文地點   |國外部專利處
    '(隔4行)    |
    '-----------------------
    '承辦       |判行
    '(隔2行)  |
    '-----------------------
    '確認已報告+Email已回存
    'end 2017/10/20
    For i = 1 To 3
        dblPrtY = m_dblTop + m_dblTitleHeight + i * 1000
        Printer.Line (m_dblLeft + 50, dblPrtY)-(m_dblLeft + intFieldWidth(0) + intFieldWidth(1), dblPrtY)
    Next
    
    'Added by Lydia 2017/10/20 新增承辦、判行、確認已報告+Email已回存之簽核欄位
    '承辦、判行(橫線)
    'Remove by Lydia 2019/03/21 取消承辦、判行(by A4011)
    'dblPrtY = m_dblTop + m_dblTitleHeight + 7 * 1000
    'Printer.Line (m_dblLeft + 50, dblPrtY)-(m_dblLeft + intFieldWidth(0) + intFieldWidth(1), dblPrtY)
    'dblPrtY = m_dblTop + m_dblTitleHeight + 8 * 1000
    'Printer.Line (m_dblLeft + 50, dblPrtY)-(m_dblLeft + intFieldWidth(0) + intFieldWidth(1), dblPrtY)
    'end 2019/03/21
    '確認已報告+Email已回存(橫線),與Anny(A4011)確認過標題下橫線不用畫
    dblPrtY = m_dblTop + m_dblTitleHeight + 11 * 1000
    Printer.Line (m_dblLeft + 50, dblPrtY)-(m_dblLeft + intFieldWidth(0) + intFieldWidth(1), dblPrtY)
    'end 2017/10/20
    
    'Move by Lydia 2017/10/20 從上方移下來
    '欄位分隔線-直線
    dblPrtX = m_dblLeft
    For i = 0 To UBound(intFieldWidth)
        dblPrtX = dblPrtX + intFieldWidth(i)
        'Modified by Lydia 2017/10/20 第1條分隔線只畫到"確認已報告+Email已回存"
        'Printer.Line (dblPrtX, dblPrtY)-(dblPrtX, 14950)
        If i = 0 Then
           Printer.Line (dblPrtX, startY)-(dblPrtX, dblPrtY)
        Else
           Printer.Line (dblPrtX, startY)-(dblPrtX, 14950)
        End If
        'end 2017/10/20
    Next
    'end 2017/10/20
    
    '頁碼
    Printer.Font.Size = 12
    strExc(0) = iPage
    dblPrtX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strExc(0)) / 2)
    dblPrtY = Printer.ScaleHeight - m_dblTop * 2.5
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)
End Sub

Private Sub PrintField3Title(cp01 As String, cp02 As String)
    '第三欄
    intLine = 1
    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) + 200
    
    Printer.Font.Size = 16
    strExc(0) = cp01 & "-" & cp02
    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500 - 100
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strExc(0)
    
    Printer.FontSize = 14
End Sub

'Added by Lydia 2015/12/31 案件性質為會稿924時，檢查一定要有相關總收文號;增加按鈕供使用者選擇;
Private Sub Command2_Click()
    Set frm060101_2.fmParent = Me
    frm060101_2.Show
    Me.Hide
End Sub

Private Sub txtCP43_GotFocus()
   TextInverse txtCP43
End Sub

Private Sub txtCP43_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Morgan 2017/9/25
Private Sub txtPayToday_GotFocus()
   TextInverse txtPayToday
   CloseIme
End Sub
'Added by Morgan 2017/9/25
Private Sub txtPayToday_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Morgan 2017/9/25
Private Sub SetPayToday()
   'Modified by Lydia 2018/09/11 改成模組
'   txtPayToday = ""
'   If txtCP118 = "Y" Then
'      '當日3點半以前預設當日扣款否則要人工輸入
'      If Text7(0) > strSrvDate(2) Or (Text7(0) = strSrvDate(2) And Val(ServerTime) <= 153000) Then
'         txtPayToday = "Y"
'      End If
'   End If
   txtPayToday.Text = Pub_FcpSetPayToday("1", Text7(0).Text, txtCP118.Text)
   'end 2018/09/11
End Sub

'Added by Lydia 2018/03/30 列印Word檔
Private Sub PrintFileList(ByVal sFileList As String, ByVal sCopy As Integer)
Dim oWordApp As Word.Application
Dim tmpArr As Variant
Dim intA As Integer, inR As Integer
Dim TempFile As String

Set oWordApp = New Word.Application
On Error GoTo ErrHnd
   
   If sCopy = 0 Or sFileList = "" Then Exit Sub
             
    tmpArr = Empty
    tmpArr = Split(sFileList, "&")
    For intA = 0 To UBound(tmpArr)
         If Trim(tmpArr(intA)) <> "" Then
              TempFile = tmpArr(intA)
              If Dir(m_AttchPath & "\" & TempFile) <> "" Then
                  For inR = 1 To sCopy
                        'PDF檔
                        If Right(UCase(TempFile), 4) = ".PDF" Then
                             PUB_PrintPDF m_AttchPath & "\" & TempFile
                        'Word檔
                        ElseIf Right(UCase(TempFile), 4) = ".DOC" Or Right(UCase(TempFile), 5) = ".DOCX" Then
                               'PUB_SetWordActivePrinter '因為是新開檔案,不用執行切換
                               oWordApp.Documents.Open FileName:=m_AttchPath & "\" & TempFile, ReadOnly:=True
                               oWordApp.ActiveDocument.PrintOut Background:=False '如果不設前景執行,可能列印未完成就關閉
                              'oWordApp.ActiveDocument.PrintOut Background:=False, Copies:=2, Collate:=True
                               oWordApp.Quit wdDoNotSaveChanges
                               Set g_WordAp = Nothing
                        Else '其他用系統預設
                              ShellExecute Me.hWnd, "print", m_AttchPath & "\" & TempFile, vbNullString, vbNullString, 1
                              Sleep 1000
                        End If
                  Next inR
              Else
                    MsgBox "找不到下列檔案:" & vbCrLf & m_AttchPath & "\" & TempFile
              End If
         End If
    Next intA
    
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox "自動列印說明書出錯，請通知電腦中心！" & vbCrLf & m_AttchPath & "\" & TempFile & _
                   vbCrLf & "錯誤訊息：" & Err.Description, vbCritical
      Resume Next
   End If
End Sub

'Added by Lydia 2019/10/25 翻譯瑕疵備註之選單
Private Sub Combo6_Validate(Cancel As Boolean)
    If Combo6.Tag <> Combo6.Text Then
       txtTF37.Text = txtTF37.Text & IIf(Trim(txtTF37.Text) <> "", "、", "") & Combo6.Text
    End If
    Combo6.Tag = Combo6.Text
End Sub

'Added by Lydia 2019/10/25
Private Sub txtTF37_Validate(Cancel As Boolean)
    '翻譯瑕疵備註
    If txtTF37.Text = "" Then Exit Sub
    If Not CheckLengthIsOK(txtTF37, 100) Then
       Cancel = True
    End If
End Sub

Private Sub txtTF37_GotFocus()
    TextInverse txtTF37
End Sub

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

'Added by Lydia 2021/01/21
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
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtRecDate_Validate(Cancel As Boolean)
   If txtRecDate.Tag <> txtRecDate.Text Then
      'Modify By Sindy 2022/5/18 + And cp(10) = "416"
      If txtRecDate = "Y" And cp(148) = "Y" And cp(10) = "416" Then
         txtEmail = "Y"
      End If
   End If
   txtRecDate.Tag = txtRecDate.Text
End Sub

'Added by Lydia 2021/01/21 實審(實體審查)發文：處理定稿、帳單和承辦單
Private Sub ProcDNfor416()
Dim m_Type As String
Dim nFrm As Form

    If Not (cp(1) = "FCP" And cp(10) = "416") Then Exit Sub
     
    If m_416Type = "0" Then '非特別情況,彈詢問
        '檢查同發文日是否有其他道發文,若無,才詢問是否要出帳單
        strExc(1) = "select cp09 from caseprogress" & _
                    " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                    " and cp27=" & DBDATE(Text7(0)) & " and cp57 is null" & _
                    " and cp09<>'" & Label3(0) & "' and cp43<>'" & Label3(0) & "'"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(1))
        If intI = 0 Then
           If MsgBox("實審是否要產生帳單？", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
               m_Type = "Y"
           Else
               m_Type = "N"
           End If
        End If
    End If
    
    If Val(cp(82)) = 0 And txtEmail.Text <> "N" Then   '與原先不同(選「出帳單」=Y)，不論是否出帳單第一次發文都要發email;
        '除了電子組在實審+主修時會先做請款信,會用這封信做後續請款流程,所以不用發通知email (=N)
        If m_AddMcRecord <> "" Then '人工維護Email
             cnnConnection.Execute m_AddMcRecord
        Else
             strExc(5) = Pub_FcpSetPayToday("2", Text7(0).Text, txtPayToday.Text) '扣款日
             Call PUB_GetFCPEmpMail416("1", strReceiveNo, m_eFlag, m_416Type, txtPAID, txtRecDate, DBDATE(Text7(0).Text), strExc(5))
        End If
    End If
    
    If m_Type = "Y" Or m_416Type = "2" Then
        '2.特殊請款(由承辦出帳單)：出定稿(產生電子檔在Typing2)，不出帳單

        '檢查表單是否已開啟，若是，則關閉
        For Each nFrm In Forms
           If StrComp(nFrm.Name, "frm060306_7", vbTextCompare) = 0 Then
              Unload frm060306_7
           End If
           If StrComp(nFrm.Name, "frm060306", vbTextCompare) = 0 Then
              Unload frm060306
              Exit For
           End If
        Next
        frm060306.Show
        frm060306.Text1.Text = pa(1)
        frm060306.Text2.Text = pa(2)
        frm060306.Text3.Text = pa(3)
        frm060306.Text4.Text = pa(4)
        frm060306.m_quy416 = True
        frm060306.Command1_Click
        If frm060306.MSHFlexGrid1.Rows >= 2 Then
           If frm060306.MSHFlexGrid1.TextMatrix(1, 2) <> "" Then
              frm060306.MSHFlexGrid1.TextMatrix(1, 0) = "v"
              Call frm060306.cmdok_Click(1)
              frm060306_7.Show
              '要產生請款單
              If cp(60) = "" And m_Type = "Y" And m_416Type <> "2" Then '尚未有請款單並且非特殊請款單
                 frm060306_7.Text1(1) = "Y"
              End If
              frm060306_7.txtPAID.Text = Me.txtPAID.Text ' 已收款
              frm060306_7.m_CallName = Me.Name  '呼叫的表單名稱
              Call frm060306_7.cmdok_Click(0)
              Unload frm060306
           End If
        End If
    End If
End Sub

'Added by Morgan 2024/11/18
Private Sub ProcDNfor447()
Dim nFrm As Form
Dim stSQL As String, intR As Integer
Dim strTo As String, strSubject As String, strContent As String, strCC As String

    If Not (cp(1) = "FCP" And cp(10) = "447") Then Exit Sub
    
    If PUB_GetST03(Text7(1)) = "F22" Then '承辦人是掛程序人員時才要自動跑請款函+帳單 --敏莉
        '檢查表單是否已開啟，若是，則關閉
        For Each nFrm In Forms
           If StrComp(nFrm.Name, "frm060306_7", vbTextCompare) = 0 Then
              Unload frm060306_7
           End If
           If StrComp(nFrm.Name, "frm060306", vbTextCompare) = 0 Then
              Unload frm060306
              Exit For
           End If
        Next
        frm060306.Show
        frm060306.Text1.Text = pa(1)
        frm060306.Text2.Text = pa(2)
        frm060306.Text3.Text = pa(3)
        frm060306.Text4.Text = pa(4)
        frm060306.m_quyAnyCP10 = "447"
        frm060306.Command1_Click
        If frm060306.MSHFlexGrid1.Rows >= 2 Then
           If frm060306.MSHFlexGrid1.TextMatrix(1, 2) <> "" Then
              frm060306.MSHFlexGrid1.TextMatrix(1, 0) = "v"
              Call frm060306.cmdok_Click(1)
              frm060306_7.Show
              'Modified by Morgan 2024/11/25 要排除特殊請款
              If cp(60) = "" And cp(148) <> "Y" Then
                 frm060306_7.Text1(1) = "Y"
              End If
              frm060306_7.txtPAID.Text = Me.txtPAID.Text ' 已收款
              frm060306_7.m_CallName = Me.Name  '呼叫的表單名稱
              Call frm060306_7.cmdok_Click(0)
              Unload frm060306
           End If
        End If
         
         'Modified by Morgan 2024/11/25 +判斷特殊請款
         strSubject = "【已送件-" & Label3(5) & "】" & IIf(cp(148) = "Y", "請進行請款", "可請款發文") & "(Email) Our Ref: " & pa(1) & "-" & pa(2) & IIf(pa(4) <> "00", "-" & pa(3) & "-" & pa(4), IIf(pa(3) <> "0", "-" & pa(3), "")) & " [INCOM.447]"
         
         If pa(150) = "4" Then
            strSubject = "【機械設計組】" & strSubject
         End If
         
         'Added by Morgan 2024/11/25
         If cp(148) = "Y" Then
            strContent = "1. 特殊請款，請進行人工請款作業。" & vbCrLf & _
                         "2. 請款定稿已存Typing2。"
         Else
         'end 2024/11/25
            strContent = "請款定稿、帳單已存Typing2。"
         End If
         
         strTo = PUB_GetFCPSalesNo(cp(1), cp(2), cp(3), cp(4), cp(10))
         'Modified by Morgan 2024/12/19 +CC給操作人--敏莉
         strCC = PUB_GetFCPProSup(strTo) & ";" & strUserNum & ";backup"
         stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                      "values('" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                      ",'" & ChgSQL(strSubject) & "','" & ChgSQL(strContent) & "'," & CNULL(strCC) & ")"
         cnnConnection.Execute stSQL, intR
        
    End If
End Sub

