VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_08 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(移轉)"
   ClientHeight    =   5892
   ClientLeft      =   5340
   ClientTop       =   408
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5892
   ScaleWidth      =   9156
   Begin VB.TextBox txtCP113 
      Height          =   270
      Left            =   3450
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2940
      Width           =   540
   End
   Begin VB.TextBox txtPayToday 
      Height          =   264
      Left            =   8295
      MaxLength       =   1
      TabIndex        =   81
      Top             =   5130
      Width           =   255
   End
   Begin VB.TextBox textCP118 
      Height          =   270
      Left            =   8040
      MaxLength       =   1
      TabIndex        =   6
      Top             =   2940
      Width           =   375
   End
   Begin VB.TextBox textCP91 
      Height          =   264
      Left            =   1185
      MaxLength       =   9
      TabIndex        =   14
      Top             =   4560
      Width           =   1092
   End
   Begin VB.TextBox textCP90 
      Height          =   264
      Left            =   1185
      MaxLength       =   9
      TabIndex        =   13
      Top             =   4320
      Width           =   1092
   End
   Begin VB.TextBox textCP89 
      Height          =   264
      Left            =   1185
      MaxLength       =   9
      TabIndex        =   12
      Top             =   4050
      Width           =   1092
   End
   Begin VB.TextBox textCP56 
      Height          =   264
      Left            =   1185
      MaxLength       =   9
      TabIndex        =   11
      Top             =   3795
      Width           =   1092
   End
   Begin VB.TextBox textCP92 
      Height          =   264
      Left            =   1185
      MaxLength       =   9
      TabIndex        =   15
      Top             =   4830
      Width           =   1092
   End
   Begin VB.TextBox textCP84 
      Alignment       =   1  '靠右對齊
      Height          =   285
      Left            =   3075
      TabIndex        =   1
      Top             =   2655
      Width           =   1035
   End
   Begin VB.TextBox textPrintMan 
      Height          =   264
      Left            =   5715
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "2"
      Top             =   3195
      Width           =   384
   End
   Begin VB.TextBox textCP22 
      Height          =   264
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   7
      Top             =   5175
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8310
      TabIndex        =   22
      Top             =   -15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6240
      TabIndex        =   20
      Top             =   -15
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7080
      TabIndex        =   21
      Top             =   -15
      Width           =   1200
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H008080FF&
      Caption         =   "變更事項(&R)"
      Height          =   400
      Left            =   5025
      Style           =   1  '圖片外觀
      TabIndex        =   19
      Top             =   -15
      Width           =   1200
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   3810
      TabIndex        =   18
      Top             =   -15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   1
      Left            =   2580
      TabIndex        =   17
      Top             =   -15
      Width           =   1200
   End
   Begin VB.TextBox textCP27 
      Height          =   264
      Left            =   795
      MaxLength       =   8
      TabIndex        =   0
      Top             =   2625
      Width           =   1320
   End
   Begin VB.TextBox textUargeDate 
      Height          =   264
      Left            =   960
      MaxLength       =   8
      TabIndex        =   3
      Top             =   2910
      Width           =   1392
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   1155
      MaxLength       =   1
      TabIndex        =   8
      Top             =   3195
      Width           =   372
   End
   Begin VB.TextBox textCP18 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5430
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2025
      Width           =   3675
   End
   Begin VB.TextBox textCF09 
      Height          =   264
      Left            =   4815
      MaxLength       =   12
      TabIndex        =   5
      Top             =   2940
      Width           =   612
   End
   Begin VB.ComboBox textCP44 
      Height          =   300
      Left            =   4800
      TabIndex        =   2
      Top             =   2640
      Width           =   1392
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   675
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   405
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1215
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1215
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5430
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   705
      Width           =   3675
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   945
      Width           =   2532
   End
   Begin VB.TextBox textTM15_S 
      Height          =   300
      Left            =   1860
      MaxLength       =   200
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   10
      Top             =   3465
      Width           =   7245
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   495
      Left            =   1170
      TabIndex        =   84
      Top             =   5100
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;882"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP44_2 
      Height          =   270
      Left            =   6210
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2640
      Width           =   2925
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "5159;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   300
      Left            =   1170
      TabIndex        =   16
      Top             =   5595
      Width           =   7815
      VariousPropertyBits=   -1467989989
      MaxLength       =   2000
      ScrollBars      =   2
      Size            =   "13779;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1185
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2325
      Width           =   7815
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13785;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44 
      Height          =   264
      Left            =   5430
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   450
      Width           =   3675
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP92_2 
      Height          =   270
      Left            =   2385
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   4830
      Width           =   6615
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "11663;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP91_2 
      Height          =   270
      Left            =   2385
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   4560
      Width           =   6615
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "11663;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP90_2 
      Height          =   270
      Left            =   2385
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   4320
      Width           =   6615
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "11663;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP89_2 
      Height          =   270
      Left            =   2385
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   4050
      Width           =   6615
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "11663;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM81 
      Height          =   264
      Left            =   1215
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   2010
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM80 
      Height          =   264
      Left            =   5430
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   1770
      Width           =   3675
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM79 
      Height          =   264
      Left            =   1215
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM78 
      Height          =   264
      Left            =   5430
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   1500
      Width           =   3675
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1215
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1485
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   264
      Left            =   5430
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1245
      Width           =   3675
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5430
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   975
      Width           =   3675
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP56_2 
      Height          =   270
      Left            =   2385
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3795
      Width           =   6615
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "11663;466"
      SpecialEffect   =   0
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
      Left            =   2550
      TabIndex        =   83
      Top             =   2970
      Width           =   765
   End
   Begin VB.Label lblPayToday 
      AutoSize        =   -1  'True
      Caption         =   "電子送件是否當日扣款:         (Y/N)"
      Height          =   180
      Left            =   6360
      TabIndex        =   82
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label43 
      AutoSize        =   -1  'True
      Caption         =   "是否電子送件:          (Y: 是)"
      Height          =   180
      Left            =   6870
      TabIndex        =   79
      Top             =   2970
      Width           =   2085
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "受讓人5 :"
      Height          =   180
      Left            =   210
      TabIndex        =   78
      Top             =   4875
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "受讓人4 :"
      Height          =   180
      Left            =   210
      TabIndex        =   76
      Top             =   4605
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "受讓人3 :"
      Height          =   180
      Left            =   210
      TabIndex        =   74
      Top             =   4365
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "受讓人2 :"
      Height          =   180
      Left            =   210
      TabIndex        =   72
      Top             =   4095
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2 :"
      Height          =   180
      Index           =   5
      Left            =   4545
      TabIndex        =   66
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3 :"
      Height          =   180
      Index           =   7
      Left            =   105
      TabIndex        =   65
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4 :"
      Height          =   180
      Index           =   13
      Left            =   4545
      TabIndex        =   64
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5 :"
      Height          =   180
      Index           =   14
      Left            =   105
      TabIndex        =   63
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   210
      TabIndex        =   62
      Top             =   5160
      Width           =   900
   End
   Begin VB.Label Label39 
      AutoSize        =   -1  'True
      Caption         =   "發文規費："
      Height          =   180
      Left            =   2145
      TabIndex        =   61
      Top             =   2715
      Width           =   900
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "(1:移轉人 2:受讓人 )"
      Height          =   180
      Left            =   6225
      TabIndex        =   60
      Top             =   3240
      Width           =   1560
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "定稿對象 :"
      Height          =   180
      Left            =   4755
      TabIndex        =   59
      Top             =   3240
      Width           =   810
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "(N:不出名)"
      Height          =   180
      Left            =   4200
      TabIndex        =   58
      Top             =   5220
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "是否出名 :"
      Height          =   180
      Left            =   2760
      TabIndex        =   57
      Top             =   5220
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "進度備註 :"
      Height          =   180
      Left            =   120
      TabIndex        =   56
      Top             =   5640
      Width           =   810
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "發文日 :"
      Height          =   180
      Left            =   105
      TabIndex        =   55
      Top             =   2625
      Width           =   630
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "催審期限 :"
      Height          =   180
      Left            =   135
      TabIndex        =   54
      Top             =   2925
      Width           =   810
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "代理人 :"
      Height          =   180
      Left            =   4125
      TabIndex        =   53
      Top             =   2685
      Width           =   630
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "列印定稿 :"
      Height          =   180
      Left            =   75
      TabIndex        =   52
      Top             =   3240
      Width           =   810
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
      Height          =   180
      Left            =   1695
      TabIndex        =   51
      Top             =   3240
      Width           =   2745
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "點數 :"
      Height          =   180
      Index           =   10
      Left            =   4545
      TabIndex        =   50
      Top             =   2055
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "大約"
      Height          =   180
      Index           =   12
      Left            =   4410
      TabIndex        =   49
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "可接獲回音"
      Height          =   180
      Left            =   5460
      TabIndex        =   48
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人1 :"
      Height          =   180
      Left            =   105
      TabIndex        =   47
      Top             =   1530
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   255
      Index           =   4
      Left            =   5835
      TabIndex        =   46
      Top             =   1755
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "承辦人 :"
      Height          =   180
      Left            =   4545
      TabIndex        =   45
      Top             =   1245
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人 :"
      Height          =   180
      Index           =   2
      Left            =   4545
      TabIndex        =   44
      Top             =   435
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Index           =   0
      Left            =   105
      TabIndex        =   43
      Top             =   675
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號 :"
      Height          =   180
      Index           =   1
      Left            =   105
      TabIndex        =   42
      Top             =   405
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質 :"
      Height          =   180
      Index           =   6
      Left            =   105
      TabIndex        =   41
      Top             =   1215
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號 :"
      Height          =   180
      Index           =   9
      Left            =   4545
      TabIndex        =   40
      Top             =   705
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員 :"
      Height          =   180
      Index           =   11
      Left            =   4545
      TabIndex        =   39
      Top             =   975
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "審定號數/申請案號 :"
      Height          =   180
      Left            =   105
      TabIndex        =   38
      Top             =   990
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱 :"
      Height          =   180
      Left            =   105
      TabIndex        =   37
      Top             =   2325
      Width           =   810
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "不一併移轉註冊號數 :"
      Height          =   180
      Left            =   135
      TabIndex        =   36
      Top             =   3465
      Width           =   1710
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "受讓人1 :"
      Height          =   180
      Left            =   210
      TabIndex        =   35
      Top             =   3840
      Width           =   720
   End
End
Attribute VB_Name = "frm020102_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/23 Form2.0已修改 textTM44/textCP13/textCP14/textTM23(申請人名).../textCP56_2(受讓人名).../cmbTM05/textCP64/textCP44_2/lstNameAgent
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
' 案件性質代號
Dim m_CP10 As String
' 申請人
Dim m_TM23 As String
'add by nickc 2007/01/03
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
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
'Add By Cheng 2002/12/12
Dim m_CP14 As String '原承辦人
'Add By Cheng 2003/03/07
Dim m_CP55 As String '移轉人
'add by nickc 2007/01/08
Dim m_CP93 As String
Dim m_CP94 As String
Dim m_CP95 As String
Dim m_CP96 As String
'Add By Cheng 2003/10/06
Public m_blnClkChgButton As Boolean '是否按下變更事項按鈕
'add by nick 2004/08/12
Dim m_CP84 As String            '發文規費
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

'add by nickc 2006/01/27
Dim m_CP110 As String
'add by nickc 2007/01/03
Dim m_textUargeDate As String
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
' 智權人員
Dim m_CP13 As String
Dim m_CP12 As String
'Dim strCP45 As String
'2010/6/10 End
Dim m_CP07 As String 'Add By Sindy 2010/12/28 法定期限
Dim m_QSP As Boolean 'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
Dim m_990CP09 As String 'Add By Sindy 2016/12/20
Dim strLD18 As String 'Add By Sindy 2019/12/20 信函總收文號
Dim m_strCF10 As String 'Add By Sindy 2020/8/12 取得主管機關
Dim m_AgentName As String 'Add By Amy 2021/12/23

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
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
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

Private Sub cmdMod_Click()
   frm020102_04.SetData 0, m_TM01, True
   frm020102_04.SetData 1, m_TM02, False
   frm020102_04.SetData 2, m_TM03, False
   frm020102_04.SetData 3, m_TM04, False
   frm020102_04.SetData 4, m_CP09, False
   frm020102_04.SetParent Me
   frm020102_04.SetParent_MainForm frm020102_01 'Add By Sindy 2018/9/25
   Me.Hide
   frm020102_04.Show
   frm020102_04.QueryData
'    m_blnClkChgButton = True
End Sub

Private Sub cmdok_Click(Index As Integer)
'Add By Cheng 2002/11/11
Dim ii As Integer
Dim strNewCP64 As String 'Add by Amy 2020/02/05 進度備註
   
   'Modify By Sindy 2010/11/19 把「確定」及「同時發文」按鈕程式碼合併
   Select Case Index
      Case 0, 1
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
                  '2005/8/2 MODIFY BY SONIA
                  'GetCu103ByCustomer Me, m_TM23
                  'Modified by Lydia 2024/07/03 改傳入變數;
                  'GetCu103ByCustomer Me, textCP56.Text
                  Call Pub_GetDataFrm020102(textCP56.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                  
                  '2005/8/2 END
                  'edit by nickc 2006/01/20
                  'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Then
                  If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                        'Modified by Lydia 2024/07/03
                        'Set frm020102_22.oNextForm = Me
                        Call frm020102_22.SetParent(Me, textCP56.Text)
                        frm020102_22.Label4.Caption = textCP56 & " " & textCP56_2 'Add By Sindy 2014/7/30
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
                  If textCP89.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                  'Modified by Lydia 2024/07/03 改傳入變數;
                  'GetCu103ByCustomer Me, textCP89.Text
                  Call Pub_GetDataFrm020102(textCP89.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                  
                  If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                        'Modified by Lydia 2024/07/03
                        'Set frm020102_22.oNextForm = Me
                        Call frm020102_22.SetParent(Me, textCP89.Text)
                        frm020102_22.Label4.Caption = textCP89 & " " & textCP89_2 'Add By Sindy 2014/7/30
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
                  If textCP90.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                  'Modified by Lydia 2024/07/03 改傳入變數;
                  'GetCu103ByCustomer Me, textCP90.Text
                  Call Pub_GetDataFrm020102(textCP90.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                  
                  If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                        'Modified by Lydia 2024/07/03
                        'Set frm020102_22.oNextForm = Me
                        Call frm020102_22.SetParent(Me, textCP90.Text)
                        frm020102_22.Label4.Caption = textCP90 & " " & textCP90_2 'Add By Sindy 2014/7/30
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
                  If textCP91.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                  'Modified by Lydia 2024/07/03 改傳入變數;
                  'GetCu103ByCustomer Me, textCP91.Text
                  Call Pub_GetDataFrm020102(textCP91.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                  
                  If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                        'Modified by Lydia 2024/07/03
                        'Set frm020102_22.oNextForm = Me
                        Call frm020102_22.SetParent(Me, textCP91.Text)
                        frm020102_22.Label4.Caption = textCP91 & " " & textCP91_2 'Add By Sindy 2014/7/30
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
                  If textCP92.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                  'Modified by Lydia 2024/07/03 改傳入變數;
                  'GetCu103ByCustomer Me, textCP92.Text
                  Call Pub_GetDataFrm020102(textCP92.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                  
                  If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                        'Modified by Lydia 2024/07/03
                        'Set frm020102_22.oNextForm = Me
                        Call frm020102_22.SetParent(Me, textCP92.Text)
                        frm020102_22.Label4.Caption = textCP92 & " " & textCP92_2 'Add By Sindy 2014/7/30
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
            
            strNewCP64 = textCP64 'Add by Amy 2020/02/05
            
            'Modify By Sindy 2011/3/9 若為電子送件則不經發文室
            'Modify By Sindy 2023/8/1 電子送件欄位值不是空白者,即為電子送件
            If (textCP118.Visible = True And textCP118 <> "") Then
               'Added by Morgan 2016/5/16 電子送件也要記錄主管機關
               If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27, , True) = False Then
                  Exit Sub
               End If
               'end 2016/5/16
               'Add by Amy 2020/02/05 +輸入收文文號
               If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
                  'Add By Sindy 2020/8/12 主管機關為經濟部智慧財產局,才做自動扣款
                  If m_CP130s = "經濟部智慧財產局" Then
                  '2020/8/12 END
                     'Add by Amy 2020/01/13
                     'If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
                        'If textCP118 = "Y" And Val(textCP84) > 0 Then
                        If Val(textCP84) > 0 Then
                           If txtPayToday.Visible = True And txtPayToday = "" Then
                              MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
                              txtPayToday.SetFocus
                              Exit Sub
                           End If
                           strExc(0) = InputBox("請輸入智慧局收文文號!!")
                           If strExc(0) = "" Then
                              Exit Sub
                           Else
                              strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & textCP64 '先保留進度備註，等檢查完後更新欄位
                           End If
                        End If
                     'End If
                     'end 2020/01/13
                  'Add By Sindy 2020/8/12
                  ElseIf txtPayToday.Visible = True And txtPayToday <> "" Then
                     txtPayToday = ""
                  End If
                  '2020/8/12 END
               End If
               'end 2020/02/05
            Else
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
            End If
            
            textCP64 = strNewCP64 'Add by Amy 2020/02/05
            
            ' 設定滑鼠游標為等待狀態
            Screen.MousePointer = vbHourglass
            ' 更新欄位輸入的內容
            OnUpdateField
            ' 存檔
          'Modify By Cheng 2002/11/06
      '      'OnSaveData
            If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
              '定稿對象 1.原申請人(TM23) 2.移轉申請人(CP56)
              'Add By Cheng 2002/11/11
              Select Case m_TM01
              Case "TB"
                 If textPrint <> "N" Then
                    If Len(Me.textPrintMan.Text) > 0 Then
                       For ii = 1 To Len(Me.textPrintMan.Text)
                          PrintLetterTB (Mid(Me.textPrintMan.Text, ii, 1))
                       Next ii
                    End If
'                 'Add By Sindy 2021/2/25
'                 Else
'                     If strLD18 <> "" Then
'                        Call PUB_TCaseAskIsPost(strLD18)
'                     End If
'                 '2021/2/25 END
                 End If
              Case Else
                 If textPrint <> "N" Then
                    If Len(Me.textPrintMan.Text) > 0 Then
                       For ii = 1 To Len(Me.textPrintMan.Text)
                          PrintLetter (Mid(Me.textPrintMan.Text, ii, 1))
                       Next ii
                    End If
'                 'Add By Sindy 2021/2/25
'                 Else
'                     If strLD18 <> "" Then
'                        Call PUB_TCaseAskIsPost(strLD18)
'                     End If
'                 '2021/2/25 END
                 End If
              End Select
              'Add By Sindy 2025/8/15
              If textPrint = "N" Then
                  If strLD18 <> "" Then
                     Call PUB_TCaseAskIsPost(strLD18)
                  End If
              End If
              '2025/8/15 END
              
            '2012/7/23 add by sonia
            '台灣案發文規費與收文規費不符時,mail給智權人員
            If textCP84.Enabled = True And m_TM10 = "000" And Val(Me.textCP84.Text) <> Val(m_CP84) Then
               'Add by Lydia 2014/10/13 內商服務業務(TC)之台灣案發文-規費與收文規費不符時,請加同時發給特殊設定人員"財務處總帳人員"
               If m_QSP = True Then
                 PUB_ChkOfficialFee m_CP09, Me.textCP84.Text, "A"
               Else
                 '2020/01/13 Modify by Amy +if 傳strCP118參數
                 If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
                    PUB_ChkOfficialFee m_CP09, Me.textCP84.Text, IIf(textCP118 = "Y", "A", "")
                 Else
                    PUB_ChkOfficialFee m_CP09, Me.textCP84.Text
                 End If
               End If
            End If
            '2012/7/23 end
            
            'Add By Sindy 2018/5/3
            If frm020102_01.bolIsEMPFlow = True Then
               frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
               frm090202_4.QueryData
            End If
            '2018/5/3 End
            
            'Add By Sindy 2025/7/11 外商發文時,增加發Mail通知承辦人及副本給判發主管
            If Left(m_CP12, 1) = "F" Then
               Call PUB_FCTSendRecvMail(m_CP09)
            End If
            '2025/7/11 END
            
            ' 設定滑鼠游標為預設
            Screen.MousePointer = vbDefault
            If Index = 0 Then '確定鍵
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
            ElseIf Index = 1 Then '同時發文鍵
               ' 呼叫第一個畫面
               frm020102_01.SetData 0, m_TM01, True
               frm020102_01.SetData 1, m_TM02, False
               frm020102_01.SetData 2, m_TM03, False
               frm020102_01.SetData 3, m_TM04, False
               frm020102_01.SetQueryFromTM
               Unload Me
               frm020102_01.Show
               frm020102_01.radio(1).Value = True
               frm020102_01.radio_Click 1
               frm020102_01.QueryData
            End If
         End If
      Case Else
   End Select
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

'Private Sub cmdTogether_Click()
''Add By Cheng 2002/11/11
'Dim ii As Integer
'
'   If CheckDataValid = True Then
'      'Add By Cheng 2002/07/15
'      '重新檢查欄位有效性
'      If TxtValidate = False Then Exit Sub
'
'      ' 設定滑鼠游標為等待狀態
'      Screen.MousePointer = vbHourglass
'      ' 更新欄位輸入的內容
'      OnUpdateField
'      ' 存檔
'    'Modify By Cheng 2002/11/06
''      'OnSaveData
'      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
'        'Add By Cheng 2002/11/11
'        Select Case m_TM01
'        Case "TB"
'           If textPrint <> "N" Then
'              If Len(Me.textPrintMan.Text) > 0 Then
'                 For ii = 1 To Len(Me.textPrintMan.Text)
'                    PrintLetterTB (Mid(Me.textPrintMan.Text, ii, 1))
'                 Next ii
'              End If
'           End If
'        Case Else
'           If textPrint <> "N" Then
'              If Len(Me.textPrintMan.Text) > 0 Then
'                 For ii = 1 To Len(Me.textPrintMan.Text)
'                    PrintLetter (Mid(Me.textPrintMan.Text, ii, 1))
'                 Next ii
'              End If
'           End If
'        End Select
'
'      ' 設定滑鼠游標為預設
'      Screen.MousePointer = vbDefault
'
'      ' 呼叫第一個畫面
'      frm020102_01.SetData 0, m_TM01, True
'      frm020102_01.SetData 1, m_TM02, False
'      frm020102_01.SetData 2, m_TM03, False
'      frm020102_01.SetData 3, m_TM04, False
'      frm020102_01.SetQueryFromTM
'      Unload Me
'      frm020102_01.Show
'      frm020102_01.radio(1).Value = True
'      frm020102_01.radio_Click 1
'      frm020102_01.QueryData
'   End If
'End Sub

'Private Sub Form_Activate()
'    'Add By Cheng 2003/10/06
'    '若有按下變更事項按鈕, 則重新讀取資料
'    'edit by nickc 2005/08/23
'    'If m_blnClkChgButton = True  Then
'    If m_blnClkChgButton = True Or (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
'        pub_ModifyCaseNum = ""
'        QueryData
''        m_blnClkChgButton = False
'    End If
'End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   
   'add by nickc 2007/01/30
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
   textCP44_2.BackColor = &H8000000F
   textCP56_2.BackColor = &H8000000F
   
   'add by nickc 2007/01/30
   textCP89_2.BackColor = &H8000000F
   textCP90_2.BackColor = &H8000000F
   textCP91_2.BackColor = &H8000000F
   textCP92_2.BackColor = &H8000000F
   
   
   MoveFormToCenter Me
'    m_blnClkChgButton = False
   'Add by nickc 2006/01/27
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Add by Amy 2021/12/23一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
    lstNameAgent.Height = 500
    lstNameAgent.Width = 1300
 
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2012/4/17
   
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
         'Add By Sindy 2012/4/17
         strSql = "SELECT * FROM ChangeEvent " & _
                  "WHERE CE01 = '" & m_CP09 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount <= 0 Then
            m_blnClkChgButton = True
         Else
            m_blnClkChgButton = False
         End If
         rsTmp.Close
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
      Else
         'Add By Sindy 2009/06/29
         ' 申請案號
         If IsNull(rsTmp.Fields("TM12")) = False Then
            textTM15 = rsTmp.Fields("TM12")
         End If
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         'textTM12 = rsTmp.Fields("TM12")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         'textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
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
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         'textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      'add b nickc 2007/01/03
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
      
      ' 正商標號數
      If IsNull(rsTmp.Fields("TM27")) = False Then
         'textTM27 = rsTmp.Fields("TM27")
      End If
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("tm77"))
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
      'add by nickc 2007/01/03
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
         'textTM12 = rsTmp.Fields("SP11")
         'Add By Sindy 2009/06/29
         textTM15 = rsTmp.Fields("SP11")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("SP12")) = False Then
         'textTM20 = TAIWANDATE(rsTmp.Fields("SP12"))
      End If
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("sp72"))
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
Dim strCP44 As String
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
      m_CP10 = ""
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 業務區別
      m_CP12 = ""
      If IsNull(rsTmp.Fields("CP12")) = False Then
         '91.6.11 MODIFY BY SONIA
         'textCP12 = GetStaffDepartment(rsTmp.Fields("CP12"))
         'textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      m_CP13 = ""
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
         m_CP13 = rsTmp.Fields("CP13")
      End If
      
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      'Add By Cheng 2002/12/12
      m_CP14 = "" & rsTmp.Fields("CP14")
      
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
      'Add By Sindy 2011/3/9
      ' 是否電子送件
      textCP118 = Empty
      If IsNull(rsTmp.Fields("CP118")) = False Then
         textCP118 = rsTmp.Fields("CP118")
      End If
      SetCPFieldOldData "CP118", textCP118, 0
      
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
      'ADD BY SONIA 2014/11/6 電子送件案預設發文日為承辦人發文日CP85
      If textCP118 = "Y" Then
         textCP27 = TAIWANDATE(rsTmp.Fields("CP85"))
      End If
      'END  2014/11/6
      
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = rsTmp.Fields("CP44")
      End If
      SetCPFieldOldData "CP44", textCP44, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      ' 彼所案號
'      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         textTM45 = rsTmp.Fields("CP45")
'         strCP45 = rsTmp.Fields("CP45")
      End If
      SetCPFieldOldData "CP45", textTM45, 0
'      SetCPFieldOldData "CP45", strCP45, 0
        'Add By Cheng 2003/03/07
        '記錄原移轉人
        m_CP55 = "" & rsTmp("CP55").Value
      SetCPFieldOldData "CP55", m_CP55, 0
      ' 移轉申請人
      textCP56 = Empty
      If IsNull(rsTmp.Fields("CP56")) = False Then
         textCP56 = rsTmp.Fields("CP56")
         textCP56_Validate False
      End If
      SetCPFieldOldData "CP56", textCP56, 0
      'add by nickc 2007/01/08
      m_CP93 = CheckStr(rsTmp.Fields("CP93"))
      m_CP94 = CheckStr(rsTmp.Fields("CP94"))
      m_CP95 = CheckStr(rsTmp.Fields("CP95"))
      m_CP96 = CheckStr(rsTmp.Fields("CP96"))
      SetCPFieldOldData "CP93", m_CP93, 0
      SetCPFieldOldData "CP94", m_CP94, 0
      SetCPFieldOldData "CP95", m_CP95, 0
      SetCPFieldOldData "CP96", m_CP96, 0
      textCP89 = Empty
      If IsNull(rsTmp.Fields("CP89")) = False Then
         textCP89 = rsTmp.Fields("CP89")
         textCP89_Validate False
      End If
      SetCPFieldOldData "CP89", textCP89, 0
      textCP90 = Empty
      If IsNull(rsTmp.Fields("CP90")) = False Then
         textCP90 = rsTmp.Fields("CP90")
         textCP90_Validate False
      End If
      SetCPFieldOldData "CP90", textCP90, 0
      textCP91 = Empty
      If IsNull(rsTmp.Fields("CP91")) = False Then
         textCP91 = rsTmp.Fields("CP91")
         textCP91_Validate False
      End If
      SetCPFieldOldData "CP91", textCP91, 0
      textCP92 = Empty
      If IsNull(rsTmp.Fields("CP92")) = False Then
         textCP92 = rsTmp.Fields("CP92")
         textCP92_Validate False
      End If
      SetCPFieldOldData "CP92", textCP92, 0
      
      
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      'add by nickc 2006/01/27
      'm_CP110 = CheckStr(rsTmp.Fields("cp110"))
      'SetCPFieldOldData "CP110", m_CP110, 0
      'Modify By Sindy 2010/9/20
      If m_CP110 = "" Then m_CP110 = CheckStr(rsTmp.Fields("cp110"))
      If m_CP110 = "" And m_TM10 = "000" Then m_CP110 = "94007,81040" 'Add By Sindy 2016/8/31 移轉(501)時,出名代理人預設為94007.林景郁和81040.閻啟泰
      SetCPFieldOldData "CP110", CheckStr(rsTmp.Fields("cp110")), 0
      '2010/9/20 End
      
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
      strSubSQL = "SELECT CP44, Max(CP27||CP09) FROM CaseProgress " & _
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
           'Modify By Cheng 2002/09/18
'            textCP44.AddItem m_AgentList(nIndex).aiName
           textCP44.AddItem m_AgentList(nIndex).aiCode
        Next nIndex
        ' 設定顯示為第一筆
        If textCP44.ListCount > 0 Then
           textCP44.ListIndex = 0
           textCP44_Validate False
        End If
    'add by nick 2004/08/12 發文規費
     If IsNull(rsTmp.Fields("CP17")) = False Then
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
      'Add by Amy 2020/01/13 電子送件一率自動扣款(A)若超過3點半發文則須人工輸入是否當日扣款
      If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
         txtPayToday = ""
         If textCP118 = "Y" Then
            'Modify by Amy 2020/08/11 發文日小於系統日,電子送件是否當日扣款設N;發文日為當天且3點半前才設Y(原只判斷3點半)
            If Val(textCP27) < strSrvDate(2) Then
               txtPayToday = "N"
            ElseIf Val(textCP27) = strSrvDate(2) And Val(ServerTime) <= 153000 Then
               txtPayToday = "Y"
            End If
            'end 2020/08/11
         End If
      End If
      'end 2020/01/13
      textCP27.Tag = textCP27.Text 'Add By Sindy 2020/8/12
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   'add by nickc 2006/01/27
   Dim tm(1 To 4) As String
   
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

   'add by nickc 2006/01/27
   tm(1) = m_TM01
   tm(2) = m_TM02
   tm(3) = m_TM03
   tm(4) = m_TM04
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
   
   'Modify By Sindy 2012/7/26
   '台灣案才需顯示出名代理人
   lstNameAgent.Clear
   If m_TM10 = "000" Then
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
      'Modify by Amy 2021/12/23 改Form2.0,bForm2設True
      PUB_SetOurAgent lstNameAgent, tm(), m_CP110, , True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2012/7/26 End
   
   'add by nickc 2007/01/03 取得催審期限的日期
   textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
   textCP27.Tag = textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   
   ' 大約?可接獲回音(欄位)
   textCF09 = Empty
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CF09")) = False Then
         textCF09 = rsTmp.Fields("CF09")
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/17 若已經從基本檔抓出來，就不重抓
   If Trim(textPrint) = "" Then
        textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   'Add By Sindy 2025/8/11 檢查卷宗區是否已有承辦放入之CUS,若有,系統不產出定稿
   If PUB_CPPChkFileExists(m_CP09, "cus") = True Then
      textPrint = "N"
   End If
   '2025/8/11 END
   
   'Add By Sindy 2011/10/28 T內商000台灣案所有案件性質加電子送件功能
   'Modify by Amy 2020/01/23 +是否電子送件
   lblPayToday.Visible = False
   txtPayToday.Visible = False
   If m_TM01 = "T" And m_TM10 = "000" Then
      Label43.Visible = True
      textCP118.Visible = True
      If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
        lblPayToday.Visible = True
        txtPayToday.Visible = True
      End If
   'end 2020/01/13
   Else
      Label43.Visible = False
      textCP118.Visible = False
   End If
   '2011/10/28 End
   
   Call PUB_TCaseEFeeRemind(m_CP09) 'Add By Sindy 2016/5/9 內商電子收文請款提醒訊息
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
    'Add By Cheng 2002/07/18
   Set frm020102_08 = Nothing
End Sub

'add by nickc 2006/01/27
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/5 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify by Amy 2021/12/23 改Form2.0,使用PUB_Num2Id會錯
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         'end 2021/12/23

         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      textCP22 = ""
   Else
      textCP22 = "N"
   End If
   'Add By Sindy 2015/7/22
   If textCP118 = "Y" And textCP22 = "N" Then
      Cancel = True
      MsgBox "電子送件時不可為不出名!!!", vbExclamation, "資料檢核"
      lstNameAgent.SetFocus
   End If
   '2015/7/22 END
End Sub

'Add By Sindy 2011/10/28
Private Sub textCP118_GotFocus()
   TextInverse textCP118
   CloseIme
End Sub

'Add By Sindy 2011/10/28
Private Sub textCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

' 是否出名
Private Sub textCP22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'edit by nickc 2006/01/27
' 是否出名
'Private Sub textCP22_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP22) = False Then
'      Select Case textCP22
'         Case " ", "N":
'         Case Else
'            Cancel = True
'            strTit = "資料檢核"
'            strMsg = "只可輸入空白或N"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textCP22_GotFocus
'      End Select
'   End If
'End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2002/12/03
    KeyAscii = UpperCase(KeyAscii)
End Sub

' 89.06.18 受讓人轉大寫
Private Sub textCP56_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 受讓人
Private Sub textCP56_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCP56_2 = Empty
   If IsEmptyText(textCP56) = False Then
        Me.textCP56 = ChangeCustomerL(Me.textCP56.Text)
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      oState = True
      'textCP56_2 = GetCustomerName(textCP56)
      textCP56_2 = GetCustomerNameAndState(textCP56, "0", oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textCP56_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "受讓人代碼<" & textCP56 & ">不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP56_GotFocus
      End If
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
      'Add by Amy 2020/01/13 當發文日有改時,電子送件案要人工輸入是否當日扣款
      If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
        If textCP27.Tag <> textCP27.Text Then
            textCP27.Tag = textCP27.Text
            If textCP118 = "Y" Then
                txtPayToday.Text = ""
            End If
        End If
      End If
      'end 2020/01/13
      'add by nickc 2007/01/11 重新算催審期限
      'Modified by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
      'If textUargeDate = m_textUargeDate Then
      If textCP27.Tag <> textCP27.Text Then
        textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
        m_textUargeDate = textUargeDate
      End If
      textCP27.Tag = textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   End If
EXITSUB:
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCP44_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCP44.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
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
   If m_TM10 <> 台灣國家代號 Then
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
         'edit by nick 2004/07/22
         If strTempName <> "" Then
            Cancel = True
            Exit Sub
         End If
      End If
      If IsEmptyText(textCP44_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44_GotFocus
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
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

Private Sub textCP84_GotFocus()
   InverseTextBox textCP84
End Sub

Private Sub textCP84_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
   
   Cancel = False
Dim nResponse
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

'add by nickc 2007/01/08
Private Sub textCP89_GotFocus()
InverseTextBox textCP89
End Sub
Private Sub textCP89_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textCP89_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCP89_2 = Empty
   If IsEmptyText(textCP89) = False Then
        Me.textCP89 = ChangeCustomerL(Me.textCP89.Text)
      Dim oState As Boolean
      oState = True
      textCP89_2 = GetCustomerNameAndState(textCP89, "0", oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textCP89_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "受讓人2代碼<" & textCP89 & ">不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP89_GotFocus
      End If
   End If
End Sub
Private Sub textCP90_GotFocus()
InverseTextBox textCP90
End Sub
Private Sub textCP90_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textCP90_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCP90_2 = Empty
   If IsEmptyText(textCP90) = False Then
        Me.textCP90 = ChangeCustomerL(Me.textCP90.Text)
      Dim oState As Boolean
      oState = True
      textCP90_2 = GetCustomerNameAndState(textCP90, "0", oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textCP90_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "受讓人3代碼<" & textCP90 & ">不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP90_GotFocus
      End If
   End If
End Sub
Private Sub textCP91_GotFocus()
InverseTextBox textCP91
End Sub
Private Sub textCP91_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textCP91_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCP91_2 = Empty
   If IsEmptyText(textCP91) = False Then
        Me.textCP91 = ChangeCustomerL(Me.textCP91.Text)
      Dim oState As Boolean
      oState = True
      textCP91_2 = GetCustomerNameAndState(textCP91, "0", oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textCP91_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "受讓人4代碼<" & textCP91 & ">不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP91_GotFocus
      End If
   End If
End Sub
Private Sub textCP92_GotFocus()
InverseTextBox textCP92
End Sub
Private Sub textCP92_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textCP92_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCP92_2 = Empty
   If IsEmptyText(textCP92) = False Then
        Me.textCP92 = ChangeCustomerL(Me.textCP92.Text)
      Dim oState As Boolean
      oState = True
      textCP92_2 = GetCustomerNameAndState(textCP92, "0", oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textCP92_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "受讓人5代碼<" & textCP92 & ">不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP92_GotFocus
      End If
   End If
End Sub

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

' 更新欄位的內容
Private Sub OnUpdateField()
   Dim strCP64 As String
   
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' 代理人
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
    'Add By Cheng 2003/03/07
    '若有輸入移轉申請人
    If Me.textCP56.Text <> "" Then
        '若移轉人與原申請人不同時
        '2009/10/19 modify by sonia 應為判斷移轉申請人與原申請人不同時
        'If ChangeCustomerL(m_CP55) <> ChangeCustomerL(m_TM23) Then
        If ChangeCustomerL(textCP56) <> ChangeCustomerL(m_TM23) Then
            '更新進度檔移轉人
            SetCPFieldNewData "CP55", ChangeCustomerL(m_TM23)
        End If
    End If
   ' 移轉申請人
   If IsEmptyText(textCP56) = False Then
      SetCPFieldNewData "CP56", textCP56 & String(9 - Len(textCP56), "0")
   Else
      SetCPFieldNewData "CP56", textCP56
   End If
   ' 進度備註
   strCP64 = textCP64
   'Modify By Cheng 2003/06/03
   'SetCPFieldNewData "CP64", textCP64 & "； " & "不一併移轉註冊號數 : " & textTM15_S
   If Me.textTM15_S.Text <> "" Then strCP64 = strCP64 & "不一併移轉註冊號數 : " & textTM15_S
'edit by nickc     2006/01/27
'   If IsEmptyText(textAgName) = False Then
'      strCP64 = strCP64 & "," & "本所出名代理人:" & textAgName
'   End If
   SetCPFieldNewData "CP64", strCP64
   
   ' 是否出名
   SetCPFieldNewData "CP22", textCP22
   'add by nickc 2006/01/27
   SetCPFieldNewData "CP110", m_CP110
   'add by nickc 2007/01/08
    '若有輸入移轉申請人
    If Me.textCP89.Text <> "" Then
        '若移轉人與原申請人不同時
        '2009/10/19 modify by sonia 應為判斷移轉申請人與原申請人不同時
        'If ChangeCustomerL(m_CP93) <> ChangeCustomerL(m_TM78) Then
        If ChangeCustomerL(textCP89) <> ChangeCustomerL(m_TM78) Then
            '更新進度檔移轉人
            SetCPFieldNewData "CP93", ChangeCustomerL(m_TM78)
        End If
    End If
   ' 移轉申請人
   If IsEmptyText(textCP89) = False Then
      SetCPFieldNewData "CP89", textCP89 & String(9 - Len(textCP89), "0")
   Else
      SetCPFieldNewData "CP89", textCP89
   End If
    '若有輸入移轉申請人
    If Me.textCP90.Text <> "" Then
        '若移轉人與原申請人不同時
        '2009/10/19 modify by sonia 應為判斷移轉申請人與原申請人不同時
        'If ChangeCustomerL(m_CP94) <> ChangeCustomerL(m_TM79) Then
        If ChangeCustomerL(textCP90) <> ChangeCustomerL(m_TM79) Then
            '更新進度檔移轉人
            SetCPFieldNewData "CP94", ChangeCustomerL(m_TM79)
        End If
    End If
   ' 移轉申請人
   If IsEmptyText(textCP90) = False Then
      SetCPFieldNewData "CP90", textCP90 & String(9 - Len(textCP90), "0")
   Else
      SetCPFieldNewData "CP90", textCP90
   End If
    '若有輸入移轉申請人
    If Me.textCP91.Text <> "" Then
        '若移轉人與原申請人不同時
        '2009/10/19 modify by sonia 應為判斷移轉申請人與原申請人不同時
        'If ChangeCustomerL(m_CP95) <> ChangeCustomerL(m_TM80) Then
        If ChangeCustomerL(textCP91) <> ChangeCustomerL(m_TM80) Then
            '更新進度檔移轉人
            SetCPFieldNewData "CP95", ChangeCustomerL(m_TM80)
        End If
    End If
   ' 移轉申請人
   If IsEmptyText(textCP91) = False Then
      SetCPFieldNewData "CP91", textCP91 & String(9 - Len(textCP91), "0")
   Else
      SetCPFieldNewData "CP91", textCP91
   End If
    '若有輸入移轉申請人
    If Me.textCP92.Text <> "" Then
        '若移轉人與原申請人不同時
        '2009/10/19 modify by sonia 應為判斷移轉申請人與原申請人不同時
        'If ChangeCustomerL(m_CP96) <> ChangeCustomerL(m_TM81) Then
        If ChangeCustomerL(textCP92) <> ChangeCustomerL(m_TM81) Then
            '更新進度檔移轉人
            SetCPFieldNewData "CP96", ChangeCustomerL(m_TM81)
        End If
    End If
   ' 移轉申請人
   If IsEmptyText(textCP92) = False Then
      SetCPFieldNewData "CP92", textCP92 & String(9 - Len(textCP92), "0")
   Else
      SetCPFieldNewData "CP92", textCP92
   End If
   'Add By Sindy 2011/3/9
   ' 是否電子送件
   SetCPFieldNewData "CP118", textCP118
   
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
End Sub

' 更新案件進度檔
'Modify By Cheng 2002/11/06
'Private Sub OnUpdateCaseProgress()
Private Function OnUpdateCaseProgress() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   
'Add By Cheng 2002/11/06
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
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateCaseProgress = False
End Function

'Modify By Cheng 2002/11/06
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strNP08 As String
Dim strNP07 As String
Dim strNP22 As String
Dim objCopyCP As ClsCopyCP
'Add By Cheng 2002/06/18
Dim ii As Integer
'Add By Cheng 2003/03/07
Dim StrSQLa As String
Dim strCP09 As String 'Add By Sindy 2010/6/10
Dim strCP44 As String 'Add By Sindy 2010/11/5
Dim bolSysDt As Boolean 'Add By Sindy 2010/12/28

'Add By Cheng 2002/11/06
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
   
   ' 更新案件進度檔
    'Modify By Cheng 2002/11/06
'   OnUpdateCaseProgress
   If OnUpdateCaseProgress = False Then GoTo ErrorHandler
    'Add By Cheng 2003/03/07
    '若有輸入移轉申請人
    If Me.textCP56.Text <> "" Then
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT"
'edit by nickc 2006/01/26
'            StrSQLa = "Update TradeMark Set TM23='" & ChangeCustomerL(Me.textCP56.Text) & "' " & _
                            " ,TM24='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "1")) & "' " & _
                            " ,TM25='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "2")) & "' " & _
                            " ,TM26='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "3")) & "' " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            'Modify By Sindy 2011/2/22
'            StrSQLa = "Update TradeMark Set TM23='" & ChangeCustomerL(Me.textCP56.Text) & "' " & _
'                            " ,TM24='" & ChgSQL(Pub_RplCu112(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "1"), m_CU112)) & "' " & _
'                            " ,TM25='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "2")) & "' " & _
'                            " ,TM26='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "3")) & "' " & _
'                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            StrSQLa = "Update TradeMark Set TM23='" & ChangeCustomerL(Me.textCP56.Text) & "' " & _
                            " ,TM24='" & ChgSQL(Pub_RplCu112(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "1"), m_CU112, ChangeCustomerL(Me.textCP56.Text))) & "' " & _
                            " ,TM25='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "2")) & "' " & _
                            " ,TM26='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "3")) & "' " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP08='" & ChangeCustomerL(Me.textCP56.Text) & "' " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
    End If
    'add by nickc 2007/01/08
    If Me.textCP89.Text <> "" Then
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM78='" & ChangeCustomerL(Me.textCP89.Text) & "' " & _
                            " ,TM82='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP89.Text), "1")) & "' " & _
                            " ,TM86='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP89.Text), "2")) & "' " & _
                            " ,TM90='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP89.Text), "3")) & "' " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP58='" & ChangeCustomerL(Me.textCP89.Text) & "' " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
   '2009/12/4 add by sonia T-158071
   Else
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM78=NULL,TM82=NULL,TM86=NULL,TM90=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP58=NULL " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
    '2009/12/4 END
    End If
    
    If Me.textCP90.Text <> "" Then
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM79='" & ChangeCustomerL(Me.textCP90.Text) & "' " & _
                            " ,TM83='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP90.Text), "1")) & "' " & _
                            " ,TM87='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP90.Text), "2")) & "' " & _
                            " ,TM91='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP90.Text), "3")) & "' " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP59='" & ChangeCustomerL(Me.textCP90.Text) & "' " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
   '2009/12/4 add by sonia T-158071
   Else
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM79=NULL,TM83=NULL,TM87=NULL,TM91=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP59=NULL " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
    '2009/12/4 END
    End If
    
    If Me.textCP91.Text <> "" Then
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM80='" & ChangeCustomerL(Me.textCP91.Text) & "' " & _
                            " ,TM84='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP91.Text), "1")) & "' " & _
                            " ,TM88='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP91.Text), "2")) & "' " & _
                            " ,TM92='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP91.Text), "3")) & "' " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP65='" & ChangeCustomerL(Me.textCP91.Text) & "' " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
   '2009/12/4 add by sonia T-158071
   Else
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM80=NULL,TM84=NULL,TM88=NULL,TM92=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP65=NULL " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
    '2009/12/4 END
    End If
    
    If Me.textCP92.Text <> "" Then
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM81='" & ChangeCustomerL(Me.textCP92.Text) & "' " & _
                            " ,TM85='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP92.Text), "1")) & "' " & _
                            " ,TM89='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP92.Text), "2")) & "' " & _
                            " ,TM93='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP92.Text), "3")) & "' " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP66='" & ChangeCustomerL(Me.textCP92.Text) & "' " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
   '2009/12/4 add by sonia T-158071
   Else
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM81=NULL,TM85=NULL,TM89=NULL,TM93=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP66=NULL " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
    '2009/12/4 END
    End If
    
   '2009/8/21 add by sonia 馬德里母案移轉子案也要改申請人
   If m_TM10 = "238" And m_CP10 = 移轉 Then
      strSql = "UPDATE TRADEMARK SET TM23=(SELECT TM23 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM24=(SELECT TM24 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM25=(SELECT TM25 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM26=(SELECT TM26 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM78=(SELECT TM78 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM79=(SELECT TM79 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM80=(SELECT TM80 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM81=(SELECT TM81 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM82=(SELECT TM82 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM83=(SELECT TM83 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM84=(SELECT TM84 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM85=(SELECT TM85 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM86=(SELECT TM86 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM87=(SELECT TM87 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM88=(SELECT TM88 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM89=(SELECT TM89 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM90=(SELECT TM90 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM91=(SELECT TM91 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM92=(SELECT TM92 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "'), " & _
                                    "TM93=(SELECT TM93 FROM TRADEMARK WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03='" & m_TM03 & "' AND TM04='" & m_TM04 & "') " & _
                  "WHERE TM01='" & m_TM01 & "' AND TM02='" & m_TM02 & "' AND TM03<>'0' AND TM04<>'00'"
      cnnConnection.Execute strSql
   End If
   '2009/8/21 END
   
    'add by nickc 2006/11/17
    Select Case m_TM01
    Case "T", "TF", "FCT"
        If textPrint <> "N" Then
            strSql = "Update TradeMark Set TM77='" & textPrint & "' Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute strSql
        End If
    Case Else
        If textPrint <> "N" Then
            strSql = "Update ServicePractice Set SP72='" & textPrint & "' " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute strSql
        End If
    End Select
   
   
   ' 若有輸入催審期限時, 新增一筆催審的記錄到下一程序檔
   If IsEmptyText(textUargeDate) = False Then
      strNP07 = "305"
      strNP22 = GetNextProgressNo()
        'Modify By Cheng 2002/12/12
        '下一程序智權人員應掛承辦人員非程序人員
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                        DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strUserNum & "'," & strNP22 & ")"
      'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                        DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & m_CP14 & "'," & strNP22 & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                        PUB_GetWorkDay1(textUargeDate, True) & "," & DBDATE(textUargeDate) & ",'" & m_CP14 & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 收達不印接洽結案單
'      '92.6.8 SONIA 加 言詞辯論, 準備程序
      Select Case strNP07
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
         Case "102", "105", "702", "708", "305", "998", "997"
         Case Else:
            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   End If
   
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
         
         ' 延展, 使用宣誓, 刊登廣告, 繳年費, 收達不印接洽結案單
'         '92.6.8 SONIA 加 言詞辯論, 準備程序
         Select Case strNP07
'            Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
            Case "102", "105", "702", "708", "305", "998", "997"
            Case Else:
               ' 列印國內案件接洽及結案記錄單
               g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
                'Add By Cheng 2004/04/08
                '新增列印接洽結案單資料
                pub_AddressListSN = pub_AddressListSN + 1
                PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
         End Select
      End If
   End If
   rsTmp.Close
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若該筆記錄是母案時, 同時對所有的子案做新增案件進度檔的工作
   If m_TM01 = "TF" And m_TM03 = "0" And m_TM04 = "00" Then
      Set objCopyCP = New ClsCopyCP
    'Modify By Cheng 2002/11/06
'      objCopyCP.CopyCaseProgress m_CP09
      If objCopyCP.CopyCaseProgress(m_CP09) = False Then GoTo ErrorHandler
      Set objCopyCP = Nothing
   End If
   
   
   'add by nick 2004/08/12 更新實際發文規費
    strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
    cnnConnection.Execute strSql
    
   'add by nick 2004/09/27 存公司負責人英文名稱
   'edit by nick 2004/10/07
   'If m_CU103 <> "" And m_TM01 <> "FCT" Then
   'edit by nickc 2006/01/20
   'If (m_CU103 <> "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) <> "") And m_TM01 <> "FCT" Then
   'edit by nickc 2007/08/10
   'If (m_CU103 <> "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) <> "" Or m_CU112 <> "") And m_TM01 <> "FCT" Then
   'Modify By Sindy 2012/10/31 +SeekCu10(1),SeekCu10(2),SeekCu10(3),SeekCu10(4),SeekCu10(5)
   If (SeekCu103(1) <> "" Or (SeekCu05(1) & SeekCu88(1) & SeekCu89(1) & SeekCu90(1)) <> "" Or SeekCu112(1) <> "" Or (SeekCu39(1) & SeekCu40(1) & SeekCu41(1)) <> "" Or SeekCu10(1) <> "") And m_TM01 <> "FCT" Then
            'edit by nickc 2005/11/22 應更新移轉申請人
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "'  Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' "
            'edit by nickc 2006/01/20
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "'  Where Cu01 = '" & Mid(ChangeCustomerL(textCP56), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP56), 9, 1) & "' "
            'edit by nickc 2007/08/10
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "',cu112='" & ChgSQL(m_CU112) & "'  Where Cu01 = '" & Mid(ChangeCustomerL(textCP56), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP56), 9, 1) & "' "
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(1)) & "',cu05='" & ChgSQL(SeekCu05(1)) & "',cu88='" & ChgSQL(SeekCu88(1)) & "',cu89='" & ChgSQL(SeekCu89(1)) & "',cu90='" & ChgSQL(SeekCu90(1)) & "',cu112='" & ChgSQL(SeekCu112(1)) & "',cu39='" & ChgSQL(SeekCu39(1)) & "',cu40='" & ChgSQL(SeekCu40(1)) & "',cu41='" & ChgSQL(SeekCu41(1)) & "',cu10='" & ChgSQL(SeekCu10(1)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textCP56.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP56.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(1)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textCP56.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP56.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   'add by nickc 2007/08/10 加多申請人也要
   If (SeekCu103(2) <> "" Or (SeekCu05(2) & SeekCu88(2) & SeekCu89(2) & SeekCu90(2)) <> "" Or SeekCu112(2) <> "" Or (SeekCu39(2) & SeekCu40(2) & SeekCu41(2)) <> "" Or SeekCu10(2) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(2)) & "',cu05='" & ChgSQL(SeekCu05(2)) & "',cu88='" & ChgSQL(SeekCu88(2)) & "',cu89='" & ChgSQL(SeekCu89(2)) & "',cu90='" & ChgSQL(SeekCu90(2)) & "',cu112='" & ChgSQL(SeekCu112(2)) & "',cu39='" & ChgSQL(SeekCu39(2)) & "',cu40='" & ChgSQL(SeekCu40(2)) & "',cu41='" & ChgSQL(SeekCu41(2)) & "',cu10='" & ChgSQL(SeekCu10(2)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textCP89.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP89.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(2)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textCP89.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP89.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(3) <> "" Or (SeekCu05(3) & SeekCu88(3) & SeekCu89(3) & SeekCu90(3)) <> "" Or SeekCu112(3) <> "" Or (SeekCu39(3) & SeekCu40(3) & SeekCu41(3)) <> "" Or SeekCu10(3) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(3)) & "',cu05='" & ChgSQL(SeekCu05(3)) & "',cu88='" & ChgSQL(SeekCu88(3)) & "',cu89='" & ChgSQL(SeekCu89(3)) & "',cu90='" & ChgSQL(SeekCu90(3)) & "',cu112='" & ChgSQL(SeekCu112(3)) & "',cu39='" & ChgSQL(SeekCu39(3)) & "',cu40='" & ChgSQL(SeekCu40(3)) & "',cu41='" & ChgSQL(SeekCu41(3)) & "',cu10='" & ChgSQL(SeekCu10(3)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textCP90.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP90.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(3)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textCP90.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP90.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(4) <> "" Or (SeekCu05(4) & SeekCu88(4) & SeekCu89(4) & SeekCu90(4)) <> "" Or SeekCu112(4) <> "" Or (SeekCu39(4) & SeekCu40(4) & SeekCu41(4)) <> "" Or SeekCu10(4) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(4)) & "',cu05='" & ChgSQL(SeekCu05(4)) & "',cu88='" & ChgSQL(SeekCu88(4)) & "',cu89='" & ChgSQL(SeekCu89(4)) & "',cu90='" & ChgSQL(SeekCu90(4)) & "',cu112='" & ChgSQL(SeekCu112(4)) & "',cu39='" & ChgSQL(SeekCu39(4)) & "',cu40='" & ChgSQL(SeekCu40(4)) & "',cu41='" & ChgSQL(SeekCu41(4)) & "',cu10='" & ChgSQL(SeekCu10(4)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textCP91.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP91.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(4)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textCP91.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP91.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(5) <> "" Or (SeekCu05(5) & SeekCu88(5) & SeekCu89(5) & SeekCu90(5)) <> "" Or SeekCu112(5) <> "" Or (SeekCu39(5) & SeekCu40(5) & SeekCu41(5)) <> "" Or SeekCu10(5) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(5)) & "',cu05='" & ChgSQL(SeekCu05(5)) & "',cu88='" & ChgSQL(SeekCu88(5)) & "',cu89='" & ChgSQL(SeekCu89(5)) & "',cu90='" & ChgSQL(SeekCu90(5)) & "',cu112='" & ChgSQL(SeekCu112(5)) & "',cu39='" & ChgSQL(SeekCu39(5)) & "',cu40='" & ChgSQL(SeekCu40(5)) & "',cu41='" & ChgSQL(SeekCu41(5)) & "',cu10='" & ChgSQL(SeekCu10(5)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textCP92.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP92.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(5)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textCP92.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP92.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   'add by nickc 2007/01/30 移轉時，要更新下一程序的智權人員
   '2013/4/18 modify by sonia 改用共用函數
   'If rsTmp.State = 1 Then rsTmp.Close
   'strSql = "SELECT * FROM Customer Where Cu01 = '" & Mid(ChangeCustomerL(textCP56), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP56), 9, 1) & "' "
   'rsTmp.CursorLocation = adUseClient
   'rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'If rsTmp.RecordCount > 0 Then
   '   pub_ChgSalesTargetIsNp m_TM01, m_TM02, m_TM03, m_TM04, CheckStr(rsTmp.Fields("cu13"))
   'End If
   pub_ChgSalesTargetIsNp m_TM01, m_TM02, m_TM03, m_TM04, PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
   '2013/4/18 end
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 列印定稿
   'Modify By Cheng 2002/06/17
'   If textPrint <> "N" Then
'      PrintLetter
'   End If
'Modify By Cheng 2002/11/11
'   Select Case m_TM01
'   Case "TB"
'      If textPrint <> "N" Then
'         If Len(Me.textPrintMan.Text) > 0 Then
'            For ii = 1 To Len(Me.textPrintMan.Text)
'               PrintLetterTB (Mid(Me.textPrintMan.Text, ii, 1))
'            Next ii
'         End If
'      End If
'   Case Else
'      If textPrint <> "N" Then
'         If Len(Me.textPrintMan.Text) > 0 Then
'            For ii = 1 To Len(Me.textPrintMan.Text)
'               PrintLetter (Mid(Me.textPrintMan.Text, ii, 1))
'            Next ii
'         End If
'      End If
'   End Select
   
'   'Add By Sindy 2010/6/10
'   '由巨京代理之移轉案，發文移轉時，請同時假收文變更代理人
'   '變更案的收文與發文時間同移轉案
'   '前次發文代理人不是Y52269
'   strSql = "SELECT CP44, Max(CP27||CP09) FROM CaseProgress " & _
'                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
'                        "CP02 = '" & m_TM02 & "' AND " & _
'                        "CP03 = '" & m_TM03 & "' AND " & _
'                        "CP04 = '" & m_TM04 & "' AND " & _
'                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null And CP27 Is Not Null Group By CP44 Order By 2 Desc, 1 "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 0 Then strCP44 = ""
'   If intI = 1 Then strCP44 = Trim(RsTemp.Fields("CP44"))
'   If Left(strCP44, 6) <> "Y52269" And Left(Trim(textCP44.Text), 6) = "Y52269" Then
'      If m_TM01 = "T" And m_TM10 = "020" And m_CP10 = "501" Then
'         strCP09 = AutoNo("B", 6)
'         '新增一筆B類
'         strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp43,cp44,cp45,cp64) " & _
'                        "values (" & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & _
'                        "," & CNULL(m_TM04) & "," & CNULL(DBDATE(textCP27)) & "," & CNULL(strCP09) & ",301," & _
'                        CNULL(m_CP12) & "," & CNULL(m_CP13) & "," & CNULL(strUserNum) & ",'N','N'," & CNULL(DBDATE(textCP27)) & ",'N'," & _
'                        CNULL(m_CP09) & ",'" & textCP44.Text & "','" & strCP45 & "','變更代理人')"
'         cnnConnection.Execute strSql
'         '新增變更事項檔
'         strSql = "insert into ChangeEvent(CE01,CE55) values('" & strCP09 & "','V')"
'         cnnConnection.Execute strSql
'      End If
'   End If
   'Modify By Sindy 2012/3/23
   Call PUB_T020InsB301(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textCP44, m_TM10, m_CP10, textCP27, m_CP12, m_CP13, textTM45)
   
   'Add by Amy 2020/01/13
   If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
        strSql = ""
        If textCP118 = "Y" And Val(textCP84) > 0 Then
           If txtPayToday <> "" Then
              strSql = ",CP118 = 'A' "
              If txtPayToday = "Y" Then
                  strSql = strSql & ",CP152 = " & CompWorkDay(2, DBDATE(textCP27))
              Else
                  strSql = strSql & ",CP152 =" & CompWorkDay(3, DBDATE(textCP27))
              End If
              strSql = "Update CaseProgress Set " & Mid(strSql, 2) & " Where CP09 = '" & m_CP09 & "' "
              cnnConnection.Execute strSql
           End If
        End If
   End If
   'end 2020/01/13
   
   'Add By Sindy 2011/3/9 若為電子送件則自動設定為不經發文室
   '以防動作為重新發文, 所以一併把發文室相關欄位清空
   If textCP118.Visible = True And textCP118 = "Y" Then
      strSql = "Update CaseProgress Set CP123=null" & _
                                                          ",CP124=null" & _
                                                          ",CP125=null" & _
                                                          ",CP28=null" & _
                                                          ",CP131=null" & _
                                                          ",CP132=null" & _
                   " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Add by Sindy 98/3/24
   If m_TM10 = "000" Then
      'Modify By Sindy 2009/04/24
      'PUB_UpdateDispatch m_CP09s, m_CP123s
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
   End If
   
   'Add by Sindy 2012/10/4 外->台,智權人員是葉雪貞及巨京,發文規費和收文規費不相同時,系統自動更改進度檔內規費費用及計算點數
   'Modified by Lydia 2015/10/16 + m_CP84
   Call PUB_TSendUpdateCP1718(m_CP09, textCP84, textPrint, m_TM10, m_CP13, m_CP84)
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'ADD BY SONIA 2015/11/24 若申請為旅狐國際及部分關係企業者，案件備註若無"不銷卷"字樣,則要加入
   If (textCP56 <> "" And InStr(strTmTRAVEL_FOXCust, Left(textCP56, 8)) > 0) Or _
      (textCP89 <> "" And InStr(strTmTRAVEL_FOXCust, Left(textCP89, 8)) > 0) Or _
      (textCP90 <> "" And InStr(strTmTRAVEL_FOXCust, Left(textCP90, 8)) > 0) Or _
      (textCP91 <> "" And InStr(strTmTRAVEL_FOXCust, Left(textCP91, 8)) > 0) Or _
      (textCP92 <> "" And InStr(strTmTRAVEL_FOXCust, Left(textCP92, 8)) > 0) Then
      strSql = "update trademark" & _
               " set tm58=decode(tm58,null,'" & ChangeTStringToTDateString(strSrvDate(2)) & "不銷卷','" & ChangeTStringToTDateString(strSrvDate(2)) & "不銷卷,'||tm58)" & _
               " Where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "'" & _
               " and (instr(tm58,'不銷卷')=0 or tm58 is null)"
      cnnConnection.Execute strSql
   End If
   'END 2015/11/24
   
   'Add By Sindy 2019/12/20 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = m_CP09
      'Added by Morgan 2025/10/15
      '定稿對象選2時要傳入受讓人編號，系統會以此判斷是否E化客戶(出申請人定稿不會觸發UpdateNoneApplLtr) Ex:T-210494
      '原程式可同時輸入12但User並沒不知道可以這樣用(桂英)，本次改為只能輸1碼，因同時出兩分定稿則要有2筆寄發文件的紀錄，需再研究。
      If textPrintMan = "2" Then
         PUB_AddLetterProgress strLD18, 0, IIf(textPrint = "N", False, True), "", False, textCP56, m_CP10, m_TM44
      Else
      'end 2025/10/15
         PUB_AddLetterProgress strLD18, 0, IIf(textPrint = "N", False, True), "", False, m_TM23, m_CP10, m_TM44
      End If
      'end 2025/10/15
   End If
   '2019/12/20 END
   Call PUB_UpdateLP19_T(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textCP27) 'Add by Sindy 2020/2/12 收據/回執設定
   
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
   
   Set rsTmp = Nothing
'Add By Cheng 2002/11/03
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

Private Sub textPrintMan_GotFocus()
TextInverse Me.textPrintMan
End Sub

Private Sub textPrintMan_KeyPress(KeyAscii As Integer)
'Add By Cheng 2002/06/17
KeyAscii = UpperCase(KeyAscii)
If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
End Sub

Private Sub textPrintMan_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPrintMan) = False Then
      'Modify By Cheng 2002/06/17
'      Select Case textPrintMan
'         Case "1", "2", "3":
'         Case Else:
'            Cancel = True
'            strTit = "檢核資料"
'            strMsg = "定稿對象只可輸入1或2或3"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textPrintMan_GotFocus
'      End Select
   
      'Add By Cheng 2002/06/17
      With Me.textPrintMan
         Select Case Len(.Text)
         Case 2
            If Mid(.Text, 1, 1) = Mid(.Text, 2, 1) Then
               MsgBox "定稿對象重覆輸入!!!", vbExclamation + vbOKOnly
               Cancel = True
               textPrintMan_GotFocus
            End If
         End Select
      End With
   End If
End Sub

' 不一併移轉註冊號數
Private Sub textTM15_S_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nCount As Integer
Dim nIndex As Integer
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
'Add By Cheng 2003/08/12
Dim strOurCase As String
Dim strNotOurCase As String

   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textTM15_S) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textTM15_S)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM15_S, nIndex)
        'Modify By Cheng 2003/08/12
'      strSQL = "SELECT * FROM TradeMark " & _
'               "WHERE TM15 = '" & textTM15_S & "' "
      strSql = "SELECT * FROM TradeMark " & _
               "WHERE TM15 = '" & strTemp & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
'         rsTmp.Close
'         Cancel = True
'         strTit = "檢核資料"
'         strMsg = "註冊號數<" & strTemp & ">不存在"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textTM15_S_GotFocus
'         GoTo EXITSUB
        strNotOurCase = strNotOurCase & strTemp & ","
      Else
        strOurCase = strOurCase & strTemp & ","
      End If
      rsTmp.Close
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM15_S, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTM15_S, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "註冊號數<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM15_S_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
    'Add By Cheng 2003/08/12
    If strOurCase <> "" Then strOurCase = "本所案件：" & Left(strOurCase, Len(strOurCase) - 1)
    If strNotOurCase <> "" Then strNotOurCase = vbCrLf & "非本所案件：" & Left(strNotOurCase, Len(strNotOurCase) - 1)
    If MsgBox(strOurCase & strNotOurCase, vbExclamation + vbOKCancel) = vbCancel Then
        Cancel = True
        textTM15_S_GotFocus
        GoTo EXITSUB
    End If
           
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 催審期限
Private Sub textUargeDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textUargeDate) = False Then
      If CheckIsTaiwanDate(textUargeDate, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "催審期限日期不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textUargeDate_GotFocus
      End If
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/23檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
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
   
   'Add By Sindy 2012/4/17
   If m_blnClkChgButton = False Then
      MsgBox "請輸入變更事項!!!", vbExclamation + vbOKOnly
      Me.cmdMod.SetFocus
      GoTo EXITSUB
   End If
   
   ' 發文日不可空白
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' 申請國家非台灣時代理人不可空白
   If m_TM10 >= "010" Then
      If IsEmptyText(textCP44) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入代理人"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 受讓人不可空白
   If IsEmptyText(textCP56) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入受讓人"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP56.SetFocus
      GoTo EXITSUB
   End If
   'Add By Cheng 2002/06/17
   ' 定稿對象
   If IsEmptyText(textPrintMan) = True Then
      If textPrint <> "N" Then
         strTit = "檢核資料"
         strMsg = "有列印定稿時, 定稿對象不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPrintMan.SetFocus
         GoTo EXITSUB
      End If
   End If
   
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
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textUargeDate_GotFocus()
   InverseTextBox textUargeDate
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textTM15_S_GotFocus()
   InverseTextBox textTM15_S
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

Private Sub textCP56_GotFocus()
   InverseTextBox textCP56
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
End Sub

'Add by Amy 2020/01/13
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
'end 2020/01/13


' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField(strPrintMan As String)
'strPrintMan : 列印對象
   Dim strTM23Nation As String
   Dim strSql As String
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 案件性質為移轉
   If m_CP10 = "501" Then
   'add by nickc 2008/01/04  因為要加入  TF 系統別
   If m_TM01 = "T" Then
      ' 申請國家為台灣
      If m_TM10 < "010" Then
         ' 申請人國籍為台灣
         'edit by nickc 2006/06/30
         'If strTM23Nation < "010" Then
         If textPrint = "1" Then
            If strPrintMan = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "02", strUserNum
               ' 回音
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
                        "'回音','" & "大約" & textCF09 & "後可接獲回音。')"
               '         "'" & "回音" & "','" & textCF09 & "')"
               cnnConnection.Execute strSql
            ElseIf strPrintMan = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "34", strUserNum
               ' 回音
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "34" & "','" & strUserNum & "'," & _
                        "'回音','" & "大約" & textCF09 & "後可接獲回音。')"
               '         "'" & "回音" & "','" & textCF09 & "')"
               cnnConnection.Execute strSql
            End If
         ' 申請人國籍非台灣
         'edit by nickc 2006/06/30
         'Else
         ElseIf textPrint = "2" Then
            If strPrintMan = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "07", strUserNum
            ElseIf strPrintMan = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "37", strUserNum
            End If
         End If
      ' 申請國家為大陸
      ElseIf m_TM10 = "020" Then
        'Modify By Cheng 2003/01/08
'         If strPrintMan = "1" Then
'            ' 清除定稿例外欄位檔原有資料
'            EndLetter "01", m_CP09, "00", strUserNum
'            ' 案件性質分類
'            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
'                     "'" & "案件性質分類" & "','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
'            cnnConnection.Execute strSQL
'            ' 回音
'            'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            '         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
'            '         "'" & "回音" & "','" & textCF09 & "')"
'            'cnnConnection.Execute strSQL
'         ElseIf strPrintMan = "2" Then
'            ' 清除定稿例外欄位檔原有資料
'            EndLetter "01", m_CP09, "33", strUserNum
'            ' 案件性質分類
'            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "33" & "','" & strUserNum & "'," & _
'                     "'" & "案件性質分類" & "','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
'            cnnConnection.Execute strSQL
'            ' 回音
'            'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            '         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
'            '         "'" & "回音" & "','" & textCF09 & "')"
'            'cnnConnection.Execute strSQL
'         End If
            '2005/3/29 modify by sonia
            '' 清除定稿例外欄位檔原有資料
            'EndLetter "01", m_CP09, "06", strUserNum
            '' 案件性質分類
            'StrSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            '         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
            '         "'" & "案件性質分類" & "','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
            'cnnConnection.Execute StrSql
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                If strPrintMan = "1" Then
                   EndLetter "01", m_CP09, "06", strUserNum
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "01" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                            "'" & "案件性質分類" & "','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
                   cnnConnection.Execute strSql
                ElseIf strPrintMan = "2" Then
                   EndLetter "01", m_CP09, "36", strUserNum
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "01" & "','" & m_CP09 & "','" & "36" & "','" & strUserNum & "'," & _
                            "'" & "案件性質分類" & "','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
                   cnnConnection.Execute strSql
                End If
            End If
            '2005/3/29 end
      End If
   ElseIf m_TM01 = "TF" Then
        EndLetter "01", m_CP09, "01", strUserNum
   End If
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter(strPrintMan As String)
'Add By Sindy 2012/1/12
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/12 End

'strPrintMan : 列印對象
   Dim strTM23Nation As String
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField strPrintMan
   
   'Add By Sindy 2012/1/12
   ET01 = "01"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/12 End
   
   ' 案件性質為移轉
   If m_CP10 = "501" Then
      'add by nickc 2008/01/04  因為要加入  TF 系統別
      If m_TM01 = "T" Then
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 列印定稿
               'Modify By Cheng 2002/04/17
   '            NowPrint m_CP09, "01", "02", False, strUserNum, 0
               If strPrintMan = "1" Then
   '               NowPrint m_CP09, "01", "02", False, strUserNum, 0
                  ET03 = "02" 'Modify By Sindy 2012/1/12
               ElseIf strPrintMan = "2" Then
   '               NowPrint m_CP09, "01", "34", False, strUserNum, 0
                  ET03 = "34" 'Modify By Sindy 2012/1/12
               End If
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            ElseIf textPrint = "2" Then
               ' 列印定稿
               'Modify Cheng 2002/06/17
   '            NowPrint m_CP09, "01", "07", False, strUserNum, 0
               If strPrintMan = "1" Then
   '               NowPrint m_CP09, "01", "07", False, strUserNum, 0
                  ET03 = "07" 'Modify By Sindy 2012/1/12
               ElseIf strPrintMan = "2" Then
   '               NowPrint m_CP09, "01", "37", False, strUserNum, 0
                  ET03 = "37" 'Modify By Sindy 2012/1/12
               End If
            End If
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            ' 列印定稿
               'Modify By Cheng 2003/01/08
   '         'Modify By Cheng 2002/06/17
   ''         NowPrint m_CP09, "01", "00", False, strUserNum, 0
   '         If strPrintMan = "1" Then
   '            NowPrint m_CP09, "01", "00", False, strUserNum, 0
   '         ElseIf strPrintMan = "2" Then
   '            NowPrint m_CP09, "01", "33", False, strUserNum, 0
   '         End If
               '2005/3/8 modify by sonia
               'NowPrint m_CP09, "01", "06", False, strUserNum, 0
               'add by nickc 2006/06/30
               If textPrint = "1" Then
                   If strPrintMan = "1" Then
   '                   NowPrint m_CP09, "01", "06", False, strUserNum, 0
                     ET03 = "06" 'Modify By Sindy 2012/1/12
                   ElseIf strPrintMan = "2" Then
   '                   NowPrint m_CP09, "01", "36", False, strUserNum, 0
                     ET03 = "36" 'Modify By Sindy 2012/1/12
                   End If
               End If
               '2005/3/8 end
         End If
      ElseIf m_TM01 = "TF" Then
   '        NowPrint m_CP09, "01", "01", False, strUserNum, 0
         ET03 = "01" 'Modify By Sindy 2012/1/12
      End If
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
         'Add By Sindy 2019/12/20 + strLD18.信函總收文號
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

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpFieldTB(strPrintMan As String)
'strPrintMan : 列印對象
   Dim strTM23Nation As String
   Dim strSql As String
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   Select Case m_CP10
'      ' 條碼申請
'      Case "802":
'         ' 清除定稿例外欄位檔原有資料
'         EndLetter "01", m_CP09, "28", strUserNum
'         ' 回音
'         strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  "VALUES ('" & "01" & "','" & m_CP09 & "','" & "28" & "','" & strUserNum & "'," & _
'                  "'" & "回音" & "','" & textCF09 & "')"
'         cnnConnection.Execute strSQL
      ' 移轉
      Case "501":
         'add by nickc 2006/06/30
         If textPrint = "1" Then
            ' 清除定稿例外欄位檔原有資料
            If strPrintMan = "1" Then
               EndLetter "01", m_CP09, "29", strUserNum
            ElseIf strPrintMan = "2" Then
               EndLetter "01", m_CP09, "38", strUserNum
            End If
        End If
'      ' 繳年費
'      Case "708":
'         ' 清除定稿例外欄位檔原有資料
'         EndLetter "01", m_CP09, "32", strUserNum
   End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetterTB(strPrintMan As String)
'Add By Sindy 2012/1/12
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/12 End
   
'strPrintMan :
   Dim strTM23Nation As String
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpFieldTB strPrintMan
   
   'Add By Sindy 2012/1/12
   ET01 = "01"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/12 End
   
   Select Case m_CP10
'      ' 條碼申請
'      Case "802":
'         ' 列印定稿
'         'Modify By Cheng 2002/06/14
''         NowPrint m_CP09, "01", "28", False, strUserNum, 0
'         NowPrint m_CP09, "01", "28", IIf(textWord.Text = "Y", True, False), strUserNum, 0
      ' 移轉
      Case "501":
         ' 列印定稿
'         NowPrint m_CP09, "01", "29", False, strUserNum, 0
         'add by nickc 2006/06/30
         If textPrint = "1" Then
            If strPrintMan = "1" Then
'               NowPrint m_CP09, "01", "29", False, strUserNum, 0
               ET03 = "29" 'Modify By Sindy 2012/1/12
            ElseIf strPrintMan = "2" Then
'               NowPrint m_CP09, "01", "38", False, strUserNum, 0
               ET03 = "38" 'Modify By Sindy 2012/1/12
            End If
         End If
'      ' 繳年費
'      Case "708":
'         ' 列印定稿
''         NowPrint m_CP09, "01", "32", False, strUserNum, 0
'         NowPrint m_CP09, "01", "32", IIf(textWord.Text = "Y", True, False), strUserNum, 0
   End Select
   
   'Add By Sindy 2012/1/12
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
      Else
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0
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
   'add by nick 2004/08/12 發文規費，申請國家台灣才檢查
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
   'edit by nickc 2006/01/27
   'If Me.textCP22.Enabled = True Then
   '   Cancel = False
   '   textCP22_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
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
   
   If Me.textCP56.Enabled = True Then
      Cancel = False
      textCP56_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'add by nickc 2007/01/08
   If Me.textCP89.Enabled = True Then
      Cancel = False
      textCP89_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP90.Enabled = True Then
      Cancel = False
      textCP90_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP91.Enabled = True Then
      Cancel = False
      textCP91_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP92.Enabled = True Then
      Cancel = False
      textCP92_Validate Cancel
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
   
   If Me.textPrint.Enabled = True Then
      Cancel = False
      textPrint_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textPrintMan.Enabled = True Then
      Cancel = False
      textPrintMan_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM15_S.Enabled = True Then
      Cancel = False
      textTM15_S_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textUargeDate.Enabled = True Then
      Cancel = False
      textUargeDate_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'add by nickc 2006/03/07
   'Modify By Sindy 2012/7/26
   'If lstNameAgent.Enabled = True Then
   If lstNameAgent.Visible = True Then
   '2012/7/26 End
       Cancel = False
       lstNameAgent_Validate Cancel
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

'edit by nickc 2006/01/27
' 91.09.02 modify by louis
'Private Sub textAgName_GotFocus()
'   InverseTextBox textAgName
'   textAgName.IMEMode = 1
'End Sub
'
'' 91.09.02 modify by louis
'' 本所出名代理人
'Private Sub textAgName_Validate(Cancel As Boolean)
'   Cancel = False
'   If CheckLengthIsOK(textAgName, 10) = False Then
'      Cancel = True
'   End If
'   If Cancel = False Then: textAgName.IMEMode = 2
'End Sub

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
