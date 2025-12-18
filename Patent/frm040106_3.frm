VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040106_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外指示信"
   ClientHeight    =   5760
   ClientLeft      =   36
   ClientTop       =   948
   ClientWidth     =   9312
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9312
   Begin VB.CommandButton Command2 
      Caption         =   "補件期限"
      Height          =   375
      Index           =   2
      Left            =   4500
      TabIndex        =   77
      Top             =   0
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4455
      Left            =   90
      TabIndex        =   52
      Top             =   1230
      Width           =   9105
      _ExtentX        =   16066
      _ExtentY        =   7853
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   420
      TabCaption(0)   =   "一般資料"
      TabPicture(0)   =   "frm040106_3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(9)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label16"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label20(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label20(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label20(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label20(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label30"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label31"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label32"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label3(9)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label3(10)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label3(11)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label3(12)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label3(13)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label33"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label34"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(15)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(16)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text7(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text7(2)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text7(9)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text7(10)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text7(14)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text7(15)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text7(16)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text7(4)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text7(6)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text7(8)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text7(5)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text7(7)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text7(17)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text7(18)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text7(0)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Check1(5)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Combo2"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Check1(4)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Check1(0)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Check1(1)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Check1(2)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Check1(3)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Text6"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Check2"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Check1(6)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Check1(7)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Check1(8)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtDate"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Check3"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "chkFix(0)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "chkFix(1)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "Check4"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).ControlCount=   58
      Begin VB.CheckBox Check4 
         Caption         =   "台灣案同日提申"
         Height          =   255
         Left            =   4350
         TabIndex        =   83
         Top             =   743
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CheckBox chkFix 
         Caption         =   "主動修正"
         Height          =   225
         Index           =   1
         Left            =   6435
         TabIndex        =   17
         Top             =   1800
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CheckBox chkFix 
         Caption         =   "國際階段修正"
         Height          =   225
         Index           =   0
         Left            =   4905
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  '靠右對齊
         Caption         =   "是否順稿:"
         Height          =   225
         Left            =   5760
         TabIndex        =   81
         Top             =   1050
         Width           =   1140
      End
      Begin VB.TextBox txtDate 
         Height          =   270
         Left            =   8010
         MaxLength       =   8
         TabIndex        =   80
         Top             =   1020
         Width           =   885
      End
      Begin VB.CheckBox Check1 
         Caption         =   "簡要說明一份"
         Height          =   255
         Index           =   8
         Left            =   7560
         TabIndex        =   79
         Top             =   1290
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "檢索報告一份"
         Height          =   225
         Index           =   7
         Left            =   2520
         TabIndex        =   13
         Top             =   1530
         Width           =   1860
      End
      Begin VB.CheckBox Check1 
         Caption         =   "PCT公開文本"
         Height          =   225
         Index           =   6
         Left            =   630
         TabIndex        =   12
         Top             =   1530
         Width           =   1590
      End
      Begin VB.CheckBox Check2 
         Caption         =   "前案檢索資料"
         Height          =   225
         Left            =   3375
         TabIndex        =   15
         Top             =   1800
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   2010
         TabIndex        =   14
         Top             =   1770
         Width           =   276
      End
      Begin VB.CheckBox Check1 
         Caption         =   "附圖說明一份"
         Height          =   255
         Index           =   3
         Left            =   4905
         TabIndex        =   9
         Top             =   1290
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "附圖一份"
         Height          =   225
         Index           =   2
         Left            =   3825
         TabIndex        =   8
         Top             =   1303
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "說明書一份"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   7
         Top             =   1290
         Width           =   1230
      End
      Begin VB.CheckBox Check1 
         Caption         =   "委托書一份"
         Height          =   255
         Index           =   0
         Left            =   630
         TabIndex        =   6
         Top             =   1290
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         Caption         =   "照片三份"
         Height          =   225
         Index           =   4
         Left            =   6435
         TabIndex        =   10
         Top             =   1303
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   960
         TabIndex        =   0
         Top             =   420
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "優先權證明書乙份"
         Height          =   225
         Index           =   5
         Left            =   4905
         TabIndex        =   11
         Top             =   1530
         Width           =   1860
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   0
         Left            =   3330
         TabIndex        =   2
         Top             =   735
         Width           =   975
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   525
         Index           =   18
         Left            =   4545
         TabIndex        =   27
         Top             =   3825
         Width           =   4425
         VariousPropertyBits=   -1467987941
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "7805;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   525
         Index           =   17
         Left            =   135
         TabIndex        =   26
         Top             =   3825
         Width           =   4275
         VariousPropertyBits=   -1467987941
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "7541;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   7
         Left            =   5460
         TabIndex        =   24
         Top             =   3090
         Width           =   975
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   5
         Left            =   5460
         TabIndex        =   22
         Top             =   2850
         Width           =   975
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   8
         Left            =   960
         TabIndex        =   25
         Top             =   3330
         Width           =   975
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   6
         Left            =   960
         TabIndex        =   23
         Top             =   3090
         Width           =   975
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   4
         Left            =   960
         TabIndex        =   21
         Top             =   2850
         Width           =   975
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   16
         Left            =   1320
         TabIndex        =   20
         Top             =   2550
         Width           =   7635
         VariousPropertyBits=   671107099
         MaxLength       =   160
         Size            =   "13467;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   15
         Left            =   1320
         TabIndex        =   19
         Top             =   2310
         Width           =   7635
         VariousPropertyBits=   671107099
         MaxLength       =   180
         Size            =   "13467;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   14
         Left            =   1320
         TabIndex        =   18
         Top             =   2070
         Width           =   7635
         VariousPropertyBits=   671107099
         MaxLength       =   160
         Size            =   "13467;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   10
         Left            =   4365
         TabIndex        =   5
         Top             =   1020
         Width           =   255
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   9
         Left            =   1485
         TabIndex        =   4
         Top             =   1020
         Width           =   255
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   2
         Left            =   6855
         TabIndex        =   3
         Top             =   750
         Width           =   975
         VariousPropertyBits=   671107097
         MaxLength       =   4
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text7 
         Height          =   300
         Index           =   3
         Left            =   1290
         TabIndex        =   1
         Top             =   735
         Width           =   975
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "順稿期限:"
         Height          =   180
         Index           =   16
         Left            =   7110
         TabIndex        =   82
         Top             =   1065
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "指定提申日:"
         Height          =   180
         Index           =   15
         Left            =   2355
         TabIndex        =   78
         Top             =   780
         Width           =   945
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "案件備註:"
         Height          =   180
         Left            =   4620
         TabIndex        =   76
         Top             =   3630
         Width           =   765
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "進度備註:"
         Height          =   195
         Left            =   135
         TabIndex        =   75
         Top             =   3630
         Width           =   780
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   13
         Left            =   2040
         TabIndex        =   74
         Top             =   3375
         Width           =   2490
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "4392;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   12
         Left            =   6510
         TabIndex        =   73
         Top             =   3120
         Width           =   2400
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "4233;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   11
         Left            =   2040
         TabIndex        =   72
         Top             =   3135
         Width           =   2490
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "4392;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   10
         Left            =   6510
         TabIndex        =   71
         Top             =   2880
         Width           =   2400
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "4233;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   9
         Left            =   2040
         TabIndex        =   70
         Top             =   2880
         Width           =   2490
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "4392;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人4"
         Height          =   180
         Index           =   6
         Left            =   4620
         TabIndex        =   69
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人2"
         Height          =   180
         Index           =   5
         Left            =   4620
         TabIndex        =   68
         Top             =   2880
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人5"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   67
         Top             =   3360
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人3"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   66
         Top             =   3120
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人1"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   65
         Top             =   2880
         Width           =   645
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(日):"
         Height          =   195
         Left            =   135
         TabIndex        =   64
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(英):"
         Height          =   195
         Left            =   135
         TabIndex        =   63
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(中):"
         Height          =   195
         Left            =   135
         TabIndex        =   62
         Top             =   2100
         Width           =   1095
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "是否一併提出實審請求:         (N:不提)"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   53
         Top             =   1800
         Width           =   2940
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "附件:"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   61
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "是否修改指示內容:       (Y:Word)"
         Height          =   180
         Index           =   4
         Left            =   2865
         TabIndex        =   60
         Top             =   1065
         Width           =   2490
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "是否列印指示信:       (N:不印)"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   59
         Top             =   1065
         Width           =   2265
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   5
         Left            =   7905
         TabIndex        =   58
         Top             =   795
         Width           =   1080
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "1905;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "案件性質:"
         Height          =   180
         Left            =   6045
         TabIndex        =   57
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "最終提申期限:"
         Height          =   180
         Index           =   9
         Left            =   135
         TabIndex        =   56
         Top             =   780
         Width           =   1125
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   6
         Left            =   2520
         TabIndex        =   55
         Top             =   480
         Width           =   6285
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "11086;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代理人:"
         Height          =   195
         Left            =   135
         TabIndex        =   54
         Top             =   450
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   1
      Left            =   8175
      TabIndex        =   31
      Top             =   0
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   7365
      TabIndex        =   30
      Top             =   0
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   2
      Left            =   9120
      TabIndex        =   32
      Top             =   6240
      Width           =   972
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2724
      MaxLength       =   2
      TabIndex        =   39
      Top             =   684
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2484
      MaxLength       =   1
      TabIndex        =   38
      Top             =   684
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1644
      MaxLength       =   6
      TabIndex        =   37
      Top             =   684
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1164
      MaxLength       =   3
      TabIndex        =   36
      Top             =   684
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "優先權資料"
      Height          =   375
      Index           =   1
      Left            =   6255
      TabIndex        =   29
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "發明人"
      Height          =   375
      Index           =   0
      Left            =   5505
      TabIndex        =   28
      Top             =   0
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '內實線
      Index           =   1
      X1              =   210
      X2              =   9030
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   90
      X2              =   9210
      Y1              =   1200
      Y2              =   1200
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   8
      Left            =   4410
      TabIndex        =   51
      Top             =   930
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2646;317"
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
      Left            =   3570
      TabIndex        =   50
      Top             =   930
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "法定期限:"
      Height          =   180
      Index           =   1
      Left            =   6090
      TabIndex        =   49
      Top             =   690
      Width           =   765
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   7
      Left            =   7170
      TabIndex        =   48
      Top             =   690
      Width           =   1800
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3175;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   4
      Left            =   7170
      TabIndex        =   47
      Top             =   930
      Width           =   1800
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3175;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   3
      Left            =   7170
      TabIndex        =   46
      Top             =   450
      Width           =   1800
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3175;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   2
      Left            =   4410
      TabIndex        =   45
      Top             =   690
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2646;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   1
      Left            =   4410
      TabIndex        =   44
      Top             =   450
      Width           =   1500
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2646;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   6090
      TabIndex        =   43
      Top             =   930
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "收款後辦案:"
      Height          =   180
      Left            =   6090
      TabIndex        =   42
      Top             =   450
      Width           =   945
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "本所期限:"
      Height          =   180
      Index           =   0
      Left            =   3570
      TabIndex        =   41
      Top             =   690
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   210
      TabIndex        =   40
      Top             =   690
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Index           =   0
      Left            =   3570
      TabIndex        =   35
      Top             =   450
      Width           =   765
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   0
      Left            =   1170
      TabIndex        =   34
      Top             =   450
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3440;317"
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
      Left            =   210
      TabIndex        =   33
      Top             =   450
      Width           =   585
   End
End
Attribute VB_Name = "frm040106_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/13 改成Form2.0 (Text7,Label3)
'Memo by Morgan 2018/5/24 指定國家相關程式沒用,已移除
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'2005/7/5整理
Option Explicit
Dim strReceiveNo As String
'edit by nickc 2007/02/02
'Dim pA(T_PA) As String, cp(T_CP) As String
Dim pa() As String, cp() As String

Dim intWhere As Integer
Dim strInventor(1 To 2) As String, strPriority(1 To 5) As String
Dim strAdddeadline(1 To 3) As String
Dim m_strCompDate As String '記錄程式計算的提申期限
'Add By Cheng 2002/08/23
Dim m_strCust1 As String '申請人1
Dim m_strCust2 As String '申請人2
Dim m_strCust3 As String '申請人3
Dim m_strCust4 As String '申請人4
Dim m_strCust5 As String '申請人5
'Add By Cheng 2003/08/06
 '申請人地址
Dim m_strAdd(1 To 5) As String
'Add By Cheng 2003/08/11
'發明人地址
Dim m_strIAdd(1 To 10) As String
'Add by Morgan 2007/12/24
Public iFrom As Integer '0=內專,1=承辦人, 2=程序(跑歷程)
Dim m_bolFMP As Boolean 'Add by Morgan 2009/12/17 是否 FMP 案
Dim m_Subject As String 'Added by Morgan 2016/5/20
Dim stCP13 As String, stCP12 As String 'Added by Morgan 2021/1/28
Dim m_TWCP09 As String, m_TWCP01 As String, m_TWCP02 As String, m_TWCP03 As String, m_TWCP04 As String 'Added by Morgan 2021/11/1 台灣未發文新申請案收文號 'Modified by Morgan 2025/2/14 +本所號
Dim m_Have020 As Boolean, m_CPto020(1 To 4) As String   'Added by Lydia 2022/06/08 澳門案是否有大陸母案
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/10/31 是否為寰華案

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
'Modified by Lydia 2015/02/02 strTxt(1 To 20) => 21
'Modified by Lydia 2022/06/08 strTxt(1 To 21) => 40
Dim strTxt(1 To 40) As String, intStep As Integer, i As Integer
Dim ii As Integer, strET06 As String, iAttCnt As Integer
Dim arrAtt() As String 'Modified by Morgan 2016/3/21 附註可能不只兩個，改動態陣列
Dim bolPic202 As Boolean 'Added by Morgan 2016/12/27
 
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   intStep = 1
   
   'Added by Morgan 2016/12/19
   bolPic202 = False
   '若有"線條清晰之圖式"的補文件未發文時附件"附圖"要改為"參考附圖"
   strExc(0) = "select 1 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='202' and cp27||cp57 is null and instr(cp64,'線條清晰之圖式')>0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      bolPic202 = True
      Check1(2).Caption = "參考附圖一份" 'Added by Morgan 2016/12/27 --品薇
   End If
   'end 2016/12/19
   
   'Added by Morgan 2025/2/6
   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " select '" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      "','代理人撰稿提申日',cp47 from caseprogress where cp43='" & cp(9) & "' and cp10='958' and cp01='P' and cp47>0"
   intStep = intStep + 1
   'end 2025/2/6
            
'Remove by Morgan 2009/9/14 改共用例外欄位<最終提申期限>
   'Add by Morgan 2006/6/6
'   If ET03 = "38" And Text7(3) <> "" Then
'      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "'" & _
'         ",'提申期限'," & TransDate(Text7(3), 2) & ")"
'         intStep = intStep + 1
   'end 2006/6/6
'   Else

   If ET03 <> "38" Then
'end 2009/9/14
      'Add by Morgan 2009/6/3
      If Text7(0) <> "" Then
         'Added by Morgan 2021/11/4
         If Check4.Value = vbChecked Then
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','台灣案同日提申','♀')"
            intStep = intStep + 1
         End If
         'end 2021/11/4
      
         'modify by sonia 2019/12/106 incCNV_CHINESE_CUN改用incCNV_CHINESE_CUN1
         strExc(0) = TranslateKeyWord(incCNV_CHINESE_CUN1, DBDATE(Text7(0)), "")
      'Modified by Morgan 2025/2/6
      '   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','指定提申期限或四日內','務必於" & strExc(0) & "當日')"
      '   intStep = intStep + 1
         
      'Else
      '   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      '      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      '      "','指定提申期限或四日內','儘速於四個工作日內')"
      '   intStep = intStep + 1
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','提申期限','務必於" & strExc(0) & "當日')"
         intStep = intStep + 1
      ElseIf Text7(3) <> "" Then
         strExc(0) = TranslateKeyWord(incCNV_CHINESE_CUN1, DBDATE(Text7(3)), "")
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','提申期限','務必於" & strExc(0) & "前（含）')"
         intStep = intStep + 1
      Else
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','提申期限','儘速於四個工作日內')"
         intStep = intStep + 1
      'end 2025/2/6
      End If
      'end 2009/6/3
      
      'Removed by Morgan 2025/2/6 移到上面
      'If Text7(3) <> "" Then
      '   'modify by sonia 2019/12/106 incCNV_CHINESE_CUN改用incCNV_CHINESE_CUN1
      '   strExc(0) = TranslateKeyWord(incCNV_CHINESE_CUN1, DBDATE(Text7(3)), "")
      '   If ET03 = "31" Then
      '      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      '         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      '         "','有無提申期限','本案即將公開，故最遲須於" & strExc(0) & "前提申，')"
      '   Else
      '      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      '         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
      '         "','有無提申期限','故須於" & strExc(0) & "前提申，')"
      '   End If
      '   intStep = intStep + 1
      'End If
      'end 2025/2/6
      
      
      'Modify by Morgan 2009/7/10 指定提申控制一樣
'Removed by Morgan 2020/3/13 委托書已改用電子檔EMail，取消此提醒。--品薇,玲玲
'      If Text7(3) <> "" Or Text7(0) <> "" Then
'
'         'Modify by Morgan 2007/12/27
'         '1.提申期限-3天>系統日,才要印註解
'         '2.收達回覆期限=系統日+10天 -->原來設定
'         '3.若收達回覆期限>提申期限-3天,收達回覆期限=提申期限-3天
'
'         '提申期限-3天
'         If Text7(3) <> "" Then
'            strExc(1) = CompDate(2, -3, Text7(3))
'         Else
'            strExc(1) = CompDate(2, -3, Text7(0))
'         End If
'
'         If Val(strExc(1)) > Val(strSrvDate(1)) Then
'            '系統日+10天
'            strExc(2) = CompDate(2, 10, strSrvDate(1))
'            If Val(strExc(2)) > Val(strExc(1)) Then
'               strExc(2) = strExc(1)
'            End If
'            'modify by sonia 2019/12/106 incCNV_CHINESE_CUN改用incCNV_CHINESE_CUN1
'            strExc(0) = TranslateKeyWord(incCNV_CHINESE_CUN1, strExc(2), "")
'            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'               "','提申期限註解','＊本案倘於" & strExc(0) & "仍未接獲委托書，則請儘速傳真告知本所')"
'            intStep = intStep + 1
'         End If
'         'end 2007/12/27
'      End If
'end 2020/3/13
      
   End If
   'Add by Lydia 2015/02/02 有勾「順稿期限」，變更內文"儘速處理提申事宜"
   '品薇-現行實務有時會有順稿的要求，故會在出指示信時於順稿期限欄位鍵入要求代理人傳回日期,內文改成"儘速處理提申事宜"
      'Modified by Morgan 2020/9/22 有指定提申日除外
      'If Check3.Value = 1 And txtDate <> "" Then
      If Check3.Value = 1 And txtDate <> "" And Text7(0) = "" Then
      'end 2020/9/23
         If pa(9) = "020" And InStr("101,102,103,109,110,112,111", Text7(2)) > 0 And InStr("31,32,38", ET03) > 0 Then
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','儘速處理提申事宜','儘速處理提申事宜')"
            
            intStep = intStep + 1
         End If
      End If
    ' end 2015/02/02
    
    strExc(0) = ""
    'Modify by Morgan 2006/6/6 加附件:PCT公開文本,檢索報告一份
    'For i = 0 To 5
    For i = 0 To Check1.Count - 1
        If Check1(i).Value = 1 Then
            strExc(0) = strExc(0) & Check1(i).Caption & ","
        End If
    Next
    If strExc(0) <> "" Then
        If Right(strExc(0), 1) = "," Then strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
        strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                                "','附件內容','" & strExc(0) & "')"
        intStep = intStep + 1
    End If
    
    '910704 Sieg 402
    iAttCnt = 0 'Add by Morgan 2009/10/7
    If Text6.Visible = True And Text6 = "" Then
         iAttCnt = iAttCnt + 1
         ReDim Preserve arrAtt(iAttCnt) 'Added by Morgan 2016/3/21
         'Add by Morgan 2005/5/25
         If Text7(2) = "109" Then
            'Modify by Morgan 2009/10/7
            'strET06 = "附註：本案請一併提出「國際初步審查」請求。"
            arrAtt(iAttCnt) = "本案請一併提出「國際初步審查」請求。"
         Else
            'Modify by Morgan 2009/10/7
            'strET06 = "附註：本案請一併提出實質審查請求。"
            arrAtt(iAttCnt) = "本案請一併提出實質審查請求。"
         End If
         '2005/5/25 end
         
         'Remove by Morgan 2009/10/7 移到下面
         'strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                                "','有無實審','" & strET06 & "')"
         'intStep = intStep + 1
    End If
    
    'Added by Morgan 2015/8/12
    If pa(9) = "020" And cp(10) = "101" Then
       If PUB_ChkCPExist(cp, "417") = True Then
         iAttCnt = iAttCnt + 1
         ReDim Preserve arrAtt(iAttCnt) 'Added by Morgan 2016/3/21
         arrAtt(iAttCnt) = "本案請辦理提早公開。"
       End If
    End If
    'end 2015/8/12
    
    'Add by Morgan 2009/10/7
    '大陸一案兩請指示信附註
    If pa(9) = "020" And InStr("101,102", cp(10)) > 0 Then
         strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,ptm04 C2" & _
            " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & cp(1) & "' and cm02='" & cp(2) & "' and cm03='" & cp(3) & "' and cm04='" & cp(4) & "'" & _
            " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & cp(1) & "' and cm06='" & cp(2) & "' and cm07='" & cp(3) & "' and cm08='" & cp(4) & "') X" & _
            ",patent,patenttrademarkmap where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 and ptm01(+)='1' and ptm02(+)=pa08"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            iAttCnt = iAttCnt + 1
            ReDim Preserve arrAtt(iAttCnt) 'Added by Morgan 2016/3/21
            arrAtt(iAttCnt) = "本案為一案兩請，故請務必與 " & RsTemp("C1") & " 之" & RsTemp("C2") & "案於同一日提出申請，並同時提出「同日申請發明專利和實用新型專利的聲明」。"
         End If
         
         'Add by Morgan 2010/5/24
         strExc(0) = "select 1 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='430' and cp57 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            iAttCnt = iAttCnt + 1
            ReDim Preserve arrAtt(iAttCnt) 'Added by Morgan 2016/3/21
            'Modify by Morgan 2010/11/8
            'arrAtt(iAttCnt) = "本案請一併提出保密審查請求，並請於早上以臨櫃方式提出申請，且於當日取得受理通知書及保密審查決定書後，立即e-mail至本所，謝謝！"
            'Modified by Morgan 2013/4/9
            'arrAtt(iAttCnt) = "本案請一併提出「向外國申請專利前的保密審查請求」，並請於早上以臨櫃方式提出申請，且於當日取得受理通知書及保密審查決定書後，立即e-mail至本所，謝謝！"
            arrAtt(iAttCnt) = "本案請依照專利法實施細則第8條第2款提出「向外國申請專利之保密審查請求」，並請於當日取得受理通知書及保密審查決定書後，立即e-mail至本所，謝謝！"
            'end 2013/4/9
         End If
         'end 2010/5/24
         
         'Added by Morgan 2016/12/19
         '若有"線條清晰之圖式"的補文件未發文時指示信要自動帶出「本所將盡快提供正式圖檔」的字眼
         If bolPic202 Then
            iAttCnt = iAttCnt + 1
            ReDim Preserve arrAtt(iAttCnt)
            arrAtt(iAttCnt) = "本所將盡快提供正式圖檔。"
         End If
         'end 2016/12/19
    End If
    
    
    'Added by Morgan 2018/11/30
    If Val(Check1(0).Tag) = 1 And Check1(0).Value = 0 Then
      iAttCnt = iAttCnt + 1
      ReDim Preserve arrAtt(iAttCnt)
      arrAtt(iAttCnt) = "本所將後補委任書予 貴方，煩請 貴方仍先按指定時間提申本案。"
    End If
    'end 2018/11/30
    
    If iAttCnt > 0 Then
         strET06 = "附註："
         If iAttCnt = 1 Then
            strET06 = strET06 & arrAtt(iAttCnt)
         Else
            For intI = 1 To iAttCnt
               If intI > 1 Then strET06 = strET06 & vbCrLf & "　　　"
               strET06 = strET06 & intI & "." & arrAtt(intI)
            Next
         End If
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                                "','有無實審','" & strET06 & "')"
         intStep = intStep + 1
    End If
    'end 2009/10/7
    
    'Add By Cheng 2003/08/06
    '大陸新申請案指示信
    'Modify by Morgan 2006/6/6 加38
    'If ET03 = "31" Or ET03 = "32" Then
    'Modified by Lydia 2022/06/08 + 大陸衍生澳門案
    If ET03 = "31" Or ET03 = "32" Or ET03 = "38" Or ET03 = "39" Or ET03 = "21" Then
      If ET03 = "38" Then
         strExc(0) = GetValue(Text7(18), "PCT優先權日", ";")
         If Val(strExc(0)) > 0 Then
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                                    "','PCT優先權日'," & TransDate(strExc(0), 2) & ")"
            intStep = intStep + 1
         End If
      End If
        GetCustAddress pa(1), pa(2), pa(3), pa(4)
        For ii = 1 To 5
            If m_strAdd(ii) <> "" Then
                strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                                        "','申請人" & ii & "地址','" & m_strAdd(ii) & "')"
                intStep = intStep + 1
            End If
        Next ii
        
         'Modify By Sindy 2014/11/24
         If strSrvDate(1) >= 專利發明人檔啟用日 Then
            If m_strIAdd(1) <> "" Then
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                                       "','創作人地址','" & m_strIAdd(1) & "')"
               intStep = intStep + 1
            End If
         Else
            'Add By Cheng 2003/08/11
            For ii = 1 To 10
                If m_strIAdd(ii) <> "" Then
                    strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                                            "','創作人" & ii & "地址','" & m_strIAdd(ii) & "')"
                    intStep = intStep + 1
                Else
                   Exit For
                End If
            Next ii
         End If
    End If
   
   'Add by Morgan 2004/5/14
   If Check2.Value = 1 Then
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                              "','前案檢索資料','隨函一併附上相關前案檢索資料，')"
      intStep = intStep + 1
   End If
   
   
   'Added by Morgan 2012/3/2
   If chkFix(0).Value = 1 Then
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                              "','有國際階段修正','♀')"
      intStep = intStep + 1
   End If
   If chkFix(1).Value = 1 Then
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                              "','有國家階段修正','♀')"
      intStep = intStep + 1
   End If
   'End 2012/3/2
   
   
   'Add by Morgan 2006/10/3
   If Text7(2) = "111" Then
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) SELECT '" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','大陸案授權公告日',PA14 FROM CASEMAP,PATENT WHERE CM10='4' AND CM01='" & cp(1) & "' AND CM02='" & cp(2) & "' AND CM03='" & cp(3) & "' AND CM04='" & cp(4) & "' AND PA01(+)=CM05 AND PA02(+)=CM06 AND PA03(+)=CM07 AND PA04(+)=CM08 AND ROWNUM<2"
      intStep = intStep + 1
   End If
   
   'Add by Morgan 2008/1/9
   If Text7(2) = "110" Then
      'Modified by Morgan 2012/9/17 +PA09
      strExc(0) = "SELECT PA01,PA02,PA03,PA04,PA11,PA12,PA09 FROM CASEMAP,PATENT WHERE CM10='4' AND CM01='" & cp(1) & "' AND CM02='" & cp(2) & "' AND CM03='" & cp(3) & "' AND CM04='" & cp(4) & "' AND PA01(+)=CM05 AND PA02(+)=CM06 AND PA03(+)=CM07 AND PA04(+)=CM08"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','大陸案申請號','" & .Fields("PA11") & "')"
         intStep = intStep + 1

         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','大陸案公開日','" & .Fields("PA12") & "')"
         intStep = intStep + 1
         
         'Added by Morgan 2012/9/17
         If .Fields("PA09") = "221" Then
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','母案EPC要印','♀')"
            intStep = intStep + 1
            
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','母案EPC不印','♀')"
            intStep = intStep + 1
            
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','母案申請國','歐洲')"
               
         Else
            
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','母案申請國','中國')"
         End If
         intStep = intStep + 1
         'end 2012/9/17

         strExc(0) = "SELECT PD07,PD05,PD06,NA03 FROM PRIDATE,NATION WHERE PD01='" & .Fields("PA01") & "' AND PD02='" & .Fields("PA02") & "'" & _
                      " AND PD03='" & .Fields("PA03") & "' AND PD04='" & .Fields("PA04") & "'" & _
                      " AND NA01(+)=PD07"
         End With
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(0) = "": strExc(1) = "": strExc(2) = "": strExc(3) = ""
            With RsTemp
            Do While Not .EOF
               '優先權國家重複只印一次
               If strExc(0) <> "" & .Fields("PD07") Then
                  '對於多筆記錄以'; '來連起來
                  strExc(1) = strExc(1) & "; " & .Fields("NA03").Value
                  strExc(0) = "" & .Fields("PD07") '前筆國家
               End If
               '優先權日
               strExc(2) = strExc(2) & "; " & .Fields("PD05")
               '優先權號
               strExc(3) = strExc(3) & "; " & .Fields("PD06")
               .MoveNext
            Loop
            End With
            strExc(1) = Mid(strExc(1), 3)
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','大陸案優先權國家','" & strExc(1) & "')"
            intStep = intStep + 1
            strExc(2) = Mid(strExc(2), 3)
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','大陸案優先權日','" & strExc(2) & "')"
            intStep = intStep + 1
            strExc(3) = Mid(strExc(3), 3)
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','大陸案優先權號','" & strExc(3) & "')"
            intStep = intStep + 1
            
         End If
      End If
   End If
   
   'Add by Morgan 2009/12/17
   If Check3.Value = 1 And txtDate <> "" Then
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','順稿期限','" & DBDATE(txtDate) & "')"
      intStep = intStep + 1
   End If
   
   'Added by Morgan 2015/4/24
   'If Text7(2) = "101" Then 'Removed by Morgan 2019/12/11 不限制，新型申請也可能 Ex:P-123913
      If PUB_ChkCPExist(cp, "414") = True Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','有恢復權利要印','♀')"
         intStep = intStep + 1
      End If
   'End If
   'end 2015/4/24
   
   'Added by Lydia 2022/06/08 大陸衍生澳門案的國外指示信定稿(分案時有建立與大陸案關聯為大陸衍生案)
   If pa(9) = "044" And m_Have020 = True And m_CPto020(1) <> "" And m_CPto020(2) <> "" Then
       Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa()) '使用個案申請人1地址
       strExc(0) = "select pa10,pa11,pa14,pa15,cp45 from patent,caseprogress " & _
                        "where pa01='" & m_CPto020(1) & "' and pa02='" & m_CPto020(2) & "' and pa03='" & m_CPto020(3) & "' and pa04='" & m_CPto020(4) & "' " & _
                        "and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp31='Y'"
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
              strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','澳門相關大陸案號','" & m_CPto020(1) & "-" & m_CPto020(2) & IIf(m_CPto020(3) <> "0", "-" & m_CPto020(3), "") & IIf(m_CPto020(4) <> "00", "-" & m_CPto020(4), "") & "')"
              intStep = intStep + 1
           If "" & RsTemp.Fields("PA10") <> "" Then
              strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','澳門相關大陸案申請日','" & RsTemp.Fields("pa10") & "')"
              intStep = intStep + 1
           End If
           If "" & RsTemp.Fields("PA11") <> "" Then
              strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','澳門相關大陸申請案案號','" & RsTemp.Fields("pa11") & "')"
              intStep = intStep + 1
           End If
           If "" & RsTemp.Fields("PA15") <> "" Then
              strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','澳門相關大陸案授權公告號','" & RsTemp.Fields("pa15") & "')"
              intStep = intStep + 1
           End If
           If "" & RsTemp.Fields("PA14") <> "" Then
              strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','澳門相關大陸案授權公告日','" & RsTemp.Fields("pa14") & "')"
              intStep = intStep + 1
           End If
           If "" & RsTemp.Fields("CP45") <> "" Then
              strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','澳門相關大陸案陸代案號','" & RsTemp.Fields("CP45") & "')"
              intStep = intStep + 1
           End If
       End If
       '另外抓代表人中+英文
       If Trim(pa(79) & pa(80)) <> "" Then
          If Trim(pa(82) & pa(83) & pa(109) & pa(110) & pa(112) & pa(113) & pa(115) & pa(116) & pa(118) & pa(119) & pa(121) & pa(122) & pa(124) & pa(125) & pa(127) & pa(128) & pa(130) & pa(131)) = "" Then
              strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','專利案代表人1中英','" & pa(79) & IIf(pa(80) <> "", "/", "") & pa(80) & "')"
          Else
              strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','專利案代表人1中英','1." & pa(79) & IIf(pa(80) <> "", "/", "") & pa(80) & "')"
          End If
          intStep = intStep + 1
          ii = 1
       End If
       If Trim(pa(82) & pa(83)) <> "" Then
           ii = ii + 1
           strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','專利案代表人" & ii & "中英','" & ii & "." & pa(82) & IIf(pa(83) <> "", "/", "") & pa(83) & "')"
           intStep = intStep + 1
       End If
       For intI = 3 To 10

           If Trim(pa(109 + (3 * (intI - 3))) & pa(110 + (3 * (intI - 3)))) <> "" Then
               ii = ii + 1
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','專利案代表人" & ii & "中英','" & ii & "." & pa(109 + (3 * (intI - 1))) & IIf(pa(110 + (3 * (intI - 1))) <> "", "/", "") & pa(110 + (3 * (intI - 1))) & "')"
               intStep = intStep + 1
           End If
       Next intI
   End If
   'end 2022/06/08
   
   'Added by Morgan 2024/1/25
   If pa(9) = "020" Then
      'Modified by Morgan 2024/1/30
      'If PUB_GetCustomerValue(pa(26), "CU182") = "Y" Then
      If cp(118) = "" Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','紙本送件才印','♀')"
         intStep = intStep + 1
      End If
   End If
   'end 2024/1/25
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(intStep - 1, strTxt) Then
   If Not ClsLawExecSQL(intStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim bolChk As Boolean, strTmp As String
   Dim oChk As CheckBox
   
   Select Case Index
      Case 0 '確定
         'Add by Morgan 2007/12/26
         If iFrom <> 1 Then
         'end 2007/12/26
            '(申請人1一定要輸入)
            If IsEmptyText(Text7(4)) = True Then
               MsgBox "請輸入申請人1", vbOKOnly + vbCritical, "檢核資料"
               Exit Sub
            End If
         
            'Add by Morgan 2009/5/27 若已有提申期限需刪除後才可作業，以免期限錯亂！
            If PUB_ChkFileNP(cp(9)) Then
               MsgBox "下一程序已有提申期限，不可作業！"
               Exit Sub
            End If
            If Text7(3) <> "" And Text7(0) <> "" Then
               MsgBox "已有指定提申日不可再輸入最終提申期限！"
               Text7(3).SetFocus
               Exit Sub
            End If
            If cp(7) <> "" Then
               If Text7(3) <> "" Then
                  If Val(DBDATE(Text7(3))) > Val(DBDATE(cp(7))) Then
                     MsgBox "最終提申期限不可晚於法定期限！"
                     Text7(3).SetFocus
                     Exit Sub
                  End If
               ElseIf Text7(0) <> "" Then
                  If Val(DBDATE(Text7(0))) > Val(DBDATE(cp(7))) Then
                     MsgBox "指定提申日不可晚於法定期限！"
                      Text7(0).SetFocus
                     Exit Sub
                  End If
               End If
            End If
            'end 2009/5/27
         End If
         '檢查輸入資料的完整性
         If CheckDataIntegrity = False Then Exit Sub
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Added by Lydia 2023/10/04 FMP案待客戶最終指示相關控管：彈提醒
         If m_bolFMP2 = True Then 'Added by Lydia 2023/10/31 只限寰華案要通知
            If PUB_ChkFMP970mail("3", cp(1), cp(2), cp(3), cp(4), strTmp) = True Then
               If strTmp <> "" Then
                  MsgBox strTmp, vbInformation
               End If
            End If
         End If
         'end 2023/10/04
         
         '是否列印指示信(或申請書)
         If Text7(9) <> "N" Then '指示信
            '是否修改指示信內容
            If Text7(10) = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If

            '選擇案件性質
            Select Case Text7(2)
               'Add by Morgan 2005/5/25 加PCT申請(109)
               'Case 發明申請, 新型申請, 設計申請
               Case 發明申請, 新型申請, 設計申請, "109", "103", "110", "112", "111"
                  'Add by Morgan 2006/6/6
                  If pa(46) = "Y" Then
                     strTmp = "38"
                  'end 2006/6/6
                  Else
                     Select Case pa(9)
                        Case "013" '香港
                           strTmp = "33"
                           
                        'Add by Morgan 2006/6/9
                        Case "044" '澳門
                           'Added by Lydia 2022/06/08 大陸衍生澳門案
                           If m_Have020 = True Then
                               strTmp = "21"
                           Else
                           'end 2022/06/08
                               strTmp = "39"
                           End If 'Added by Lydia 2022/06/08
                        Case Else
                           If strPriority(1) <> "" Then
                              strTmp = "32"
                           Else
                              strTmp = "31"
                           End If
                     End Select
                  End If
                  
                  'Add by Morgan 2009/10/14
                  '若只勾選委托書則定稿特別
                  If Check1(0).Value = vbChecked Then
                     intI = 0
                     For Each oChk In Check1
                        intI = intI + oChk.Value
                     Next
                     If intI = 1 Then
                        strTmp = "42"
                     End If
                  End If
                  
               '2008/2/4 add by sonia 實體審查PCT案定稿不同
               Case 實體審查
                  If pa(46) = "Y" Then
                     strTmp = "38"
                  Else
                     strTmp = "30"
                  End If
               '2008/2/4 end
               '2009/10/26 add by sonia大陸實用新型的申請檢索報告
               Case "421"
                  If pa(9) = "020" And pa(8) = "2" Then
                     strTmp = "43"
                  Else
                     strTmp = "30"
                  End If
               '2009/10/26 end
               Case Else
                  strTmp = "30"
            End Select
            
            StartLetter "02", strTmp
            'Modified by Morgan 2016/5/20
            '指示信電子化
            'NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, 0
            If iFrom = 0 And Left(Pub_StrUserSt03, 1) <> "F" Then
               NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum, , , , , , , , , , , , , strReceiveNo
               If bolChk Then
                  frm1105_1.m_RecNo = strReceiveNo
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & cp(10) & ".DATA.PDF"
                  frm1105_1.m_Subject = m_Subject
                  frm1105_1.Show
               End If
            Else
               NowPrint strReceiveNo, "02", strTmp, bolChk, strUserNum
            End If
            'end 2016/5/20
         End If
A0:
         frm040106_1.Show
         ' 90.07.11 modify by louis (回第一個畫面清除)
         frm040106_1.Clear
         Unload Me
      Case 1 '回前畫面
         frm040106_1.Show
         Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean
'Modified by Lydia 2015/02/02 strTxt(1 To 20) => 21
Dim strTxt(1 To 21) As String, intStep As Integer
Dim i As Integer, strTmp As String, intMax As Long, varTmp As Variant
Dim nIndex As Integer
'Add By Cheng 2002/06/27
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
FormSave = True
cnnConnection.BeginTrans

   'Modify by Morgan 2007/12/24
   '若從承辦人作業來時只存案件名稱
   If iFrom = 1 Then
      Select Case pa(1)
         Case "P"
            strTxt(1) = "UPDATE PATENT SET PA05=" & CNULL(ChgSQL(Text7(14))) & ",PA06=" & CNULL(ChgSQL(Text7(15))) & _
               ",PA07=" & CNULL(ChgSQL(Text7(16))) & "  WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
               cnnConnection.Execute strTxt(1)
      End Select
   Else
      '911028 nick 移到下面
      '   'edit by nickc 2007/02/02 不用 dll 了
      'intMax = objPublicData.GetNextProgressNo
      intMax = GetNextProgressNo
      If Combo2 <> "" Then
         cp(44) = Combo2
         'MODIFY BY SONIA 90.10.21
         If Len(cp(44)) < 9 Then cp(44) = Left(Combo2.Text, 9) & String(9 - Len(Combo2.Text), "0")
         'edit by nickc 2007/02/02 不用 dll 了
         'If Not objPublicData.GetCaseThatCode(cp) Then cp(45) = ""
         If Not ClsPDGetCaseThatCode(cp) Then cp(45) = ""
      Else
         cp(45) = ""
      End If
      
      '1
      '更新基本檔
      strExc(0) = ""
      '92.5.29 CANCEL BY SONIA
      'Select Case Text7(2)
      '   Case "101", "102", "103", "105", "301", "302", "303", "304", "305", "306", "307", "803"
      '      If pa(1) = "P" Then
      '         strExc(0) = "PA10=" & Val(TransDate(ServerDate - 19110000, 2)) & ","
      '      End If
      '   Case Else
      '      strExc(0) = ""
      'End Select
      '92.5.29 END
      If strInventor(2) <> "" Then
         varTmp = Split(strInventor(2), ",")
         'Modify By Sindy 2014/11/12
         If strSrvDate(1) >= 專利發明人檔啟用日 Then
            '全部刪除,重新新增
            strSql = "delete from patentInventor where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4))
            Pub_SeekTbLog strSql 'Add By Sindy 2017/8/23
            cnnConnection.Execute strSql
            For i = 0 To UBound(varTmp)
               strSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
                        CNULL(pa(1)) & "," & CNULL(pa(2)) & "," & CNULL(pa(3)) & "," & CNULL(pa(4)) & "," & i + 1 & ",'" & varTmp(i) & "')"
               Pub_SeekTbLog strSql 'Add By Sindy 2017/8/23
               cnnConnection.Execute strSql
            Next i
         Else
         '2014/11/12 END
            For i = 0 To UBound(varTmp)
               'Modified by Morgan 2014/12/9 補單引號
               strExc(0) = strExc(0) & "PA" & Format(i + 60) & "='" & varTmp(i) & "',"
            Next
         End If
      End If
      ' 90.07.18 modify by louis (存檔時存九碼)
      For nIndex = 4 To 8
         If IsEmptyText(Text7(nIndex)) = False Then
            Text7(nIndex) = Text7(nIndex) & String(9 - Len(Text7(nIndex)), "0")
         End If
      Next nIndex
      Select Case pa(1)
         Case "P"
            'Modified by Morgan 2011/11/17 +地址欄位也會有單引號
            strTxt(1) = "UPDATE PATENT SET " & strExc(0) & "PA05=" & CNULL(ChgSQL(Text7(14))) & ",PA06=" & CNULL(ChgSQL(Text7(15))) & _
               ",PA07=" & CNULL(ChgSQL(Text7(16))) & ",PA26=" & CNULL(Text7(4)) & ",PA27=" & CNULL(Text7(5)) & _
               ",PA28=" & CNULL(Text7(6)) & ",PA29=" & CNULL(Text7(7)) & ",PA30=" & CNULL(Text7(8)) & _
               ",PA31=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(4).Text, "1"))) & ",PA36=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(4).Text, "2"))) & ",PA41=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(4).Text, "3"))) & _
               ",PA32=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(5).Text, "1"))) & ",PA37=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(5).Text, "2"))) & ",PA42=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(5).Text, "3"))) & _
               ",PA33=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(6).Text, "1"))) & ",PA38=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(6).Text, "2"))) & ",PA43=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(6).Text, "3"))) & _
               ",PA34=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(7).Text, "1"))) & ",PA39=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(7).Text, "2"))) & ",PA44=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(7).Text, "3"))) & _
               ",PA35=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(8).Text, "1"))) & ",PA40=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(8).Text, "2"))) & ",PA45=" & CNULL(ChgSQL(PUB_GetCustEachAdd(Me.Text7(8).Text, "3"))) & _
               ",PA91=" & CNULL(ChgSQL(Text7(18))) & "  WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
               'Add By Cheng 2002/11/06
               cnnConnection.Execute strTxt(1)
      End Select
      '更新案件進度檔
      'Modify By Cheng 2002/08/07
      '取消"鑑定報告對照"頁籤
      '91.10.28 MODIFY BY SONIA 不可上發文日
      'strTxt(2) = "UPDATE CASEPROGRESS SET cp27=" & CNULL(TransDate(ServerDate - 19110000, 2)) & _
      '   ",cp44=" & CNULL(cp(44)) & ",CP45=" & CNULL(cp(45)) & _
      '   ",cp64=" & CNULL(ChgSQL(Text7(17))) & _
      '   " WHERE CP09='" & strReceiveNo & "'"
      strTxt(2) = "UPDATE CASEPROGRESS SET cp44=" & CNULL(cp(44)) & ",CP45=" & CNULL(cp(45)) & _
         ",cp64=" & CNULL(ChgSQL(Text7(17))) & _
         " WHERE CP09='" & strReceiveNo & "'"
       'Add By Cheng 2002/11/06
       cnnConnection.Execute strTxt(2)
      '91.10.28 END
      intStep = 3
      '7
      'Add By Cheng 2002/12/10
      '若更改案件性質
      If Text7(2) <> cp(10) Then
         If Left(Text7(2), 1) = "3" Then
            strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP09," & _
               "CP10,CP20,CP32) VALUES ('" & pa(1) & "','" & pa(2) & _
               "','" & pa(3) & "','" & pa(4) & "','" & AutoNo("B", 6) & "','" & Text7(2) & _
               "','N','N')"
           'Add By Cheng 2002/11/06
           cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
            If Text7(2) = "301" Or Text7(2) = "302" Or Text7(2) = "303" Then
               strTxt(intStep) = "UPDATE PATENT SET PA08='" & Right(Text7(2), 1) & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
               'Add By Cheng 2002/11/06
               cnnConnection.Execute strTxt(intStep)
               intStep = intStep + 1
            End If
         ElseIf Text7(2) = 舉發 Then
            strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP09," & _
               "CP10,CP20,CP32) VALUES ('" & pa(1) & "','" & pa(2) & _
               "','" & pa(3) & "','" & pa(4) & "','" & AutoNo("B", 6) & "','" & 舉發 & _
               "','N','N')"
           'Add By Cheng 2002/11/06
           cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
            strTxt(intStep) = "UPDATE PATENT SET PA23='3' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
           'Add By Cheng 2002/11/06
           cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
         End If
      End If
      
      'Modify by Amy 2014/04/17 +, strPriority(5)
      If Not ClsPDSavePriority(pa, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)) Then
           'Add By Cheng 2002/11/06
           GoTo ErrorHandler
      End If
   '8
      If pa(1) = "P" Then
         i = 專利
      Else
         i = 0
      End If
         
'Remove by Morgan 2009/12/11
'      'Add By Cheng 2002/12/10
'      '若有輸入補件期限
'      If strAdddeadline(1) <> "" Then
'
'         'Modify by Morgan 2004/9/14
'         '改新增至案件進度檔
'
'         Dim varAddDeadLineTemp1, varAddDeadLineTemp2, varAddDeadLineTemp3
'
'         varAddDeadLineTemp1 = Split(strAdddeadline(1), ",")
'         varAddDeadLineTemp2 = Split(strAdddeadline(2), ",")
'         varAddDeadLineTemp3 = Split(strAdddeadline(3), ",")
'
'         For i = LBound(varAddDeadLineTemp1) To UBound(varAddDeadLineTemp1)
'            '2008/11/26 MODIFY BY SONIA 發現文件未存入CP64,加存CP43如此發文時才能更新
'            'strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14" & _
'                  ",CP20,CP32) VALUES ('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'" & _
'                  "," & strSrvDate(1) & "," & varAddDeadLineTemp1(i) & "," & varAddDeadLineTemp2(i) & _
'                  ",'" & AutoNo("B", 6) & "','" & 補文件 & "','" & cp(12) & "','" & cp(13) & "','" & cp(14) & "','N','N')"
'            strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14" & _
'                  ",CP20,CP32,CP43,CP64) VALUES ('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'" & _
'                  "," & strSrvDate(1) & "," & varAddDeadLineTemp1(i) & "," & varAddDeadLineTemp2(i) & _
'                  ",'" & AutoNo("B", 6) & "','" & 補文件 & "','" & cp(12) & "','" & cp(13) & "','" & cp(14) & "','N','N','" & strReceiveNo & "','" & varAddDeadLineTemp3(i) & "')"
'            '2008/11/26 END
'            cnnConnection.Execute strSQL
'
'         Next
'
'         '2004/9/14 END
'
'      End If
'end 2009/12/11
      
      'Add by Morgan 2009/5/27
      '若有輸最終提申期限或指定提申日時要新增下一程序
      'Modify by Morgan 2010/11/1 年費不必--敏惠
      If Text7(3) <> "" Or Text7(0) <> "" And Text7(2) <> "605" Then
         '最終
         If Text7(3) <> "" Then
            strExc(1) = DBDATE(Text7(3))
            strExc(3) = "996"
         '指定
         Else
            strExc(1) = DBDATE(Text7(0))
            strExc(3) = "995"
         End If
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = " insert into nextprogress a (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
            " values('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','" & strExc(3) & "'" & _
            "," & strExc(2) & "," & strExc(1) & ",'" & strUserNum & "',GETNP22)"
         cnnConnection.Execute strSql, intI
         
         'Added by Morgan 2016/12/19
         '若有"線條清晰之圖式"的補文件未發文時自動設定繪圖人員同新案, 若有指定提申時更新所限為該日期
         strSql = "update caseprogress set cp06=" & IIf(strExc(3) = "995", strExc(2), "cp06") & ",cp29=nvl(cp29,'" & cp(29) & "'),cp107=nvl(cp107,'Y')" & _
            " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='202' and cp27||cp57 is null and instr(cp64,'線條清晰之圖式')>0"
         cnnConnection.Execute strSql, intI
         If intI > 0 Then
            strSql = "update engineerprogress set ep14=nvl(ep14," & strSrvDate(1) & "),ep17=nvl(ep17," & strSrvDate(1) & ")" & _
               " where ep02 in (select cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='202' and cp27||cp57 is null and instr(cp64,'線條清晰之圖式')>0)"
            cnnConnection.Execute strSql, intI
         End If
         'end 2016/12/19
         
         'Added by Morgan 2021/11/1
         '台灣案同日提申
         If Check4.Value = vbChecked And m_TWCP09 <> "" Then
            strExc(0) = "與" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)) & "指定於" & ChangeWStringToTDateString(DBDATE(Text7(0))) & "同日提申"
            
            'EMail通知承辦人
            'Modified by Morgan 2022/3/25 +副本給負責送件的程序人員 --韻丞
            'Added by Morgan 2025/1/24
            If strSrvDate(1) >= P業務區劃分啟用日 Then
               strExc(1) = PUB_GetPHandler(m_TWCP01 & m_TWCP02 & m_TWCP03 & m_TWCP04)
            Else
            'end 2025/1/24
            
               strExc(1) = Pub_GetSpecMan("A1")
            End If
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc09,mc13)" & _
               " select '" & strUserNum & "' mc01,cp14 mc02,to_char(sysdate,'yyyymmdd') mc03" & _
               ",to_char(sysdate,'hh24miss') mc04,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'" & strExc(0) & "' mc07" & _
               ",'" & strExc(1) & "',cp09 mc13" & _
               " from caseprogress where cp09='" & m_TWCP09 & "'"
            cnnConnection.Execute strSql, intI
            
            '歷程新增聯絡
            strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep10)" & _
                     " select cp09 eep01,nvl(max(eep02),0)+1 eep02,'" & strUserNum & "' eep03,'" & EMP_聯絡 & "' eep04" & _
                     ",cp14 eep05,to_char(sysdate,'yyyymmdd') eep06,to_char(sysdate,'hh24miss') eep07,'本案" & strExc(0) & "' eep08" & _
                     ",'" & strExc(1) & "' eep10 from caseprogress,empelectronprocess where cp09='" & m_TWCP09 & "'" & _
                     " and eep01(+)=cp09 group by cp09,cp14"
            cnnConnection.Execute strSql, intI
            
            '設定指定送件日
            strSql = "update caseprogress set cp141='3',cp142=" & DBDATE(Text7(0)) & ",cp164='1' where cp09='" & m_TWCP09 & "'"
            cnnConnection.Execute strSql, intI
            
         End If
         'end 2021/11/1
      End If
      
      'Add by Morgan 2009/12/17
      If pa(9) <> "000" And Check3.Value = 1 And txtDate <> "" Then
         strExc(2) = DBDATE(txtDate)
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(2) & _
            " where np01='" & cp(9) & "' and np06 is null and np07='994'"
         cnnConnection.Execute strSql, intI
         If intI = 0 Then
            'Added by Morgan 2018/6/13 外專人員操作時抓FCP管制人--敏莉
            If Left(Pub_StrUserSt03, 1) = "F" Then
               strExc(1) = PUB_GetFCPHandler(cp(1), cp(2), cp(3), cp(4))
            Else
            'end 2018/6/13
               'Added by Morgan 2025/1/24
               If strSrvDate(1) >= P業務區劃分啟用日 Then
                  strExc(1) = PUB_GetPHandler(cp(1) & cp(2) & cp(3) & cp(4))
               Else
               'end 2025/1/24
               
                  strExc(1) = Pub_GetSpecMan("G") '順稿管制人
                  
               End If 'Added by Morgan 2025/1/24
            End If 'Added by Morgan 2018/6/13
            
            strSql = "insert into nextprogress(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22)" & _
               " SELECT '" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & _
               "','994'," & strExc(2) & "," & strExc(2) & ",'" & strExc(1) & "',NP22" & _
               " FROM (SELECT NVL(MAX(NP22),0)+1 NP22 FROM NEXTPROGRESS) X"
            cnnConnection.Execute strSql, intI
         End If
      End If
      
   End If
   
   'Added by Morgan 2016/5/20
   '指示信電子化
   If iFrom = 0 And Left(Pub_StrUserSt03, 1) <> "F" Then
      m_Subject = "請代為提出" & Label3(5) & "申請" & IIf(cp(45) <> "", " Y/R:" & cp(45) & ";", "") & " O/R:" & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
      If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
         'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
         'strExc(2) = Pub_GetSpecMan("PS4")
         strExc(2) = PUB_GetLetterJudgeNew("2", pa(1), cp(10), pa(9))
         PUB_AddAppForm strReceiveNo, True, strExc(2), m_Subject
      End If
   End If
   'end 2016/5/20
   
    'Add By Cheng 2002/11/06
    cnnConnection.CommitTrans
    Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    FormSave = False
End Function

Private Sub Combo2_Click()
   Combo2_Validate False
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
   ' 90.08.09 modify by louis
   If StrLength(Combo2.Text) > 9 Then
      Cancel = True
      MsgBox "代理人代碼錯誤", vbOKOnly + vbCritical, "檢核資料"
      Combo2.SelStart = 0
      Combo2.SelLength = Len(Combo2.Text)
      Exit Sub
   End If
   If Len(Me.Combo2.Text) <= 0 Then
      MsgBox "代理人欄不可為空白!!!", vbExclamation
      Cancel = True
      Exit Sub
   End If
   If Combo2.Text <> "" Then
      If Not ChgType(12) Then Cancel = True
   End If
   'Add by Morgan 2004/9/14 檢查客戶/代理人是否不再使用
   If Cancel = False Then
      If PUB_CheckStatus(Combo2.Text) = False Then Cancel = True
   End If
End Sub

Private Sub Command2_Click(Index As Integer)
   Select Case Index
      Case 0 '發明人
         ModifyInventor strInventor(1), strInventor(2)
      Case 1 '優先權資料
         'Moidify by Amy 2014/04/17 +, strPriority(5)
         ModifyPriority strPriority(1), strPriority(2), strPriority(3), pa(8), , pa(1) & pa(2) & pa(3) & pa(4), pa(9), , strPriority(4), strPriority(5)
        'Add By Cheng 2002/12/10
      Case 2 '補件期限
         
         'Modify by Morgan 2009/12/11
         'frm880003.PA09 = pa(9)
         'frm880003.CP10 = Text7(2)
         ''2008/11/26 ADD BY SONIA P指示信預設本所期限為系統日+1個月,且抓工作天
         ''2009/6/26 MODIFY BY SONIA 改為系統日+2個月
         'frm880003.m_txtAddDeadline = PUB_GetWorkDay1(CompDate(1, 2, strSrvDate(1)), True)
         ''2008/11/26 END
         'ModifyAddDeadline strAdddeadline(1), strAdddeadline(2), strAdddeadline(3)
         strExc(1) = PUB_GetWorkDay1(CompDate(1, 2, strSrvDate(1)), True)
         ModifyAddDeadline1 cp(9), strExc(1)
         'end 2009/12/11
   End Select
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
ReDim cp(1 To TF_CP) As String
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   intWhere = 國內
   With frm040106_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   ReadPatent
   '顯示收文號
   Label3(0) = strReceiveNo
   '91.11.10 ADD BY SONIA
   Select Case Text7(2)
      'Memo by Morgan 2009/10/14 香港短期專利112的附件寫死在定稿內,畫面勾選無效
      'Modify by Morgan 2008/1/8 加 112
      Case 發明申請, 新型申請, 112
         'Add by Morgan 2006/6/6 加PCT
         If pa(46) = "Y" Then
            Check1(0).Value = 1
            Check1(1).Value = 1
            Check1(2).Value = 1
            
            'Modified by Morgan 2012/3/2
            'Check1(6).Value = 1
            'Check1(7).Value = 1
            chkFix(0).Visible = True
            chkFix(1).Visible = True
            'End 2012/3/2
         Else
         'end 2006/6/6
           Check1(0).Value = 1
           Check1(1).Value = 1
           Check1(2).Value = 1
         End If
      Case 設計申請
           Check1(0).Value = 1
           Check1(2).Value = 1
           Check1(3).Value = 1
           Check1(8).Value = 1 'Add by Morgan 2009/10/12 --郭
      'Add by Morgan 2005/5/25 --郭
      Case "109"
           Check1(0).Value = 1
           Check1(1).Value = 1
           Check1(2).Value = 1
   End Select
   '91.11.10 END
   
   Check1(0).Tag = Check1(0).Value   'Added by Morgan 2018/11/30 紀錄預設,若取消勾選時指示信加備註
   
   'Added by Morgan 2021/11/4
   'Modified by Morgan 2022/4/15 +109 PCT申請 --品薇
   If (pa(9) = "020" Or pa(9) = "056") And (cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103" Or cp(10) = "109") Then
      Check4.Visible = True
   Else
      Check4.Visible = False
   End If
   'end 2021/11/4
   
   'Add by Morgan 2007/12/25
   If iFrom = 1 Then
      Me.Command2(0).Enabled = False
      Me.Command2(2).Enabled = False
      Me.Command2(1).Enabled = False
      If cp(44) = "" Then
         Combo2.Text = ""
         Label3(6).Caption = ""
      End If
      Combo2.Enabled = False
      Text7(9).Enabled = False
      Text7(10).Enabled = False
      Text7(4).Enabled = False
      Text7(5).Enabled = False
      Text7(6).Enabled = False
      Text7(7).Enabled = False
      Text7(8).Enabled = False
      Text7(17).Enabled = False
      Text7(18).Enabled = False
   'Add by Morgan 2009/6/1
   Else
      If PUB_ChkFileNP(cp(9)) Then MsgBox "下一程序已有提申期限，若為重寫指示信，要先刪除後才可作業！"
   End If
   
   'Add by Morgan 2009/12/17
   'Modified by Morgan 2021/1/28
   'If Left(cp(12), 1) = "F" And pa(10) <> "000" Then
   stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
   stCP12 = GetSalesArea(stCP13)
   'Modified by Lydia 2023/06/20 + And pa(9) <> "000"
   If Left(stCP12, 1) = "F" And pa(9) <> "000" Then
   'end 2021/1/28
      m_bolFMP = True
   Else
      m_bolFMP = False
   End If
   'Added by Lydia 2023/10/31 判斷寰華案
   m_bolFMP2 = False
   If m_bolFMP = True Then
      m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4))
   End If
   'end 2023/10/31
   'FMP新案預設順稿期限
   strExc(0) = "select np08 from nextprogress where np01='" & cp(9) & "' and np06 is null and np07='994'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Check3.Value = 1
      Check3.Enabled = False
      txtDate = TransDate(RsTemp(0), 1)
   ElseIf m_bolFMP And (cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103") Then
      strExc(1) = PUB_GetDeadLine(DBDATE(cp(5)), DBDATE(cp(7)), 5)
      If strExc(1) <> "" Then
         Check3.Value = 1
         txtDate = TransDate(strExc(1), 1)
      End If
   End If
   'end 2009/12/17

    'Added by Lydia 2022/06/08 澳門案是否有大陸母案 (大陸衍生澳門案)
    If pa(1) = "P" And pa(9) = "044" And m_bolFMP = True And InStr("101,102,103,109,110,111,112", cp(10)) > 0 Then
       strExc(0) = pa(1): strExc(1) = pa(2): strExc(2) = pa(3): strExc(3) = pa(4)
       m_Have020 = Cls003GetCaseMap(strExc, 5)
       If m_Have020 = True Then
          m_CPto020(1) = strExc(4): m_CPto020(2) = strExc(5): m_CPto020(3) = strExc(6): m_CPto020(4) = strExc(7)
          '預設的順稿勾選拿掉
          Check3.Value = 0
          Check3.Enabled = False
       End If
    End If
    'end 2022/06/08
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2021/11/2
   'Set frm040106_3 = Nothing 'Removed by Morgan 2021/12/13 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub ReadPatent()
 Dim Lbl As Object, txt As Object, i As Integer
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
   strInventor(1) = ""
   Select Case pa(1)
      Case "P"
         '讀取專利基本檔
         If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
            '顯示案件名稱中英日
            For i = 5 To 7
               Text7(i + 9) = pa(i)
            Next
            '顯示案件備註
            Text7(18) = pa(91)
            '顯示申請人1,2,3,4,5
            For i = 26 To 30
               If pa(i) <> "" Then
                  strInventor(1) = strInventor(1) & pa(i) & ","
                  Text7(i - 22) = pa(i)
                  ChgType (i)
               End If
            Next
         End If
         'Add By Cheng 2002/08/23
         m_strCust1 = "" & Me.Text7(4).Text
         m_strCust2 = "" & Me.Text7(5).Text
         m_strCust3 = "" & Me.Text7(6).Text
         m_strCust4 = "" & Me.Text7(7).Text
         m_strCust5 = "" & Me.Text7(8).Text
         
      Case "PS"
         If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
            For i = 5 To 7
               Text7(i + 9) = pa(i)
            Next
            Text7(18) = pa(18)
            If pa(8) <> "" Then
               strInventor(1) = strInventor(1) & pa(i) & ","
               Text7(4) = pa(8)
               ChgType (26)
            End If
            If pa(58) <> "" Then
               strInventor(1) = strInventor(1) & pa(i) & ","
               Text7(5) = pa(58)
               ChgType (27)
            End If
            If pa(59) <> "" Then
               strInventor(1) = strInventor(1) & pa(i) & ","
               Text7(6) = pa(59)
               ChgType (28)
            End If
         End If
         'Add By Cheng 2002/08/23
         m_strCust1 = "" & Me.Text7(4).Text
         m_strCust2 = "" & Me.Text7(5).Text
         m_strCust3 = "" & Me.Text7(6).Text
         m_strCust4 = ""
         m_strCust5 = ""
   End Select
   If Right(strInventor(1), 1) = "," Then strInventor(1) = Left(strInventor(1), Len(strInventor(1)) - 1)
   '顯示申請國家
   If pa(9) <> "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetNation(pA(9), strExc(0)) Then Label3(8) = strExc(0)
      If ClsPDGetNation(pa(9), strExc(0)) Then Label3(8) = strExc(0)
   End If
   
   cp(9) = strReceiveNo
   '讀取案件進度檔
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.ReadCaseProgressDatabase(cp(), intWhere) Then
   If ClsPDReadCaseProgressDatabase(cp(), intWhere) Then
      '顯示智權人員
      If cp(13) <> "" Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(cp(13), strExc(0)) Then Label3(1) = strExc(0)
         If ClsPDGetStaff(cp(13), strExc(0)) Then Label3(1) = strExc(0)
      End If
      '顯示本所期限
      Label3(2) = cp(6)
      '顯示法定期限
      Label3(7) = cp(7)
      
      'Add by Morgan 2009/7/24
      '最終提申預設為法限
      Text7(3) = DBDATE(cp(7))
      
      '顯示案件性質
      If cp(10) <> "" Then
         Text7(2) = cp(10)
         ChgType (2)
      End If
        'Add By Cheng 2002/12/10
        '若案件性質為發明(101), 新型(102), 設計(103), 改請(3xx), 異議(801), 才開放可修改案件性質
        Select Case cp(10)
        Case "101", "102", "103", "301", "302", "303", "304", "305", "306", "307", "801"
            Me.Text7(2).Enabled = True
        Case Else
            Me.Text7(2).Enabled = False
        End Select
      
      '顯示進度備註
      Text7(17) = cp(64)
      '顯示代理人
      'Modify by Morgan 2008/10/16 +若進度檔已有代理人則預設
      'Modified by Lydia 2016/10/27 +新案有申請人指定國外代理人檔則預設 => cp(9), pa(9), pa(26)
      AddAgent Combo2, cp, , cp(44), cp(116), cp(9), pa(9), pa(26)
   End If
   
   ChkCp10
   GetCaseFee pa(1), pa(9), cp(10)
   
   If pa(1) = "P" Then
      i = 專利
   Else
      i = 0
   End If
   
   'Modify by Amy 2014/04/17 +, strPriority(5)
   If Not ClsPDReadPriority(pa, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)) Then
      
   End If
   
   ' 90.07.05 modify by louis
   If pa(9) < "010" Then
      ' 是否列印指示信預設為N
      Text7(9) = "N": Text7(9).Enabled = False
      Select Case cp(10)
         Case "101", "102", "103", "104", "105":
            ' 是否列印通知函預設為N
            Text7(21) = "N": Text7(21).Enabled = False
         Case Else:
      End Select
   End If
   
'Removed by Morgan 2021/2/26 移到下面
'   Text6 = ""
'   'ADD BY SONIA 2015/4/27
'   If pa(9) = "020" Then
'      intI = 1
'      strExc(0) = "select 1 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='416' and cp57 is null"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'      Else
'         Text6 = "N"
'      End If
'   End If
'   '2015/4/27 END
'end 2021/2/26
   
   'Add by  Morgan 2004/5/14
   If cp(10) = "416" And pa(9) = "020" Then
      Check2.Visible = True
   Else
      Check2.Visible = False
   End If
   
   'Modify by Morgan 2005/5/25 加"109"
   'If cp(10) = "101" And pa(9) = "020" Then
   'Modify by Morgan 2008/7/17 加澳門發明 --郭
   'Modify by Morgan 2013/1/11 加澳門外觀設計 --品薇 Ex.P-103941
   'Modified by Morgan 2021/2/26 +澳門新型--品薇 Ex.P-122796
   If (pa(9) = "020" And cp(10) = "101") Or (pa(9) = "056" And cp(10) = "109") Or (pa(9) = "044" And cp(10) = "101") Or (pa(9) = "044" And cp(10) = "102") Or (pa(9) = "044" And cp(10) = "103") Then
      Text6.Visible = True
      Label20(6).Visible = True
      'Added by Morgan 2021/2/26
      strExc(0) = "select 1 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='416' and cp57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Text6 = ""
      Else
         Text6 = "N"
      End If
      'end 2021/2/26
   Else
      Text6.Visible = False
      Label20(6).Visible = False
   End If
   
   '91.11.19 ADD BY SONIA
   Text7(10) = "Y"
   '91.11.19 END

End Sub

Private Sub ChkCp10()
 Dim bolChk As Boolean
   bolChk = False
   Select Case Text7(2)
      Case "101", "102", "103", "104", "105"
         Command2(1).Enabled = True
         bolChk = True
      Case Else
         Command2(1).Enabled = False
         If Left(Text7(2), 1) = "3" Then bolChk = True
   End Select
End Sub

Private Sub GetCaseFee(ByVal CF01 As String, ByVal CF02 As String, ByVal CF03 As String)
   m_strCompDate = ""
   intI = 1
   strExc(0) = "SELECT CF05,CF11 FROM CASEFEE WHERE CF01='" & CF01 & "' AND " & _
      "CF02='" & CF02 & "' AND CF03='" & CF03 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields(1)) Then
         '91.11.19 MODIFY BY SONIA 不預設在畫面上改在發文才做
         'Text7(3) = TransDate(CompDate(2, Val(rsTemp.Fields(1)), TransDate(ServerDate - 19110000, 2)), 1)
         'm_strCompDate = "" & Me.Text7(3).Text
         'Text7(3) = TransDate(CompDate(2, Val(rsTemp.Fields(1)), TransDate(ServerDate - 19110000, 2)), 1)
         'm_strCompDate = "" & Me.Text7(3).Text
         '91.11.19 END
      End If
   End If
End Sub

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String, bolTmp As Boolean
   ChgType = False
   Select Case i
        'Add By Cheng 2002/12/10
        Case 2 '案件性質
           If pa(9) = 台灣國家代號 Then
              bolTmp = False
           Else
              bolTmp = True
           End If
           'edit by nickc 2007/02/02 不用 dll 了
           'If objPublicData.GetCaseProperty(pA(1), Text7(i), strTempName, BolTmp) Then
           If ClsPDGetCaseProperty(pa(1), Text7(i), strTempName, bolTmp) Then
              Label3(5) = strTempName
              ChgType = True
           End If
      Case 12 '代理人
         strExc(1) = Combo2.Text
         'Modify By Cheng 2002/07/08
         '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'         If objPublicData.GetAgent(strExc(1), strTempName) Then
         If PUB_GetAgentName(pa(1), strExc(1), strTempName) Then
            Combo2.Text = strExc(1)
            Label3(6) = strTempName
            ChgType = True
         Else
            Label3(6) = ""
         End If
      Case 26, 27, 28, 29, 30 '申請人
         strExc(1) = Text7(i - 22).Text
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GETCUSTOMER(strExc(1), strTempName) Then
         If ClsPDGetCustomer(strExc(1), strTempName) Then
            Text7(i - 22).Text = strExc(1)
            Label3(i - 17) = strTempName
            ChgType = True
         Else
            Label3(i - 17) = ""
         End If
   End Select
End Function

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text6_GotFocus()
   TextInverse Text6
End Sub

Private Sub Text7_GotFocus(Index As Integer)
   TextInverse Text7(Index)
End Sub

Private Sub Text7_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 1, 4, 5, 6, 7, 8
         KeyAscii = UpperCase(KeyAscii)
      Case 10, 13, 23
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 19
         KeyAscii = UpperCase(KeyAscii)
         If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 9, 20, 21
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub Text7_Validate(Index As Integer, Cancel As Boolean)
 Dim i As Integer
   Select Case Index
      'Add By Cheng 2002/12/10
      Case 2 '案件性質
         If Text7(Index) <> "" Then
            '若案件性質已改變
            If cp(10) <> Text7(Index) Then
               Select Case Text7(Index)
                  Case "301", "302", "303", "304", "305", "306", "307", "803"
                     If ChgType(Index) Then
                        Command2(1).Enabled = False
                     Else
                        Cancel = True
                     End If
                  Case Else
                     MsgBox "案件性質只可為改請程序之案件性質或舉發 !", vbCritical
                     Text7(Index) = cp(10)
                     Cancel = True
               End Select
            Else
               ' 90.08.29 modify by louis
               If IsEmptyText(Text7(2)) = False Then
                  GetCaseFee pa(1), pa(9), Text7(2)
               End If
            End If
         Else
            MsgBox "案件性質不可空白 !", vbCritical
            Cancel = True
         End If
      Case 3 '提申期限
         If Text7(Index) <> "" Then
            If Not ChkDate(Text7(Index)) Then
               Cancel = True
            End If
         End If
      Case 4, 5, 6, 7, 8 '申請人
         If Text7(Index) <> "" Then
            If Not ChgType(Index + 22) Then Cancel = True
         Else
            Label3(Index + 5) = ""
         End If
         'Add By Cheng 2002/08/22
         If Cancel = False Then
            Select Case Index
            Case 4
               If Me.Text7(Index).Text <> m_strCust1 Then
                  If Not PUB_EditCustOk(Me.Label3(0).Caption, pa(1), pa(2), pa(3), pa(4)) Then Cancel = True
               End If
            Case 5
               If Me.Text7(Index).Text <> m_strCust2 Then
                  If Not PUB_EditCustOk(Me.Label3(0).Caption, pa(1), pa(2), pa(3), pa(4)) Then Cancel = True
               End If
            Case 6
               If Me.Text7(Index).Text <> m_strCust3 Then
                  If Not PUB_EditCustOk(Me.Label3(0).Caption, pa(1), pa(2), pa(3), pa(4)) Then Cancel = True
               End If
            Case 7
               If Me.Text7(Index).Text <> m_strCust4 Then
                  If Not PUB_EditCustOk(Me.Label3(0).Caption, pa(1), pa(2), pa(3), pa(4)) Then Cancel = True
               End If
            Case 8
               If Me.Text7(Index).Text <> m_strCust5 Then
                  If Not PUB_EditCustOk(Me.Label3(0).Caption, pa(1), pa(2), pa(3), pa(4)) Then Cancel = True
               End If
            End Select
         End If
         
      Case 16
         If Text7(14) = "" And Text7(15) = "" And Text7(16) = "" Then
            MsgBox "案件名稱不可同時空白 !", vbCritical
            Text7(14).SetFocus
         End If
   End Select
   If Cancel = True Then
      TextInverse Text7(Index)
      Text7(Index).SetFocus 'Added by Morgan 2021/12/13
   End If
End Sub

'Add By Cheng 2002/03/08
Private Function CheckDataIntegrity() As Boolean
Dim Cancel As Boolean
Cancel = False
'Add by Morgan 2007/12/26
If iFrom <> 1 Then
'end 2007/12/26
   If Me.Combo2.Text = "" Then
      MsgBox "代理人欄位不可空白!!!", vbExclamation + vbOKOnly
      Me.Combo2.SetFocus
      GoTo IntegrityOrNot
   End If
   
   '檢查代理人欄位
   Combo2_Validate Cancel
   If Cancel = True Then
      Me.Combo2.SetFocus
      GoTo IntegrityOrNot
   End If
End If
CheckDataIntegrity = True
Exit Function

IntegrityOrNot:
CheckDataIntegrity = False
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

   'Added by Morgan 2021/12/13 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/13

For Each objTxt In Text7
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

'Add by Morgan 2004/9/14
If Combo2.Enabled = True Then
   Cancel = False
   Combo2_Validate Cancel
   If Cancel = True Then
      Combo2.SetFocus
      Exit Function
   End If
End If

'Add by Morgan 2009/12/17
'若勾選要順稿則一定要輸順稿期限
If Check3.Value = 1 Then
   If txtDate = "" Then
      MsgBox "若有勾選要順稿則一定要輸順稿期限！"
      txtDate.SetFocus
      Exit Function
   Else
      txtDate_Validate Cancel
      If Cancel = True Then
         txtDate.SetFocus
         Exit Function
      End If
   End If
End If
'end 2009/12/17

'Added by Morgan 2016/1/4
'大陸案申請人國籍必須為臺灣才可主張臺灣優先權--郭
'Modified by Morgan 2018/1/24 +大陸、香港、澳門國籍也可以可主張臺灣優先權 P-119443--郭
If pa(9) = "020" And InStr(strPriority(1), "000") > 0 Then
   'Modified by Morgan 2025/9/11 舊提醒取消改用新的
   'strExc(1) = ""
   'For ii = 0 To 4
   '   If pa(26 + ii) <> "" Then
   '      strExc(2) = Left(pa(26 + ii) & "000", 9)
   '      strExc(0) = "select cu01||cu02 from customer where cu01='" & Left(strExc(2), 8) & "' and cu02='" & Mid(strExc(2), 9) & "' and cu10>'010' and cu10<>'020' and cu10<>'044' and cu10<>'013'"
   '      intI = 1
   '      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   '      If intI = 1 Then
   '         strExc(1) = "Y"
   '         Exit For
   '      End If
   '   End If
   'Next
   'If strExc(1) = "Y" Then
   '   MsgBox "大陸案申請人國籍必須為臺灣才可主張臺灣優先權！", vbCritical
   '   Exit Function
   'End If
   If PUB_ChkCNTWPriority(pa(), strPriority(1), strPriority(3)) = False Then
      If PUB_CNPriorityMsg() = vbNo Then
         Exit Function
      End If
   End If
   'end 2025/9/11
End If
'end 2016/1/4

'Added by Morgan 2021/11/1
'台灣案同日提申檢查
m_TWCP09 = ""
If Check4.Value = vbChecked Then
   If Text7(0).Text = "" Then
      MsgBox "勾選【" & Check4.Caption & "】必須輸入指定提申日！", vbCritical
      Text7(0).SetFocus
      Exit Function
   End If
   
   '檢查未發文的台灣案發明/新型/設計/衍生設計申請案
   'Modified by Morgan 2023/12/28 +抓國外案(台灣可能後收文) Ex:P-132614
   'Modified by Morgan 2025/2/14 +cp01,cp02,cp03,cp04
   strExc(0) = "select cp09,cp01,cp02,cp03,cp04 from (" & _
      " select cm05,cm06,cm07,cm08 from casemap where cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "' and cm10='0'" & _
      " union select cm01,cm02,cm03,cm04 from casemap where cm05='" & pa(1) & "' and cm06='" & pa(2) & "' and cm07='" & pa(3) & "' and cm08='" & pa(4) & "' and cm10='0'" & _
      ") X,caseprogress,patent where cp01(+)=cm05 and cp02(+)=cm06 and cp03(+)=cm07 and cp04(+)=cm08" & _
      " and cp27||cp57 is null and cp10 in (101,102,103,125) and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='000'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_TWCP09 = RsTemp("cp09")
      'Added by Morgan 2025/2/14
      m_TWCP01 = RsTemp("cp01")
      m_TWCP02 = RsTemp("cp02")
      m_TWCP03 = RsTemp("cp03")
      m_TWCP04 = RsTemp("cp04")
      'end 2025/2/14
   Else
      MsgBox "沒有未發文的台灣案【發明/新型/設計/衍生設計】申請程序，不應勾選【" & Check4.Caption & "】！", vbCritical
      Exit Function
   End If
End If
'end 2021/11/1

'Added by Morgan 2023/8/1
If pa(1) = "P" And pa(9) = "020" And (cp(10) = "701" Or cp(10) = "708") Then
   If PUB_CNP701708Check(cp(55) & "," & cp(93) & "," & cp(94) & "," & cp(95) & "," & cp(96), cp(56) & "," & cp(89) & "," & cp(90) & "," & cp(91) & "," & cp(92), True) = False Then
      Exit Function
   End If
End If
'end 2023/8/1

TxtValidate = True
End Function

'Add By Cheng 2002/06/24
Private Function GetCF15(strCF01 As String, strCF02 As String, strCF03 As String) As Boolean
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetCF15 = False
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
StrSQLa = "Select CF15 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   If IsNull(rsA.Fields(0).Value) = False Then
      GetCF15 = False
   End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2003/08/06
'取得申請人地址(縣市)
Private Sub GetCustAddress(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String)
Dim ii As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim blnTaiwan As Boolean '申請人國籍是否為台灣

   blnTaiwan = False
   For ii = 1 To 5
       m_strAdd(ii) = ""
   Next ii
   For ii = 1 To 10
       m_strIAdd(ii) = ""
   Next ii
   'Modified by Lydia 2022/06/07 +英文地址
   StrSQLa = "Select '1', CU23, CU10, CU24, CU25, CU26, CU27, CU28, CU102 From Customer, Patent Where CU01=substr(PA26,1,8) And CU02=substr(PA26,9,1) And " & ChgPatent(strPA01 & strPA02 & strPA03 & strPA04)
   StrSQLa = StrSQLa & " Union Select '2', CU23, CU10, CU24, CU25, CU26, CU27, CU28, CU102 From Customer, Patent Where CU01=substr(PA27,1,8) And CU02=substr(PA27,9,1) And " & ChgPatent(strPA01 & strPA02 & strPA03 & strPA04)
   StrSQLa = StrSQLa & " Union Select '3', CU23, CU10, CU24, CU25, CU26, CU27, CU28, CU102 From Customer, Patent Where CU01=substr(PA28,1,8) And CU02=substr(PA28,9,1) And " & ChgPatent(strPA01 & strPA02 & strPA03 & strPA04)
   StrSQLa = StrSQLa & " Union Select '4', CU23, CU10, CU24, CU25, CU26, CU27, CU28, CU102 From Customer, Patent Where CU01=substr(PA29,1,8) And CU02=substr(PA29,9,1) And " & ChgPatent(strPA01 & strPA02 & strPA03 & strPA04)
   StrSQLa = StrSQLa & " Union Select '5', CU23, CU10, CU24, CU25, CU26, CU27, CU28, CU102 From Customer, Patent Where CU01=substr(PA30,1,8) And CU02=substr(PA30,9,1) And " & ChgPatent(strPA01 & strPA02 & strPA03 & strPA04)
   StrSQLa = StrSQLa & " Order By 1 "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       While Not rsA.EOF
           '若為第一個申請人
           'Modified by Morgan 2020/4/17 改各自判斷--品薇 Ex:P124104
           'If "" & rsA.Fields(0).Value = "1" Then
               blnTaiwan = False
           'end 2020/4/17
               '若國籍為台灣
               If "" & rsA.Fields(2).Value < "010" Then
                   blnTaiwan = True
               End If
               
           'End If 'Removed by Morgan 2020/4/17
           '若國籍為台灣
           'Modify by Morgan 2005/5/25 PCT要完整地址
           'If blnTaiwan = True Then
           If blnTaiwan = True And Text7(2) <> "109" Then
               If InStr("" & rsA.Fields(1).Value, "市") > 0 Then
                   m_strAdd(rsA.Fields(0).Value) = Left("" & rsA.Fields(1).Value, InStr("" & rsA.Fields(1).Value, "市"))
               ElseIf InStr("" & rsA.Fields(1).Value, "縣") > 0 Then
                   m_strAdd(rsA.Fields(0).Value) = Left("" & rsA.Fields(1).Value, InStr("" & rsA.Fields(1).Value, "縣"))
               End If
           '若國籍非台灣
           Else
               m_strAdd(rsA.Fields(0).Value) = "" & rsA.Fields(1).Value
           End If
           rsA.MoveNext
       Wend
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   'Add By Cheng 2003/08/11
   '若申請人國籍非台灣
   '2009/4/10 MODIFY BY SONIA 不管申請人國籍,發明人國籍台灣者帶中華民國,非台灣者帶地址
   'If blnTaiwan = False Then
      'Modify By Sindy 2014/11/12
      If strSrvDate(1) >= 專利發明人檔啟用日 Then
         m_strIAdd(1) = ""
         StrSQLa = "Select pi05,IN07,IN11,na03 From Inventor,PatentInventor,nation" & _
                   " Where pi01='" & strPA01 & "' And pi02='" & strPA02 & "' And pi03='" & strPA03 & "' And pi04='" & strPA04 & "'" & _
                   " And substr(Pi06,1,8)=IN01(+) And substr(Pi06,9,2)=IN02(+)" & _
                   " And in11=na01(+)" & _
                   " order by pi05 asc"
      Else
      '2014/11/12 END
         'Memo by Lydia 2021/08/17 刪除舊程式碼：專利發明人在專利基本檔60~69
      End If
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      strExc(10) = 0 'Add By Sindy 2018/5/28
      If rsA.RecordCount > 0 Then
         rsA.MoveFirst
         While Not rsA.EOF
            strExc(10) = strExc(10) + 1 'Add By Sindy 2018/5/28
            '2009/4/10 MODIFY BY SONIA 發明人國籍台灣者帶中華民國,非台灣者帶地址
            'If "" & rsA.Fields(0).Value = "1" Then
            '   m_strIAdd(rsA.Fields(0).Value) = "　　　　地　　　　　址：" & rsA.Fields(1).Value
            'ElseIf "" & rsA.Fields(0).Value = "2" Then
            '   m_strIAdd(rsA.Fields(0).Value) = "" & rsA.Fields(1).Value
            'End If
            'Modify By Sindy 2014/11/24
            If strSrvDate(1) >= 專利發明人檔啟用日 Then
               'Add By Sindy 2018/5/28 FMP抓國籍
               If PUB_ChkIsFMP(strPA01, strPA02, strPA03, strPA04) = True Then
                  If "" & rsA.Fields(2).Value < "010" Then
                     If m_strIAdd(1) = "" Then
                        m_strIAdd(1) = IIf(rsA.RecordCount > 1, strExc(10) & ".", "") & "中華民國"
                     Else
                        'm_strIAdd(1) = m_strIAdd(1) & vbCrLf & "　　　　　　　　　　　　中華民國"
                        m_strIAdd(1) = m_strIAdd(1) & vbCrLf & IIf(rsA.RecordCount > 1, strExc(10) & ".", "") & "中華民國"
                     End If
                  Else
                     If m_strIAdd(1) = "" Then
                        m_strIAdd(1) = IIf(rsA.RecordCount > 1, strExc(10) & ".", "") & rsA.Fields(3).Value
                     Else
                        m_strIAdd(1) = m_strIAdd(1) & vbCrLf & IIf(rsA.RecordCount > 1, strExc(10) & ".", "") & rsA.Fields(3).Value
                     End If
                  End If
               '抓中文地址
               Else
                  If "" & rsA.Fields(2).Value < "010" Then
                     If m_strIAdd(1) = "" Then
                        m_strIAdd(1) = "中華民國"
                     Else
                        'm_strIAdd(1) = m_strIAdd(1) & vbCrLf & "　　　　　　　　　　　　中華民國"
                        m_strIAdd(1) = m_strIAdd(1) & vbCrLf & "中華民國"
                     End If
                  Else
                     If m_strIAdd(1) = "" Then
                        m_strIAdd(1) = "" & rsA.Fields(1).Value 'Modified by Morgan 2018/7/12 沒地址不用帶--品薇 Ex:P-120361
                     Else
                        m_strIAdd(1) = m_strIAdd(1) & vbCrLf & rsA.Fields(1).Value
                     End If
                  End If
               End If
            Else
            '2014/11/24 END
               If "" & rsA.Fields(2).Value < "010" Then
                  m_strIAdd(rsA.Fields(0).Value) = "中華民國"
               Else
                  m_strIAdd(rsA.Fields(0).Value) = "" & rsA.Fields(1).Value
               End If
               '2009/4/10 END
            End If
            rsA.MoveNext
         Wend
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   'End If
End Sub
'Add by Morgan 2006/6/6 讀出字串內某提示辭後的值
'p_Content:字串,p_PreWord:提示辭,p_SplitChar:分隔字元
Private Function GetValue(p_Content As String, p_PreWord As String, p_SplitChar As String) As String
   Dim iPos1 As Integer, iPos2 As Integer, sRtn As String
   iPos1 = InStr(p_Content, p_PreWord)
   If iPos1 > 0 Then
      iPos2 = InStr(iPos1, p_Content, p_SplitChar)
      If iPos2 > iPos1 + Len(p_PreWord) Then
         GetValue = Mid(p_Content, iPos1 + Len(p_PreWord), iPos2 - iPos1 - Len(p_PreWord))
      End If
   End If
End Function

Private Sub Check3_Click()
   If Check3.Value = 0 Then
      txtDate = ""
   End If
End Sub

Private Sub txtDate_GotFocus()
   TextInverse txtDate
   CloseIme
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
   Dim strTmp As String
   If txtDate <> "" Then
      If ChkDate(txtDate) = False Then
         Cancel = True
      Else
         If Text7(0) <> "" And Val(DBDATE(txtDate)) > Val(DBDATE(Text7(0))) Then
            MsgBox "順稿期限不可大於指定提申期限！"
            Cancel = True
         ElseIf Text7(3) <> "" And Val(DBDATE(txtDate)) > Val(DBDATE(Text7(3))) Then
            MsgBox "順稿期限不可大於最終提申期限！"
            Cancel = True
         End If
      End If
   End If
End Sub
