VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm05010401_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "一般來函"
   ClientHeight    =   5772
   ClientLeft      =   276
   ClientTop       =   936
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form25"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   8520
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   6480
      TabIndex        =   62
      Top             =   60
      Width           =   1170
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   345
      Index           =   0
      Left            =   5655
      TabIndex        =   61
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   7665
      TabIndex        =   63
      Top             =   60
      Width           =   825
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4152
      Left            =   144
      TabIndex        =   44
      Top             =   1632
      Width           =   8292
      _ExtentX        =   14626
      _ExtentY        =   7324
      _Version        =   393216
      TabsPerRow      =   4
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm05010401_3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFee1s"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFee2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblProperty"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label19"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label28"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label29"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblNextCaseProperty"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label31"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label32"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label16"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label111"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label18"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblPromoter"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label10"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblAno"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label15"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label17"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label21"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label7"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label30"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label34"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label35"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label36"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label40"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label44"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblCaseField(30)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblFee1"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label12(0)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label12(1)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label12(3)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label12(2)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label12(4)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label12(5)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label12(6)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label12(7)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label12(8)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label12(9)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label48"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label50"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label49"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label47"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "lblEno"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Label12(11)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Label12(10)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtCaseField(3)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtCaseField(1)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtCaseField(16)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtCaseField(2)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtCaseField(8)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtCaseField(4)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtCaseField(5)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtCaseField(6)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "txtCaseField(0)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtCaseField(9)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtCaseField(13)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtCaseField(15)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txtCaseField(14)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtCaseField(12)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtCaseField(17)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txtCaseField(25)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txtCaseField(29)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txtCaseField(30)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txtCaseField(10)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txtCaseField(7)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txtCaseField(11)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txtCaseField(33)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txtCaseField(31)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txtCaseField(32)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txtCaseField(34)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "txtCaseField(24)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "txtCaseField(35)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "txtCaseField(36)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "txtCaseField(37)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "txtCaseField(38)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "txtCaseField(26)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "txtCaseField(41)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "txtCaseField(40)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "txtCaseField(42)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "txtCaseField(43)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "Frame2"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "cmdCountry"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "Text8"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "Text7"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "Text5(1)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "Text5(0)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "txtAbandonFee"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "txtFiles"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).Control(89)=   "Text37"
      Tab(0).Control(89).Enabled=   0   'False
      Tab(0).ControlCount=   90
      TabCaption(1)   =   "對造／國外ID／其他"
      TabPicture(1)   =   "frm05010401_3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(1)=   "Label23"
      Tab(1).Control(2)=   "Label24"
      Tab(1).Control(3)=   "Label25"
      Tab(1).Control(4)=   "Label26"
      Tab(1).Control(5)=   "Label27"
      Tab(1).Control(6)=   "Label37"
      Tab(1).Control(7)=   "Label43"
      Tab(1).Control(8)=   "Label33"
      Tab(1).Control(9)=   "Label42"
      Tab(1).Control(10)=   "Label46"
      Tab(1).Control(11)=   "txtCaseField(18)"
      Tab(1).Control(12)=   "txtCaseField(19)"
      Tab(1).Control(13)=   "txtCaseField(20)"
      Tab(1).Control(14)=   "txtCaseField(21)"
      Tab(1).Control(15)=   "txtCaseField(22)"
      Tab(1).Control(16)=   "txtCaseField(23)"
      Tab(1).Control(17)=   "txtCaseField(39)"
      Tab(1).Control(18)=   "Label52"
      Tab(1).Control(19)=   "Label53"
      Tab(1).Control(20)=   "txtPA22"
      Tab(1).Control(21)=   "txtPA14"
      Tab(1).Control(22)=   "txtPA15"
      Tab(1).Control(23)=   "CmdAFID03(0)"
      Tab(1).Control(24)=   "CmdAFID03(1)"
      Tab(1).Control(25)=   "CmdAFID03(2)"
      Tab(1).Control(26)=   "CmdAFID03(3)"
      Tab(1).Control(27)=   "CmdAFID03(4)"
      Tab(1).Control(28)=   "txtIDSPt(2)"
      Tab(1).Control(29)=   "txtIDSFee(2)"
      Tab(1).Control(30)=   "txtIDSPt(1)"
      Tab(1).Control(31)=   "txtIDSFee(1)"
      Tab(1).ControlCount=   32
      TabCaption(2)   =   "修圖資料"
      TabPicture(2)   =   "frm05010401_3.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtCaseField(28)"
      Tab(2).Control(1)=   "txtCaseField(27)"
      Tab(2).Control(2)=   "Label39"
      Tab(2).Control(3)=   "Label38"
      Tab(2).ControlCount=   4
      Begin VB.TextBox txtIDSFee 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   1
         Left            =   -68328
         MaxLength       =   6
         TabIndex        =   54
         Top             =   2808
         Width           =   765
      End
      Begin VB.TextBox txtIDSPt 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   1
         Left            =   -67404
         MaxLength       =   3
         TabIndex        =   55
         Top             =   2808
         Width           =   375
      End
      Begin VB.TextBox txtIDSFee 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   2
         Left            =   -68328
         MaxLength       =   6
         TabIndex        =   56
         Top             =   3108
         Width           =   765
      End
      Begin VB.TextBox txtIDSPt 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   2
         Left            =   -67404
         MaxLength       =   3
         TabIndex        =   57
         Top             =   3108
         Width           =   375
      End
      Begin VB.TextBox Text37 
         Height          =   270
         Left            =   6840
         MaxLength       =   6
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   3870
         Width           =   735
      End
      Begin VB.TextBox txtFiles 
         Height          =   270
         Left            =   7680
         MaxLength       =   2
         TabIndex        =   145
         Top             =   900
         Width           =   375
      End
      Begin VB.CommandButton CmdAFID03 
         Caption         =   "申5"
         Height          =   270
         Index           =   4
         Left            =   -71040
         TabIndex        =   144
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton CmdAFID03 
         Caption         =   "申4"
         Height          =   270
         Index           =   3
         Left            =   -71540
         TabIndex        =   143
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton CmdAFID03 
         Caption         =   "申3"
         Height          =   270
         Index           =   2
         Left            =   -72050
         TabIndex        =   142
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton CmdAFID03 
         Caption         =   "申2"
         Height          =   270
         Index           =   1
         Left            =   -72550
         TabIndex        =   141
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton CmdAFID03 
         Caption         =   "申請人1"
         Height          =   270
         Index           =   0
         Left            =   -73380
         TabIndex        =   140
         Top             =   2160
         Width           =   825
      End
      Begin VB.TextBox txtPA15 
         Height          =   270
         Left            =   -70620
         MaxLength       =   20
         TabIndex        =   52
         Top             =   2460
         Width           =   2535
      End
      Begin VB.TextBox txtPA14 
         Height          =   270
         Left            =   -73380
         MaxLength       =   8
         TabIndex        =   51
         Top             =   2460
         Width           =   972
      End
      Begin VB.TextBox txtAbandonFee 
         Height          =   285
         Left            =   7230
         TabIndex        =   27
         Top             =   2310
         Width           =   825
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   6165
         MaxLength       =   2
         TabIndex        =   19
         Top             =   1470
         Width           =   285
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   6705
         MaxLength       =   2
         TabIndex        =   20
         Top             =   1470
         Width           =   285
      End
      Begin VB.TextBox txtPA22 
         Height          =   270
         Left            =   -73380
         MaxLength       =   20
         TabIndex        =   53
         Top             =   2790
         Width           =   2535
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1305
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   32
         Top             =   2595
         Width           =   285
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   7515
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   15
         Top             =   1185
         Width           =   285
      End
      Begin VB.CommandButton cmdCountry 
         Caption         =   "指定國家"
         Height          =   285
         Left            =   4500
         TabIndex        =   42
         Top             =   3435
         Width           =   1005
      End
      Begin VB.Frame Frame2 
         Height          =   375
         Left            =   1050
         TabIndex        =   151
         Top             =   840
         Width           =   4275
         Begin VB.TextBox Text12 
            Height          =   270
            Left            =   2970
            MaxLength       =   7
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   75
            Width           =   975
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Left            =   840
            MaxLength       =   2
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   75
            Width           =   375
         End
         Begin VB.TextBox Text11 
            Height          =   270
            Left            =   1950
            MaxLength       =   2
            TabIndex        =   9
            Top             =   75
            Width           =   375
         End
         Begin VB.OptionButton Option4 
            Caption         =   "文到           天"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton Option4 
            Caption         =   "        月"
            Height          =   180
            Index           =   1
            Left            =   1710
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   120
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton Option4 
            Caption         =   "                      日"
            Height          =   225
            Index           =   2
            Left            =   2700
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   105
            Width           =   1515
         End
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "ＩＤＳ報價:  1. 第一階段                  (           P)"
         Height          =   180
         Left            =   -70344
         TabIndex        =   160
         Top             =   2856
         Width           =   3552
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "2. 第二階段                  (           P)"
         Height          =   180
         Left            =   -69288
         TabIndex        =   159
         Top             =   3156
         Width           =   2508
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   43
         Left            =   7740
         TabIndex        =   157
         Top             =   2010
         Width           =   330
         VariousPropertyBits=   671107099
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   42
         Left            =   6780
         TabIndex        =   156
         Top             =   2010
         Width           =   825
         VariousPropertyBits=   671107099
         Size            =   "1455;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   40
         Left            =   6780
         TabIndex        =   153
         Top             =   1740
         Width           =   825
         VariousPropertyBits=   671107099
         Size            =   "1455;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   41
         Left            =   7740
         TabIndex        =   154
         Top             =   1740
         Width           =   330
         VariousPropertyBits=   671107099
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   26
         Left            =   5130
         TabIndex        =   34
         Top             =   2760
         Width           =   2190
         VariousPropertyBits=   671107099
         Size            =   "3863;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   900
         Index           =   39
         Left            =   -73968
         TabIndex        =   58
         Top             =   3120
         Width           =   4500
         VariousPropertyBits=   -1467987941
         MaxLength       =   170
         ScrollBars      =   2
         Size            =   "7937;1587"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   38
         Left            =   5445
         TabIndex        =   24
         Top             =   1740
         Width           =   330
         VariousPropertyBits=   671107099
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   37
         Left            =   2310
         TabIndex        =   22
         Top             =   1740
         Width           =   330
         VariousPropertyBits=   671107099
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   36
         Left            =   4050
         TabIndex        =   23
         Top             =   1740
         Width           =   825
         VariousPropertyBits=   671107099
         Size            =   "1455;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   35
         Left            =   900
         TabIndex        =   21
         Top             =   1740
         Width           =   825
         VariousPropertyBits=   671107099
         Size            =   "1455;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   24
         Left            =   4050
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1470
         Width           =   825
         VariousPropertyBits=   671107103
         Size            =   "1455;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   34
         Left            =   5430
         TabIndex        =   31
         Top             =   2295
         Width           =   330
         VariousPropertyBits=   671107099
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   32
         Left            =   2310
         TabIndex        =   29
         Top             =   2295
         Width           =   330
         VariousPropertyBits=   671107099
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   31
         Left            =   900
         TabIndex        =   28
         Top             =   2295
         Width           =   825
         VariousPropertyBits=   671107099
         Size            =   "1455;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   33
         Left            =   4050
         TabIndex        =   30
         Top             =   2295
         Width           =   825
         VariousPropertyBits=   671107099
         Size            =   "1455;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   11
         Left            =   4725
         TabIndex        =   33
         Top             =   2595
         Width           =   3480
         VariousPropertyBits=   671107099
         Size            =   "6138;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   7
         Left            =   900
         TabIndex        =   25
         Top             =   2010
         Width           =   825
         VariousPropertyBits=   671107099
         Size            =   "1455;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   10
         Left            =   4050
         TabIndex        =   26
         Top             =   2010
         Width           =   825
         VariousPropertyBits=   671107103
         Size            =   "1455;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   30
         Left            =   5085
         TabIndex        =   14
         Top             =   1185
         Width           =   915
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1614;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   29
         Left            =   7335
         TabIndex        =   40
         Top             =   3150
         Width           =   285
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   28
         Left            =   -73530
         TabIndex        =   60
         Top             =   750
         Width           =   1245
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "2196;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   27
         Left            =   -73530
         TabIndex        =   59
         Top             =   450
         Width           =   1245
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "2196;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   25
         Left            =   6975
         TabIndex        =   5
         Top             =   570
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
         Index           =   23
         Left            =   -73380
         TabIndex        =   50
         Top             =   1860
         Width           =   6555
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "11562;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   405
         Index           =   17
         Left            =   1035
         TabIndex        =   43
         Top             =   3720
         Width           =   5730
         VariousPropertyBits=   -1467987941
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "10107;714"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   22
         Left            =   -73380
         TabIndex        =   49
         Top             =   1560
         Width           =   6555
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "11562;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   21
         Left            =   -73380
         TabIndex        =   48
         Top             =   1260
         Width           =   6555
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "11562;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   20
         Left            =   -73080
         TabIndex        =   47
         Top             =   960
         Width           =   6255
         VariousPropertyBits=   671107099
         MaxLength       =   100
         Size            =   "11033;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   19
         Left            =   -73080
         TabIndex        =   46
         Top             =   660
         Width           =   6255
         VariousPropertyBits=   671107099
         MaxLength       =   100
         Size            =   "11033;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   18
         Left            =   -73080
         TabIndex        =   45
         Top             =   360
         Width           =   6255
         VariousPropertyBits=   671107099
         MaxLength       =   100
         Size            =   "11033;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   12
         Left            =   900
         TabIndex        =   35
         Top             =   2880
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   6
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   14
         Left            =   1440
         TabIndex        =   38
         Top             =   3165
         Width           =   510
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "900;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   15
         Left            =   4140
         TabIndex        =   39
         Top             =   3150
         Width           =   1590
         VariousPropertyBits=   671107099
         Size            =   "2805;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   13
         Left            =   4140
         TabIndex        =   36
         Top             =   2880
         Width           =   1005
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1773;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   9
         Left            =   7335
         TabIndex        =   37
         Top             =   2880
         Width           =   285
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "503;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   0
         Left            =   1035
         TabIndex        =   1
         Top             =   285
         Width           =   690
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "1217;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   6
         Left            =   900
         TabIndex        =   16
         Top             =   1470
         Width           =   825
         VariousPropertyBits=   671107099
         Size            =   "1455;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   5
         Left            =   3780
         TabIndex        =   4
         Top             =   570
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
         Index           =   4
         Left            =   1215
         TabIndex        =   3
         Top             =   570
         Width           =   915
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1614;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   8
         Left            =   2310
         TabIndex        =   17
         Top             =   1470
         Width           =   330
         VariousPropertyBits=   671107099
         Size            =   "582;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   2
         Left            =   3105
         TabIndex        =   13
         Top             =   1185
         Width           =   915
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1614;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   16
         Left            =   1035
         TabIndex        =   41
         Top             =   3435
         Width           =   2940
         VariousPropertyBits=   671107099
         MaxLength       =   40
         Size            =   "5186;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   1
         Left            =   5175
         TabIndex        =   2
         Top             =   285
         Width           =   645
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "1138;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   3
         Left            =   1035
         TabIndex        =   12
         Top             =   1185
         Width           =   915
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1614;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "RCE報價：                   (         P)"
         Height          =   180
         Index           =   10
         Left            =   5910
         TabIndex        =   158
         Top             =   2055
         Width           =   2355
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "IDS報價：                   (         P)"
         Height          =   180
         Index           =   11
         Left            =   5970
         TabIndex        =   155
         Top             =   1785
         Width           =   2280
      End
      Begin VB.Label lblEno 
         Caption         =   "英國脫歐案專利號數："
         Height          =   180
         Left            =   3195
         TabIndex        =   152
         Top             =   2790
         Width           =   1845
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "來函期限："
         Height          =   180
         Left            =   135
         TabIndex        =   150
         Top             =   930
         Width           =   900
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "通知函判發人:"
         Height          =   180
         Left            =   6840
         TabIndex        =   149
         Top             =   3690
         Width           =   1125
      End
      Begin MSForms.Label Label50 
         Height          =   180
         Left            =   7650
         TabIndex        =   148
         Top             =   3900
         Width           =   600
         VariousPropertyBits=   27
         Caption         =   "XXX"
         Size            =   "1058;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "引證前案檔案數量："
         Height          =   180
         Left            =   5970
         TabIndex        =   146
         Top             =   930
         Width           =   1620
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "自動發證公告號："
         Height          =   180
         Left            =   -72120
         TabIndex        =   139
         Top             =   2505
         Width           =   1440
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "自動發證公告日："
         Height          =   180
         Left            =   -74880
         TabIndex        =   138
         Top             =   2505
         Width           =   1440
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "報價備註："
         Height          =   180
         Left            =   -74880
         TabIndex        =   137
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "點數："
         Height          =   180
         Index           =   9
         Left            =   4905
         TabIndex        =   136
         Top             =   1785
         Width           =   540
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "點數："
         Height          =   180
         Index           =   8
         Left            =   1770
         TabIndex        =   135
         Top             =   1785
         Width           =   540
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         Caption         =   "補未收費程序："
         Height          =   180
         Index           =   7
         Left            =   2760
         TabIndex        =   134
         Top             =   1785
         Width           =   1260
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         Caption         =   "補虧損："
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   133
         Top             =   1785
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         Caption         =   "代理人費用："
         Height          =   180
         Index           =   5
         Left            =   2940
         TabIndex        =   132
         Top             =   1515
         Width           =   1080
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "點數："
         Height          =   180
         Index           =   4
         Left            =   4905
         TabIndex        =   131
         Top             =   2340
         Width           =   540
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "點數："
         Height          =   180
         Index           =   2
         Left            =   1770
         TabIndex        =   130
         Top             =   2340
         Width           =   540
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         Caption         =   "修正費："
         Height          =   180
         Index           =   3
         Left            =   3300
         TabIndex        =   129
         Top             =   2340
         Width           =   720
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         Caption         =   "面詢費："
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   128
         Top             =   2340
         Width           =   735
      End
      Begin VB.Label Label12 
         Alignment       =   1  '靠右對齊
         Caption         =   "製圖費："
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   127
         Top             =   2055
         Width           =   735
      End
      Begin VB.Label lblFee1 
         AutoSize        =   -1  'True
         Caption         =   "領證費："
         Height          =   180
         Left            =   135
         TabIndex        =   125
         Top             =   1500
         Width           =   720
      End
      Begin VB.Label lblCaseField 
         Caption         =   "約定期限："
         Height          =   180
         Index           =   30
         Left            =   4140
         TabIndex        =   124
         Top             =   1230
         Width           =   900
      End
      Begin VB.Label Label44 
         Alignment       =   1  '靠右對齊
         Caption         =   "放棄專利權費："
         Height          =   180
         Left            =   5970
         TabIndex        =   123
         Top             =   2355
         Width           =   1260
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "自動發證證書號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   122
         Top             =   2835
         Width           =   1440
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "是否為證書勘誤：       (Y:是)"
         Height          =   180
         Left            =   5892
         TabIndex        =   120
         Top             =   3168
         Width           =   2244
      End
      Begin VB.Label Label39 
         Caption         =   "修圖法定期限："
         Height          =   255
         Left            =   -74880
         TabIndex        =   119
         Top             =   750
         Width           =   1455
      End
      Begin VB.Label Label38 
         Caption         =   "修圖本所期限："
         Height          =   255
         Left            =   -74880
         TabIndex        =   118
         Top             =   450
         Width           =   1455
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "國外ID號數："
         Height          =   180
         Left            =   -74880
         TabIndex        =   117
         Top             =   2205
         Width           =   1080
      End
      Begin VB.Label Label36 
         Caption         =   "是否修改通知函內容：         (Y:Word)"
         Height          =   180
         Left            =   5175
         TabIndex        =   116
         Top             =   615
         Width           =   2895
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "案件目前准駁:         (1:准 , 2:駁)"
         Height          =   195
         Left            =   135
         TabIndex        =   115
         Top             =   2640
         Width           =   2625
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "專利權是否存在：      (Y/N)"
         Height          =   180
         Left            =   6120
         TabIndex        =   114
         Top             =   1230
         Width           =   2115
      End
      Begin VB.Label Label30 
         Caption         =   "機關文號："
         Height          =   255
         Left            =   135
         TabIndex        =   113
         Top             =   3450
         Width           =   1005
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱（日）："
         Height          =   180
         Left            =   -74880
         TabIndex        =   112
         Top             =   1905
         Width           =   1440
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱（英）："
         Height          =   180
         Left            =   -74880
         TabIndex        =   111
         Top             =   1605
         Width           =   1440
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "對造名稱（中）："
         Height          =   180
         Left            =   -74880
         TabIndex        =   110
         Top             =   1305
         Width           =   1440
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱（日）："
         Height          =   180
         Left            =   -74880
         TabIndex        =   109
         Top             =   1005
         Width           =   1800
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱（英）："
         Height          =   180
         Left            =   -74880
         TabIndex        =   108
         Top             =   705
         Width           =   1800
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "對造案件名稱（中）："
         Height          =   180
         Left            =   -74880
         TabIndex        =   107
         Top             =   405
         Width           =   1800
      End
      Begin VB.Label Label7 
         Caption         =   "EPC："
         Height          =   255
         Left            =   4005
         TabIndex        =   106
         Top             =   3450
         Width           =   645
      End
      Begin VB.Label Label21 
         Caption         =   "是否算案件數：            （N：不算）"
         Height          =   255
         Left            =   135
         TabIndex        =   105
         Top             =   3180
         Width           =   2895
      End
      Begin VB.Label Label17 
         Caption         =   "是否列印通知函：        （N：不印）"
         Height          =   180
         Left            =   2295
         TabIndex        =   104
         Top             =   615
         Width           =   2820
      End
      Begin VB.Label Label15 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "點數："
         Height          =   180
         Left            =   1770
         TabIndex        =   103
         Top             =   1515
         Width           =   540
      End
      Begin VB.Label lblAno 
         Caption         =   "美國讓渡登記號："
         Height          =   180
         Left            =   3195
         TabIndex        =   102
         Top             =   2640
         Width           =   1440
      End
      Begin VB.Label Label10 
         Caption         =   "對造號數："
         Height          =   180
         Left            =   3210
         TabIndex        =   101
         Top             =   3150
         Width           =   900
      End
      Begin MSForms.Label lblPromoter 
         Height          =   255
         Left            =   2040
         TabIndex        =   100
         Top             =   2910
         Width           =   780
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "1376;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "核准通知書影本：     （Y/N）"
         Height          =   180
         Left            =   5895
         TabIndex        =   99
         Top             =   2925
         Width           =   2310
      End
      Begin VB.Label Label111 
         Alignment       =   1  '靠右對齊
         Caption         =   "讓渡費："
         Height          =   180
         Left            =   3300
         TabIndex        =   98
         Top             =   2055
         Width           =   720
      End
      Begin VB.Label Label14 
         Caption         =   "來函性質："
         Height          =   180
         Left            =   135
         TabIndex        =   97
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label16 
         Caption         =   "官方發文日："
         Height          =   180
         Left            =   135
         TabIndex        =   96
         Top             =   615
         Width           =   1080
      End
      Begin VB.Label Label32 
         Caption         =   "法定期限："
         Height          =   180
         Left            =   135
         TabIndex        =   95
         Top             =   1230
         Width           =   900
      End
      Begin VB.Label Label31 
         Alignment       =   1  '靠右對齊
         Caption         =   "本所期限："
         Height          =   180
         Left            =   2160
         TabIndex        =   94
         Top             =   1230
         Width           =   900
      End
      Begin MSForms.Label lblNextCaseProperty 
         Height          =   255
         Left            =   5895
         TabIndex        =   93
         Top             =   300
         Width           =   2310
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "4075;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label29 
         Caption         =   "承辦人："
         Height          =   255
         Left            =   135
         TabIndex        =   92
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label Label28 
         Caption         =   "下一程序："
         Height          =   180
         Left            =   4185
         TabIndex        =   91
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label19 
         Caption         =   "承辦期限："
         Height          =   180
         Left            =   3210
         TabIndex        =   90
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label Label9 
         Caption         =   "進度備註："
         Height          =   255
         Left            =   135
         TabIndex        =   89
         Top             =   3690
         Width           =   1005
      End
      Begin MSForms.Label lblProperty 
         Height          =   255
         Left            =   1800
         TabIndex        =   88
         Top             =   300
         Width           =   2130
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "3757;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblFee2 
         Caption         =   " (其他費用,含第        至        年年費)"
         Height          =   180
         Left            =   4905
         TabIndex        =   87
         Top             =   1515
         Width           =   2730
      End
      Begin VB.Label lblFee1s 
         BackColor       =   &H80000010&
         Height          =   180
         Left            =   180
         TabIndex        =   126
         Top             =   1560
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin MSForms.ComboBox cboCaseName 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1080
      TabIndex        =   0
      Top             =   420
      Width           =   7350
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "12965;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label41 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3240
      TabIndex        =   121
      Top             =   1350
      Width           =   1005
   End
   Begin MSForms.Label lblNation 
      Height          =   255
      Left            =   6030
      TabIndex        =   86
      Top             =   750
      Width           =   2400
      VariousPropertyBits=   27
      Size            =   "4233;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   9
      Left            =   5310
      TabIndex        =   85
      Top             =   750
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "申請國家："
      Height          =   195
      Left            =   4320
      TabIndex        =   84
      Top             =   750
      Width           =   915
   End
   Begin MSForms.Label lblSales 
      Height          =   255
      Left            =   6120
      TabIndex        =   83
      Top             =   1350
      Width           =   2310
      VariousPropertyBits=   27
      Size            =   "4075;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblAgent 
      Height          =   255
      Left            =   1800
      TabIndex        =   82
      Top             =   750
      Width           =   2445
      VariousPropertyBits=   27
      Size            =   "4313;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   6
      Left            =   5310
      TabIndex        =   81
      Top             =   1050
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   80
      Top             =   450
      Width           =   915
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   195
      Left            =   135
      TabIndex        =   79
      Top             =   750
      Width           =   735
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   2
      Left            =   945
      TabIndex        =   78
      Top             =   750
      Width           =   780
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   8
      Left            =   1305
      TabIndex        =   77
      Top             =   1350
      Width           =   1095
   End
   Begin MSForms.Label lblCaseProperty 
      Height          =   255
      Left            =   6030
      TabIndex        =   76
      Top             =   1050
      Width           =   2400
      VariousPropertyBits=   27
      Size            =   "4233;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   0
      Left            =   915
      TabIndex        =   75
      Top             =   120
      Width           =   1920
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   4
      Left            =   945
      TabIndex        =   74
      Top             =   1050
      Width           =   1365
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   5
      Left            =   3240
      TabIndex        =   73
      Top             =   1050
      Width           =   1005
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   7
      Left            =   5310
      TabIndex        =   72
      Top             =   1350
      Width           =   735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   0
      Left            =   4320
      TabIndex        =   71
      Top             =   1350
      Width           =   900
   End
   Begin VB.Label lblIssue 
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   195
      Left            =   2430
      TabIndex        =   70
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "本所號："
      Height          =   255
      Index           =   0
      Left            =   135
      TabIndex        =   69
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號："
      Height          =   195
      Left            =   135
      TabIndex        =   68
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日："
      Height          =   195
      Left            =   135
      TabIndex        =   67
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   195
      Index           =   0
      Left            =   4320
      TabIndex        =   66
      Top             =   1050
      Width           =   915
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請號："
      Height          =   255
      Left            =   2910
      TabIndex        =   65
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblCaseField 
      Height          =   255
      Index           =   1
      Left            =   3660
      TabIndex        =   64
      Top             =   120
      Width           =   1965
   End
End
Attribute VB_Name = "frm05010401_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/6 改成Form2.0 (txtCaseField,cboCaseName,lblAgent,lblSales...
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'重整 by Morgan 2005/11/11
'Modify by Morgan 2008/5/12 公開費已不再輸入,因畫面空間有限，故該欄位亦不保留
Option Explicit

'intLastRow上一次反白的Row,blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'intLeaveKind離開時，2:結束  1:回上一畫面  0:確定
Dim intLeaveKind As Integer
'StrCountry存放指定國家  strMoneyCountry存放繳費國家 strMoney存放費用
Dim strCountry As String, strMoneyCountry As String, strMoney As String
Dim strFagentNo As String 'Added by Morgan 2020/8/7
' 系統別
Dim m_PA01 As String
' 國家
Dim m_PA09 As String
Dim m_PA08 As String '專利種類
Dim m_PA26 As String '申請人1
Dim m_intTab As Integer '記錄頁籤值
Dim m_NP07 As String
Dim bolDo As Boolean
Dim m_blnClosed As Boolean '是否閉卷
Dim m_blnCancelClosed As Boolean '是否取消閉卷
Dim m_strCloseDate As String '閉卷日期
Dim m_NewCP09 As String
Dim m_varTemp
Dim m_i  As Integer
Dim m_dobDateAdd  As Double
Dim m_strStartDate As String
Dim m_blnCompNextDate As Boolean '是否繼續計算下一次的期限
Dim m_strDate As String
Dim m_strDate1 As String
Dim m_blnCustReturnSheet As Boolean '判斷是否列印案件回覆單
Dim bolOnlyCustReturnSheet As Boolean '判斷是否只列印案件回覆單
Dim stCP09 As String, stCP14 As String, stCP27 As String
Dim m_ET02 As String
Dim m_bolEPC7Up As Boolean '是否超過7個成員國
Dim m_strRestEPCMember As String '其他為指定的成員國
Dim m_bolSaveCheck As Boolean '判斷是否是存檔檢查
'一案兩請之新型相關資料
Dim m_bolIsDualApp As Boolean, m_stCaseNo As String, m_stCertNo As String, m_stAppNo As String, m_stCaseName As String, m_stUPA(1 To 4) As String
Dim m_bolActive As Boolean 'Active事件是否已觸發
Dim m_bolIsNP107N As Boolean '是否有答辯不續辦
Dim m_strNP07 As String '不續辦案件性質 Added by Morgan 2013/4/18
Dim m_strRetSheet2NP07 As String '第二張回覆單案件性質
'Add by Morgan 2006/3/27
Dim m_str222MailCP14 As String '告建議性處分承辦人
Dim m_str222MailCP09 As String '告建議性處分收文號
Dim stCP12 As String, stCP13 As String
Dim m_strNP22 As String '下一程序流水號
Dim m_PA57 As String     '2008/11/28 add by sonia 1912通知已轉他所記錄是否閉卷
Dim m_CP14ST06 As String '2010/1/20 add by sonia 承辦人所別
Dim m_CustX07166 As Boolean   '2012/11/26 add by sonia 是否順德(含關係企業)專利案件
Dim m_specialCust As Boolean  '2013/8/12 add by sonia 申請人是否為X69011010華碩 2013/9/12因新日興及義隆也是只出定稿不分析,故改變數名稱
Dim m_SimpleReportCust As Boolean '2015/8/21 add by sonia 先出簡單報告定稿客戶
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim strChoseBase As String 'Added by Lydia 2017/05/09 被視為未主張的基礎案
Dim strBasePD06 As String  'Added by Lydia 2017/05/09 被視為未主張的基礎案(只有優先權號)
Dim m_bolAddLP As Boolean, m_strCP10 As String, m_strLD18 As String, m_str1998CP09 As String 'Added by Morgan 2018/7/2
Dim m_bolJudgerAlert As Boolean 'Added by Morgan 2018/11/12
Dim m_bolReKeyInOK As Boolean 'Added by Morgan 2020/7/30 是否與2次確認期限一致
Dim m_bolIDSPrice As Boolean, m_bolRCE As Boolean 'Added by Morgan 2020/12/24
Dim m_bolAdd217BCP As Boolean 'Added by Morgan 2021/6/3 是否內部收文公開費
Dim m_CustX69365 As Boolean 'Added by Morgan 2021/10/5 是否長庚醫院案件
Dim m_IDNGrant As Boolean 'Added by Morgan 2021/11/4 是否印尼發明/新型核准,年費法限,年費所限
Dim m_iNextIndex As Integer, m_iNoStopIdx 'Added by Morgan 2021/12/9
Dim m_Close413 As String 'Added by Morgan 2023/3/3 自請撤回413詢問是否閉卷
Dim m_bolWebQuery As Boolean 'Added by Morgan 2023/6/6 美國通知審查中是否官網查詢
Dim m_intEstMonths As Integer 'Added by Morgan 2023/6/6 通知審查中預估審查結果月數
Dim m_bolBPFCase As Boolean '是否寶齡富錦 Added by Morgan 2023/6/27
Dim m_USCaseNo As String 'Added by Morgan 2023/12/11 相關美國案本所案號(管制IDS期限)
Dim m_Alert As String 'Added by Morgan 2024/7/12 提醒訊息(泰國公開費可發文)

Private Sub cmdCountry_Click()
   ModifyMoneyCountry strCountry, strMoneyCountry, strMoney, strFagentNo
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim i As Integer, strTmp As String

   Select Case Index
      Case 0
      
         If CheckDataValid() = False Then GoTo EXITSUB
         Screen.MousePointer = vbHourglass
         i = 23
         If i = 23 Then
            '重新檢查欄位有效性
            m_bolSaveCheck = True
            If TxtValidate = False Then
               m_bolSaveCheck = False
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
            'Add by Amy 2018/03/20  判斷有未收款彈訊息
            If InStr("1001,1002,1006", txtCaseField(0)) > 0 Then
              If Pub_B911NotPay(field(1), field(2), field(3), field(4)) = True Then
                MsgBox "此案有未收款！", vbExclamation
              End If
            End If
            'end 2018/03/20
            m_bolSaveCheck = False
            '申請案核准且為自動發證時提示輸入公告日及證書號
            If AutoIssue() = True Then
               'Added by Morgan 2012/4/26
               '西班牙發明核准要輸入公告日及公告號
               If field(8) = "1" And field(9) = "211" Then
                  If txtPA14 = "" Or txtPA15 = "" Then
                     MsgBox "西班牙發明核准要輸入公告日及公告號!!", vbExclamation
                     SSTab1.Tab = 1
                     If txtPA14 = "" Then
                        txtPA14.SetFocus
                     Else
                        txtPA15.SetFocus
                     End If
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  End If
               Else
               'end 2012/4/26
               
                  If txtPA14 = "" Or txtPA22 = "" Then
                     SSTab1.Tab = 1
                     If MsgBox("公告日或證書號空白,是否要輸入?", vbYesNo + vbDefaultButton1 + vbExclamation, "自動發證國家核准確認") = vbYes Then
                        If txtPA14 = "" Then
                           txtPA14.SetFocus
                        Else
                           txtPA22.SetFocus
                        End If
                        Screen.MousePointer = vbDefault
                        Exit Sub
                     End If
                  End If
               End If 'Added by Morgan 2012/4/26
            End If
            '檢查是否發明案核准且為一案兩請,先作韓國
            m_bolIsDualApp = False
            If txtCaseField(0) = 核准 And cp(10) = 發明申請 And field(9) = "012" Then
               m_bolIsDualApp = PUB_IsDualApply(field, m_stUPA, m_stCaseNo, m_stCertNo, m_stAppNo, m_stCaseName)
            End If
            
            ''歐盟設計核准
            'Remove by Morgan 2007/12/5 歐盟設計官方核准函已不告知公告日-- 甄妮
            'If txtCaseField(0) = 核准 And lblCaseField(9) = "239" And cp(10) = 設計申請 Then
            '   If txtPublic.Text = "" Then
            '      SSTab1.Tab = 0
            '      MsgBox "歐盟設計核准必須輸入公告與否!", vbExclamation
            '      txtPublic.SetFocus
            '      Screen.MousePointer = vbDefault
            '      Exit Sub
            '   End If
            'End If
            '檢查約定期限
            '2008/11/28 MODIFY BY SONIA 改為核駁,最終核駁都要輸
            'If m_bolIsNP107N = True Then
            '   If txtCaseField(5) = "" And IsEmptyText(txtCaseField(30)) = True Then
            '      SSTab1.Tab = 0
            '      MsgBox "核駁且有答辯不續辦需輸入約定期限!", vbExclamation
            '      txtCaseField(30).SetFocus
            '      Screen.MousePointer = vbDefault
            '      Exit Sub
            '   End If
            'End If
            '2008/12/19 MODIFY BY SONIA有下一程序者才控制
            'If txtCaseField(30).Visible = True Then
            'Modified by Morgan 2024/12/6 改控制enable，這樣駐點才不會亂跳
            'If txtCaseField(30).Visible = True And txtCaseField(1) <> "" Then
            If txtCaseField(30).Enabled And txtCaseField(1) <> "" Then
            'end 2024/12/6
               If IsEmptyText(txtCaseField(30)) = True Then
                  SSTab1.Tab = 0
                  'Modified by Morgan 2021/9/2 應該是約定下一程序的收文性質
                  'MsgBox "核駁請輸入約定期限 !", vbExclamation
                  MsgBox lblNextCaseProperty & "請輸入約定期限 !", vbExclamation
                  'end 2021/9/2
                  txtCaseField(30).SetFocus
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
            End If
            '2008/11/28 END
            'Add by Morgan 2008/6/12
            If txtCaseField(1) = "601" Then
               If Trim(txtCaseField(6)) = "" Then
                  MsgBox "下一程序為【領證】時，【領證費】欄位不可空白！"
                  txtCaseField(6).SetFocus
                  Screen.MousePointer = vbDefault
                  Exit Sub
               ElseIf Trim(txtCaseField(8)) = "" Then
                  MsgBox "下一程序為【領證】時，【點數】欄位不可空白！"
                  txtCaseField(8).SetFocus
                  Screen.MousePointer = vbDefault
                  Exit Sub
               End If
               
               'Added by Morgan 2023/5/12
               '游本俊〈X75231000〉及游翊〈X75231010〉的CFP案件之領證費及年費在程序人員報價後，系統請另外再多加服務費NT$1,000並跳訊息告知操作人員。
               If InStr("X75231000,X75231010", field(26)) > 0 Then
                  If MsgBox("客戶游本俊〈X75231000〉及游翊〈X75231010〉因長年在中國大陸，所以委辦案件付款都是由本所收據金額直接換算當時美金在匯至本所華南銀行。近日發現華南銀行扣除手續費後會有虧損情況，故此兩客戶的CFP案件之領證費及年費請" & vbCrLf & vbCrLf & "另外再多加 NT$1,000(0點)，本次報價是否已調整？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                     txtCaseField(8).SetFocus
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  End If
               End If
               'end 2023/5/12
            End If
            
            '2008/11/28 add by sonia 1912通知已轉他所詢問是否閉卷
            m_PA57 = ""
            If txtCaseField(0) = "1912" And Not m_blnClosed Then
               If MsgBox("通知已轉他所，是否要閉卷？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                  m_PA57 = "Y"
               End If
            End If
            '2008/11/28 End
         
            'Added by Morgan 2023/3/3 自請撤回413詢問是否閉卷
            m_Close413 = "N"
            If cp(10) = "413" And field(57) = "" Then
               '新申請案不必詢問直接閉卷
               strExc(0) = "select cp27 from caseprogress where cp09='" & cp(43) & "' and cp10 in (" & NewCasePtyList & ") "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  m_Close413 = "Y"
               Else
                  If MsgBox("此為自請撤回之核准, 請問是否要閉卷？", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
                     m_Close413 = "Y"
                  End If
               End If
            End If
            'end 2023/3/3
         
            '詢問是否計算結餘
            '2008/11/28 MODIFY BY SONIA 加 1006最終核駁,1912通知已轉他所
            'Modify by Morgan 2010/3/10 +建議性處分書 1220
            'modify by sonia 2024/12/20 1912通知已轉他所改為自動上可結餘不必詢問
            If ((txtCaseField(0) = 核准 Or txtCaseField(0) = 核駁 Or txtCaseField(0) = "1006" Or txtCaseField(0) = "1220") And (cp(10) = 復審 Or cp(10) = 再發行)) Or txtCaseField(0) = 專利權消滅 Then
               '2011/11/8 modify by sonia TF子案不可結餘故加傳本所案號
               'Pub_EndModCashMsg field(9)
               Pub_EndModCashMsg field(9), field(1), field(2), field(3), field(4)
            End If
            'add by sonia 2024/12/20  1912通知已轉他所改為自動上可結餘不必詢問
            If txtCaseField(0) = "1912" Then
               bolEndModCash = True  '自動上結餘日
            End If
            'end 2024/12/20
         
            'add by sonia 2025/4/18 案件僅變更401、讓與701，於核准時詢問是否計算結餘
            If (cp(10) = "401" Or cp(10) = "701") And field(16) = "1" Then
               strExc(0) = "Select * From Caseprogress WHERE CP01='" & field(1) & "' AND CP02='" & field(2) & "' AND CP03='" & field(3) & "' AND CP04='" & field(4) & "' And Cp09<>'" & cp(9) & "' and cp09<'B' and cp60 is not null and Cp59 Is Null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  Pub_EndModCashMsg field(9), field(1), field(2), field(3), field(4)
               End If
            End If
            'end 2025/4/18
         
            '若本案已閉卷
            m_blnCancelClosed = False
            If m_blnClosed Then
               'Add by Morgan 2007/5/8 若來函有期限但已閉卷
               If txtCaseField(3) <> "" Then
                  If MsgBox("本案目前為閉卷狀態，為管制期限將於存檔時取消閉卷，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  End If
                  m_blnCancelClosed = True
               Else
               'end 2007/5/8
                  If m_strCloseDate = "" Then
                      strExc(10) = "此案已於 ??年 ??月 ??日 閉卷，是否取消閉卷? "
                  Else
                      strExc(10) = "此案已於 " & Left(m_strCloseDate, 4) - 1911 & "年" & Mid(m_strCloseDate, 5, 2) & "月" & Right(m_strCloseDate, 2) & "日閉卷，是否取消閉卷? "
                  End If
                  
                  If MsgBox(strExc(10), vbExclamation + vbYesNo) = vbYes Then
                      m_blnCancelClosed = True
                  Else
                      m_blnCancelClosed = False
                  End If
               End If
            End If
         
            '2012/11/26 add by sonia 順德(及關係企業)要先簡單報告,承辦人也要預設
            If m_CustX07166 = False Then m_CustX07166 = PUB_CheckX07166Remind(field(1), cp(9), txtCaseField(0).Text)
            '2015/8/21 add by sonia 加入判斷是否為要先簡單報告的客戶
            If m_CustX07166 = True Then m_SimpleReportCust = True  '順德(及關係企業)也是要先簡單報告的客戶
            If m_SimpleReportCust = False Then m_SimpleReportCust = PUB_SimpleReportCust(field(1), txtCaseField(0).Text, ChangeCustomerL(field(26)), ChangeCustomerL(field(27)), ChangeCustomerL(field(28)), ChangeCustomerL(field(29)), ChangeCustomerL(field(30)))
            'Added by Morgan 2021/10/5 長庚醫院案件
            m_CustX69365 = PUB_ChkIsX69365Case(field(1), field(2), field(3), field(4))
            'If m_CustX69365 = True Then m_SimpleReportCust = True '簡單報告 Removed by Morgan 2022/3/28 取消轉公文,修改OA發文管制日(所限) --黃教威
            
            'end 2021/10/5
            If m_SimpleReportCust = True Then txtCaseField(5) = "" '要出定稿
            '2015/8/21 end
            
            'Added by Morgan 2018/7/3
            '配合 CFP電子化,整合簡單報告條件
            If m_SimpleReportCust Then
               'Modified by Morgan 2021/10/6 來函性質改用常數判斷
               'If Not (txtCaseField(0) = "1002" Or txtCaseField(0) = "1006" Or txtCaseField(0) = "1220" Or txtCaseField(0) = "1206" Or txtCaseField(0) = "1209") Then
               If InStr(PatentOAPtyList, txtCaseField(0).Text) = 0 Then
               'end 2021/10/5
                  m_SimpleReportCust = False
               End If
            End If
            'end 2018/7/3
            
            '2013/1/22 modify by sonia 加入1220建議性處份書
            'Modified by Morgan 2021/10/6 來函性質改用常數判斷
            'If m_CustX07166 = True And InStr("1002,1006,1206,1209,1220", txtCaseField(0).Text) > 0 Then
            If m_CustX07166 = True And InStr(PatentOAPtyList, txtCaseField(0).Text) > 0 Then
            'end 2021/10/5
               txtCaseField(5) = ""
               '2013/2/8 add by sonia
               'cancel by sonia 2015/4/7
               'If txtCaseField(12) <> "84012" And txtCaseField(12) <> "93003" And txtCaseField(12) <> "97034" Then
               '   txtCaseField(12) = "84012"
               '   CheckKeyIn 12
               '   MsgBox "本案之承辦人改為粘竺儒 !!!", vbInformation
               'End If
               'end 2015/4/7
               '2013/2/8 end
               'add by sonia 2015/4/20 專利處再改需求,若工程師非粘竺儒84012、楊佳蓉A0039則改為楊佳蓉A0039,並顯示訊息
               'modify by sonia 2016/4/27 再取消84012
               'If txtCaseField(12) <> "84012" And txtCaseField(12) <> "A0039" Then
               'modify by sonia 2017/3/7 楊佳蓉產假改蔡順興94019
               'If txtCaseField(12) <> "A0039" Then
               '   txtCaseField(12) = "A0039"
               '   CheckKeyIn 12
               '   MsgBox "本案之承辦人改為楊佳蓉 !!!", vbInformation
               'End If
               'end 2017/3/7
               'Removed by Morgan 2025/7/17 取消--郭
               'If txtCaseField(12) <> "94019" Then
               '   txtCaseField(12) = "94019"
               '   CheckKeyIn 12
               '   MsgBox "本案之承辦人改為蔡順興 !!!", vbInformation
               'End If
               'end 2025/7/17
               'end 2015/4/20
            End If
            '2012/11/26 END
               
            'Added by Lydia 2017/05/09 後案官方來函性質「視為未主張」，若有兩個以上的優先權，則 show出該兩個優先權讓user勾選哪一個被視為未主張，若只有一個優先權，就直接自優先權資料處移至案件備註(刪除PriDate,寫案件備註)，該案若有以優先權計算期限的，則請重新計算期限。
            strChoseBase = ""
            strBasePD06 = ""
            If txtCaseField(0) = "1918" Then
               Set RsTemp = PUB_ReadPDStateNew(field, cp(10))
               If RsTemp.RecordCount = 1 Then
                  strChoseBase = RsTemp.Fields("優先權號") & "|" & RsTemp.Fields("優先權日") & "|" & RsTemp.Fields("PD07")
               ElseIf RsTemp.RecordCount > 1 Then
                   Set frm880012.grdDataList.Recordset = RsTemp
                   Set frm880012.fmParent = Me
                   frm880012.iTyp = "4"
                   frm880012.Show vbModal
                   If Me.Tag = "" Then
                      MsgBox "請選擇一個優先權資料!"
                      Exit Sub
                   Else
                      strChoseBase = Me.Tag
                      Me.Tag = ""
                   End If
               End If
               'Added by Lydia 2025/08/06 因為檢查出上線後無觸發控制，所以另外寫提醒; ex.CFP-032840
               If strChoseBase = "" Then
                  MsgBox "請選擇一個優先權資料!"
                  Exit Sub
               End If
               'end 2025/08/06
            End If
            'end 2017/05/09
            
            'Add By Sindy 2022/7/21
            If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
               If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
                  Screen.MousePointer = vbDefault 'Added by Morgan 2023/6/6
                  Exit Sub
               End If
            End If
            '2022/7/21 END
            
            'Added by Morgan 2023/6/6
            m_bolWebQuery = False
            m_intEstMonths = 0
            If txtCaseField(0) = "1905" Then
               If field(9) = "221" And PUB_ChkCPExist(cp(), "1209") = False Then
                  strExc(1) = "接獲擴大檢索報告"
               Else
                  strExc(1) = "有審查結果"
               End If
               
               strExc(0) = ""
               Do
                  strExc(0) = InputBox("請輸入預估月數：" & vbCrLf & vbCrLf & "(預計再過?個月才會" & strExc(1) & ")", lblProperty, "?")
                  If strExc(0) = "" Then
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  Else
                     m_intEstMonths = Val(strExc(0))
                     If m_intEstMonths > 0 Then
                        Exit Do
                     Else
                        MsgBox "請輸入大於0的數字！", vbExclamation
                     End If
                  End If
               Loop
               
               If field(9) = 美國國家代號 Then
                  If MsgBox("是否為官網預估時間？", vbYesNo + vbDefaultButton1) = vbYes Then
                     m_bolWebQuery = True
                  End If
               End If
            End If
            'end 2023/6/6
            
            m_Alert = "" 'Added by Morgan 2024/7/1
            
            If SaveDatabase Then
            
               If m_Alert <> "" Then MsgBox m_Alert, vbInformation 'Added by Morgan 2024/7/1
               
               'Add by Morgan 2007/1/18 申請人為"福興"時彈訊息
               If InStr(field(26) & field(27) & field(28) & field(29) & field(30), "X43179") > 0 And InStr("1002,1006,1202,1203,1205,1206,1209", txtCaseField(0).Text) > 0 Then
                  MsgBox "請影印一份OA交付智權人員【" & GetStaffName(stCP13) & "】！"
               End If
               'end 2007/1/18
         
                '若承辦人是王協理且未發文則要發EMail通知
                'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
                If stCP14 = "99050" And stCP27 = "" Then
                    Call PUB_SendMail(strUserNum, "99050", stCP09, "分案通知")
                End If
                
               'Add by Morgan 2006/3/27
               If m_str222MailCP14 <> Empty Then
                  PUB_SendMail strUserNum, m_str222MailCP14, m_str222MailCP09, field(1) & "-" & field(2) & "-" & field(3) & "-" & field(4) & "文件齊備通知！", "", "本案已收到建議性處分書，系統自動上文件齊備日！"
               End If
                
               m_ET02 = m_NewCP09 '預設新收文號
               
               If txtCaseField(5) <> "N" Then
                  'Added by Morgan 2018/7/3
                  '同時簡單報告及回覆單兩個定稿時只能到定稿維護修改內容
                  If m_bolAddLP And m_SimpleReportCust And m_blnCustReturnSheet And txtCaseField(25).Text = "Y" Then
                     MsgBox "為配合轉PDF檔至卷宗區，簡單報告請到定稿維護修改內容!!", vbInformation, "CFP電子化"
                     txtCaseField(25).Text = ""
                  End If
                  'end 2018/7/3
               
                  Select Case txtCaseField(0)
                  
                     Case 核准
                        '一般
                        strTmp = "00"
                        Select Case cp(10)
                           'Added by Morgan 2019/1/23
                           Case 431 'PPH
                              strTmp = "15"
                           'end 2019/1/23
                           Case 變更
                              strTmp = "21"
                              'Add by Morgan 2008/1/14
                              If field(9) = 美國國家代號 Then
                                 strTmp = "22"
                              End If
                           Case 自請撤回
                              strTmp = "13"
                              
                           'Modify by Morgan 2007/8/9 加繼承,定稿一樣
                           'Modify by Morgan 2007/10/3 加授權
                           Case 讓與, 繼承, 授權
                              Select Case field(9)
                                 Case 美國國家代號
                                    strTmp = "22"
                                 Case Else
                                    strTmp = "02"
                              End Select
                                                                                 
                           '2005/5/19 MODIFY BY SONIA 加分割
                           '2007/8/3 加 424請求繼續審查
                           'Modify by Morgan 2008/10/15 +改請案
                           'Modify by Sonia  2013/11/15 加復審805(CFP-023821)
                           'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
                           'Modified by Lydia 2016/08/27 +438 再考量試行計畫(AFCP2.0)
                           'Modified by Morgan 2019/11/1 +CA申請 122
                           Case 發明申請, 新型申請, 設計申請, 追加申請, 答辯, CIP申請, CPA申請, CA申請, 再發行, 訴願, 分割, "424", "301", "302", "303", "805", "126", "438", "105"
                              
                              '歐盟設計
                              'Removed by Morgan 2020/9/28 取消(不會通知核准,直接輸證書)--玫音
                              'If field(9) = "239" And m_PA08 = "3" Then
                              '   strTmp = "01"
                              'Else
                              'end 2020/9/28
                              
                                 '自動發證-申請案核准且無下一程序
                                 If txtCaseField(1) = "" Then
                                       'Added by Morgan 2012/4/26
                                       '西班牙發明案核准出公告定稿
                                       If field(9) = "211" And field(8) = "1" Then
                                          strTmp = "26"
                                       Else
                                       'end 2012/4/26
                                          strTmp = "16"
                                       End If 'Added by Morgan 2012/4/26
                                 Else
                                    Select Case field(9)
                                       Case 美國國家代號
                                          '發明
                                          If m_PA08 = "1" Then
                                             '初審
                                             'Modified by Morgan 2019/11/4 +CA申請 Ex:CFP-029157-1-00 --甄妮
                                             If cp(10) = 發明申請 Or cp(10) = 追加申請 Or cp(10) = 分割 Or cp(10) = CIP申請 Or cp(10) = CA申請 Then
                                                strTmp = "03"
                                                '有製圖費
                                                If Val(txtCaseField(7)) > 0 Then
                                                   If cp(10) = 發明申請 Or cp(10) = CIP申請 Or cp(10) = CA申請 Then
                                                      strTmp = "23"
                                                   End If
                                                End If
                                             '再審
                                             'Modified by Lydia 2016/08/27 +438 再考量試行計畫(AFCP2.0)
                                             ElseIf (cp(10) = 答辯 Or cp(10) = "424" Or cp(10) = "126" Or cp(10) = "438") Then '2007/8/3 加 424請求繼續審查 'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
                                                strTmp = "04"
                                                '有製圖費
                                                If Val(txtCaseField(7)) > 0 Then
                                                   strTmp = "25"
                                                End If
                                             End If
                                          '設計
                                          ElseIf m_PA08 = "3" Then
                                                strTmp = "05"
                                                '有製圖費
                                                If Val(txtCaseField(7)) > 0 Then
                                                   strTmp = "20"
                                                End If
                                          End If
                                          
   'Removed by Morgan 2013/12/16 改用新報價,不再使用
   '                                       'Added by Morgan 2013/10/18 本段程式 103/1/1 以後刪除(定稿也刪)
   '                                       If Val(strSrvDate(1)) < 20140101 Then
   '                                          If strTmp = "03" Then
   '                                             strTmp = "51"
   '                                          ElseIf strTmp = "04" Then
   '                                             strTmp = "52"
   '                                          ElseIf strTmp = "05" Then
   '                                             strTmp = "53"
   '                                          ElseIf strTmp = "20" Then
   '                                             strTmp = "54"
   '                                          ElseIf strTmp = "23" Then
   '                                             strTmp = "55"
   '                                          ElseIf strTmp = "25" Then
   '                                             strTmp = "56"
   '                                          End If
   '                                       End If
   '                                       'end 2013/10/18
   'end 2013/12/16
                                          
                                       Case "102" '加拿大
                                          '發明
                                          If m_PA08 = "1" Then
                                              strTmp = "06"
                                          End If
                                          
                                       Case "012" '韓國
                                          '發明
                                          If m_PA08 = "1" Then
                                             strTmp = "07"
                                             '一案兩請
                                             If m_bolIsDualApp = True Then
                                                strTmp = "17"
                                             End If
                                          '設計
                                          ElseIf m_PA08 = "3" Then
                                             strTmp = "08"
                                          End If
                                          
                                       Case "019" '泰國
                                          '設計
                                          If m_PA08 = "3" Then
                                             strTmp = "11"
                                          End If
                                          
                                       Case "201" '英國
                                          '發明
                                          If m_PA08 = "1" Then
                                             strTmp = "09"
                                          End If
                                          
                                       Case "221" 'EPC
                                         ' Modified by Lydia 2014/12/10 EPC告准函原需輸入7國以上才會帶出所有可指定國，更改為不限定輸入國家數量。
                                          m_bolEPC7Up = PUB_CheckEPC(strMoneyCountry, cp(1) & cp(2) & cp(3) & cp(4), m_strRestEPCMember)
                                          '七國以上
                                         ' If m_bolEPC7Up Then
                                             m_bolEPC7Up = True
                                             strTmp = "14"
                                         ' Else
                                         '    strTmp = "10"
                                         ' End If
                                          
                                       Case Else
                                       
                                    End Select
                                 End If
                                 
                              'End If 'Removed by Morgan 2020/9/28
                           Case Else
                              '2008/7/18 add by sonia 其他案件性質用無期限定稿24,否則會用00印掛號字樣
                              strTmp = "24"
                              '2008/7/18 end
                        End Select
                        
                        'Add by Morgan 2008/5/7 新增領證報價通知
                        If Val(strSrvDate(1)) > 20080810 And txtCaseField(1) = "601" Then
                           'Modified by Morgan 2015/2/9
                           '要報價但沒有定稿時提醒(目前程式有預設,不會發生)
                           If strTmp = "" Then
                              MsgBox "本案要報價但沒有系統的定稿，請注意！", vbExclamation
                           Else
                           'end 2015/2/9
                              
                              'Added by Morgan 2021/10/15 寶齡富錦 Y55435 案件
                              If ChangeCustomerS(field(75)) = "Y55435" Then
                                 strTmp = "99"
                              End If
                              'end 2021/10/15
                        
                              PUB_AddLetterCache m_NewCP09, m_strNP22, m_ET02, "03", strTmp, , m_strLD18
                              StartLetter1 strTmp, m_NewCP09, m_strNP22
                              '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                              strExc(0) = CompWorkDay(5, strSrvDate(1))
                              strExc(1) = DBDATE(txtCaseField(2))
                              If Val(strExc(0)) >= Val(strExc(1)) Then
                                 PUB_Cache2Letter m_NewCP09, m_strNP22, False, IIf(Me.txtCaseField(25).Text = "Y", True, False)
                              End If
                           End If
                           
                        'Added by Morgan 2021/11/5 +印尼發明/新型核准
                        ElseIf m_IDNGrant And m_strNP22 <> "" Then
                           PUB_AddLetterCache m_NewCP09, m_strNP22, m_ET02, "03", strTmp, , m_strLD18
                           StartLetter1 strTmp, m_NewCP09, m_strNP22
                           '若[系統日+5個工作天>=所限]時，不必讓智權人員確認，直接列印
                           strExc(0) = CompWorkDay(5, strSrvDate(1))
                           If Val(strExc(0)) >= Val(m_strDate1) Then
                              PUB_Cache2Letter m_NewCP09, m_strNP22, False, IIf(Me.txtCaseField(25).Text = "Y", True, False)
                           End If
                        'end 2021/11/5
                        Else
                           StartLetter "03", m_ET02, strTmp
                           NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                        End If
                        'end 2008/5/7
                     
                     '2008/11/28 MODIFY BY SONIA 加 1006
                     'Modify by Morgan 2010/3/10 +建議性處分書 1220
                     Case 核駁, "1006", "1220"
                        '檢查下一程序有答辯不續辦才出通知函
                        If m_bolIsNP107N = True Then
                           'Modify by Morgan 2006/4/7 定稿統一
                           strTmp = "18"
                           
                           'Added by Morgan 2022/5/19 寶齡富錦 Y55435 案件
                           '最終核駁
                           If txtCaseField(0) = "1006" Then
                              If ChangeCustomerS(field(75)) = "Y55435" Then
                                 strTmp = "99"
                              End If
                            End If
                           'end 2022/5/19
                           
                           StartLetter "03", m_ET02, strTmp
                           NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                           
                        'End If 'Removed by Morgan 2018/7/3 只會有一種定稿
                        
                        '2012/11/26 ADD BY SONIA 順德及其關係企業加定稿,也要印案件回覆單
                        '2013/1/22 modify by sonia 加入1220建議性處份書
                        '2015/8/21 modify by sonia 加入要簡單報告客戶
                        'Modified by Morgan 2018/7/3 來函性質判斷移到上面(要在存檔前才能控制新增D類收文轉官文(1998))
                        'If (m_CustX07166 = True Or m_SimpleReportCust = True) And (txtCaseField(0) = "1002" Or txtCaseField(0) = "1006" Or txtCaseField(0) = "1220") Then
                        ElseIf m_SimpleReportCust Then
                        'end 2018/7/3
                           
                           'Modified by Morgan 2018/10/2 回覆單改先產生且不開Word否則E化後維護畫面會帶錯內容
                           If m_blnCustReturnSheet = True Then
                              StartLetter "03", m_ET02, "12"
                              NowPrint m_ET02, "03", "12", False, strUserNum, , , , , , , , , , , , , m_strLD18
                              g_LD18 = "" 'Added by Morgan 2018/10/2
                           End If
                           
                           strTmp = "19"
                           StartLetter "03", m_ET02, strTmp
                           NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_str1998CP09
                           
                        'End If 'Removed by Morgan 2018/7/3 只會有一種定稿
                        '2012/11/26 END
                        
                        '2013/8/12 ADD BY SONIA 特殊客戶只印定稿不分析,也要印案件回覆單
                        'Modified by Morgan 2018/7/3 配合CFP電子化,改用28(含回覆單的簡單報告)
                        'If m_specialCust = True Then
                        '   StartLetter "03", m_ET02, "19"
                        '   NowPrint m_ET02, "03", "19", IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum
                        '   If m_blnCustReturnSheet = True Then
                        '      StartLetter "03", m_ET02, "12"
                        '      NowPrint m_ET02, "03", "12", IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum
                        '   End If
                        ElseIf m_specialCust = True Then
                           strTmp = "28"
                           StartLetter "03", m_ET02, strTmp
                           NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                           
                        ElseIf m_blnCustReturnSheet = True Then
                           StartLetter "03", m_ET02, "12"
                           NowPrint m_ET02, "03", "12", False, strUserNum, , , , , , , , , , , , , m_strLD18
                           g_LD18 = "" 'Added by Morgan 2018/10/2
                        'end 2018/7/3
                        End If
                        '2013/8/12 END
                        
                     'Added by Morgan 2019/1/25 核發(先用優先權證明書,其他有用到再改或加定稿)
                     Case "1008"
                        strTmp = "24"
                        If cp(10) = "443" Then strTmp = "25" 'Added by Morgan 2023/6/14 +申請紙本專利證書
                        NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                     'end 2019/1/25
                     
                     Case 其他來函   '1902其他來函
                        If txtCaseField(29) = "Y" Then
                           strTmp = "00"
                           NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                        
                        'Added by Morgan 2021/1/11
                        ElseIf m_blnCustReturnSheet Then
                           StartLetter "03", m_ET02, "12"
                           NowPrint m_ET02, "03", "12", False, strUserNum, , , , , , , , , , , , , m_strLD18
                           g_LD18 = ""
                        End If
                        'end 2021/1/11
                        
                     '2008/11/11 ADD BY SONIA通知已轉他所1912
                     '2008/11/28 ADD BY SONIA通知審查中1905
                     'Add By Sindy 2009/06/17'公告異議期滿通知1223
                     '2009/11/11 MODIFY BY SONIA 加註冊登記1607
                     '2010/7/16 modify by sonia 加1606專利權公告作廢
                     '2015/05/20 Modified by Lydia +EPC通知審查中1905未收文檢索報告之定稿
                     Case "1005", "1912", "1905", "1223", "1607", "1606" '核准先行通知1005,通知已轉他所1912,通知審查中1905,公告異議期滿通知1223,註冊登記1607,專利權公告作廢1606
                        strTmp = "00" '預設
                        'Modified by Morgan 2015/6/24 +EPC判斷
                        If txtCaseField(0) = "1905" Then
                           If field(9) = "221" Then
                              If PUB_ChkCPExist(cp(), "1209") = False Then strTmp = "01"
                           'Added by Morgan 2023/6/6
                           'Modified by Morgan 2023/9/28 +And m_bolWebQuery = True
                           ElseIf field(9) = "101" And m_bolWebQuery = True Then
                              strTmp = "02"
                           'end 2023/6/6
                           End If
                        End If
                        'end 2015/05/20
                        
                        'Added by Lydia 2016/12/26 英國核准先行通知(1005)定稿
                        If field(9) = "201" And txtCaseField(0) = "1005" Then strTmp = "02"
                        
                        StartLetter "03", m_ET02, strTmp
                        NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                     
                     Case "1213"   '初步審查合格通知1213
                        strTmp = "14"
                        StartLetter "03", m_ET02, strTmp
                        NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                     
                     'Added by Lydia 2017/05/09 視為未主張
                     Case "1918"
                        strTmp = "00"
                        StartLetter "03", m_ET02, strTmp
                        NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                     'end 2017/05/09
                     
                     
                     'Modify by Morgan 2005/11/10 重整
                     Case Else
                        'Added by Morgan 2018/11/7 +PCT進EPC通知修正定稿--禧佩
                        If txtCaseField(0) = "1201" And field(9) = "221" And field(46) = "Y" Then
                           strTmp = "00"
                           StartLetter "03", m_ET02, strTmp
                           NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                           
                        'Added by Morgan 泰國通知繳公開費
                        ElseIf txtCaseField(0) = "1236" And field(9) = "019" Then
                           strTmp = "00"
                           StartLetter "03", m_ET02, strTmp
                           NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                        '只印案件回覆單
                        Else
                           '2012/11/26 ADD BY SONIA 順德及其關係企業加定稿
                           '2015/8/21 modify by sonia 加入要簡單報告客戶
                           'Modified by Morgan 2018/7/3 來函性質判斷移到上面(要在存檔前才能控制新增D類收文轉官文(1998))
                           'If (m_CustX07166 = True Or m_SimpleReportCust = True) And (txtCaseField(0) = "1206" Or txtCaseField(0) = "1209") Then
                           If m_SimpleReportCust Then
                           'end 2018/7/3
                              'ADD BY SONIA 2014/6/26 EPC檢索報告改內容故加處理狀況27, CFP-026300
                              If field(9) = "221" And txtCaseField(0) = "1209" Then
                                 strTmp = "27"
                                 StartLetter "03", m_ET02, strTmp
                                 NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_str1998CP09
                              Else
                              'END 2014/6/26
                                 strTmp = "19"
                                 StartLetter "03", m_ET02, strTmp
                                 NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_str1998CP09
                              End If
                              
                           End If
                           '2012/11/26 END
                           
                           '2013/8/12 ADD BY SONIA 特殊客戶只印定稿不分析
                           If m_specialCust = True Then
                           'Modified by Morgan 2018/7/3 配合CFP電子化,改用28(含回覆單的簡單報告)
                           '   StartLetter "03", m_ET02, "19"
                           '   NowPrint m_ET02, "03", "19", IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum
                           'End If
                           ''2013/8/12 END
                           'If m_blnCustReturnSheet = True Then
                              strTmp = "28"
                              StartLetter "03", m_ET02, strTmp
                              NowPrint m_ET02, "03", strTmp, IIf(Me.txtCaseField(25).Text = "Y", True, False), strUserNum, , , , , , , , , , , , , m_strLD18
                           ElseIf m_blnCustReturnSheet Then
                           'end 2018/7/3
                              StartLetter "03", m_ET02, "12"
                              NowPrint m_ET02, "03", "12", False, strUserNum, , , , , , , , , , , , , m_strLD18
                              g_LD18 = "" 'Added by Morgan 2018/10/2
                           End If
                           
                        End If 'Added by Morgan 2018/11/7
                  End Select
                  
                  'Added by Morgan 2018/7/3 CFP電子化
                  If m_bolAddLP And txtCaseField(25).Text = "Y" And m_strLD18 = g_LD18 Then
                     frm1105_1.m_RecNo = m_strLD18
                     frm1105_1.m_PdfName = PUB_CaseNo2FileName(cp(1), cp(2), cp(3), cp(4)) & "." & m_strCP10 & ".CUS.PDF"
                     frm1105_1.Show
                  End If
                  'end 2018/7/3
                
               'Modify by Morgan 2005/11/10 重整
               '只印案件回覆單
               ElseIf m_blnCustReturnSheet = True Then
                  StartLetter "03", m_ET02, "12"
                  NowPrint m_ET02, "03", "12", False, strUserNum, , , , , , , , , , , , , m_strLD18
                  g_LD18 = "" 'Added by Morgan 2018/10/2
               End If
               
               'Add by Morgan 2007/3/22
               If txtCaseField(4).Text <> "" And txtCaseField(0).Text = "1001" And Me.Text7.Text = "1" And Me.Text7.Text <> Me.Text7.Tag Then
                  PUB_SameCaseCheck1 cp(), 1, DBDATE(txtCaseField(4).Text)
               End If
               'end 2007/3/22
               
               'Add By Sindy 2016/10/7
               If Me.m_strIR01 <> "" Then
                  bolLeave = True
                  Unload frm05010401_1
                  Unload frm05010401_2
                  intLeaveKind = 0 'Added by Morgan 2020/7/29
                  Unload Me
                  'Modify By Sindy 2022/5/20
                  'frm04010519.GoNext
                  Forms(0).Tmpfrm04010519.GoNext
                  Set Forms(0).Tmpfrm04010519 = Nothing
                  '2022/5/20 END
               Else
               '2016/10/7 END
                  bolLeave = True
                  intLeaveKind = 0
                  Unload Me
               End If
            Else
                MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
            End If
         End If
         Screen.MousePointer = vbDefault
      Case 1, 2
         If Index = 2 Then
            intLeaveKind = 2
         Else
            intLeaveKind = 1
         End If
         Unload Me
   End Select
EXITSUB:
End Sub

Private Sub StartLetter(ByVal ET01 As String, ET02 As String, ByVal ET03 As String)
Dim strTxt(1 To 25) As String, iStep As Integer, strTmp As Variant
Dim strTemp1 As String, strStartDate As String, strTemp As Variant
Dim bolTmp As Boolean, StrExt1 As String, StrExt2 As String, i As Integer
Dim iEPC As Integer 'EPC 指定國家順序
Dim iPos As Integer '字元搜尋位置
Dim Jjj As Integer
   
   Jjj = 1
         
   EndLetter ET01, ET02, ET03, strUserNum
   
   'Added by Morgan 2023/12/11
   If m_USCaseNo <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','美國案本所案號','" & m_USCaseNo & "')"
      Jjj = Jjj + 1
      
      If Val(txtIDSFee(1)) > 0 Then
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','IDS報價1','" & txtIDSFee(1) & "')"
         Jjj = Jjj + 1
      End If
      
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','IDS報價2','" & txtIDSFee(2) & "')"
      Jjj = Jjj + 1
   End If
   'end 2023/12/11
      
   If Trim(txtCaseField(6)) <> "" Then
      'Modified by Morgan 2021/6/1 +公開費
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','" & IIf(txtCaseField(0) = "1236", "公開費", "領證費") & "','" & Val(txtCaseField(6)) & "')"
      Jjj = Jjj + 1
   End If
   
   If Trim(txtCaseField(7)) <> "" Then
       strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
          "','製圖費','" & Val(txtCaseField(7)) & "')"
       Jjj = Jjj + 1
   End If
   
   If Trim(txtCaseField(10)) <> "" Then
       strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
          "','讓渡費','" & Val(txtCaseField(10)) & "')"
       Jjj = Jjj + 1
   End If
   
   'modify by Morgan 2008/5/12 變數名稱"點數"改為"領證費點數"
   If Trim(txtCaseField(8)) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','領證費點數','" & Val(txtCaseField(8)) & "')"
      Jjj = Jjj + 1
   End If
   'end 2008/5/12
   
   Select Case lblCaseField(9)
      Case "221" 'EPC
         'Added by Morgan 2023/8/10
         If ClsPDReadCountry(專利, field(), strExc(0), True, False, True) = True Then
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','EPC指定國家','" & strExc(0) & "')"
            Jjj = Jjj + 1
         End If
         'end 2023/8/10
         Select Case txtCaseField(0)
            Case 核准
               '加入epc 核准國家領證費用
               If Trim(strMoneyCountry) <> "" And Trim(strMoney) <> "" Then
                  strTmp = Split(strMoneyCountry, ",")
                  strTemp = Split(strMoney, ",")
                  StrExt1 = ""
                  StrExt2 = ""
                  For i = 0 To UBound(strTmp)
                     If UBound(strTmp) <> UBound(strTemp) Then
                        If UBound(strTmp) < UBound(strTemp) Then
                           If Val(strTemp(i)) <> 0 Then
                              '指定國家抓中文
                              iEPC = iEPC + 1
                              'Modify by Morgan 2006/8/10 瑞士要加列茲敦斯登-- 禧佩
                              StrExt1 = StrExt1 & str(iEPC) & "." & GetNationName(strTmp(i), 0) & IIf(strTmp(i) = "205", "/列茲敦斯登", "") & "：新台幣 "
                              StrExt1 = StrExt1 & strTemp(i) & " 元整。" & Chr$(13)
                              StrExt2 = str(Val(StrExt2) + Val(strTemp(i)))
                           End If
                        Else
                           If i >= UBound(strTemp) Then
                           Else
                              If Val(strTemp(i)) <> 0 Then
                                 '指定國家抓中文
                                 iEPC = iEPC + 1
                                 'Modify by Morgan 2006/8/10 瑞士要加列茲敦斯登-- 禧佩
                                 StrExt1 = StrExt1 & str(iEPC) & "." & GetNationName(strTmp(i), 0) & IIf(strTmp(i) = "205", "/列茲敦斯登", "") & "：新台幣 "
                                 StrExt1 = StrExt1 & strTemp(i) & " 元整。" & Chr$(13)
                                 StrExt2 = str(Val(StrExt2) + Val(strTemp(i)))
                              End If
                           End If
                        End If
                     Else
                        If Val(strTemp(i)) <> 0 Then
                           '指定國家抓中文
                           iEPC = iEPC + 1
                           'Modify by Morgan 2006/8/10 瑞士要加列茲敦斯登-- 禧佩
                           StrExt1 = StrExt1 & str(iEPC) & "." & GetNationName(strTmp(i), 0) & IIf(strTmp(i) = "205", "/列茲敦斯登", "") & "：新台幣 "
                           StrExt1 = StrExt1 & strTemp(i) & " 元整。" & Chr$(13)
                           StrExt2 = str(Val(StrExt2) + Val(strTemp(i)))
                        End If
                     End If
                  Next i
                  strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','費用','" & StrExt2 & "')"
                  Jjj = Jjj + 1
                  strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','EPC核准國家領證費用','" & StrExt1 & "')"
                  Jjj = Jjj + 1
                  
                  'Modify by Morgan 2008/5/12 變數名稱由"領證費 + 製圖費"改為"費用合計"
                  strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                     "','費用合計','" & Val(txtCaseField(6)) + Val(StrExt2) & "')"
                  Jjj = Jjj + 1
                  'end 2008/5/12
                  
                  'EPC其他成員國
                  If m_bolEPC7Up = True Then
                     StrExt2 = PUB_GetNationName(m_strRestEPCMember)
                     StrExt2 = Replace(StrExt2, ",", "、")
                     i = 0: iPos = 0
                     Do
                        iPos = i
                        i = i + 1
                        i = InStr(i, StrExt2, "、")
                     Loop While i > 0
                     If iPos > 0 Then StrExt2 = Left(StrExt2, iPos - 1) & "及" & Mid(StrExt2, iPos + 1)
                     strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','EPC其他成員國','" & StrExt2 & "')"
                     Jjj = Jjj + 1
                  End If
               End If
            
         End Select
         
      Case "101" '美國
          'Add by Morgan 2005/11/11
          If CheckStr(txtCaseField(4)) <> "" Then
             strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','准駁日'," & DBDATE(txtCaseField(4)) & ")"
             Jjj = Jjj + 1
          End If
          
          'Modify by Morgan 2008/5/12 變數名稱由"領證費 + 製圖費"改為"費用合計"
          If CheckStr(txtCaseField(6)) <> "" Or CheckStr(txtCaseField(7)) <> "" Then
             strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                "','費用合計','" & Val(txtCaseField(6)) + Val(txtCaseField(7)) & "')"
             Jjj = Jjj + 1
          End If
          
         'Remove by Morgan 2008/5/12 改直接寫在定稿內由"讓渡費"控制是否出現該段文字
         '讓渡費
'         If Len(Me.txtCaseField(10).Text) > 0 Then
'           If cp(10) <> "101" And cp(10) <> "107" Then
'                 strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                 "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'                 "','是否有讓渡費','如要辦理讓渡，亦請一併告之，讓渡費計新台幣" & Format(Me.txtCaseField(10).Text, DDollar) & "元整。" & vbCrLf & "')"
'                 Jjj = Jjj + 1
'           Else
'                 strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                 "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
'                 "','是否有讓渡費','如要辦理讓渡，亦請一併告之，讓渡費計新台幣" & Format(Me.txtCaseField(10).Text, DDollar) & "元整。" & "')"
'                 Jjj = Jjj + 1
'           End If
'         End If
         'end 2008/5/12
         
      Case Else '其他國家
      
         Select Case ET03
            
            Case "14"
               '形式審查合格通知
               Select Case lblCaseField(9)
                  Case "042"
                     If m_PA08 = "3" Then  '越南設計
                        strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                          "','列印備註','公開後6個月')"
                     Else
                        strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                          "','列印備註','之後')"
                     End If
                     Jjj = Jjj + 1
                  Case Else
                     strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                       "','列印備註','之後')"
                     Jjj = Jjj + 1
              End Select

            Case Else
            
               'Add by Morgan 2005/2/17 歐盟設計用
               'Removed by Morgan 2020/9/28 取消(不會通知核准,直接輸證書)--玫音
               'If txtPublic.Text <> "" Then
               '   strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               '      "','公告與否','" & IIf(txtPublic.Text = "1", "已", "將") & "')"
               '   Jjj = Jjj + 1
               'Else
               'end 2020/9/28
               
                  'Added by Morgan 2012/4/26
                  If txtPA14 <> "" Then
                     If Val(DBDATE(txtPA14)) > Val(strSrvDate(1)) Then
                        strExc(1) = "將"
                     Else
                        strExc(1) = "已"
                     End If
                     strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
                        "','公告與否','" & strExc(1) & "')"
                     Jjj = Jjj + 1
                  End If
                  'end 2012/4/26
                  
               'End If 'Removed by Morgan 2020/9/28
         End Select
         
         '一案兩請
         If m_bolIsDualApp = True Then
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','放棄專利權費','" & Format(txtAbandonFee) & "')"
            Jjj = Jjj + 1
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','費用合計','" & Format(Val(txtAbandonFee) + Val(txtCaseField(6))) & "')"
            Jjj = Jjj + 1
         
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','新型專利號數','" & m_stCertNo & "')"
            Jjj = Jjj + 1
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','新型本所號','" & m_stCaseNo & "')"
            Jjj = Jjj + 1
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','新型申請案號','" & m_stAppNo & "')"
            Jjj = Jjj + 1
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               " VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
               "','新型案件名稱','" & m_stCaseName & "')"
            Jjj = Jjj + 1
         End If
     End Select
  
   'Modify by Morgan 2006/5/23 有輸都要存
   'Add by Morgan 2005/5/20 核駁通知(有答辯不續辦)-德國,日本
   If txtCaseField(30) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','約定期限','" & DBDATE(txtCaseField(30)) & "')"
      Jjj = Jjj + 1
   End If
   
   If CheckStr(txtCaseField(2)) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','本所期限','" & DBDATE(txtCaseField(2)) & "')"
      Jjj = Jjj + 1
   End If
   If CheckStr(txtCaseField(3)) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','法定期限','" & DBDATE(txtCaseField(3)) & "')"
      Jjj = Jjj + 1
   End If
         
   If txtCaseField(1) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','下一程序','" & txtCaseField(1) & "')"
      Jjj = Jjj + 1
   End If
   If m_strRetSheet2NP07 <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','下一程序2','" & m_strRetSheet2NP07 & "')"
      Jjj = Jjj + 1
   End If
   If lblNextCaseProperty <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','下一程序名稱','" & lblNextCaseProperty & "')"
      Jjj = Jjj + 1
   End If
   
                  
   '美國讓渡/繼承登記號
   If txtCaseField(11) <> "" Then
      i = InStr(1, txtCaseField(11), "/", 1)
      If i > 0 Then
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','補文件 V 1','" & Mid(txtCaseField(11), 1, i - 1) & "')"
         Jjj = Jjj + 1
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','補文件 V 2','" & Right(txtCaseField(11), Val(Len(txtCaseField(11)) - i)) & "')"
         Jjj = Jjj + 1
      End If
   End If
   
   'Add by Morgan 2007/8/9
   'Modify by Morgan 2007/10/3 加授權
   'Modify by Morgan 2008/1/14 加變更
   'Modified by Morgan 2022/8/5 其他國家也適用 Ex:CFP-25478--禧佩
   'If m_PA09 = 美國國家代號 And (cp(10) = 繼承 Or cp(10) = 授權 Or cp(10) = 變更) And txtCaseField(0).Text = 核准 Then
   If (cp(10) = 繼承 Or cp(10) = 授權 Or cp(10) = 變更) And txtCaseField(0).Text = 核准 Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','相關收文性質','" & GetCaseTypeName(cp(1), cp(10)) & "')"
      Jjj = Jjj + 1
   End If
   'end 2007/8/9
   
   'Add by Morgan 2007/8/24
   If Text5(0).Enabled = True And Text5(0) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','含年費','(含第" & Text5(0) & "至" & Text5(1) & "年年費)')"
      Jjj = Jjj + 1
   End If
   'end 2007/8/24
   
   'Add by Morgan 2011/4/20
   '歐盟設計公告作廢通知函
   If field(9) = "239" And field(8) = "3" And "1606" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " select '" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','撤銷生效日',max(np09) from nextprogress where np02='" & field(1) & "'" & _
         " and np03='" & field(2) & "' and np04='" & field(3) & "'" & _
         " and np05='" & field(4) & "' and np07='607'"
      Jjj = Jjj + 1
   End If
   
   'Added by Morgan 2013/4/18
   If m_bolIsNP107N = True And m_strNP07 <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','閉卷程序','" & GetCaseTypeName(m_PA01, m_strNP07) & "')"
      Jjj = Jjj + 1
   End If
   'end 2013/4/18
   
   'Added by Lydia 2017/05/09 視為未主張
   If txtCaseField(0) = "1918" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
          "','未主張優先權號','" & strBasePD06 & "')"
       Jjj = Jjj + 1
   End If
   'end 2017/05/09
   
   'Added by Morgan 2023/6/6
   If txtCaseField(0) = "1905" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
          "','預估月數','" & m_intEstMonths & "')"
       Jjj = Jjj + 1
   End If
   'end 2023/6/6
   
   'Added by Morgan 2023/10/23
   '欣興電子轉公文特殊內容
   If m_SimpleReportCust = True And Left(field(26), 8) = "X1583316" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
          "','欣興不印','♀')"
      Jjj = Jjj + 1
       strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
          "','欣興才印','♀')"
      Jjj = Jjj + 1
   End If
   'end 2023/10/23
   
   If Not ClsLawExecSQL(Jjj - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Function SaveDatabase() As Boolean
Dim StrSQLa As String, m_PrintCForm As String
Dim rsA As New ADODB.Recordset
Dim strNA49 As String '是否自動發註冊證
Dim strDateS(3) As String
Dim strNP22  As String
Dim strTemp As String, strTemp1 As String, strTemp2 As String, dobDateAdd As Double
Dim varTemp As Variant
Dim strDate As String, strDate1 As String, strStartDate As String
Dim strTxt(1 To 30) As String, iStep As Integer
Dim lMax As Long, i As Integer
Dim bolNP22 As Boolean, NP22(1 To 3) As String, iNP22 As Integer
Dim strReceiveNo As String '總收文號
Dim strCe(99) As String, bolChk As Boolean
Dim strTmp(1 To 5) As String
Dim intStep As Integer
Dim strMsg As String
Dim strPromoteDate As String '2010/1/19 add by sonia
Dim bolUpdatePA20 As Boolean 'Add by Morgan 2010/5/26
Dim arrData() As String 'Added by Lydia 2017/05/09
Dim bolSavPdf As Boolean 'Added by Morgan 2018/10/2
Dim bolAddNP As Boolean 'Added by Morgan 2023/5/23
Dim st307Msg As String 'Added by Morgan 2024/6/18

On Error GoTo ErrorHandler

   SaveDatabase = True

   cnnConnection.BeginTrans

   iStep = 1: intStep = 1
   '若來函性質屬於爭議程序(18XX), 則更新專利基本檔PA19為"Y"
   If Left(Me.txtCaseField(0).Text, 2) = "18" Then
      field(19) = "Y"
   End If
   
   '若來函性質為"專利權消滅"(1604)
   If Me.txtCaseField(0).Text = "1604" Then
      field(17) = "N"
      field(57) = "Y"
      '94.1.10 ADD BY SONIA
      field(59) = "89"
      field(58) = TransDate(lblCaseField(8), 2)
      '94.1.10 END
   End If

   'Modify by Morgan 2006/3/23 加最終核駁1006
   'Modify by Morgan 2010/3/10 +建議性處分書 1220
   If txtCaseField(0) = 核准 Or txtCaseField(0) = 核駁 Or txtCaseField(0) = "1006" Or txtCaseField(0) = "1220" Then
      field(17) = text8
      field(16) = Text7
      
      '2007/8/3 MODIFY BY SONIA 加424請求繼續審查 CFP-016746
      'Modify by Morgan 2008/5/23 +122 CA申請 CFP-019053
      'Modify by Sonia  2013/11/15 加復審805(CFP-023821)
      'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
      'Modified by Lydia 2016/08/26 +438 再考量試行計畫(AFCP2.0)
      'Modified by Morgan 2020/12/18 改寫函數判斷以便共用及修改
      'If (cp(10) >= "101" And cp(10) <= "105") Or cp(10) = "107" Or cp(10) = "113" Or cp(10) = "114" Or cp(10) = "122" Or cp(10) = "126" Or (cp(10) >= "301" And cp(10) <= "307") Or cp(10) = "424" Or cp(10) = "438" Or cp(10) = "501" Or cp(10) = "802" Or cp(10) = "805" Then
      If PUB_ChkIsRltPty(cp(1), cp(10), field(9)) = True Then
         field(20) = txtCaseField(4)
         bolUpdatePA20 = True 'Add by Morgan 2010/5/26
      End If
      
      '92.1.18 ADD BY SONIA 存'美國讓渡登記號'於大陸申請案號欄
      'Modify by Morgan 2007/8/9 加繼承
      'If field(9) = "101" And cp(10) = 讓與 And txtCaseField(0) = 核准 Then
      'Modify by Morgan 2007/10/3 加授權
      'Modify by Morgan 2008/1/14 加變更
      If field(9) = "101" And (cp(10) = 讓與 Or cp(10) = 繼承 Or cp(10) = 授權 Or cp(10) = 變更) And txtCaseField(0) = 核准 Then
         cp(30) = txtCaseField(11)
      End If
      '92.1.18 END
      
      'Modify by Morgan 2009/10/16 改判斷來函性質更新進度檔准駁,日期若沒輸則用系統日
      'If cp(24) = "" Then cp(24) = Text7
      'If cp(25) = "" Then cp(25) = txtCaseField(4)
      If txtCaseField(0) = 核准 Then
         cp(24) = "1"
      Else
         cp(24) = "2"
      End If
      If txtCaseField(4) <> "" Then
         cp(25) = txtCaseField(4)
      Else
         cp(25) = strSrvDate(1)
      End If
      'end 2009/10/16
      
      strTxt(iStep) = GetCPSQL(cp())
      cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
      
      strTxt(iStep) = "update nextprogress set np06='Y' where np01='" & cp(9) & "' and np06 is null and np07=" & 催審
      cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
      
      If txtCaseField(0) = 核准 Then
         '若來函性質為核淮時, 其案件性質為"101"~107,113,114,501,或3開頭, 2007/8/3加424請求繼續審查,才要寫入PermitRecord
         'Modify by Morgan 2008/5/23 +122 CA申請
         'Modify by Sonia  2013/11/15 加復審805(CFP-023821)
         'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
         'Modified by Lydia 2016/08/26 +438 再考量試行計畫(AFCP2.0)
         If (cp(10) >= "101" And cp(10) <= "107" And cp(10) <> "106") Or cp(10) = "113" Or cp(10) = "114" Or cp(10) = "501" Or cp(10) = "424" Or cp(10) = "126" Or cp(10) = "805" Or cp(10) = "122" Or Left(cp(10), 1) = "3" Or cp(10) = "438" Then
            strTxt(iStep) = "DELETE FROM PERMITRECORD WHERE PR01 = " & CNULL(cp(1)) & " AND " & _
               "PR02 = " & CNULL(cp(2)) & " AND PR03 = " & CNULL(cp(3)) & " AND " & _
               "PR04 = " & CNULL(cp(4))
            cnnConnection.Execute strTxt(iStep)
            iStep = iStep + 1
   
            strTxt(iStep) = "insert into permitrecord (pr01,pr02,pr03,pr04,pr05) values(" + _
               CNULL(cp(1)) + "," + CNULL(cp(2)) + "," + CNULL(cp(3)) + "," + CNULL(cp(4)) + "," + CNULL(TransDate(txtCaseField(4), 2)) + ")"
            cnnConnection.Execute strTxt(iStep)
            iStep = iStep + 1
         End If
         '2013/5/15 add by sonia 日本新申請案核准,若該案有讓渡701時同時上核准(Trigger會自動更新其催審上Y)
         If field(9) = "011" And (cp(10) >= "101" And cp(10) <= "103") Then
            strTxt(iStep) = "update caseprogress set cp24='1',cp25=" & IIf(txtCaseField(4) = "", strSrvDate(1), CNULL(TransDate(txtCaseField(4), 2), True)) & " where cp10='701' and cp27 is not null and cp24 is null and " & _
                            "cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'"
            cnnConnection.Execute strTxt(iStep)
            iStep = iStep + 1
         End If
         '2013/5/15 end
      End If
   '2009/10/20 add by sonia 申請檢索報告421的檢索報告來函要更新申請檢索報告的催審期限為Y
   '2009/11/3 modify by sonia 加申請技術評價書423
   ElseIf txtCaseField(0) = 檢索報告 Then
      If (cp(10) = "421" Or cp(10) = "423") Then
         strTxt(iStep) = "update nextprogress set np06='Y' where np01='" & cp(9) & "' and np06 is null and np07=" & 催審
         cnnConnection.Execute strTxt(iStep)
         iStep = iStep + 1
         '2009/11/3 add by sonia 同時更新申請檢索報告421或申請技術評價書423的核准
         strTxt(iStep) = "update caseprogress set cp24='1',cp25=" & IIf(txtCaseField(4) = "", strSrvDate(1), CNULL(TransDate(txtCaseField(4), 2), True)) & " where cp09='" & cp(9) & "' "
         cnnConnection.Execute strTxt(iStep)
         iStep = iStep + 1
         '2009/11/3 end
      End If
      
      'Added by Morgan 2012/3/21
      'EPC發明
      If field(9) = "221" And field(8) = "1" Then
         strSql = "update nextprogress set np06='Y' where np02='" & field(1) & "' and np03='" & field(2) & "' and np04='" & field(3) & "' and np05='" & field(4) & "' and np06 is null and np07='" & 檢索報告 & "'"
         cnnConnection.Execute strSql, intI
      End If
      'end 2012/3/21
      
   '2009/10/20 end
   
   'Added by Morgan 2023/6/13
   ElseIf txtCaseField(0) = "1008" Then
      strTxt(iStep) = "update nextprogress set np06='Y' where np01='" & cp(9) & "' and np06 is null and np07=" & 催審
      cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
   'end 2023/6/13
   End If

   'Add by Morgan 2004/5/28
   '暫改以便收CPS案件
   'Memo by Lydia 2020/12/01 以field() 更新 Patent
   If field(1) <> "CPS" Then
      strTxt(iStep) = GetPASQL(field())
       cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
   End If
   
   m_PrintCForm = ""
   'edit by nickc 2007/02/02
   'Dim strDataTemp(1 To T_CP) As String
   Dim strDataTemp() As String
   ReDim strDataTemp(1 To TF_CP) As String
   
   strDataTemp(1) = cp(1)
   strDataTemp(2) = cp(2)
   strDataTemp(3) = cp(3)
   strDataTemp(4) = cp(4)
   strDataTemp(5) = strSrvDate(1)
   strDataTemp(6) = txtCaseField(2)
   strDataTemp(7) = txtCaseField(3)
   strDataTemp(9) = 主管機關來函
   'Added by Lydia 2015/10/02 部份案件性質之核准1001改為核發1008
   If txtCaseField(0) = "1001" And InStr(Patent1001Display, cp(10)) > 0 Then
       strDataTemp(10) = "1008"
   Else
       strDataTemp(10) = txtCaseField(0)
   End If
   'end 2015/10/02
   
   'Added by Lydia 2016/12/26 英國核准先行通知(1005),不輸入下一程序,法定期限及本所期限欄存定稿例外欄位
   If field(9) = "201" And strDataTemp(10) = "1005" Then
      strDataTemp(6) = "": strDataTemp(7) = ""
      strDataTemp(142) = TransDate(txtCaseField(3), 2) 'Added by Morgan 2021/2/26 暫存輸入的法限以便於2次確認檢查
   End If
   'end 2016/12/26
   
   '2009/12/30 MODIFY BY SONIA CFP-022728
   'strDataTemp(13) = cp(13)
   'strDataTemp(12) = cp(12)
   strDataTemp(13) = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
   strDataTemp(12) = GetSalesArea(strDataTemp(13))
   '2009/12/30 END
   strDataTemp(14) = txtCaseField(12)
   strDataTemp(26) = txtCaseField(14)
   
   strDataTemp(16) = ""
   strDataTemp(17) = ""
   strDataTemp(18) = ""

   '2007/8/3 MODIFY BY SONIA 加424請求繼續審查
   'Modify by Morgan 2008/5/23 +122 CA申請
   'Modify by Sonia  2013/11/15 加復審805(CFP-023821)
   'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
   'Modified by Lydia 2016/08/26 +438 再考量試行計畫(AFCP2.0)
   If (cp(10) >= "101" And cp(10) <= "107" And cp(10) <> "106") Or cp(10) = "113" Or cp(10) = "114" Or cp(10) = "501" Or cp(10) = "424" Or cp(10) = "126" Or cp(10) = "122" Or Left(cp(10), 1) = "3" Or cp(10) = "805" Or cp(10) = "438" Then
      If txtCaseField(0) = 核准 Then
      
'Removed by Morgan 2012/9/6 費用應該掛在專利證書--秀玲
'         strNA49 = ""
'         StrSQLa = "SELECT NA49,NA53,NA54 FROM NATION WHERE NA01='" & m_PA09 & "'"
'         rsA.CursorLocation = adUseClient
'         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount > 0 Then
'            Select Case m_PA08
'            Case "1"
'               strNA49 = "" & rsA.Fields(0).Value
'            Case "2"
'               strNA49 = "" & rsA.Fields(1).Value
'            Case "3"
'               strNA49 = "" & rsA.Fields(2).Value
'            End Select
'         End If
'         If rsA.State <> adStateClosed Then rsA.Close
'         Set rsA = Nothing
'         '若來函性質不為"核准"(1001)或申請國家非自動發註冊證
'         If strNA49 = "Y" Then
'            If txtCaseField(6) <> "" Or txtCaseField(7) <> "" Then
'               strDataTemp(16) = Val(txtCaseField(6)) + Val(txtCaseField(7))
'               strDataTemp(17) = Val(txtCaseField(6)) + Val(txtCaseField(7)) - (Val(txtCaseField(8)) * 1000)
'            End If
'            strDataTemp(18) = txtCaseField(8)
'         End If
'end 2012/9/6

         '核准直接上發文日
         strDataTemp(27) = strSrvDate(1)
         m_PrintCForm = "N"
      End If
   End If
   
   If strNA49 = "" Then
      strTemp1 = "N"
   Else
      strTemp1 = ""
   End If
   strDataTemp(20) = strTemp1
   strDataTemp(32) = strTemp1
   strDataTemp(36) = txtCaseField(15)
   strDataTemp(37) = txtCaseField(18)
   strDataTemp(38) = txtCaseField(19)
   strDataTemp(39) = txtCaseField(20)
   strDataTemp(40) = txtCaseField(21)
   strDataTemp(41) = txtCaseField(22)
   strDataTemp(42) = txtCaseField(23)
   strDataTemp(48) = txtCaseField(13)
   strDataTemp(43) = cp(9)
   '2008/8/26 modify by sonia 櫃台收文日改存 cp119
   strDataTemp(119) = ChangeTStringToWString(Me.lblCaseField(8).Caption)
   '2008/8/26 end
   
   '2008/10/24 MODIFY BY SONIA CP64仍存
   'Add by Morgan 2010/3/24 領證報價定稿要帶進度備註故不再紀錄(看櫃台收文日欄位就好)
   If txtCaseField(1) = "601" Then
      strDataTemp(64) = txtCaseField(17).Text
      'Added by Morgan 2023/6/12 俄羅斯申請紙本證書固定報價8,000(3)
      If field(9) = "023" Then
         strDataTemp(64) = "申請紙本證書8,000(3);" & strDataTemp(64)
      End If
      'end 2023/6/12
      'Modify by Morgan 2010/3/30 代理人費用一律帶到備註
      'If Val(txtCaseField(35)) > 0 Or Val(txtCaseField(35)) > 0 Then
         strExc(1) = ""
         If Val(txtCaseField(35)) > 0 Then
            strExc(1) = strExc(1) & "+補虧損" & Val(txtCaseField(35)) & "(" & txtCaseField(37) & ")"
         End If
         If Val(txtCaseField(36)) > 0 Then
            strExc(1) = strExc(1) & "+補未收費程序" & Val(txtCaseField(36)) & "(" & txtCaseField(38) & ")"
         End If
         strDataTemp(64) = "領證費" & txtCaseField(6) & "(" & txtCaseField(8) & ")=代理人費用" & txtCaseField(24) & "+點數" & (1000 * (Val(txtCaseField(8)) - Val(txtCaseField(37)) - Val(txtCaseField(38)))) & strExc(1) & ";" & strDataTemp(64)
      'End If
   Else
   'end 2010/3/24
      If txtCaseField(17) = "" Then
         strDataTemp(64) = "櫃台收文日：" & lblCaseField(8).Caption
      Else
         strDataTemp(64) = "櫃台收文日：" & lblCaseField(8).Caption & "，" & Me.txtCaseField(17).Text
      End If
   End If
   
   strDataTemp(144) = txtCaseField(39) 'Added by Morgan 2021/10/28 報價備註有輸就回寫，工程師才看的到 --禧佩
   
   'Added by Morgan 2023/12/11
   If m_USCaseNo <> "" Then
      If Val(txtIDSFee(1)) > 0 Then
         strDataTemp(64) = "IDS報價:1.第一階段 " & txtIDSFee(1) & "(" & txtIDSPt(1) & "P), 2.第二階段 " & txtIDSFee(2) & "(" & txtIDSPt(2) & "P);" & strDataTemp(64)
      ElseIf Val(txtIDSPt(2)) > 0 Then
         strDataTemp(64) = "IDS報價:第二階段 " & txtIDSFee(2) & "(" & txtIDSPt(2) & "P);" & strDataTemp(64)
      End If
   'Added by Morgan 2024/6/11 非報價定稿有輸入時帶到報價備註(OA有前案時要報給工程師) Ex:CFP-33147 --禧佩
   Else
      If Val(txtIDSFee(1)) > 0 Then
         strDataTemp(144) = "IDS報價:1.第一階段 " & txtIDSFee(1) & "(" & txtIDSPt(1) & "P), 2.第二階段 " & txtIDSFee(2) & "(" & txtIDSPt(2) & "P);" & strDataTemp(144)
      ElseIf Val(txtIDSPt(2)) > 0 Then
         strDataTemp(144) = "IDS報價:第二階段 " & txtIDSFee(2) & "(" & txtIDSPt(2) & "P);" & strDataTemp(144)
      End If
   End If
   'end 2023/12/11
   
   'Add by Morgan 2010/7/19
   strDataTemp(133) = DBDATE(lblCaseField(8))
   
   '2008/11/28 ADD BY SONIA 加存約定期限於CP64,方能印在C類接洽單上
   'Removed by Morgan 2021/9/2 改都放在 NP23 ( C類接洽單列印也有一併修改 )
   'If txtCaseField(30).Text <> "" Then
   '   strDataTemp(64) = strDataTemp(64) & "，約定期限：" & txtCaseField(30).Text
   'End If
   'end 2021/9/2
   '2008/11/28 END
   
   'Added by Lydia 2020/11/19 CFP英國脫歐案管制：於進度備註CP64加註'英國脫歐案專利號數'
   If field(1) = "CFP" And field(9) = "239" And strDataTemp(10) = "1608" And txtCaseField(26).Visible = True And Trim(txtCaseField(26)) <> "" Then
     strDataTemp(64) = strDataTemp(64) & "，英國脫歐案專利號數：" & Trim(txtCaseField(26).Text)
     strDataTemp(30) = Trim(txtCaseField(26))
   End If
   'end 2020/11/19
   'Added by Lydia 2020/12/01 CFP英國脫歐案管制：輸在英國案則註冊號數就直接更新基本檔
   If field(1) = "CFP" And field(9) = "201" And strDataTemp(10) = "1608" And txtCaseField(26).Visible = True And Trim(txtCaseField(26)) <> "" Then
       strSql = "Update Patent set PA22=" & CNULL(ChgSQL(Trim(txtCaseField(26)))) & " where pa01='" & field(1) & "' and pa02='" & field(2) & "' and pa03='" & field(3) & "' and pa04='" & field(4) & "' "
       cnnConnection.Execute strSql
   End If
   'end 2020/12/01
   
   '92.2.27 ADD BY SONIA 核駁, 通知要求遻取, 通知提供前案 為不算案件數 '92.3.20 檢索報告
   'Modify by Morgan 2006/3/23 加最終核駁1006
   'Modify by Morgan 2010/3/10 +建議性處分書 1220
   If txtCaseField(0) = 核駁 Or txtCaseField(0) = "1006" Or txtCaseField(0) = "1220" Or txtCaseField(0) = 通知提供前案 Or txtCaseField(0) = 通知要求選取 Or txtCaseField(0) = 檢索報告 Then
      strDataTemp(26) = "N"
   End If
   '92.2.27 END
   '92.3.7 ADD BY SONIA
   If txtCaseField(0) = "1908" Then
      m_PrintCForm = "N"
      strDataTemp(27) = strSrvDate(1)
      strDataTemp(26) = "N"
      strDataTemp(32) = ""
      strDataTemp(20) = ""
      strDataTemp(16) = Val(txtCaseField(6))
      strDataTemp(17) = Val(txtCaseField(6)) - (Val(txtCaseField(8)) * 1000)
      strDataTemp(18) = Val(txtCaseField(8))
   End If
   If txtCaseField(29) = "Y" Then
      m_PrintCForm = "N"
   End If

   '92.3.23 ADD BY SONIA 承辦人為程序時, 不印C類接洽單且直接上發文日
   If txtCaseField(12) <> "" Then
      If GetStaffDepartment(txtCaseField(12)) = "P12" Then
         m_PrintCForm = "N"
         strDataTemp(27) = strSrvDate(1)
         'Add by Morgan 2008/8/8 改印定稿時上發文日
         If txtCaseField(0) = 核准 And Val(strSrvDate(1)) > 20080810 And txtCaseField(1) = "601" Then
            strDataTemp(27) = ""
            'Added by Morgan 2018/10/4 承辦期限=系統日+5個工作天, 固定設要出通知函
            strDataTemp(48) = CompWorkDay(5, strSrvDate(1))
            txtCaseField(5) = ""
            'end 2018/10/4
         'Added by Morgan 2022/1/27
         '印尼核准有報價定稿不上發文日 Ex:CFP-030831
         ElseIf m_IDNGrant And txtCaseField(5) = "" Then
            strDataTemp(27) = ""
         'end 2022/1/27
         End If
      'Add by Morgan 2009/5/8
      '輸核准時若承辦為工程師(非程序)則不出客戶定稿,不上發文日
      ElseIf txtCaseField(0) = 核准 Then
         strDataTemp(27) = ""
         txtCaseField(5) = "N"
      'end 2009/5/8
      End If
   End If
   
   'Add by Morgan 2011/4/20
   'EPC收到檢索報告時要更新實審及指定費的期限(准駁通知日+6個月)
   'Modified by Morgan 2022/12/26 未公開才要更新 Ex:CFP-033058--玫音
   If field(9) = "221" And txtCaseField(0) = "1209" And txtCaseField(4) <> "" And Val(field(12)) = 0 Then
      strDataTemp(64) = "准駁通知日：" & txtCaseField(4) & ";" & strDataTemp(64)
      strExc(1) = CompDate(1, 6, txtCaseField(4))  '法限
      strDateS(0) = ""
      strDateS(1) = field(1)
      strDateS(2) = field(9)
      strDateS(3) = strExc(1)
      GetCtrlDT strDateS
      strExc(2) = PUB_GetWorkDay1(strDateS(0), True) '所限
      strSql = "update caseprogress set cp06=" & strExc(2) & ",cp07=" & strExc(1) & _
         " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "'" & _
         " and cp04='" & cp(4) & "' and cp10 in ('416','215') and cp27||cp57 is null" & _
         " and cp07<" & strExc(1)
      cnnConnection.Execute strSql, intI
      
      strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & _
         " where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "'" & _
         " and np05='" & cp(4) & "' and np07 in ('416','215') and np06 is null" & _
         " and np09<" & strExc(1)
      cnnConnection.Execute strSql, intI
   End If
   
   '2011/5/26 ADD BY SONIA 核駁1002,最終核駁1006的費用點數不印C類接洽單改存CP144
   'If m_PrintCForm <> "N" Then   'CANCEL BY SONIA 2014/6/25 CFP-026462 掛程序
      'Modify by Morgan 2012/7/20 +建議性處分書 1220
      'modify by sonia 2018/8/21 +依職權電話通知修正1225
      If InStr("1002,1006,1201,1203,1205,1206,1209,1220,1225,1401,1307,1801,1802", txtCaseField(0)) > 0 And txtCaseField(6) <> "" Then
         '2011/6/2 modify by sonia 加報價備註欄
         'Modified by Morgan 2025/5/28
         'strDataTemp(144) = lblNextCaseProperty & "費用" & Format(txtCaseField(6), "#,###") & "(" & txtCaseField(8) & "P);" & txtCaseField(39)
         strDataTemp(144) = lblNextCaseProperty & "費用" & Format(txtCaseField(6), "#,###") & "(" & txtCaseField(8) & "P);" & strDataTemp(144)
      End If
   'End If
   '2011/5/26 END
   
    'Modified by Morgan 2018/10/4 從下面移上來併入新增語法
    '智權人員存最近收文A類接洽記錄單的智權人員
    stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
    stCP12 = GetSalesArea(stCP13)
    strDataTemp(13) = stCP13
    strDataTemp(12) = stCP12
    'end 2018/10/4
    
    'Added by Morgan 2023/6/6
    If txtCaseField(0) = "1905" Then
      If m_bolWebQuery Then
         strExc(0) = "官網"
      Else
         strExc(0) = "官方"
      End If
      strExc(0) = strExc(0) & "預估月數：" & m_intEstMonths & ";"
      strDataTemp(64) = strExc(0) & strDataTemp(64)
    End If
    'end 2023/6/6
    
   '新增案件進度檔
   strTxt(iStep) = GetCPSQL(strDataTemp(), False)
   cnnConnection.Execute strTxt(iStep)
   iStep = iStep + 1
   
   m_NewCP09 = strDataTemp(9)
   'Added by Morgan 2018/10/4 承辦期限要單獨更新(因C類收文 Trigger 會自動上齊備日並設定承辦期限)
   If strDataTemp(48) <> "" Then
      strTxt(iStep) = "Update Caseprogress Set CP48='" & strDataTemp(48) & "' Where CP09='" & strDataTemp(9) & "' "
      cnnConnection.Execute strTxt(iStep), intI
      iStep = iStep + 1
   End If
   'end 2018/10/4
   
   '2010/1/20 add by sonia 承辦人為分所人員以系統日的下一個工作天上齊備日
   If m_CP14ST06 <> "1" And strDataTemp(27) = "" Then
      strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & CompWorkDay(2, strSrvDate(1), 0) & " WHERE EP02='" & strDataTemp(9) & "'"
      cnnConnection.Execute strSql
   'Add by Morgan 2010/10/1
   ElseIf strDataTemp(48) = "" Then
      strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & strDataTemp(9) & "'"
      cnnConnection.Execute strSql
   End If
   '2010/1/20 end
    
   'Add by Morgan 2005/1/20 韓國發明核准且為一案兩請則新增新型案自請撤回下一程序
   'Modify by Morgan 2005/5/13 改自請撤回(413)為放棄專利權(429)
   m_strRetSheet2NP07 = ""
   If m_bolIsDualApp = True Then
      lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax
      strSql = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22) values (" + _
         CNULL(strDataTemp(9)) + "," + CNULL(m_stUPA(1)) + "," + CNULL(m_stUPA(2)) + "," + CNULL(m_stUPA(3)) + _
         "," + CNULL(m_stUPA(4)) + ",429," + TransDate(txtCaseField(2), 2) + "," & TransDate(txtCaseField(3), 2) & _
         "," + CNULL(PUB_GetAKindSalesNo(m_stUPA(1), m_stUPA(2), m_stUPA(3), m_stUPA(4))) + "," & lMax & ")"
      cnnConnection.Execute strSql
      m_strRetSheet2NP07 = "429"
   End If
    
   '抓最新的AB類發文代理人更新
   Pub_UpdateFromMaxCP27 cp(1), cp(2), cp(3), cp(4)
    
   stCP09 = strDataTemp(9): stCP14 = strDataTemp(14): stCP27 = strDataTemp(27)
   
   'Added by Morgan 2018/2/8 +土耳其發明案實審法限為檢索報告通知日起算3個月
   If field(9) = "235" And field(8) = "1" And txtCaseField(0) = "1209" And txtCaseField(4) <> "" Then
      'Added by Morgan 2020/2/17 有回覆檢索報告期限則設為該期限，沒有時才用檢索報告通知日+3個月 Ex:CFP-030467 --甄妮
      If txtCaseField(1) = "218" Then
         strExc(1) = TransDate(txtCaseField(3), 2)
         strExc(2) = TransDate(txtCaseField(2), 2)
      Else
      'end 2020/2/17
      
         strExc(1) = CompDate(1, 3, txtCaseField(4))
         strDateS(1) = field(1)
         strDateS(2) = field(9)
         strDateS(3) = strExc(1)
         GetCtrlDT strDateS
         '所限
         strExc(2) = strDateS(0)
         strExc(2) = PUB_GetWorkDay1(strExc(2), True)
         
      End If 'Added by Morgan 2020/2/17
      
      
      strSql = "Select cp09,cp27 From caseprogress Where cp01='" & cp(1) & "' AND cp02='" & cp(2) & "' AND cp03='" & cp(3) & "' AND cp04='" & cp(4) & "' AND cp10='416' and cp57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If IsNull(RsTemp("cp27")) Then
            strSql = "update caseprogress set cp06=" & strExc(2) & ",cp07=" & strExc(1) & " where cp09='" & RsTemp("cp09") & "'"
            cnnConnection.Execute strSql
         End If
      Else
         strSql = "Select NP01,NP07,NP22 From Nextprogress Where NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP07='416' AND NP06 IS NULL  ORDER BY NP22 DESC"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " WHERE NP01='" & RsTemp("NP01") & "' AND NP22=" & RsTemp("NP22")
         Else
            strSql = "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               " select '" & strDataTemp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','416'," & strExc(2) & "," & strExc(1) & ",'" & strDataTemp(13) & "',newNP22 from dual,(select nvl(max(np22),0)+1 newNP22 from nextprogress)"
         End If
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2018/2/8
   
   lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax
   bolNP22 = False
   iNP22 = 1

   '2007/8/3 MODIFY BY SONIA 加424請求繼續審查
   'Modify by Morgan 2008/5/23 +122 CA申請
   'Modify by Sonia  2013/11/15 加復審805(CFP-023821)
   'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
   'Modified by Lydia 2016/08/26 +438 再考量試行計畫(AFCP2.0)
   If (cp(10) >= "101" And cp(10) <= "107" And cp(10) <> "106") Or cp(10) = "113" Or cp(10) = "114" Or cp(10) = "501" Or cp(10) = "424" Or cp(10) = "126" Or cp(10) = "122" Or Left(cp(10), 1) = "3" Or cp(10) = "805" Or cp(10) = "438" Then
      If txtCaseField(0) = 核准 Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, 年費, strTemp2) = 0 Then
         'Modified by Morgan 2013/10/23
         'If ClsPDGetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, 年費, strTemp2) = 0 Then
         If ClsPDGetNationTaxEx(Val(field(8)), field(9), strTemp, strTemp1, 年費, strTemp2, , , field(10), field(21), field(72)) = 0 Then
            varTemp = Split(strTemp1, ",")
               m_varTemp = Split(strTemp1, ",")
            i = GetMoneyYears(field(72))
               m_i = i
            
            If i > UBound(varTemp) + 1 Then GoTo Nextstep
            
            dobDateAdd = varTemp(i - 1)
            If Not GetNP07(field(9), field(8), strExc(10)) Then Exit Function
            
            m_NP07 = strExc(10)
            'Add by Morgan 2005/3/16 年費要減一年
            If m_NP07 = "605" Then dobDateAdd = dobDateAdd - 1
            m_dobDateAdd = dobDateAdd
            strStartDate = GetStartDate(strTemp, cp(), field())
            m_strStartDate = strStartDate
            If strStartDate <> "" Then
               'Modified by Morgan 2019/12/4 統一用函數(規則集中)
               'strStartDate = CompDate(0, dobDateAdd, strStartDate)
               ''法定期限不必減一天
               'strDate = strStartDate
               '
               ''Add by Morgan 2011/2/23 沙烏地阿拉伯,年費期限為每年3/31
               'If field(9) = "021" Then
               '   strDate = Mid(strDate, 1, 4) & "0331"
               'End If
               ''2011/2/23 END
               strDate = DBDATE(GetFeeNextDate(m_strStartDate, m_dobDateAdd, field(9), field(8)))
               'end 2019/12/4
            Else
               GoTo Nextstep
            End If
            
            'Added by Morgan 2021/11/4
            '印尼發明/新型核准
            '核准日6個月內須繳交自申請日起算累計至核准日次年之年費，之後則逐年提前於屆滿前1個月(申請日)繳交
            If m_IDNGrant And m_NP07 = "605" Then
               'Modified by Morgan 2022/3/22
               'strDate = CompDate(1, 6, txtCaseField(4))
               strDate = GetIDN1st605FeeDate(txtCaseField(4))
               'end 2022/3/22
            End If
            'end 2021/11/4
            
            '本所期限
            strExc(1) = field(1)
            strExc(2) = field(9)
            strExc(3) = TransDate(strDate, 2)
            GetCtrlDT strExc
            strDate1 = DBDATE(strExc(0))
            
           '2008/4/8 add by sonia CFP-020190第1次年費若小於領證期限應為准後繳則不必掛年費期限
           '2008/10/23 cancel by sonia 改在領證分案限制一定要勾選年費期限,並於領證發文時一定要輸年費年度
           'If strDate < TransDate(txtCaseField(3), 2) Then
           '    GoTo NextStep
           'End If
           '2008/10/23 end
           '2008/4/8 END
           m_strDate = strDate
           m_strDate1 = strDate1
           m_blnCompNextDate = IIf(Left(strDate, 4) < Left(strSrvDate(1), 4) = True, True, False)
           
            '2005/8/3 ADD BY SONIA
            strExc(0) = "SELECT CP09 FROM CASEPROGRESS WHERE CP10=" & m_NP07 & " and CP01=" + CNULL(cp(1)) + _
               " and CP02=" + CNULL(cp(2)) + " and CP03=" + CNULL(cp(3)) + " and CP04=" + CNULL(cp(4)) & " and CP27 IS NULL AND CP57 IS NULL "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strTxt(iStep) = "update CASEPROGRESS set CP06=" + PUB_GetWorkDay1(m_strDate1, True) + ",CP07=" + m_strDate + " WHERE CP10=" & m_NP07 & _
                  " and CP01=" + CNULL(cp(1)) + _
                  " and CP02=" + CNULL(cp(2)) + " and CP03=" + CNULL(cp(3)) + " and CP04=" + CNULL(cp(4)) & " and CP27 IS NULL AND CP57 IS NULL "
               cnnConnection.Execute strTxt(iStep)
               iStep = iStep + 1
            Else
            '2005/8/3 END
            
               strExc(0) = "SELECT NP22 FROM NEXTPROGRESS WHERE np07=" & m_NP07 & " and np02=" + CNULL(cp(1)) + _
                  " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05=" + CNULL(cp(4)) & " and np06 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  '判斷是否要更新下一程序
                  If blnUpdateNP(cp(1) & cp(2) & cp(3) & cp(4)) = True Then
                      'Added by Morgan 2021/11/4
                      '印尼發明/新型核准
                      If m_IDNGrant And m_NP07 = "605" Then
                           m_strDate1 = PUB_GetWorkDay1(m_strDate1, True)
                           m_strNP22 = lMax
                      End If
                      'end 2021/11/4
                      
                      '智權人員存最近收文A類接洽記錄單的智權人員
                      '若本所期限非工作天則抓最近的工作天
                      strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22) values (" + _
                         CNULL(strDataTemp(9)) + "," + CNULL(cp(1)) + "," + CNULL(cp(2)) + "," + CNULL(cp(3)) + _
                         "," + CNULL(cp(4)) + "," + CNULL(m_NP07) + "," & Val(PUB_GetWorkDay1(m_strDate1, True)) & "," & Val(m_strDate) & _
                         "," + CNULL(PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))) + "," & lMax & ")"
                      cnnConnection.Execute strTxt(iStep)
                  End If
                  iStep = iStep + 1
                  lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax
                  
               Else
                  strTxt(iStep) = "update nextprogress set np08=" + PUB_GetWorkDay1(m_strDate1, True) + ",np09=" + m_strDate + " WHERE np07=" & m_NP07 & _
                     " and np02=" + CNULL(cp(1)) + _
                     " and np03=" + CNULL(cp(2)) + " and np04=" + CNULL(cp(3)) + " and np05=" + CNULL(cp(4)) & " and np06 is null"
                  cnnConnection.Execute strTxt(iStep)
                  iStep = iStep + 1
               End If
            End If
         End If
      End If
   End If
   
Nextstep:

   If txtCaseField(21) <> "" Then
      strTemp = txtCaseField(21)
   ElseIf cp(41) <> "" Then
      strTemp = txtCaseField(22)
   Else
      strTemp = txtCaseField(23)
   End If

   If txtCaseField(1) <> "" Then
      If txtCaseField(1) = 催審 Or txtCaseField(1) = 提申 Or txtCaseField(1) = 收達 Then
            '智權人員存最近收文A類接洽記錄單的智權人員
         strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np13,np14,np15,np22) values (" + _
            CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
            "," + CNULL(strDataTemp(4)) + "," + CNULL(txtCaseField(1)) + "," & CNULL(TransDate(txtCaseField(2), 2)) & "," & _
            CNULL(TransDate(txtCaseField(3), 2)) & "," + CNULL(PUB_GetAKindSalesNo(strDataTemp(1), strDataTemp(2), strDataTemp(3), strDataTemp(4))) + "," + CNULL(txtCaseField(16)) + "," + _
            CNULL(ChgSQL(strTemp)) + "," + CNULL(txtCaseField(17)) + "," & lMax & ")"
        cnnConnection.Execute strTxt(iStep)
         iStep = iStep + 1
      Else
         strExc(0) = txtCaseField(17)
         If Text5(0).Enabled = True And Text5(0) <> "" Then
            strExc(0) = IIf(strExc(0) <> "", "；", "") & "含第" & Text5(0) & "至" & Text5(1) & "年年費；"
         End If
         
         '答辯同時要寫約定期限
         strExc(1) = ""
         'Modified by Morgan 2016/3/3 +126 期末拋棄
         'Modified by Lydia 2016/08/28 +438 再考量試行計畫(AFCP2.0)
         'Modified by Morgan 2021/9/2 這裡不必限制案件性質，有日期就回寫，存檔前檢查控制就好，否則增加案件性質又要改
         'If (txtCaseField(1) = "107" Or txtCaseField(1) = "126" Or txtCaseField(1) = "438") And txtCaseField(30) <> "" Then
         If txtCaseField(30) <> "" Then
            strExc(1) = DBDATE(txtCaseField(30))
         End If
         
         'Added by Morgan 2023/5/23
         '泰國通知繳公開費
         bolAddNP = True
         If field(9) = "019" And txtCaseField(0) = "1236" And txtCaseField(1) = "217" Then
            '已收文
            If PUB_ChkCPExist(field, "217", , strTemp, , , strDate) Then
               bolAddNP = False
               m_bolAdd217BCP = False
               If strDate = "" Then
                  strSql = "update caseprogress set cp06=" & DBDATE(txtCaseField(2)) & ",cp07=" & DBDATE(txtCaseField(3)) & " where cp09='" & strTemp & "'"
                  cnnConnection.Execute strSql, intI
                  m_Alert = "公開費可發文！" 'Added by Morgan 2024/7/12
               End If
               
               'Added by Morgan 2024/7/12
               strSql = "update caseprogress set cp27=19221111 where cp09='" & strDataTemp(9) & "'"
               cnnConnection.Execute strSql, intI
               'end 2024/7/12
            
            ElseIf PUB_ChkNPExist(cp, "217", 0, strNP22, strTemp) Then
               bolAddNP = False
               strSql = "Update NextProgress Set NP08=" & DBDATE(txtCaseField(2)) & ",NP09=" & DBDATE(txtCaseField(3)) & " Where NP22=" & strNP22 & " and NP01='" & strTemp & "'"
               cnnConnection.Execute strSql, intI
               m_strNP22 = strNP22
            End If
         End If
         If bolAddNP Then
         'end 2023/5/23
         
             '智權人員存最近收文A類接洽記錄單的智權人員
             strTxt(iStep) = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np13,np14,np15,np22,np23) values (" + _
                CNULL(strDataTemp(9)) + "," + CNULL(strDataTemp(1)) + "," + CNULL(strDataTemp(2)) + "," + CNULL(strDataTemp(3)) + _
                "," + CNULL(strDataTemp(4)) + "," + CNULL(txtCaseField(1)) + "," & CNULL(TransDate(txtCaseField(2), 2)) & "," & _
                CNULL(TransDate(txtCaseField(3), 2)) & "," + CNULL(PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))) + "," + CNULL(txtCaseField(16)) + "," + _
                CNULL(ChgSQL(strTemp)) + "," + CNULL(ChgSQL(strExc(0))) + "," & lMax & "," & CNULL(strExc(1), True) & ")"
            cnnConnection.Execute strTxt(iStep)
             iStep = iStep + 1
             
             m_strNP22 = lMax 'Add by Morgan 2008/5/8
             
         End If
         
         'Added by Morgan 2021/6/3 內部收文公開費
         If m_bolAdd217BCP Then
            strExc(1) = AutoNo("B", 6)
            strTxt(iStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
               "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP30,CP32,CP43) VALUES " & _
               "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & _
               "," & DBDATE(txtCaseField(2)) & "," & DBDATE(txtCaseField(3)) & _
               ",'" & strExc(1) & "','217','90'," & CNULL(stCP12) & "," & CNULL(stCP13) & _
               ",'" & strUserNum & "','N','N','" & m_strNP22 & "','N','" & cp(9) & "') "
         
            cnnConnection.Execute strTxt(iStep), intI
            iStep = iStep + 1
            strSql = "update nextprogress set np06='Y',np24='" & strExc(1) & "' where np01='" & strDataTemp(9) & "' and np07='217' and np22=" & m_strNP22
            cnnConnection.Execute strSql, intI
         End If
         'end 2021/6/3
         
      End If
      
      bolNP22 = True
      NP22(iNP22) = lMax
      iNP22 = iNP22 + 1
'      lMax = lMax + 1
        lMax = ClsLawGetMax   'edit by nickc 2007/02/05 不用 dll 了  objLawDll.GetMax
   End If
'92.6.8王協理同意掛期限但不單獨印接洽結案單, 和領證印在一起, 以防二張接洽單收文後漏發文
'91.12.28 cancel by sonia 王協理說先不掛公開費期限故不印接洽結案單, 智權人員收文時直接含公開費用
   'Modify by Morgan 2007/3/27 96/4/1 以後不再掛期限--郭
   
   'end 2007/3/27
   
    'Added by Lydia 2015/04/10 申請人可在該國有多筆識別番號->call frm880021
'   '更新申請人國外ID對照檔
'   If Me.txtCaseField(26).Text <> "" Then
'      StrSQLa = "Select AFID03 FROM APPLICANTFOREIGNID WHERE AFID01='" & Left(m_PA26 & "00000000", 8) & "' AND AFID02='" & m_PA09 & "'"
'      rsA.CursorLocation = adUseClient
'      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'         If Me.txtCaseField(26).Text <> "" & rsA.Fields(0).Value Then
'            strTxt(iStep) = "UPDATE APPLICANTFOREIGNID SET AFID='" & Me.txtCaseField(26).Text & _
'               "' WHERE AFID01='" & Left(m_PA26 & "00000000", 8) & "' AND AFID02='" & m_PA09 & "'"
'            cnnConnection.Execute strTxt(iStep)
'            iStep = iStep + 1
'         End If
'      Else
'         strTxt(iStep) = "INSERT INTO APPLICANTFOREIGNID VALUES('" & Left(m_PA26 & "00000000", 8) & "','" & m_PA09 & "','" & Me.txtCaseField(26).Text & "')"
'        cnnConnection.Execute strTxt(iStep)
'         iStep = iStep + 1
'      End If
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'   Else
'      strTxt(iStep) = "DELETE FROM APPLICANTFOREIGNID WHERE AFID01='" & Left(m_PA26 & "00000000", 8) & "' AND AFID02='" & m_PA09 & "'"
'        cnnConnection.Execute strTxt(iStep)
'      iStep = iStep + 1
'   End If
    'end 2015/04/10
    
   '若有輸入修圖資料時, 新增一筆"B"類案件進度資料
   If Me.txtCaseField(27).Text <> "" And Me.txtCaseField(28).Text <> "" Then
        stCP13 = PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))
        stCP12 = GetSalesArea(stCP13)
        'Modified by Morgan 2016/3/3 承辦人&繪圖人員改87025(72006)
        strTxt(iStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
            "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP64,CP29) VALUES " & _
            "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & _
            strSrvDate(1) & "," & TransDate(txtCaseField(27), 2) & "," & TransDate(txtCaseField(28), 2) & _
            ",'" & AutoNo("B", 6) & "','" & 其他 & "','90'," & CNULL(stCP12) & "," & CNULL(stCP13) & _
            ",'87025','N','N','N','" & cp(9) & "','修圖','87025') "
      
        cnnConnection.Execute strTxt(iStep)
      iStep = iStep + 1
   End If

    '變更事項將變更基本檔或進度檔的動作, 移至變更發文時做
   strReceiveNo = cp(9)
   If cp(10) = 變更 And Me.txtCaseField(0).Text = 核准 Then
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
         strExc(2) = ""
         strExc(3) = ""
   
         '申請日 10
         If strCe(2) <> "" Then
            strExc(1) = strExc(1) & "申請日 : " & strCe(2) & ","
            If intCaseKind = 專利 Then
               strExc(2) = strExc(2) & "PA10=" & strCe(2) & ","
            Else
               strExc(2) = strExc(2) & "SP10=" & strCe(2) & ","
            End If
            strExc(3) = strExc(3) & "CE03='1',"
         End If
         
         '申請人 26-30
         bolChk = False
         For i = 4 To 8
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
         If bolChk = True Then
            strExc(1) = strExc(1) & "申請人 : "
            For i = 4 To 8
               If strCe(i) <> "" Then
                  strExc(1) = strExc(1) & strCe(i) & ","
                  'edit by nickc 2007/02/02 不用 dll 了
                  'If objPublicData.GetCustomerNameAndAddress(strCe(i), strTmp(5), strTmp(1), strTmp(2), strTmp(3)) Then
                  If ClsPDGetCustomerNameAndAddress(strCe(i), strTmp(5), strTmp(1), strTmp(2), strTmp(3)) Then
                     If intCaseKind = 專利 Then
                        strExc(2) = strExc(2) & "PA" & i + 27 & "=" & CNULL(ChgSQL(strTmp(1))) & ",PA" & i + 32 & "=" & CNULL(ChgSQL(strTmp(2))) & ",PA" & i + 37 & "=" & CNULL(ChgSQL(strTmp(3))) & ","
                     End If
                  End If
               End If
               If intCaseKind = 專利 Then
                  strExc(2) = strExc(2) & "PA" & i + 22 & "=" & CNULL(ChangeCustomerL(strCe(i))) & ","
               End If
            Next
            If intCaseKind <> 專利 Then
               strExc(2) = strExc(2) & "SP08=" & CNULL(strCe(4)) & "," & "SP58=" & CNULL(strCe(5)) & "," & "SP59=" & CNULL(strCe(6)) & ","
            End If
            strExc(3) = strExc(3) & "CE09='1',"
         Else
            '申請地址 31-45
            bolChk = False
            For i = 23 To 37
               If strCe(i) <> "" Then
                  bolChk = True
                  Exit For
               End If
            Next
            If bolChk = True Then
               strExc(1) = strExc(1) & "申請地址 : "
               For i = 23 To 37
                  If strCe(i) <> "" Then
                     strExc(1) = strExc(1) & strCe(i) & ","
                  End If
                  strExc(2) = strExc(2) & "PA" & i + 8 & "=" & CNULL(strCe(i)) & ","
               Next
               strExc(3) = strExc(3) & "CE38='1',"
            End If
         End If
         
         '專利商標種類代號 08
         If strCe(39) <> "" Then
            strExc(1) = strExc(1) & "專利商標種類代號 : " & strCe(39) & ","
            If intCaseKind = 專利 Then
               strExc(2) = strExc(2) & "PA08='" & strCe(39) & "',"
            End If
            strExc(3) = strExc(3) & "CE40='1',"
         End If
   
         '案件名稱 05-07
         bolChk = False
         For i = 41 To 43
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
         If bolChk = True Then
            strExc(1) = strExc(1) & "案件名稱 : "
            For i = 41 To 43
               If strCe(i) <> "" Then
                  strExc(1) = strExc(1) & strCe(i) & ","
               End If
               If intCaseKind = 專利 Then
                  strExc(2) = strExc(2) & "PA" & i - 36 & "=" & CNULL(strCe(i)) & ","
               Else
                  strExc(2) = strExc(2) & "SP" & i - 36 & "=" & CNULL(strCe(i)) & ","
               End If
            Next
            strExc(3) = strExc(3) & "CE44='1',"
         End If
   
         '代表人 79-84
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
   
         If bolChk Then
            strExc(1) = strExc(1) & "代表人 : "
            For i = 10 To 15
               If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
               If intCaseKind = 專利 Then
                  strExc(2) = strExc(2) & "PA" & i + 69 & "=" & CNULL(strCe(i)) & ","
               End If
            Next
            For i = 68 To 91
               If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
               If intCaseKind = 專利 Then
                  strExc(2) = strExc(2) & "PA" & i + 41 & "=" & CNULL(strCe(i)) & ","
               End If
            Next
            If intCaseKind <> 專利 Then
               strExc(2) = strExc(2) & "SP42=" & CNULL(strCe(10)) & ","
            End If
            strExc(3) = strExc(3) & "CE16='1',"
         End If
         
         '代表人中譯文
         If Not bolChk Then
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
            If bolChk Then
               strExc(1) = strExc(1) & "代表人中譯文 : "
               If intCaseKind = 專利 Then
                  strExc(2) = strExc(2) & "PA79=" & CNULL(strCe(63)) & ",PA82=" & CNULL(strCe(64)) & "," & _
                     "PA109=" & CNULL(strCe(92)) & ",PA112=" & CNULL(strCe(93)) & ",PA115=" & CNULL(strCe(94)) & "," & _
                     "PA118=" & CNULL(strCe(95)) & ",PA121=" & CNULL(strCe(96)) & ",PA124=" & CNULL(strCe(97)) & "," & _
                     "PA127=" & CNULL(strCe(98)) & ",PA130=" & CNULL(strCe(99)) & ","
               End If
               For i = 63 To 64
                  If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
               Next
               For i = 92 To 99
                  If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
               Next
               strExc(3) = strExc(3) & "CE65='1',"
            End If
         End If
   
         If strExc(1) <> "" Then
            For i = 2 To 3
               If Right(strExc(i), 1) = "," Then strExc(i) = Left(strExc(i), Len(strExc(i)) - 1)
            Next
            strTxt(intStep) = "UPDATE CHANGEEVENT SET " & strExc(3) & " WHERE CE01='" & strReceiveNo & "'"
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
         End If
      End If
   End If
   
'Remove by Morgan 2009/10/16 改上面程式處理
'   If cp(10) = 讓與 And Me.txtCaseField(0).Text = 核准 Then
'      Select Case field(9)
'         Case 美國國家代號
'            strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25='" & DBDATE(txtCaseField(4)) & "' WHERE CP09='" & strReceiveNo & "'"
'         Case Else
'            strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25='" & strSrvDate(1) & "' WHERE CP09='" & strReceiveNo & "'"
'      End Select
'      cnnConnection.Execute strTxt(intStep)
'      intStep = intStep + 1
'   End If
   
   '2008/11/28 add by sonia 1912通知已轉他所閉卷更新
   If m_PA57 = "Y" Then
   '2008/11/28 end
      'Modified by Morgan 2019/1/17 lblCaseField(8) => dbdate(lblCaseField(8))
      strSql = "Update Patent Set PA57 = 'Y',PA17='N', PA58 =" & DBDATE(lblCaseField(8)) & ", PA59 = '99' Where  " & ChgPatent(field(1) & field(2) & field(3) & field(4))
      cnnConnection.Execute strSql
   '若要取消閉卷, 則更新基本檔閉卷及其相關欄位為NULL
   ElseIf m_blnCancelClosed = True Then
      strSql = "Update Patent Set PA57 = Null, PA58 = Null, PA59 = Null Where  " & ChgPatent(field(1) & field(2) & field(3) & field(4))
      cnnConnection.Execute strSql
   End If
   
   'Added by Morgan 2023/3/3
   '自請撤回413核准閉卷
   If m_Close413 = "Y" And field(57) = "" Then
      strSql = "Update Patent Set PA57='Y',PA59='09',PA58=" & strSrvDate(1) & " Where  " & ChgPatent(field(1) & field(2) & field(3) & field(4))
      cnnConnection.Execute strSql, intI
      
      strExc(2) = AutoNo("B", 6) 'B類總收文號
      strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp43,CP44,cp45,cp46,cp57,cp58,cp116) values " & _
         " ('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & _
         ",'" & strExc(2) & "','913','" & stCP12 & "','" & stCP13 & "','" & strUserNum & "','N','N'," & strSrvDate(1) & _
         ",'N','" & cp(9) & "','" & cp(44) & "','" & cp(45) & "','" & cp(46) & "'," & strSrvDate(1) & ",'09','" & cp(116) & "')"
      cnnConnection.Execute strSql, intI
   End If
   'end 2023/3/3
   
   'Add by Morgan 2005/1/3 若為申請案核准且為自動發證時公告日證書號回寫基本檔
   If AutoIssue() = True Then
      'Modified by Morgan 2012/4/26 +PA15
      strSql = "Update Patent Set PA14=" & IIf(txtPA14 = "", "PA14", Val(txtPA14) + 19110000) & IIf(txtPA15 = "", "", ",PA15='" & ChgSQL(txtPA15) & "'") & ",PA22=" & IIf(txtPA22 = "", "PA22", "'" & txtPA22 & "'") & " Where  " & ChgPatent(field(1) & field(2) & field(3) & field(4))
      cnnConnection.Execute strSql
      
      'Added by Morgan 2012/5/25
      '自動發證國家核准時下一程序新增1603 ,期限=核准日+6個月
      strExc(1) = CompDate(1, 6, txtCaseField(4))
      strExc(2) = PUB_GetWorkDay1(strExc(1), True)
      strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
         "NP07,NP08,NP09,NP10,NP22) select '" & strDataTemp(9) & "','" & cp(1) & "'" & _
         ",'" & cp(2) & "','" & cp(3) & "','" & cp(4) & "',1603," & strExc(2) & "," & strExc(1) & _
         ",'" & strUserNum & "',NP22 from dual,(select nvl(max(np22),0)+1 NP22 from nextprogress)"
      cnnConnection.Execute strSql, intI
      'end 2012/5/25
   End If
   
   'Add by Morgan 2005/5/20
   '更新結餘
   '2008/11/28 MODIFY BY SONIA 加 1006,最終核駁,1912通知已轉他所
   'Modify by Morgan 2010/3/10 +建議性處分書 1220
   'modify by sonia 2025/4/18 +變更401、讓與701
   If ((txtCaseField(0) = 核准 Or txtCaseField(0) = 核駁 Or txtCaseField(0) = "1006" Or txtCaseField(0) = "1220") And (cp(10) = 復審 Or cp(10) = 再發行 Or cp(10) = 變更 Or cp(10) = 讓與)) Or txtCaseField(0) = 專利權消滅 Or txtCaseField(0) = "1912" Then
      Pub_UpdateEndModCash field(1), field(2), field(3), field(4)
   End If
   
   'Add by Morgan 2006/3/27 美國收到建議性處分書時,若有未齊備的'告建議性處分'內部收文則上系統日並發Mail通知承辦人
   m_str222MailCP14 = Empty: m_str222MailCP09 = Empty
   If field(9) = "101" And txtCaseField(0) = "1220" Then
      strSql = "select cp09, cp14 from caseprogress,engineerprogress" & _
         " where cp01='" & cp(1) & "' and  cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'" & _
         " and substr(cp09,1,1)='B' and cp10='222' and cp27 is null and ep02(+)=cp09 and ep06 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
      If intI = 1 Then
         m_str222MailCP09 = "" & RsTemp.Fields(0)
         m_str222MailCP14 = "" & RsTemp.Fields(1)
         strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & m_str222MailCP09 & "'"
         cnnConnection.Execute strSql
         If PUB_IfSetCP48(m_str222MailCP09) Then 'Add by Morgan 2010/10/1 新規則承辦期限隔日凌晨算
            '2010/1/19 add by sonia 更新承辦期限
            strPromoteDate = Pub_GetHandleDay(cp(1), field(9), "222")
            If strPromoteDate <> "" Then
               strSql = "Update CaseProgress Set CP48=" & CNULL(strPromoteDate) & " Where CP09='" & m_str222MailCP09 & "' "
               cnnConnection.Execute strSql
            End If
            '2010/1/19 end
         End If 'Add by Morgan 2010/10/1
      End If
   End If
                
   '2009/4/20 ADD BY SONIA 美專母案准駁時同時更新接CIP,CA或分割或CPA(但限設計)案之期限
   'Modify by Morgan 2010/3/10 +建議性處分書 1220
   'Modify by Morgan 2010/5/26 +判斷案件的准駁才要
   'If field(9) = 美國國家代號 And (Me.txtCaseField(0).Text = 核准 Or Me.txtCaseField(0).Text = 核駁 Or Me.txtCaseField(0).Text = "1006" Or Me.txtCaseField(0).Text = "1220") Then
   If bolUpdatePA20 And txtCaseField(3) <> "" And field(9) = 美國國家代號 And (Me.txtCaseField(0).Text = 核准 Or Me.txtCaseField(0).Text = 核駁 Or Me.txtCaseField(0).Text = "1006" Or Me.txtCaseField(0).Text = "1220") Then
      '2012/1/16 MODIFY BY SONIA 慧汶說要取消設計限制,CFP-024595分割案暫不送件欲等母案CFP-022261答辯結果才決定
      'strSql = "SELECT CP06,CP07,CP09 FROM CASEPROGRESS WHERE cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03<>'0' " & _
            "AND CP10 IN ('113','122') AND CP57 IS NULL AND CP27 IS NULL UNION " & _
            "SELECT CP06,CP07,CP09 FROM CASEPROGRESS,PATENT P1 WHERE cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03<>'0' " & _
            "AND CP10='114' AND CP57 IS NULL AND CP27 IS NULL " & _
            "AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND '3'=PA08 UNION " & _
            "SELECT CP06,CP07,CP09 FROM CASEPROGRESS,PATENT,DivisionCase WHERE DC05='" & cp(1) & "' and DC06='" & cp(2) & "' and DC07='" & cp(3) & "' and DC08='" & cp(4) & "'" & _
            "AND DC01=CP01 AND DC02=CP02 AND DC03=CP03 AND DC04=CP04 AND CP10='307' AND CP57 IS NULL AND CP27 IS NULL"
      'Modifiedby Morgan 2024/6/19 分割改下面共用函直接更新,並有修正語法(後兩句的patent多餘且有缺連結語法)
      'strSql = "SELECT CP06,CP07,CP09 FROM CASEPROGRESS WHERE cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03<>'0' " & _
            "AND CP10 IN ('113','122') AND CP57 IS NULL AND CP27 IS NULL UNION " & _
            "SELECT CP06,CP07,CP09 FROM CASEPROGRESS,PATENT P1 WHERE cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03<>'0' " & _
            "AND CP10='114' AND CP57 IS NULL AND CP27 IS NULL " & _
            "AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 UNION " & _
            "SELECT CP06,CP07,CP09 FROM CASEPROGRESS,PATENT,DivisionCase WHERE DC05='" & cp(1) & "' and DC06='" & cp(2) & "' and DC07='" & cp(3) & "' and DC08='" & cp(4) & "'" & _
            "AND DC01=CP01 AND DC02=CP02 AND DC03=CP03 AND DC04=CP04 AND CP10='307' AND CP57 IS NULL AND CP27 IS NULL"
      strSql = "SELECT CP06,CP07,CP09 FROM CASEPROGRESS WHERE cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03<>'0' " & _
            "AND CP10 IN ('113','122') AND CP57 IS NULL AND CP27 IS NULL" & _
            " UNION SELECT CP06,CP07,CP09 FROM CASEPROGRESS P1 WHERE cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03<>'0' " & _
            "AND CP10='114' AND CP57 IS NULL AND CP27 IS NULL "
      'end 2024/6/14
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
      If intI = 1 Then
         '無期限直接更新
         If IsNull(RsTemp.Fields(0)) And IsNull(RsTemp.Fields(1)) Then
            strSql = "UPDATE CASEPROGRESS SET CP06=" & TransDate(txtCaseField(2), 2) & ", CP07=" & TransDate(txtCaseField(3), 2) & " WHERE CP09='" & RsTemp.Fields(2) & "'"
            cnnConnection.Execute strSql
         '期限不同則先詢問是否更新
         ElseIf RsTemp.Fields(0) <> TransDate(txtCaseField(2), 2) Or RsTemp.Fields(1) <> TransDate(txtCaseField(3), 2) Then
            If MsgBox("接續案或分割案之本所期限為" & TransDate(RsTemp.Fields(0), 1) & "，法定期限為" & TransDate(RsTemp.Fields(1), 1) & "，是否要更新？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
               strSql = "UPDATE CASEPROGRESS SET CP06=" & TransDate(txtCaseField(2), 2) & ", CP07=" & TransDate(txtCaseField(3), 2) & " WHERE CP09='" & RsTemp.Fields(2) & "'"
               cnnConnection.Execute strSql
            End If
         End If
      End If
   End If
   '2009/4/20 END
   
   'Added by Morgan 2025/7/28
   '美國答辯核准若有告建議性處分未發文時自動上111111 --郭
   If field(9) = 美國國家代號 And cp(10) = "107" And txtCaseField(0).Text = 核准 Then
      strSql = "update caseprogress set cp27=19221111 where cp43='" & cp(9) & "' and cp10='222' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql, intI
   End If
   'end 2025/7/28
   
   'Added by Morgan 2024/6/19
   If txtCaseField(1) <> "" Then
      If InStr(CFP分割案抓母案期限的性質, txtCaseField(1)) > 0 Then
         strSql = "SELECT CP09 FROM DivisionCase,CASEPROGRESS" & _
            " WHERE DC05='" & cp(1) & "' and DC06='" & cp(2) & "' and DC07='" & cp(3) & "' and DC08='" & cp(4) & "'" & _
            "AND CP01(+)=DC01 AND CP02(+)=DC02 AND CP03(+)=DC03 AND CP04(+)=DC04 AND CP10='307' AND CP57 IS NULL AND CP27 IS NULL"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            Do While Not RsTemp.EOF
               st307Msg = st307Msg & PUB_Update307Ref(RsTemp("cp09"))
               If st307Msg <> "" Then
                  If Right(st307Msg, 1) <> vbCrLf Then st307Msg = st307Msg & vbCrLf
               End If
               RsTemp.MoveNext
            Loop
         End If
      End If
   End If
   'end 2024/6/19
   
   'Add by Morgan 2009/9/16
   'CFP通知修正,申復時以收文日+6個月更新相關總收文號的催審期限(機關來函才要,代理人來函不必--郭)
   If txtCaseField(0) = "1201" Or txtCaseField(0) = "1202" Then
      PUB_UpdateChkResultDate CompDate(1, 6, strSrvDate(1)), cp, m_NewCP09, txtCaseField(0), cp(9)
   End If
   
   'Added by Lydia 2016/10/19 PCT案輸入國際初步審查報告(1216),將催審-實體審查上Y
   If field(9) = "056" And txtCaseField(0) = "1216" Then
      strSql = "update nextprogress set np06='Y' where (np01,np22) = (select np01,np22 from nextprogress,caseprogress " & _
                        "where np02='" & field(1) & "' and np03='" & field(2) & "' and np04='" & field(3) & "' and np05='" & field(4) & "' " & _
                        "and np07='411' and np06 is null and np01=cp09(+) and cp10='416')"
      cnnConnection.Execute strSql
   End If
   'end 2016/10/19
   
   'Added by Morgan 2020/8/17
   '輸入1607註冊登記，指定國註冊費224催審期限上Y
   If txtCaseField(0) = "1607" And cp(10) = "224" Then
      strSql = "update nextprogress set np06='Y' where np01='" & cp(9) & "' and np07='" & 催審 & "' and np06 is null"
      cnnConnection.Execute strSql, intI
   End If
   'end 2020/8/17
   
   'Add by Morgan 2010/5/12
   '台灣案加速審查通知
   If txtCaseField(0) = 核准 And Text7.Text = "1" And Text7.Text <> Me.Text7.Tag Then
      strExc(0) = "select na28,na29 from nation where na01='" & field(9) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '發明或新型有訂實審期限的國家
         If field(8) = "1" Or (field(8) = "2" And Not IsNull(RsTemp("na28")) And RsTemp("na29") > 0) Then
            '台灣案已收文通知實審日且相關收文號尚無結果(不必管是否曾收文加速審查)
            '承辦人相同
            'Modify by Morgan 2010/5/25 承辦人離職不用
            'Modified by Morgan 2022/7/21 排除台灣案已收到審查意見通知函者--陳玲玲
            'Modified by Morgan 2024/9/13 排除已閉卷案件--郭 Ex:P-128865
            strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo" & _
               " from casemap,patent,caseprogress a" & _
               " where cm01='" & cp(1) & "' and cm02='" & cp(2) & "'" & _
               " and cm03='" & cp(3) & "' and cm04='" & cp(4) & "'" & _
               " and pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08" & _
               " and (pa16 is null or pa16='2') and pa09='000' and pa08='1' and pa57||pa108 is null" & _
               " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='1204'" & _
               " and exists(select * from caseprogress b,staff where b.cp09=a.cp43 and b.cp24 is null" & _
               " and b.cp14='" & cp(14) & "' and st01(+)=b.cp14 and st04='1')" & _
               " and not exists(select * from caseprogress b where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='1202')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
               strExc(2) = strExc(1) & " 已核准，台灣發明案 " & RsTemp(0) & " 符合提出加速審查之條件.."
               strExc(3) = "台灣發明案 " & RsTemp(0) & " 仍在審查中，惟其相對應 " & strExc(1) & " 已核准" & _
                  "，故台灣發明案符合提出加速審查之條件，若欲辦理，請洽承辦工程師 '||st02||'" & _
                  "('||st01||') 研討內容及評估費用。"
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                  " select '" & strUserNum & "','" & strDataTemp(13) & "',to_char(sysdate,'yyyymmdd')" & _
                  ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','" & strExc(3) & "',st01" & _
                  " from staff where st01='" & cp(14) & "'"
               cnnConnection.Execute strSql, intI
            End If
         End If
      End If
   End If
   'end 2010/5/12
   
   'Add by Morgan 2010/5/26
   '歐盟設計核准若有集體案時也要一併核准並新增來函
   If txtCaseField(0) = 核准 And cp(10) = "103" And field(9) = "239" Then
      strExc(0) = "select cp03,cp09 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp04='" & cp(4) & "' and cp57 is null and cp27>0 and cp10='105'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         Do While Not .EOF
            strExc(1) = "cp01"
            strExc(2) = "cp01"
            For intI = 2 To TF_CP
               Select Case intI
                  Case 60, 65, 66, 67, 68, 69, 70, 61, 62, 63, 87, 88
                  Case Else
                     strExc(1) = strExc(1) & ",cp" & Format(intI, "00")
                     If intI = 3 Then
                        strExc(2) = strExc(2) & ",'" & .Fields("cp03") & "'"
                     ElseIf intI = 9 Then
                        strExc(2) = strExc(2) & ",'" & AutoNo("C", 6) & "'"
                     ElseIf intI = 43 Then
                        strExc(2) = strExc(2) & ",'" & .Fields("cp09") & "'"
                     Else
                        strExc(2) = strExc(2) & ",cp" & Format(intI, "00")
                     End If
               End Select
            Next
            strSql = "Insert into caseprogress(" & strExc(1) & ") select " & strExc(2) & " from caseprogress where cp09='" & strDataTemp(9) & "'"
            
            cnnConnection.Execute strSql, intI
            
            strSql = "update caseprogress set cp24='1',cp25=" & DBDATE(txtCaseField(4)) & " where cp09='" & .Fields("cp09") & "'"
            cnnConnection.Execute strSql, intI
            
            strExc(0) = ""
            If txtPA14 <> "" Then
               strExc(0) = strExc(0) & ",PA14=" & DBDATE(txtPA14)
            End If
            If txtPA22 <> "" Then
               strExc(0) = strExc(0) & ",PA22='" & Replace(txtPA22, "-0001", "-" & Format(Val(.Fields("cp03")) + 1, "0000")) & "'"
            End If
            
            strSql = "update patent set pa16='1',pa20=" & DBDATE(txtCaseField(4)) & strExc(0) & " where pa01='" & cp(1) & "'" & _
               " and pa02='" & cp(2) & "' and pa03='" & .Fields("cp03") & "' and pa04='" & cp(4) & "'"
            cnnConnection.Execute strSql, intI
            
            .MoveNext
         Loop
         
         End With
      End If
   End If
   'end 2010/5/26
   
   'Added by Lydia 2017/05/09 後案官方來函性質「視為未主張」1918
   If txtCaseField(0) = "1918" Then
       '目前主張國內優先權發文後，被主張的前案會閉卷，往後若輸入來函性質為視為未主張且閉卷原因為88被主張國內優先權的，請系統自動取消前案之閉卷。
       If cp(10) = "121" Then
          Set RsTemp = PUB_ReadPDStateNew(field, cp(10), True)
          If RsTemp.RecordCount <> 0 Then
             RsTemp.MoveFirst
             Do While Not RsTemp.EOF
                strExc(0) = "" & RsTemp.Fields("本所案號")
                Call ChgCaseNo(strExc(0), strExc)
                strSql = "UPDATE PATENT SET PA57=NULL,PA58=NULL,PA59=NULL WHERE PA01='" & strExc(1) & "' AND PA02='" & strExc(2) & "' AND PA03='" & strExc(3) & "' AND PA04='" & strExc(4) & "' AND PA57='Y' "
                cnnConnection.Execute strSql
                RsTemp.MoveNext
             Loop
          End If
       End If
       '自優先權資料處移至案件備註
       If strChoseBase <> "" Then
          arrData = Split(strChoseBase, ";")
          strExc(4) = ""
          For intI = 0 To UBound(arrData)
             If Trim(arrData(intI)) <> "" Then
                Call PUB_GetPD060507(Trim(arrData(intI)), strExc(1), strExc(2), strExc(3)) '區分優先權資料
                strSql = "DELETE FROM PRIDATE WHERE PD01='" & field(1) & "' AND PD02='" & field(2) & "' AND PD03='" & field(3) & "' AND PD04 ='" & field(4) & "' "
                If strExc(1) <> "" Then strSql = strSql & "AND PD06='" & strExc(1) & "' "
                If strExc(2) <> "" Then strSql = strSql & "AND PD05=" & TransDate(strExc(2), 2) & " "
                If strExc(3) <> "" Then strSql = strSql & "AND PD07='" & strExc(3) & "' "
                cnnConnection.Execute strSql
                strBasePD06 = strBasePD06 & IIf(Len(strBasePD06) > 0, "、", "") & strExc(1)
                '備註的部份請詳列視為未主張的優先權國家、日期及優先權號
                strExc(4) = strExc(4) & IIf(Len(strExc(4)) > 0, "、", "") & IIf(strExc(3) <> "", PUB_GetNationName(strExc(3)) & ", ", "") & IIf(strExc(2) <> "", strExc(2) & ", ", "") & IIf(strExc(1) <> "", strExc(1) & ", ", "")
                strExc(4) = IIf(Right(strExc(4), 2) = ", ", Mid(strExc(4), 1, Len(strExc(4)) - 2), strExc(4))
             End If
          Next
          strSql = "UPDATE PATENT SET PA91=PA91||'" & ChangeTStringToTDateString(strSrvDate(2)) & " 視為未主張的優先權資料:" & strExc(4) & ";' WHERE PA01='" & field(1) & "' AND PA02='" & field(2) & "' AND PA03='" & field(3) & "' AND PA04='" & field(4) & "'  "
          cnnConnection.Execute strSql
       End If
       '更新公開和實審期限
       strExc(5) = PUB_GetFirstPriDate(field)
       strExc(9) = ""
       
         '公開或實審期限的相關總收文號用申請程序的收文號
         strSql = "select cp09 from caseprogress WHERE CP01='" & field(1) & "' AND CP02='" & field(2) & "' AND CP03='" & field(3) & "' AND CP04='" & field(4) & "' and instr('" & NewCasePtyList & "',cp10)>0 and cp159=0 order by cp05 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strExc(9) = RsTemp(0)
         End If
       PUB_UpdCfpDate2 field(1), field(2), field(3), field(4), strExc(5), strExc(9)
   End If
   'end 2017/05/09
   
'Removed by Morgan 2020/3/10 取消,與下面的通知重複,期限不用更新,因為無法確認是否要提--禧佩
'   'Add by Morgan 2010/6/14
'   '澳洲,日本,韓國,英國,美國及EPC核駁或檢索報告來時若有馬來西亞案且有通知提供前案未發文時發Mail通知工程師
'   'Modified by Morgan 2020/3/9 +PCT,+核准,改判斷實審未發文(原判斷1205通知提供前案),+更新實審期限
'   If InStr("015,011,012,201,101,221,056", field(9)) > 0 And (txtCaseField(0) = "1209" Or (InStr("1001,1002,1006,1220", txtCaseField(0)) > 0 And bolUpdatePA20)) Then
'      strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo,cp14,cp09,cp06" & _
'         " from caserelation,caseprogress where cr01='" & cp(1) & "'" & _
'         " and cr02='" & cp(2) & "' and cr03='" & cp(3) & "' and cr04='" & cp(4) & "'" & _
'         " and cp01(+)=cr05 and cp02(+)=cr06 and cp03(+)=cr07 and cp04(+)=cr08" & _
'         " and cp10='416' and cp27 is null and exists(select * from patent where pa01(+)=cp01" & _
'         " and pa02=cp02 and pa03=cp03 and pa04=cp04 and pa09='018')"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         strExc(2) = RsTemp("CNo") & "(馬來西亞)，已有相關案之審查結果" & Replace(lblCaseField(0), " ", "") & "(" & lblNation & ")。"
'         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'            " VALUES ( '" & strUserNum & "','" & RsTemp("cp14") & "',to_char(sysdate,'yyyymmdd')" & _
'            ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','如旨')"
'         cnnConnection.Execute strSql, intI
'
'         'Added by Morgan 2020/3/9 更新實審所限=系統日+10工作天
'         strExc(1) = CompWorkDay(11, strSrvDate(1))
'         If strExc(1) < RsTemp("cp06") Then
'            strSql = "update caseprogress set cp06=" & strExc(1) & " where cp09='" & RsTemp("cp09") & "'"
'            cnnConnection.Execute strSql, intI
'         End If
'         'end 2020/3/9
'      End If
'   End If
'end 2020/3/10
   
   'Added by Lydia 2017/09/29 CFP案件於程序輸入核駁，核准及檢索報告等案件性質時，請一併檢查各案之相關案件有收文實審或提供前案(前述二項案件性質承辦人為工程師時)但尚未發文的案件請發訊息通知該案承辦工程師
   'Modified by Morgan 2017/10/20 +控制非案件的准駁不通知 Ex:CFP-029780
   'If InStr("1001,1002,1209", txtCaseField(0)) > 0 Then
   If txtCaseField(0) = "1209" Or (InStr("1001,1002,1006,1220", txtCaseField(0)) > 0 And bolUpdatePA20) Then
      'Modified by Lydia 2018/03/14 +未取消收文
      'Modified by Lydia 2019/02/21 相關案件之申請國家為EPC則不通知(ex.CFP-030476發文通知CFP-030466)
      'strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo,cp14 " & _
                  "from caserelation,caseprogress,staff " & _
                  "where cr01='" & cp(1) & "' and cr02='" & cp(2) & "' and cr03='" & cp(3) & "' and cr04='" & cp(4) & "' " & _
                  "and cp01(+)=cr05 and cp02(+)=cr06 and cp03(+)=cr07 and cp04(+)=cr08 and cp14=st01(+) " & _
                  "and cp10 in ('416','207') and cp158=0 and cp159=0 and st03='P11'"
      'Moddified by Lydia 2025/04/21 相關案件已閉卷就不用通知 by 玫音;ex.CFP-032033輸入最終核駁，通知相關案CFP-032031 => and pa57 is null
      strExc(0) = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo,cp14 " & _
                  "from caserelation,caseprogress,staff,patent " & _
                  "where cr01='" & cp(1) & "' and cr02='" & cp(2) & "' and cr03='" & cp(3) & "' and cr04='" & cp(4) & "' " & _
                  "and cp01(+)=cr05 and cp02(+)=cr06 and cp03(+)=cr07 and cp04(+)=cr08 and cp14=st01(+) " & _
                  "and cp10 in ('416','207') and cp158=0 and cp159=0 and st03='P11' " & _
                  "and pa01(+)=cp01 and pa02(+)=cp02  and pa03(+)=cp03 and pa04(+)=cp04 and pa09<>'221' and pa57 is null "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         strExc(4) = Format(ServerTime, "000000")   'Added by Lydia 2018/03/14 PK不可重複(秒)
         Do While Not RsTemp.EOF
            strExc(4) = Format(Val(strExc(4)) + 1, "000000") 'Added by Lydia 2018/03/14
            strExc(2) = RsTemp("CNo") & "相關案件" & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) <> "000", "-" & cp(3) & "-" & cp(4), "") & "已有審查報告，請確認實審或提供前案是否可發文。"
            'Modified by Lydia 2018/03/14
            'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " VALUES ( '" & strUserNum & "','" & RsTemp("cp14") & "',to_char(sysdate,'yyyymmdd')" & _
               ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','如旨')"
            'Modified by Morgan 2021/4/27 +CC給程序 --林禧佩
            strExc(1) = PUB_GetCFPHandler(RsTemp("CNo"))
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                        " VALUES ( '" & strUserNum & "','" & RsTemp("cp14") & "'," & strSrvDate(1) & ", " & Val(strExc(4)) & ",'" & strExc(2) & "','如旨','" & strExc(1) & "')"
            cnnConnection.Execute strSql, intI
            RsTemp.MoveNext
         Loop
      End If
   End If
   'end 2017/09/29
   
   'add by sonia 2017/12/26土耳其235發明案或EPC221案 輸"核准"時下一程序掛"商業使用"核准日+3年為法限,本所=法定-2月
   If txtCaseField(0) = 核准 And txtCaseField(1) = "601" Then
      'modify by sonia 2020/4/4 +cp(4)="00"即EPC子案輸核准不可更新CFP-029945-0-39
      'modify by sonia 2020/7/24 土耳其加新型案
      If (field(8) = "1" Or field(8) = "2") And field(9) = "235" And cp(4) = "00" Then
         strExc(1) = CompDate(0, 3, txtCaseField(4))   '法限
         strExc(2) = CompDate(1, -2, strExc(1))        '本所
         strExc(2) = PUB_GetWorkDay1(strExc(2), True)
         strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
            "NP07,NP08,NP09,NP10,NP22) select '" & strDataTemp(9) & "','" & cp(1) & "'" & _
            ",'" & cp(2) & "','" & cp(3) & "','" & cp(4) & "',930," & strExc(2) & "," & strExc(1) & _
            "," & CNULL(PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))) & ",NP22 from dual,(select nvl(max(np22),0)+1 NP22 from nextprogress)"
         cnnConnection.Execute strSql, intI
      End If
      If field(9) = "221" Then    '期限掛土耳其子案
         strExc(1) = CompDate(0, 3, txtCaseField(4))   '法限
         strExc(2) = CompDate(1, -2, strExc(1))        '本所
         strExc(2) = PUB_GetWorkDay1(strExc(2), True)
         strSql = "select pa01,pa02,pa03,pa04 from patent" & _
            " where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & cp(3) & "'" & _
            " and pa09='235' and pa57 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
               "NP07,NP08,NP09,NP10,NP22) select '" & strDataTemp(9) & "','" & RsTemp.Fields("pa01") & "'" & _
               ",'" & RsTemp.Fields("pa02") & "','" & RsTemp.Fields("pa03") & "','" & RsTemp.Fields("pa04") & "',930," & strExc(2) & "," & strExc(1) & _
               "," & CNULL(PUB_GetAKindSalesNo(cp(1), cp(2), cp(3), cp(4))) & ",NP22 from dual,(select nvl(max(np22),0)+1 NP22 from nextprogress)"
            cnnConnection.Execute strSql, intI
         End If
      End If
   End If
   'end 2017/12/26
   
   'Added by Morgan 2023/12/11
   '核准IDS管制
   If txtCaseField(0) = "1001" Then
      If m_USCaseNo <> "" Then
         PUB_SetUsIDS field(1), field(2), field(3), field(4), strDataTemp(9), txtCaseField(4), field(9), txtCaseField(0), cp(10), True
      End If
   Else
   'end 2023/12/11
      PUB_SetUsIDS field(1), field(2), field(3), field(4), strDataTemp(9), txtCaseField(4), field(9), txtCaseField(0), cp(10)   'Added by Morgan 2020/12/18 美國IDS期限管制
   End If
   
   'Added by Morgan 2020/8/7
   'EPC核准新增子案進度紀錄指定國註冊費代理人
   If cmdCountry.Enabled = True And field(9) = "221" And txtCaseField(0) = "1001" And strMoneyCountry <> "" And strFagentNo <> "" Then
      'Modified by Morgan 2020/8/26 +傳案件性質才不會清除承辦人
      If Not PUB_SaveCountry(1, intCaseKind, cp(1) & cp(2) & cp(3) & cp(4), strDataTemp(9), strMoneyCountry, strFagentNo, , txtCaseField(0)) Then
         GoTo ErrorHandler
      End If
   End If
   'end 2020/8/7
   
   
   'Added by Morgan 2024/12/6
   'EPC檢索報告在通知公開後,下一程序管制檢索報告公開1238(檢索報告收文日＋2個月)--郭
   If field(9) = "221" And txtCaseField(0) = "1209" Then
      If PUB_ChkCPExist(field, "1207") = True Then
         strExc(1) = CompDate(1, 2, strDataTemp(5))
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = " insert into nextprogress a (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
            " values('" & strDataTemp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','1238'" & _
            "," & strExc(2) & "," & strExc(1) & ",'" & strUserNum & "',GETNP22)"
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2024/12/6
   
   'Add by Sindy 2016/10/7
   If m_strIR01 <> "" Then
      'Modified by Morgan 2020/7/21 +收文號
      'Modified by Morgan 2020/8/18 有法限才要傳
      'Modify By Sindy 2022/6/16 F2外專不做2次確認 + And Left(Pub_StrUserSt03, 2) <> "F2"
      If txtCaseField(3) <> "" And Left(Pub_StrUserSt03, 2) <> "F2" Then
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm05010401_1", strDataTemp(9), m_bolReKeyInOK
      Else
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm05010401_1"
      End If
   End If
   '2016/10/7 END
   
   'Added by Morgan 2018/7/2 CFP電子化
   If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) Then
      m_strLD18 = strDataTemp(9)
      m_strCP10 = strDataTemp(10)
      
      'Modified by Morgan 2018/11/12 判發人改抓畫面輸入
      'strExc(1) = PUB_GetLetterJudgeNew("1", field(1), m_strCP10, field(9), cp(10))
      strExc(1) = Text37
      'end 2018/11/12
      '有定稿或有C類接洽單(工程師寫信)都算是有客戶函,(定稿存檔時會判斷若為回覆單時會更新為無通知函(LP10='')及回覆單(LP40='Y'))
      PUB_AddLetterProgress m_strLD18, 2 + Val(txtFiles), IIf(txtCaseField(5) <> "N" Or m_PrintCForm <> "N", True, False), strExc(1), IIf(Val(txtCaseField(3)) > 0, True, False), field(26), m_strCP10, field(75)
      '若要簡單報告且為工程師承辦(有印C類接洽單)時,新增D類收文轉公文(1998)與之相對應副件同C類來函
      If m_PrintCForm <> "N" And m_SimpleReportCust And txtCaseField(5) <> "N" Then
         m_str1998CP09 = AutoNo("D", 6)
         strSql = "INSERT INTO CASEPROGRESS(cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp27" & _
            ",cp32,cp43) SELECT cp01,cp02,cp03,cp04,cp05,cp06,cp07,'" & m_str1998CP09 & "','1998',cp12,cp13,'" & strUserNum & "'" & _
            ",'N','N'," & strSrvDate(1) & ",'N',cp09 FROM CASEPROGRESS WHERE CP09='" & m_strLD18 & "'"
         cnnConnection.Execute strSql, intI
         strExc(1) = PUB_GetLetterJudgeNew("1", field(1), "1998", field(9), m_strCP10)
         PUB_AddLetterProgress m_str1998CP09, 0, True, strExc(1), IIf(Val(txtCaseField(3)) > 0, True, False), field(26), m_strCP10, field(75)
      End If
      
      'Added by Morgan 2021/10/5 長庚醫院案件
      If m_CustX69365 = True Then
         PUB_SetX69365CaseOACP06 m_strLD18 '設定長庚醫院案件OA發文管制日(所限)
         
         'Removed by Morgan 2022/3/28 取消轉公文,改同其他3家直接報告,但本所期限改為 +14天-3個工作天 --黃教威
         'PUB_SetX69365Case1998CP06 m_str1998CP09 '設定長庚醫院案件轉公文管制日(所限)i
         'end 2022/3/28
      End If
      'end 2021/10/5
         
      m_bolAddLP = True
      
      'Added by Morgan 2018/10/2 EMail通知承辦人
      If m_PrintCForm <> "N" And Left(strDataTemp(12), 1) <> "F" Then
         If txtCaseField(12) <> "" Then
            'Modified by Morgan 2023/6/27 寶齡富錦工程師承辦的來函要CC給韻如
            Pub_COrderInform strDataTemp(9), , IIf(m_bolBPFCase, IIf(txtCaseField(12) = "A0029", "", "A0029"), "")
            bolSavPdf = True
            
         'Added by Morgan 2023/10/27
         '若原承辦人離職通知柏翰分案，副本程序人員。若柏翰請假時，改游經理，副本給柏翰。
         Else
            strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
            If CheckIsPersonRest("99050", strSrvDate(1), Right(Format(ServerTime, "00:00:00"), 5)) Then
               strExc(2) = strExc(1) & "(" & lblProperty & ")案原承辦人已離職，需重新分案。"
               strExc(3) = "本案原承辦人已離職，需重新分案。" & vbCrLf & "分案主管今天請假，請代為處理。"
               'Modified by Morgan 2025/2/21
               'strExc(4) = "73022"
               pub_PMan = Pub_GetSpecMan("專利處特定編號")
               strExc(4) = pub_PMan
               'end 2025/2/19
               strExc(5) = "99050;" & strUserNum
            Else
               strExc(2) = strExc(1) & "(" & lblProperty & ")案原承辦人已離職，請重新分案。"
               strExc(3) = "本案原承辦人已離職，請重新分案。"
               strExc(4) = "99050"
               strExc(5) = strUserNum
            End If
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
               " values('" & strUserNum & "','" & strExc(4) & "',to_char(sysdate,'yyyymmdd')" & _
               ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','" & strExc(3) & "','" & strExc(5) & "')"
            cnnConnection.Execute strSql, intI
            bolSavPdf = True
         'end 2023/10/27
         End If
      End If
      
   End If
   'end 2018/7/2
   
   'Add By Sindy 2025/5/26 CFP設計案收到OA(1002=核駁、1006=最終核駁、1201=通知修正)時
   '                       請發mail通知翔龍副理跟82018月嬌主任
   strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
   strExc(2) = strExc(1) & " 已" & lblProperty
   If cp(1) = "CFP" And m_PA08 = "3" And _
      (txtCaseField(0) = 核駁 Or txtCaseField(0) = "1006" Or txtCaseField(0) = 通知修正) Then
      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc13)" & _
         " values('" & strUserNum & "','" & Pub_GetSpecMan("設定繪圖人員通知對象") & ";82018',to_char(sysdate,'yyyymmdd')" & _
         ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','','" & m_NewCP09 & "')"
      cnnConnection.Execute strSql, intI
   End If
   '2025/5/26 END
   
   cnnConnection.CommitTrans
   
   If st307Msg <> "" Then MsgBox st307Msg, vbInformation 'Added by Morgan 2024/6/19
   
On Error GoTo ErrorHandler2

    '先預設不列印案件回覆單
    m_blnCustReturnSheet = False
    'End
   If SaveDatabase And bolNP22 Then
      'Removed by Morgan 2018/10/1 99/5/1 起就已不列印
      'For i = 1 To iNP22 - 1
      '   g_PrtForm001.PrintForm NP22(i), cp(1), cp(2), cp(3), cp(4)
      'Next
      'end 2018/10/1
      
      '設定要列印案件回覆單
      m_blnCustReturnSheet = True
      'End
   End If
   
   '列印C類接洽記錄單 92.1.28 ADD BY SONIA
   'Modify by Morgan 2004/5/17
   '1213 來函不印 C 類接洽記錄單
   'If m_PrintCForm <> "N" Then g_PrtForm001.PrintCFForm strDataTemp(9)
   If m_PrintCForm <> "N" Then
      'Add by Morgan 2008/5/6 核駁加印大/小個體,費用(P)
      strTemp = ""
      'Modify by Morgan 2008/5/30 +案件性質 --郭
      'If txtCaseField(0) = "1002" Then
      '2008/11/28 ADD BY SONIA 加1006
      If InStr("1002,1006,1201,1203,1205,1206,1209,1401,1307,1801,1802", txtCaseField(0)) > 0 Then
         'Modified by Morgan 2023/3/25
         'If Label41 <> "" Then
         '   strTemp = Label41
         'ElseIf InStr(1, field(91), "小個體", 1) > 0 Then
         '   strTemp = "小個體"
         'End If
         strTemp = Label41
         'end 2023/3/25
         
'2011/5/26 CANCEL BY SONIA 6/3下午4:00啟用
'         If txtCaseField(6) <> "" Then
'            strTemp = strTemp & IIf(strTemp <> "", ", ", "") & lblNextCaseProperty & "費用" & Format(txtCaseField(6), "#,###") & "(" & txtCaseField(8) & "P)"
'         End If
      End If
      
      'Modify by Morgan 2004/12/20 要判斷來函的性質
      'If cp(10) <> "1213" Then
      If txtCaseField(0) <> "1213" Then
         g_PrtForm001.PrintCFForm strDataTemp(9), strTemp, bolSavPdf
      End If
   End If
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
ErrorHandler2:
   SaveDatabase = False

End Function

Private Sub ReadAllData()
Dim rt As Boolean, i As Integer, varSaveCursor, strTemp As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim arrCF209() As String 'Added by Morgan 2023/3/25

   On Error GoTo ErrHnd
   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   
   ReDim cp(TF_CP) As String
   cp(9) = frm05010401_2.grdDataList.TextMatrix(frm05010401_2.grdDataList.row, 0)
   If PUB_ReadAllData(cp(), field(), intCaseKind, intPWhere) Then
   
      If cp(1) = 馬德里案 Then
         lblCaseField(0) = cp(1) + " - " + Left(cp(2), 5) + _
            IIf(Right(cp(2), 1) = "0", "", " - " + Right(cp(2), 1)) + _
            IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + _
            IIf(cp(4) = "00", "", " - " + cp(4))
      Else
         lblCaseField(0) = MergeString(cp(1), cp(2), cp(3), cp(4))
      End If
      Select Case intPCaseKind
                   Case 專利
                              lblCaseField(1) = field(11)
                              lblCaseField(2) = field(26)
                              lblCaseField(9) = field(9)
                   Case 商標
                              lblCaseField(1) = field(12)
                              lblCaseField(2) = field(23)
                              lblCaseField(9) = field(10)
                   Case Else
                              lblCaseField(1) = field(11)
                              lblCaseField(2) = field(8)
                              lblCaseField(9) = field(9)
      End Select
      m_PA01 = field(1)
      m_PA09 = field(9)
      lblCaseField(4) = cp(9)
      lblCaseField(6) = cp(10)
      lblCaseField(7) = cp(13)
      '機關文號
      txtCaseField(16) = cp(8)
      '記錄專利種類
      m_PA08 = "" & field(8)
      m_PA26 = "" & field(26)
      
      txtCaseField(12) = cp(14)
      CheckKeyIn 12
      lblCaseField(8) = frm05010401_1.txtCaseCode(3)
      lblCaseField(5) = TransDate(cp(5), 1)
      
      SetNameToCombo cboCaseName, field(5), field(6), field(7)
      'Modified by Morgan 2012/5/28 +307
      'Modified by Lydia 2016/08/26 +438 再考量試行計畫(AFCP2.0)
      If field(9) <> EPC指定國家 Or ((lblCaseField(6) < "101" Or lblCaseField(6) > "105") And lblCaseField(6) <> "107" And lblCaseField(6) <> "307" And lblCaseField(6) <> "438") Then
         cmdCountry.Enabled = False
      Else
         cmdCountry.Enabled = True
      End If
      If cmdCountry.Enabled Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.ReadCountry(intCaseKind, cp(), strCountry, , False) = False Then GoTo Err1
         If ClsPDReadCountry(intCaseKind, cp(), strCountry, , False) = False Then
            GoTo err1
         'Add by Morgan 2009/5/6
         '將國家代碼依照英文名稱排序以便於輸入領證費用
         ElseIf strCountry <> "" Then
            strExc(0) = "select na01 from nation where instr('" & strCountry & "',na01)>0 order by na04"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strCountry = RsTemp.GetString(adClipString, , , ",")
               '去掉最後一個逗號
               strCountry = Left(strCountry, Len(strCountry) - 1)
            End If
         End If
      Else
         strCountry = ""
      End If
      '顯示目前准駁
      Me.Text7.Text = "" & field(16)
      Me.Text7.Tag = Me.Text7.Text 'Add by Morgan 2007/3/22
      '顯示專用權是否存在
      Me.text8.Text = "" & field(17)
      Label41 = ""
      'Modify by Morgan 2006/9/21 加法國
      'Modified by Morgan 2023/3/25
      'If field(9) = "101" Or field(9) = "102" Or field(9) = "203" Then
      '      If InStr(1, field(91), "大個體", 1) > 0 Then
      '         Label41 = "大個體"
      '      'Added by Morgan 2013/10/18
      '      ElseIf InStr(1, field(91), "微個體", 1) > 0 Then
      '         Label41 = "微個體"
      '      'end 2013/10/18
      '      End If
      '   strExc(9) = "小個體"
      '   For intI = 1 To 5
      '      If field(25 + intI) <> "" Then
      '         If PUB_GetAD03(field(25 + intI), field(9)) = "N" Then
      '            strExc(9) = "大個體"
      '            Exit For
      '         End If
      '      End If
      '   Next
      '   strExc(8) = Label41
      '   If strExc(8) = "" Then strExc(8) = "小個體"
      '   If strExc(8) = "大個體" Or strExc(9) = "大個體" Then 'Added by Morgan 2013/10/18
      '      If strExc(8) <> strExc(9) Then
      '         MsgBox "本案客戶減免設定為【" & strExc(9) & "】與基本檔不同，請檢查資料是否有誤！"
      '      End If
      '   End If 'Added by Morgan 2013/10/18
      'Modified by Morgan 2024/12/9
      If InStr(CFP_ChkEntity, field(9)) > 0 Then
         ReDim arrCF209(2) As String
         arrCF209(0) = "大個體"
         arrCF209(1) = "小個體"
         arrCF209(2) = "微個體"
         PUB_SetEntityOpt field(1), field(9), field(8), arrCF209
         If strSrvDate(1) >= PA179啟用日 Then
            If Val(field(179)) > 0 Then
               Label41 = arrCF209(Val(field(179)) - 1)
            End If
         Else
            If InStr(1, field(91), "大個體", 1) > 0 Then
               Label41 = arrCF209(0)
            ElseIf InStr(1, field(91), "小個體", 1) > 0 Then
               Label41 = arrCF209(1)
            ElseIf InStr(1, field(91), "微個體", 1) > 0 Then
               Label41 = arrCF209(2)
            End If
         End If
         
         'Modified by Morgan 2024/12/9 個體別順序會因國家有所不同,且客戶設定目前只設定是否可減免,所以改回只檢查大小個體
         'strExc(9) = arrCF209(1)
         'For intI = 1 To 5
         '   If field(25 + intI) <> "" Then
         '      If PUB_GetAD03(field(25 + intI), field(9)) = "N" Then
         '         strExc(9) = arrCF209(0)
         '         Exit For
         '      End If
         '   End If
         'Next
         'strExc(8) = Label41
         'If strExc(8) = "" Then strExc(8) = arrCF209(1)
         'If strExc(8) = arrCF209(0) Or strExc(9) = arrCF209(0) Then
         strExc(9) = "小個體"
         For intI = 1 To 5
            If field(25 + intI) <> "" Then
               If PUB_GetAD03(field(25 + intI), field(9)) = "N" Then
                  strExc(9) = "大個體"
                  Exit For
               End If
            End If
         Next
         strExc(8) = Label41
         If strExc(8) = "" Then strExc(8) = "小個體"
         If strExc(8) = "大個體" Or strExc(9) = "大個體" Then
         'end 2024/12/9
            If strExc(8) <> strExc(9) Then
               MsgBox "本案客戶減免設定為【" & strExc(9) & "】與基本檔不同，請檢查資料是否有誤！"
            End If
         End If
      'end 2023/3/25
      End If
      'Added by Lydia 2020/11/19 因為第一頁空間不足,所以"美國讓渡登記號"採用視情況顯示
      'Modified by Lydia 2020/12/01 改成美國案就顯示
      'If field(9) = 美國國家代號 And (cp(10) = 讓與 Or cp(10) = 繼承 Or cp(10) = 授權 Or cp(10) = 變更) Then
      If field(9) = 美國國家代號 Then
         m_iNoStopIdx = 11 'Added by Morgan 2021/12/16
         lblAno.Visible = True: txtCaseField(11).Visible = True
      End If
      'end 2020/11/19
      'Added by Lydia 2020/12/01 CFP英國脫歐案管制：改成歐盟案或英國案就顯示
      If field(9) = "239" Or field(9) = "201" Then
         m_iNoStopIdx = 26 'Added by Morgan 2021/12/16
         lblEno.Visible = True: txtCaseField(26).Visible = True
      End If
      'end 2020/12/01
      
   'Added by Lydia 2015/04/10 申請人可在該國有多筆識別番號
'      '顯示國外ID號數
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
'      StrSQLa = "Select AFID03 FROM APPLICANTFOREIGNID WHERE AFID01='" & Left(m_PA26 & "00000000", 8) & "' AND AFID02='" & m_PA09 & "'"
'      rsA.CursorLocation = adUseClient
'      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'        Me.txtCaseField(26).Text = "" & rsA.Fields(0).Value
'        CmdAFID03.Enabled = True
'      End If
'
'      If rsA.State <> adStateClosed Then rsA.Close
'      Set rsA = Nothing
      
      CmdAFID03(0).Enabled = True

      For intI = 1 To 4
          If Len(field(26 + intI)) > 0 Then
             CmdAFID03(intI).Visible = True
          Else
             CmdAFID03(intI).Visible = False
          End If
      Next intI
      
      '是否閉卷
       If "" & field(57) = "Y" Then
           m_blnClosed = True
       Else
           m_blnClosed = False
       End If
       m_strCloseDate = "" & field(58)
       Label50 = ""
   Else
err1:
      bolLeave = True
      intLeaveKind = 1
      Unload Me
   End If
   Screen.MousePointer = varSaveCursor
   Exit Sub
ErrHnd:
   ErrorMsg
   Screen.MousePointer = varSaveCursor
End Sub

Private Sub Form_Activate()
   '控制只執行一次
   If m_bolActive = True Then Exit Sub
   m_bolActive = True
   
   blnOKtoShow = True
   ReadAllData
   'Add by Morgan 2010/10/1 新規則承辦期限隔日凌晨算
   If Not PUB_IfSetCP48() Then
      txtCaseField(13).Enabled = False
   End If
   'end 2010/10/1
   PUB_CheckCaseBillMemo field(1) & field(2) & field(3) & field(4) 'Add by Morgan 2008/6/9
   m_CustX07166 = False         '2012/11/26 add by sonia
   m_specialCust = False        '2013/8/12  add by sonia
   m_SimpleReportCust = False   '2015/8/21  add by sonia

   txtCaseField(0).SetFocus
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   bolLeave = False
   intLeaveKind = 1
   txtCaseField_Change 1
   If intPCaseKind = 專利 And intPWhere = 國外_CF Then
      Label3.Caption = "櫃台收文日："
   End If
   Me.SSTab1.Tab = 0
   m_intTab = Me.SSTab1.Tab
   bolDo = True
   
   'Add By Sindy 2017/12/28
   m_strIR01 = frm05010401_2.m_strIR01
   m_strIR02 = frm05010401_2.m_strIR02
   m_strIR03 = frm05010401_2.m_strIR03
   m_strIR04 = frm05010401_2.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/28 END
   
   lblProperty.BackColor = SSTab1.BackColor
   lblNextCaseProperty.BackColor = SSTab1.BackColor
   lblPromoter.BackColor = SSTab1.BackColor
   txtCaseField(24).BackColor = SSTab1.BackColor
   'Add by Amy 2014/09/17 承辦期限欄位隱藏
   Label19.Visible = False
   txtCaseField(13).Visible = False
   txtCaseField(13).Enabled = False
   'end 2014/09/17
   
   'Added by Lydia 2020/11/19
   lblAno.Visible = False: txtCaseField(11).Visible = False '因為第一頁空間不足,所以"美國讓渡登記號"採用視條件顯示
   lblEno.Visible = False: txtCaseField(26).Visible = False 'CFP英國脫歐案
   lblEno.Top = lblAno.Top
   txtCaseField(26).Top = txtCaseField(11).Top
   'end 2020/11/19
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bolLeave = False Then
      If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Cancel = 1
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2010/5/12
   If intLeaveKind = 1 Then
      frm05010401_2.Show
   Else
      'Add By Sindy 2016/10/13
      If Me.m_strIR01 = "" Then
      '2016/10/13 END
         Unload frm05010401_2
         If intLeaveKind = 2 Then
            Unload frm05010401_1
         Else
            frm05010401_1.Show
            frm05010401_1.Clear
         End If
      End If
   End If
   Set frm05010401_3 = Nothing
End Sub

Private Sub lblCaseField_Change(Index As Integer)
Dim strTemp As String, strCusTemp As String
   
   Select Case Index
      Case 2
         strCusTemp = lblCaseField(Index)
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GETCUSTOMER(strCusTemp, strTemp) Then
         If ClsPDGetCustomer(strCusTemp, strTemp) Then
            lblCaseField(Index) = strCusTemp
            lblAgent.Caption = strTemp
         End If
      Case 6
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(cp(1), lblCaseField(Index), strTemp) Then
         If ClsPDGetCaseProperty(cp(1), lblCaseField(Index), strTemp) Then
            lblCaseProperty = strTemp
         End If
      Case 7
         lblSales = GetStaffName(lblCaseField(Index), True)
      Case 9
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetNation(lblCaseField(Index), strTemp) Then
         If ClsPDGetNation(lblCaseField(Index), strTemp) Then
            lblNation.Caption = strTemp
         End If
   End Select
End Sub

Private Sub lblFee1_Click()
   If lblFee1.Tag = "Y" Then
      PUB_LabelActive lblFee1, lblFee1s, False
      'Modified by Lydia 2017/07/03 +系統別
      'If PUB_GetOldPrice(field(26), field(9), field(8), txtCaseField(1), RsTemp) = True Then
      If PUB_GetOldPrice(field(26), field(9), field(8), txtCaseField(1), RsTemp, , , , , field(1)) = True Then
         PUB_LabelActive lblFee1, lblFee1s
         Set frm880014.grdDataList.Recordset = RsTemp
         Set frm880014.fmParent = Me
         'Modified by Morgan 2017/7/6 要判斷案件性質
         If txtCaseField(1) = "107" Then
            frm880014.Tag = "1" 'Added by Lydia 2017/07/03 改變grid的欄位寬度
         Else
            frm880014.Tag = ""
         End If
         'end 2017/7/6
         frm880014.Show vbModal
      End If
   End If
End Sub

Private Sub lblFee1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseDown lblFee1, lblFee1s
End Sub

Private Sub lblFee1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_LabelMouseUp lblFee1, lblFee1s
End Sub

Private Sub Option4_Click(Index As Integer)
   If Index = 0 Then
      Text10.SetFocus
   Else
   
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If Me.SSTab1.Tab = 2 Then
      If m_PA09 <> 美國國家代號 Then
         Me.txtCaseField(27).Text = Empty
         Me.txtCaseField(28).Text = Empty
         MsgBox "若申請國家非美國, 或未輸入製圖費, 則無法切換至" & """修圖資料""" & "頁籤!!!", vbExclamation + vbOKOnly
         Me.SSTab1.Tab = m_intTab
      End If
   Else
      m_intTab = Me.SSTab1.Tab
   End If
End Sub

Private Sub Text10_Change()
   If Text10 <> "" Then
      Option4(0).Value = True
      Text11 = ""
      Text12 = ""
   End If
End Sub

Private Sub Text10_GotFocus()
    TextInverse Text10
    CloseIme
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub Text11_Change()
   If Text11 <> "" Then
      Option4(1).Value = True
      Text10 = ""
      Text12 = ""
   End If
End Sub

Private Sub Text11_GotFocus()
  TextInverse Text11
  CloseIme
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text12_Change()
   If Text12 <> "" Then
      Option4(2).Value = True
      Text10 = ""
      Text11 = ""
   End If
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
   CloseIme
End Sub

'Added by Morgan 2020/7/16
Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Or Text12 = "" Then Exit Sub
   If Val(DBDATE(Text12)) < Val(strSrvDate(1)) Then
      MsgBox "來函期限不可小於系統日 !", vbCritical
      Cancel = True
   Else
      '轉民國年
      txtCaseField(3) = TransDate(Text12, 1)
   End If
End Sub

'Added by Morgan 2018/11/12
Private Sub Text37_GotFocus()
   TextInverse Text37
End Sub
'Added by Morgan 2018/11/12
Private Sub Text37_Validate(Cancel As Boolean)
   Label50 = ""
   If Text37 <> "" Then
      If ClsPDGetStaff(Text37, strExc(1)) = True Then
         Label50 = strExc(1)
      Else
         Cancel = True
      End If
   End If
End Sub

Private Sub Text5_GotFocus(Index As Integer)
   TextInverse Text5(Index)
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
   If Index = 1 Then
      If Text5(1) <> "" Then
         If Val(Text5(0)) > Val(Text5(1)) Then
            MsgBox "繳費年度錯誤，請重新輸入 !", vbCritical
            Text5_GotFocus Index
            Cancel = True
         End If
      End If
   End If
End Sub

'Add by Morgan 2005/1/118 新增放棄專利權費,定稿用
Private Sub txtAbandonFee_GotFocus()
   TextInverse txtAbandonFee
End Sub

Private Sub txtAbandonFee_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
            Case 5
               KeyAscii = UpperCase(KeyAscii)
               'Added by Morgan 2018/10/4 核准有領證期限一定要出客戶函--玫音
               If KeyAscii <> Asc("N") And KeyAscii <> 8 Then
                  KeyAscii = 0
                  Beep
               ElseIf Asc("N") = KeyAscii Then
                  If txtCaseField(0) = 核准 And txtCaseField(1) = "601" Then
                     KeyAscii = 0
                     Beep
                  End If
               End If
               'end 2018/10/4
            'Modified by Lydia 2020/11/19 +英國脫歐案專利號數26
            Case 9, 12, 14, 26
               KeyAscii = UpperCase(KeyAscii)
            'Add By Cheng 2002/07/24
            Case 25, 29 '是否修改通知函內容 及 是否為美專證書勘誤
               KeyAscii = UpperCase(KeyAscii)
               If KeyAscii <> 89 And KeyAscii <> 8 Then
                  KeyAscii = 0
               End If
            'Add by Morgan 2008/5/6
            'Modify by Morgan 2008/11/17 +面詢費點數32,修正費點數34
            'Modify by Morgan 2010/5/5 +補虧損點數27,補未收費程序點數38
            'Modified by Morgan 2021/9/23 +IDS報價40, 41,RCE報價 42, 43
            Case 6, 8, 32, 34, 37, 38, 40, 41, 42, 43
               If KeyAscii <> 8 And Not (KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Then
                  KeyAscii = 0
               End If
   End Select
End Sub

Private Sub txtCaseField_Change(Index As Integer)
'Add By Cheng 2002/07/16
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   Select Case Index
             Case 0 '來函性質
               m_IDNGrant = False 'Added by Morgan 2021/11/4
               
               If Len(txtCaseField(Index)) <> 4 Then Exit Sub 'Added by Morgan 2019/5/29
               
               'Add by Morgan 2007/8/24
               Text5(0).Enabled = False
               Text5(1).Enabled = False
               'end 2007/8/24
                                                
'Remove by Morgan 2008/5/12 公開費已不再輸入

                        'Add By Cheng 2002/07/23  '92.9.30 加 "501" BY SONIA '2007/8/3 加"424"請求繼續審查 BY SONIA
                        'Modify by Morgan 2008/5/23 +122 CA申請
                        '設定目前准駁欄值
                        'Modified by Morgan 2012/3/8 取消 802,804
                        'Modify by Sonia  2013/11/15 加復審805(CFP-023821)
                        'If (cp(10) >= "101" And cp(10) <= "107" And cp(10) <> "106") Or cp(10) = "113" Or cp(10) = "114" Or cp(10) = "501" Or cp(10) = "424" Or (cp(10) >= "301" And cp(10) <= "307") Or cp(10) = "802" Or cp(10) = "804" Or cp(10) = "122" Then
                        'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
                        'Modified by Lydia 2016/08/28 +438 再考量試行計畫(AFCP2.0)
                        'Modified by Morgan 2020/12/18 改寫函數判斷以便共用及修改
                        'If (cp(10) >= "101" And cp(10) <= "107" And cp(10) <> "106") Or cp(10) = "113" Or cp(10) = "114" Or cp(10) = "122" Or cp(10) = "126" Or (cp(10) >= "301" And cp(10) <= "307") Or cp(10) = "424" Or cp(10) = "438" Or cp(10) = "501" Or cp(10) = "805" Then
                        If PUB_ChkIsRltPty(cp(1), cp(10), field(9)) = True Then
                        'end 2020/12/18
                           If Me.txtCaseField(Index).Text = 核准 Then
                              Me.Text7.Text = "1"
                              'Add by Morgan 2007/8/24 准後繳國家領證同時繳年費
                              strExc(0) = "select 1 from nation where na01='" & field(9) & "' and decode('" & field(8) & "','1',na56,'2',na57,na58)='Y'"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 Text5(0).Enabled = True
                                 Text5(1).Enabled = True
                                 If field(9) = "017" And (field(8) = "1" Or field(8) = "2") Then m_IDNGrant = True 'Added by Morgan 2021/11/4 印尼發明/新型核准
                              End If
                              'end 2007/8/24
                           'Modify by Morgan 2006/3/23 加最終核駁1006
                           'Modify by Morgan 2010/3/10 +建議性處分書 1220
                           ElseIf InStr("1002,1006,1220", txtCaseField(Index)) > 0 Then
                              Me.Text7.Text = "2"
                           Else
                              Me.Text7.Text = "" & field(16)
                           End If
                        Else
                           Me.Text7.Text = "" & field(16)
                        End If
                        
'Removed by Morgan 2012/3/8
'                        '設定專用權是否存在欄值
'                        If cp(10) = "804" Then
'                           If Me.txtCaseField(index).Text = 核准 Then
'                              Me.Text8.Text = "Y"
'                           '2008/11/28 MODIFY BY SONIA 加 1006
'                           'ElseIf Me.txtCaseField(Index).Text = 核駁 Then
'                           'Modify by Morgan 2010/3/10 +建議性處分書 1220
'                           ElseIf InStr("1002,1006,1220", txtCaseField(index)) > 0 Then
'                              Me.Text8.Text = "N"
'                           Else
'                              Me.Text8.Text = "" & field(17)
'                           End If
'                        Else
'                           Me.Text8.Text = "" & field(17)
'                        End If
                        
                        'Add by Morgan 2005/5/20
                        'Modify by Morgan 2006/4/7 加EPC,印尼,韓國,馬來西亞
                        m_bolIsNP107N = False
                        '2008/11/28 MODIFY BY SONIA 加1006
                        'Modify by Morgan 2010/3/10 +建議性處分書1220
                        If InStr("1002,1006,1220", txtCaseField(Index)) > 0 Then
'2007/7/27 modify by sonia 因加越南CFP17724與慧汶討論後不限制國家及專利種類
'                           Select Case field(9)
'                              '全部
'                              Case "012", "017" '韓國,印尼
'                                 m_bolIsNP107N = IsNP107N
'                              '發明
'                              Case "221", "231" 'EPC,德國
'                                 If field(8) = "1" Then
'                                    m_bolIsNP107N = IsNP107N
'                                 End If
'                              '發明,設計
'                              Case "011" '日本
'                                 If field(8) <> "2" Then
'                                    m_bolIsNP107N = IsNP107N
'                                 End If
'                              '發明,新型
'                              Case "018" '馬來西亞
'                                 If field(8) <> "3" Then
'                                    m_bolIsNP107N = IsNP107N
'                                 End If
'                           End Select
                           m_bolIsNP107N = IsNP107N
'2007/7/27 END
                        End If
                        '2008/11/28 MODIFY BY SONIA 改核駁都要輸約定期限
                        'If m_bolIsNP107N = True Then
                        '2008/12/19 MODIFY BY SONIA 加檢索報告1209及通知提供前案1205
                        'Modify by Morgan 2010/3/10 +建議性處分書1220
                        '2012/11/26 modify by sonia +1206通知要求選取(順德定稿用)
                        'Removed by Morgan 2021/9/2 約定期限要寫在NP23,改以下一程序控制
                        'If InStr("1002,1006,1205,1206,1209,1220", txtCaseField(Index)) > 0 Then
                        ''2008/11/28 END
                        '   txtCaseField(30).Visible = True
                        '   lblCaseField(30).Visible = True
                        'Else
                        '   txtCaseField(30).Visible = False
                        '   lblCaseField(30).Visible = False
                        'End If
                        'end 2021/9/2
                        '2005/5/20 end
                        
                        'Added by Morgan 2021/6/1
                        txtCaseField(6) = ""
                        txtCaseField(8) = ""
                        'end 2021/6/1
                        
                        'Modify by Morgan 2008/5/30
                        'If txtCaseField(0) = "1002" Then
                        '2008/11/28 MODIFY BY SONIA 加1006
                        'Modify by Morgan 2012/7/20 +建議性處分書1220
                        'Modified by Morgan 2019/3/13 +依職權電話通知修正1225
                        If InStr("1002,1006,1201,1203,1205,1206,1209,1220,1401,1307,1801,1802,1225", txtCaseField(0)) > 0 Then
                           lblFee1 = "費用："
                           lblFee2.Visible = False
                           Text5(0).Visible = False
                           Text5(1).Visible = False
                           Label12(0).Visible = False
                           txtCaseField(7).Visible = False
                           'Add by Morgan 2008/11/13 +面詢費,修正費
                           Label12(1).Visible = False
                           Label12(2).Visible = False
                           Label12(3).Visible = False
                           Label12(4).Visible = False
                           txtCaseField(31).Visible = False
                           txtCaseField(32).Visible = False
                           txtCaseField(33).Visible = False
                           txtCaseField(34).Visible = False
                           'end 2008/11/13
                           
                           'add by Morgan 2010/4/8
                           txtCaseField(35).Text = ""
                           txtCaseField(35).Enabled = False
                           txtCaseField(35).BackColor = SSTab1.BackColor
                           txtCaseField(36).Text = ""
                           txtCaseField(36).Enabled = False
                           txtCaseField(36).BackColor = SSTab1.BackColor
                           'end 2010/4/8
                           
                        'Added by Morgan 2021/6/1 +通知繳公開費1236
                        ElseIf txtCaseField(0) = "1236" Then
                           lblFee1 = "公開費："
                           lblFee2.Visible = False
                           'Modified by Morgan 2024/7/12 改報價--禧佩
                           'txtCaseField(6) = "10000"
                           txtCaseField(6) = "16000"
                           'end 2024/7/12
                           txtCaseField(8) = "3"
                        'end 2021/6/1
                           
                        Else
                           lblFee1 = "領證費："
                           'Added by Morgan 2021/11/4
                           '印尼發明/新型核准:核准日6個月內須繳交自申請日起算累計至核准日次年之年費
                           If m_IDNGrant Then
                              lblFee1 = "年費："
                              Text5(0) = "1"
                              Text5(0).Enabled = False
                              'Text5(1).Enabled = False 'Removed by Morgan 2022/9/28
                              SetINDYear
                           End If
                           'end 2021/11/4
                           
                           lblFee2.Visible = True
                           Text5(0).Visible = True
                           Text5(1).Visible = True
                           Label12(0).Visible = True
                           txtCaseField(7).Visible = True
                           'Add by Morgan 2008/11/13 +面詢費,修正費
                           Label12(1).Visible = True
                           Label12(2).Visible = True
                           Label12(3).Visible = True
                           Label12(4).Visible = True
                           txtCaseField(31).Visible = True
                           txtCaseField(32).Visible = True
                           txtCaseField(33).Visible = True
                           txtCaseField(34).Visible = True
                           'end 2008/11/13
                           
                           'add by Morgan 2010/4/8
                           txtCaseField(35).Enabled = True
                           txtCaseField(35).BackColor = vbWhite
                           txtCaseField(36).Enabled = True
                           txtCaseField(36).BackColor = vbWhite
                           'end 2010/4/8
                        End If
                        '2008/11/28 ADD BY SONIA
                        'If txtCaseField(Index).Text = "1905" Then txtCaseField(25).Text = "Y" 'Removed by Morgan 2023/6/6 要修改的定稿內容改用問的，不必再預設修改--玫音
                        '2008/11/28 END
                        '2013/8/12 add by sonia
                        m_specialCust = PUB_CheckspecialCust(cp(9), txtCaseField(0).Text)
                        If m_specialCust Then
                           txtCaseField(12) = strUserNum
                           CheckKeyIn 12
                        End If
                        '2013/8/12 END
                        
                        'Added by Morgan 2018/11/12
                        Text37 = "": Label50 = "": Text37.Enabled = False
                        m_bolJudgerAlert = False
                        If Len(txtCaseField(0)) = 4 Then
                           Text37 = PUB_GetLetterJudgeNew("1", field(1), txtCaseField(0), field(9), cp(10))
                           If Text37 <> "" Then Text37_Validate False
                           'Removed by Morgan 2019/6/24 取消,因報價定稿一般都非當日列印,改CFP程序可自行點選王副總案件判發--郭
                           ''若判發人王副總71011請假且郭雅娟79075也請假時提醒程序要將判發人改為輸入程序人員的職代
                           'If Text37 = "71011" Then
                           '    If CheckIsPersonRest("71011", strSrvDate(1), Format(ServerTime \ 100, "00:00")) = True Then
                           '         If CheckIsPersonRest("79075", strSrvDate(1), Format(ServerTime \ 100, "00:00")) = True Then
                           '            m_bolJudgerAlert = True
                           '            Text37.Enabled = True
                           '         End If
                           '    End If
                           'End If
                           'end 2019/6/24
                        End If
                        'end 2018/11/12
                        
                        'Added by Lydia 2020/11/19 CFP英國脫歐案管制：若有英國再註冊來函時，除非英國新案已收文否則再註冊仍輸在歐盟案
                        'Remove by Lydia 2020/12/01 改成歐盟案或英國案就顯示
                        'lblEno.Visible = False: txtCaseField(26).Visible = False
                        'If field(1) = "CFP" And field(9) = "239" And txtCaseField(Index) = "1608" Then
                        '    lblEno.Visible = True: txtCaseField(26).Visible = True
                        'End If
                        ''end 2020/11/19
                        'end 2020/12/01
             Case 1
                        lblNextCaseProperty = ""
                        'Add by Morgan 2008/5/9
                        SetFee
                        
                        'Added by Morgan 2021/9/2 以下一程序性質控制是否要輸約定期限 --郭
                        '107 答辯
                        '208 選取
                        '218 回覆檢索報告
                        '421 申請檢索報告
                        '424 請求繼續審查
                        '438再考量試行計畫(AFCP2.0)
                        If Len(txtCaseField(Index)) = 3 And InStr(CFPAppDatePtyList, txtCaseField(Index)) > 0 Then
                           lblNextCaseProperty = GetCaseTypeName(m_PA01, txtCaseField(1)) 'Added by Morgan 2024/12/5
                           'Modified by Morgan 2024/12/6 改控制enable，這樣駐點才不會亂跳
                           'm_iNoStopIdx = 30 'Added by Morgan 2021/12/9
                           'txtCaseField(30).Visible = True
                           'lblCaseField(30).Visible = True
                           txtCaseField(30).Enabled = True
                           'end 2024/12/6
                        Else
                           txtCaseField(30) = ""
                           'Modified by Morgan 2024/12/6 改控制enable，這樣駐點才不會亂跳
                           'txtCaseField(30).Visible = False
                           'lblCaseField(30).Visible = False
                           txtCaseField(30).Enabled = False
                           'end 2024/12/6
                        End If
                        'end 2021/9/2
             
             Case 4
               SetINDYear 'Added by Morgan 2021/11/4
               
             Case 29
                        If txtCaseField(29) = "Y" Then
                           txtCaseField(12) = strUserNum
                           txtCaseField(14) = "N"
                           CheckKeyIn 12
                        End If
                        
             'Add by Morgan 2010/3/24
             Case 6, 8
               txtCaseField(24) = Val(txtCaseField(6)) - 1000 * (Val(txtCaseField(8)) - Val(txtCaseField(38))) - Val(txtCaseField(35)) - Val(txtCaseField(36))
             Case 35, 36
               txtCaseField(6) = Val(txtCaseField(24)) + 1000 * Val(txtCaseField(8)) + Val(txtCaseField(35)) + Val(txtCaseField(36))
             Case 38 'Add by Morgan 2010/10/21
               txtCaseField(8) = txtCaseField(8) - Val(txtCaseField(38).Tag) + Val(txtCaseField(38))
               txtCaseField(38).Tag = txtCaseField(38)
             Case Else
   End Select
   
End Sub
'Add by Morgan 2008/5/9
'設定費用及點數
Private Sub SetFee()
Dim strMsg As String
   
   lblFee1.Tag = ""
   lblFee1.BackColor = &H8000000F
   lblFee1s.Visible = False
   If txtCaseField(1) = "601" Then
            
      'Add by Morgan 2010/2/5 美國案檢查是否有收文提早公開
      If field(9) = "101" And field(8) = "1" Then
         If PUB_ChkCPExist(field, "417") Then
            strMsg = "※本案有收文【提早公開】，領證費報價不應含公開費！"
         End If
      End If
      'end 2010/2/5
      
      'Added by Morgan 2025/3/12
      '俄羅斯發明及新型輸入核准時，領證之其他費用欄位，請預設含第1至第5年年費並帶入通知客戶的定稿。--禧佩
      If field(9) = "023" And (field(8) = "1" Or field(8) = "2") Then
         strExc(0) = PUB_Getnexttimes(cp(1), cp(2), cp(3), cp(4), strExc(1)) '抓起繳年度,目前設定發文為第3年
         Text5(0) = strExc(1)
         Text5(1) = "5"
      End If
      'end 2025/3/12
      
      strExc(0) = "select yf06,yf07,YF08 from patentyearfee where yf01='" & field(9) & "'" & _
         " and yf02='" & field(8) & "' and yf04='" & txtCaseField(1) & "' and yf05='1'"
      'Added by Morgan 2023/3/25
      If strSrvDate(1) >= PA179啟用日 Then
         If field(179) = "1" Then '大個體
            strExc(0) = strExc(0) & " and yf03='Y00000002'"
         ElseIf field(179) = "3" Then '微個體
            strExc(0) = strExc(0) & " and yf03='Y00000003'"
         Else
            strExc(0) = strExc(0) & " and yf03='Y00000000'"
         End If
      Else
      'end 2023/3/25
      
         If InStr(field(91), "大個體") > 0 Then
            strExc(0) = strExc(0) & " and yf03='Y00000002'"
         'Added by Morgan 2013/10/18
         ElseIf InStr(field(91), "微個體") > 0 Then
            strExc(0) = strExc(0) & " and yf03='Y00000003'"
         'end 2013/10/18
         Else
            strExc(0) = strExc(0) & " and yf03='Y00000000'"
         End If
         
      End If 'Added by Morgan 2023/3/25
      
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         
         '領證費
         If "" & RsTemp(0) <> "" Or "" & RsTemp(1) <> "" Then 'Added by Morgan 2025/3/12 沒設定時空白，否則存檔檢查會沒用
            txtCaseField(6) = Val(Format("" & RsTemp(0))) + Val(Format("" & RsTemp(1)))
         End If
         '點數
         If "" & RsTemp(0) <> "" Then 'Added by Morgan 2025/3/12 沒設定時空白，否則存檔檢查會沒用
            txtCaseField(8) = Val(Format("" & RsTemp(0))) / 1000
         End If
         If Not IsNull(RsTemp(2)) Then
            strMsg = RsTemp(2) & vbCrLf & vbCrLf & strMsg
         End If
      End If
      
      If strMsg <> "" Then
         MsgBox strMsg, , "報價提醒！"
      End If
      
      'Add by Morgan 2008/5/16
      'Modified by Morgan 2019/3/29 +系統別
      If PUB_GetOldPrice(field(26), field(9), field(8), txtCaseField(1), , , , , , field(1)) = True Then
         lblFee1.Tag = "Y"
         lblFee1.BackColor = &HC0FFC0
         lblFee1s.Visible = True
      End If
      
   'Added by Lydia 2017/07/03 抓核駁費用
   ElseIf InStr("1002,1006", txtCaseField(0)) > 0 And txtCaseField(1) = "107" Then
         strExc(0) = PUB_GetYF0607(field(9), field(8), "Y00000000", txtCaseField(1), "1", "1", "2", strExc(1), strExc(2))
         If Val(strExc(0)) > 0 Then
            '領證費
            txtCaseField(6) = Val(strExc(1)) + Val(strExc(2))
            '點數
            txtCaseField(8) = Val(strExc(1)) / 1000
         End If
         If PUB_GetOldPrice(field(26), field(9), field(8), txtCaseField(1), , , , , , field(1)) = True Then
            lblFee1.Tag = "Y"
            lblFee1.BackColor = &HC0FFC0
            lblFee1s.Visible = True
         End If
   'end 2017/07/03
   
   End If

End Sub

Private Sub txtCaseField_LostFocus(Index As Integer)
'Add By Cheng 2002/11/22
Dim strTemp As String
Dim strTemp1 As String

   Select Case Index
      '91.11.27 ADD BY SONIA
      Case 0 '來函性質
         'Add by Morgan 2011/8/8
         txtCaseField(1).Locked = False
         txtCaseField(3).Locked = False
         txtCaseField(2).Locked = False
         txtCaseField(30).Locked = False
         txtCaseField(6).Locked = False
         txtCaseField(8).Locked = False
         'end 2011/8/8
         
         'If m_PA09 = 美國國家代號 Then
         '   Select Case txtCaseField(0)
         '   Case 核駁
         '      txtCaseField(1) = 答辯
         '   Case 通知要求選取
         '      txtCaseField(1) = 選取
         '   End Select
         '   txtCaseField_Validate (1), False
         'End If
      '91.11.27 END
         '92.1.18 ADD BY SONIA
         If lblCaseField(9) = 美國國家代號 Then
            '92.9.15 MODIFY BY SONIA
            'If lblCaseField(6) = 發明申請 Or lblCaseField(6) = 新型申請 Or lblCaseField(6) = 設計申請 Or lblCaseField(6) = 追加申請 Or lblCaseField(6) = 聯合申請 Or lblCaseField(6) = 答辯 Then
            '2005/5/19 MODIFY BY SONIA 加分割
            '2007/8/30 MODIFY BY SONIA 加請求繼續審查
            'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
            'Modified by Lydia 2016/08/27 +438 再考量試行計畫(AFCP2.0)
            If lblCaseField(6) = 發明申請 Or lblCaseField(6) = 新型申請 Or lblCaseField(6) = 設計申請 Or lblCaseField(6) = 追加申請 Or lblCaseField(6) = 聯合申請 Or lblCaseField(6) = 答辯 Or lblCaseField(6) = CIP申請 Or lblCaseField(6) = CPA申請 Or lblCaseField(6) = 再發行 Or lblCaseField(6) = 分割 Or lblCaseField(6) = "424" Or lblCaseField(6) = "126" Or lblCaseField(6) = "438" Then
            '92.9.15 END
              'Modify By Cheng 2003/06/06
      '         txtCaseField(10) = 15000
               'txtCaseField(10) = 17000'Removed by Morgan 2012/3/30 定稿原則不帶故不必再預設--郭
            'Modify by Morgan 2007/8/9 加繼承
            'ElseIf lblCaseField(6) = 讓與 Then
            'Modify by Morgan 2007/10/3 加授權
            'Modify by Morgan 2008/1/14 加變更
            ElseIf (lblCaseField(6) = 讓與 Or lblCaseField(6) = 繼承 Or lblCaseField(6) = 授權 Or lblCaseField(6) = 變更) Then
               txtCaseField(11) = "第號/第格"
            End If
         End If
         '92.1.18 END
         
         '92.10.21 add by sonia
         'Modify by Morgan 2006/3/23 加最終核駁1006
         'Modify by Morgan 2010/3/10 +建議性處分書 1220
         If txtCaseField(14) = "" Then  '2010/7/16 SONIA by sonia 加入txtCaseField(14) = ""條件,否則前已做的控制又被改掉
            If txtCaseField(0) = 核准 Or txtCaseField(0) = 核駁 Or txtCaseField(0) = "1006" Or txtCaseField(0) = 通知補文件 Or txtCaseField(0) = "1005" _
            Or txtCaseField(0) = 通知申請案號 Or txtCaseField(0) = 通知修正 Or txtCaseField(0) = 通知補充說明 Or txtCaseField(0) = 通知提供前案 _
            Or txtCaseField(0) = 通知要求選取 Or txtCaseField(0) = 通知公開 Or txtCaseField(0) = 通知公告 Or txtCaseField(0) = 檢索報告 _
            Or txtCaseField(0) = 通知證書號數 Or txtCaseField(0) = 專利證書 Or txtCaseField(0) = 其他來函 Or txtCaseField(0) = "1908" _
            Or txtCaseField(0) = "1220" _
            Then
               txtCaseField(14) = "N"
            Else
               txtCaseField(14) = ""
            End If
         End If  '2010/7/16 ADD BY SONIA
         '92.10.21 end
         
         'Add by Morgan 2004/9/7 從 Validate 移來以免存檔前檢查又帶出下出下一程序
         '93.10.4 MODIFY BY SONIA 加入分割
         'If (lblCaseField(6) = 發明申請 Or lblCaseField(6) = 新型申請 Or lblCaseField(6) = 設計申請 Or lblCaseField(6) = 追加申請 Or lblCaseField(6) = 聯合申請 Or lblCaseField(6) = 答辯 Or lblCaseField(6) = CIP申請 Or lblCaseField(6) = CPA申請 Or lblCaseField(6) = 再發行) Then
         '2005/4/20 MODIFY BY SONIA 加入訴願
         'If (lblCaseField(6) = 發明申請 Or lblCaseField(6) = 新型申請 Or lblCaseField(6) = 設計申請 Or lblCaseField(6) = 追加申請 Or lblCaseField(6) = 聯合申請 Or lblCaseField(6) = 答辯 Or lblCaseField(6) = CIP申請 Or lblCaseField(6) = CPA申請 Or lblCaseField(6) = 再發行 Or lblCaseField(6) = 分割) Then
         'Modify by Morgan 2006/6/5 加請求繼續審查 424
         'Modify by Morgan 2007/9/21 加改請(3字頭的都要)
         'Modify by Morgan 2008/5/23 +122 CA申請
         'Modify by Sonia  2013/11/15 加復審805(CFP-023821)
         'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
         'Modified by Lydia 2016/08/27 +438 再考量試行計畫(AFCP2.0)
         If (lblCaseField(6) = 發明申請 Or lblCaseField(6) = 新型申請 Or lblCaseField(6) = 設計申請 Or lblCaseField(6) = 追加申請 Or lblCaseField(6) = 聯合申請 Or lblCaseField(6) = 答辯 Or lblCaseField(6) = CIP申請 Or lblCaseField(6) = CPA申請 Or lblCaseField(6) = 再發行 Or lblCaseField(6) = 分割 Or lblCaseField(6) = 訴願 Or lblCaseField(6) = "424" Or lblCaseField(6) = "805" Or lblCaseField(6) = "122" Or Left(lblCaseField(6), 1) = "3" Or lblCaseField(6) = "126" Or lblCaseField(6) = "438") Then
         '2005/4/20END
         '93.10.4 END
            'Modify by Morgan 2011/8/8
            'If IsEmptyText(txtCaseField(1)) Then
            '   'Modify by Morgan 2005/1/18 判斷核准且非自動發證國家才要帶
            '   'txtCaseField(1) = GetNextProgress(m_PA01, m_PA09, txtCaseField(0))
            '   If Not (txtCaseField(Index) = "1001" And PUB_AutoIssue(m_PA09, m_PA08) = True) Then
            '自動發證國家加控制不可輸入下一程序及期限,也不可輸入領證費及點數--秀玲 8/4 Mail(電話確認)
            'Modified by Morgan 2012/4/24
            'If (txtCaseField(Index) = "1001" And PUB_AutoIssue(m_PA09, m_PA08, field(10)) = True) Then
            If (txtCaseField(Index) = "1001" And PUB_AutoIssue(m_PA09, m_PA08, field(10), field) = True) Then
               txtCaseField(1).Text = ""
               txtCaseField(1).Locked = True
               txtCaseField(3).Text = ""
               txtCaseField(3).Locked = True
               txtCaseField(2).Text = ""
               txtCaseField(2).Locked = True
               txtCaseField(30).Text = ""
               txtCaseField(30).Locked = True
               
               If Not m_IDNGrant Then  'Added by Morgan 2021/11/4 印尼發明/新型核准除外(要繳年費)
                  txtCaseField(6).Text = ""
                  txtCaseField(6).Locked = True
                  txtCaseField(8).Text = ""
                  txtCaseField(8).Locked = True
               End If
               
            ElseIf IsEmptyText(txtCaseField(1)) Then
            'end 2011/8/8
            
                  txtCaseField(1) = GetNextProgress(m_PA01, m_PA09, txtCaseField(0))
                  
            '   End If 'Remove by Morgan 2011/8/8
            End If
         End If
         If IsEmptyText(txtCaseField(1)) = False Then
            lblNextCaseProperty = GetCaseTypeName(m_PA01, txtCaseField(1))
         End If
         '2004/9/7 end
         '94.1.10 ADD BY SONIA
         If txtCaseField(0) = 專利權消滅 Then
            text8 = "N"
         End If
         '94.1.10 END
         'Add by Morgan 2004/12/2
      
         
      'Add By Cheng 2002/11/22
      Case 4 '准駁通知日
      
      'Cancel by Morgan 2003/12/01
      
      '    If Me.txtCaseField(Index).Text <> "" Then
      '        If txtCaseField(0) = 核駁 And m_PA09 = 美國國家代號 Then
      '            strTemp = ChangeWStringToWDateString(DBDATE(txtCaseField(Index)))
      '            strTemp1 = DateAdd("M", 3, strTemp)
      '            '預設法定期限
      '            txtCaseField(3) = ChangeWDateStringToTString(strTemp1)
      '            '預設本所期限
      '            txtCaseField(2) = ChangeWDateStringToTString(DateAdd("D", -14, strTemp1))
      '        End If
      '        '91.11.27 ADD BY SONIA
      '        If txtCaseField(0) = 通知要求選取 And m_PA09 = 美國國家代號 Then
      '            strTemp = ChangeWStringToWDateString(DBDATE(txtCaseField(Index)))
      '            strTemp1 = DateAdd("M", 1, strTemp)
      '            '預設法定期限
      '            txtCaseField(3) = ChangeWDateStringToTString(strTemp1)
      '            '預設本所期限
      '            txtCaseField(2) = ChangeWDateStringToTString(DateAdd("D", -14, strTemp1))
      '        End If
      '        '91.11.27 END
      '    End If
      
      'End 2003/12/01
      
      Case 7 '製圖費
         '若申請國家為美國
      '   '若申請國家為美國且有輸入製圖費
         'Modify By Cheng 2002/07/31
      '   If m_PA09 = 美國國家代號 And IsNumeric(Me.txtCaseField(Index).Text) Then
         '91.11.3 MODIFY BY SONIA
         'If m_PA09 = 美國國家代號 Then
         If m_PA09 = 美國國家代號 And txtCaseField(Index) <> "" Then
         '91.11.3 END
            '預設修圖本所期限
            If Me.txtCaseField(27).Text = "" Then
               If Me.txtCaseField(2).Text <> "" Then
                  strExc(0) = TransDate(CompDate(2, 21, strSrvDate(1)), 1)
                  If Val(strExc(0)) > Val(txtCaseField(2)) Then
                     txtCaseField(27).Text = txtCaseField(2)
                  Else
                     txtCaseField(27).Text = strExc(0)
                      'Add By Cheng 2003/12/08
                      '本所期限若非工作天則抓最近工作天
                      Me.txtCaseField(27).Text = TransDate(PUB_GetWorkDay1(Me.txtCaseField(27).Text, True), 1)
                  End If
               End If
            End If
            
            'Modify by Morgan 2010/8/11 百年蟲
            'If Me.txtCaseField(27).Text > Me.txtCaseField(2).Text Then
            If Val(txtCaseField(27)) > Val(txtCaseField(2)) Then
               Me.txtCaseField(27).Text = Me.txtCaseField(2).Text
            End If
            '預設修圖法定期限
            If Me.txtCaseField(28).Text = "" Then
               If Me.txtCaseField(2).Text <> "" Then
                  Me.txtCaseField(28).Text = txtCaseField(2).Text
               End If
            End If
            'Modify by Morgan 2010/8/11 百年蟲
            'If Me.txtCaseField(28).Text > Me.txtCaseField(2).Text Then
            If Val(txtCaseField(28)) > Val(txtCaseField(2)) Then
               Me.txtCaseField(28).Text = Me.txtCaseField(2).Text
            End If
            'Modify by Morgan 2010/8/11 百年蟲
            'If Me.txtCaseField(27).Text > Me.txtCaseField(28).Text Then
            If Val(txtCaseField(27)) > Val(txtCaseField(28)) Then
               Me.txtCaseField(27).Text = Me.txtCaseField(28).Text
              'Add By Cheng 2003/12/08
              '本所期限若非工作天則抓最近工作天
              Me.txtCaseField(27).Text = TransDate(PUB_GetWorkDay1(Me.txtCaseField(27).Text, True), 1)
            End If
         Else
            Me.txtCaseField(27).Text = Empty
            Me.txtCaseField(28).Text = Empty
         End If
      Case 28 '修圖法定期限
         If Me.txtCaseField(28).Text <> "" Then
            'Modify by Morgan 2010/8/11 百年蟲
            'If Me.txtCaseField(27).Text > Me.txtCaseField(28).Text Then
            If Val(txtCaseField(27)) > Val(txtCaseField(28)) Then
               MsgBox "修圖本所期限不可大於修圖法定期限!!!", vbExclamation + vbOKOnly
               Me.SSTab1.Tab = 2
               Me.txtCaseField(27).SetFocus
            End If
         End If
   End Select
   
   'Added by Morgan 2021/12/9
   If m_iNoStopIdx <> -1 Then
      SetNextStopIndex Index
   End If
   'end 2021/12/9
End Sub

Private Sub SetNextStopIndex(pIndex As Integer)
   Dim objTxt As Object
   
   m_iNextIndex = -1
   For Each objTxt In txtCaseField
      If objTxt.Enabled = True And objTxt.Visible = True And objTxt.TabStop = True Then
         If objTxt.TabIndex > txtCaseField(pIndex).TabIndex Then
            If m_iNextIndex = -1 Then
               m_iNextIndex = objTxt.Index
            ElseIf objTxt.TabIndex < txtCaseField(m_iNextIndex).TabIndex Then
               m_iNextIndex = objTxt.Index
            End If
         End If
      End If
   Next
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
Dim strDays As String
'Add By Cheng 2002/07/25
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   If CheckKeyIn(Index) = -1 Then
      Cancel = True
   End If
   
   '90.07.01 modify by louis
   Select Case Index
      Case 0:
         If m_bolSaveCheck Then Exit Sub      'Modify by Morgan 2004/12/2 判斷是否存檔前檢查
         
         'Added by Morgan 2018/11/12 控制輸入的案件性質就是存檔的這樣才能預設通知函的判發人
         If txtCaseField(0) = "1001" And InStr(Patent1001Display, cp(10)) > 0 Then
            MsgBox "案件性質" & cp(10) & "之核准請改輸1008核發！", vbExclamation
            Cancel = True
            Exit Sub
         End If
         'end 2018/11/12
         
         '91.7.22 MODIFY BY SONIA
         'txtCaseField(9) = "": txtCaseField(5) = "": txtCaseField(14) = ""
         'txtCaseField(9) = "": txtCaseField(14) = ""
         '91.7.22 END
         
         If bolDo Then
            If txtCaseField(13).Enabled Then 'Add by Morgan 2010/10/1
               '2010/1/20 modify by sonia 承辦人為北所人員以系統日計算承辦期限,分所人員以系統日的下一個工作天計算
               'txtCaseField(13) = TransDate(Pub_GetHandleDay(m_PA01, m_PA09, txtCaseField(0)), 1)
               If m_CP14ST06 <> "1" Then
                  txtCaseField(13) = TransDate(Pub_GetHandleDay(m_PA01, m_PA09, txtCaseField(0), CompWorkDay(2, strSrvDate(1), 0), IIf(txtCaseField(2) = "", "", TransDate(txtCaseField(2), 2))), 1)
               Else
                  txtCaseField(13) = TransDate(Pub_GetHandleDay(m_PA01, m_PA09, txtCaseField(0), , IIf(txtCaseField(2) = "", "", TransDate(txtCaseField(2), 2))), 1)
               End If
               '2010/1/20 end
            End If
         End If
         
         ' 下一程序 modify by sonia 91.12.26
         'If IsEmptyText(txtCaseField(1)) Then txtCaseField(1) = GetNextProgress(m_PA01, m_PA09, txtCaseField(0))
         '92.1.16 modify by sona
         'If lblCaseField(9) = 美國國家代號 And (lblCaseField(6) = 發明申請 Or lblCaseField(6) = 新型申請 Or lblCaseField(6) = 設計申請 Or lblCaseField(6) = 追加申請 Or lblCaseField(6) = 聯合申請 Or lblCaseField(6) = 答辯) Then
         '92.9.15 MODIFY BY SONIA
         'If (lblCaseField(6) = 發明申請 Or lblCaseField(6) = 新型申請 Or lblCaseField(6) = 設計申請 Or lblCaseField(6) = 追加申請 Or lblCaseField(6) = 聯合申請 Or lblCaseField(6) = 答辯) Then
         
         'Remove by Morgan 2004/9/7 移到 LostFocus 做以免存檔前檢查又帶出下出下一程序
   '      If (lblCaseField(6) = 發明申請 Or lblCaseField(6) = 新型申請 Or lblCaseField(6) = 設計申請 Or lblCaseField(6) = 追加申請 Or lblCaseField(6) = 聯合申請 Or lblCaseField(6) = 答辯 Or lblCaseField(6) = CIP申請 Or lblCaseField(6) = CPA申請 Or lblCaseField(6) = 再發行) Then
   '      '92.9.15 END
   '      '92.1.16 end
   '         If IsEmptyText(txtCaseField(1)) Then txtCaseField(1) = GetNextProgress(m_PA01, m_PA09, txtCaseField(0))
   '      End If
   '      '91.12.26 END
   '      If IsEmptyText(txtCaseField(1)) = False Then
   '         lblNextCaseProperty = GetCaseTypeName(m_PA01, txtCaseField(1))
   '      End If
         '2004/9/7 end
         
   '2009/11/11 MODIFY BY SONIA 整理合併並加1607註冊登記
   '      '2008/11/28 ADD BY SONIA
   '      If txtCaseField(0) = "1905" Or txtCaseField(0) = "1912" Then
   '         txtCaseField(12) = strUserNum
   '         txtCaseField(14) = "N"
   '         CheckKeyIn 12
   '      End If
   '      '2008/11/28 END
   '
   '      'Add By Sindy 2009/06/17
   '      '來函性質為1223時，承辦人預設為操作人員
   '      If txtCaseField(0) = "1223" Then
   '         txtCaseField(12) = strUserNum
   '         CheckKeyIn 12
   '      End If
   '      '2009/06/17 End
   '      'Add by Morgan 2004/12/20 初步審查核可預設輸入人員
   '      If txtCaseField(0) = "1213" Then
   '        txtCaseField(12) = strUserNum
   '        lblPromoter.Caption = strUserName
   '      End If
         Select Case txtCaseField(0)
            '2010/7/16 modify by sonia 加1606專利權公告作廢
            'Modified by Morgan 2021/6/3 +1236
            Case "1905", "1912", "1223", "1213", "1607", "1606", "1236"
               txtCaseField(12) = strUserNum
               txtCaseField(14) = "N"
               CheckKeyIn 12
            'Added by Lydia 2017/05/09 新增C類官方來函性質「視為未主張」(代號：1918)，可用在主張國內(121)、國際優先權(106)及優惠期(123)
            Case "1918"
               If InStr("106,121,123", cp(10)) = 0 And cp(10) <> "" Then
                  MsgBox "視為未主張，只可用在主張國內優先權、國際優先權及優惠期"
                  Cancel = True
                  Exit Sub
               End If
            'end 2107/04/05
         End Select
   '2009/11/11 END
         
         If txtCaseField(0) = 核准 Then
            txtCaseField(9) = "Y"
            txtCaseField(12) = strUserNum
            txtCaseField(14) = "N"
            CheckKeyIn 12
            
         'Added by Morgan 2019/1/25 核發預設不出定稿,承辦人設輸入人員--慧汶 Ex:CFP-030780
         ElseIf txtCaseField(0) = "1008" Then
            'txtCaseField(5) = "N" 'Removed by Morgan 2021/5/13 改預設要通知客戶(P案核發後會發出客戶通知函，CFP調整作法與P案一致)--玫音
            txtCaseField(12) = strUserNum
            txtCaseField(14) = "N"
            CheckKeyIn 12
         'end 2019/1/25
         
   'Remove by Morgan 2008/5/12 公開費已不再輸入
   '         'Modify By Cheng 2002/07/24
   '         '若申請國家為美國(101), 專利種類為發明(1), 且來函性質為核准(1001)時, 增加公開費欄
   ''         'Add By Cheng 2002/03/07
   ''         '若來函性質為核准(1001)時, 增加公開費欄
   ''         If Me.txtCaseField(Index).Text = 核准 Then
   '          '91.12.26 MODIFY BY SONIA 應判斷申請國家為美國發明及點選案件性質為發明申請或答辯且來函性質為核准(1001)時,才增加公開費欄
   '          'If m_PA09 = 美國國家代號 And m_PA08 = "1" And Me.txtCaseField(Index).Text = 核准 Then
   '          '2005/5/19 MODIFY BY SONIA 加分割
   '          'If m_PA09 = 美國國家代號 And m_PA08 = "1" And (cp(10) = 發明申請 Or cp(10) = 答辯) And Me.txtCaseField(Index).Text = 核准 Then
   '          'Modify by Morgan 2005/8/17 加CIP申請113
   '          '2007/8/3 MODIFY BY SONIA 加請求繼續審查424
   '          If m_PA09 = 美國國家代號 And m_PA08 = "1" And (cp(10) = 發明申請 Or cp(10) = 答辯 Or cp(10) = 分割 Or cp(10) = CIP申請 Or cp(10) = "424") And Me.txtCaseField(Index).Text = 核准 Then
   '          '2005/5/19 END
   '          '91.12.26 END
   '            Me.Label33.Visible = True
   '            Me.txtCaseField(24).Visible = True
   '         Else
   '            Me.Label33.Visible = False
   '            Me.txtCaseField(24).Visible = False
   '         End If
   'end 2008/5/12
   
         End If
   
         'Added by Lydia 2016/12/26 英國核准先行通知(1005),預設不列印通知函,承辦人為操作人員
         If field(9) = "201" And txtCaseField(0) = "1005" Then
            txtCaseField(12) = strUserNum
            lblPromoter.Caption = strUserName
         End If
         'end 2016/12/26
          
         'Added by Morgan 2021/3/12
         '寶齡富錦 Y55435 案件下列來函承辦人預設韻如
         '1202審查意見來函、1002核駁、1006最終核駁、1201通知修正、1209檢索報告、1205通知提供前案、1206通知要求選取、1203通知補充說明
         m_bolBPFCase = False
         If field(75) = "Y55435" And txtCaseField(0) <> "" Then
            If InStr("1202,1002,1006,1201,1209,1205,1206,1203", txtCaseField(0).Text) > 0 Then
               txtCaseField(5) = "N"
               'Modified by Morgan 2023/6/27 預設最新收文的工程師--郭
               'txtCaseField(12) = "A0029"
               If PUB_GetLastEng(cp(1), cp(2), cp(3), cp(4), strExc(1)) Then
                  txtCaseField(12) = strExc(1)
               Else
                  txtCaseField(12) = "A0029"
               End If
               m_bolBPFCase = True
               'end 2023/6/27
               CheckKeyIn 12
            End If
         End If
         'end 2021/3/12
                        
         'Modify by Morgan 2006/3/23 加最終核駁1006
         'Modify by Morgan 2010/3/10 +建議性處分書 1220
         If InStr("1002,1006,1220", txtCaseField(0)) > 0 Then
            'Add by Morgan 2006/3/23 若為美國答辯核駁時提醒程序人員
            'Modified by Morgan 2016/3/3 +126 期末拋棄
            'Modified by Lydia 2016/08/27 +438 再考量試行計畫(AFCP2.0)
            If txtCaseField(0) = 核駁 And field(9) = 美國國家代號 And (cp(10) = 答辯 Or cp(10) = "126" Or cp(10) = "438") Then
               If MsgBox("請再確認是否應為最終核駁！", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                  Cancel = True
                  Exit Sub
               End If
            End If
            
            txtCaseField(5) = "N"
            'Add by Morgan 2005/5/20 核駁且有答辯不續辦要印通知函
            If m_bolIsNP107N = True Then
               txtCaseField(5) = ""
            End If
            '2005/5/20 end
            
            '2013/8/12 ADD BY SONIA 特殊客戶只印定稿不分析,也要印案件回覆單
            If m_specialCust = True Then
               txtCaseField(5) = ""
            End If
            '2013/8/12 end
            
            '91.12.10 add by sonia
            txtCaseField(14) = "N"
            '91.12.10 end
            'Add By Cheng 2002/07/25
            If Me.txtCaseField(12).Text = "" Then
               '2007/8/3 MODIFY BY SONIA 加424請求繼續審查
               'Modify by Sonia  2013/11/15 加復審805(CFP-023821)
               'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
               'Modified by Lydia 2016/08/28 +438 再考量試行計畫(AFCP2.0)
               StrSQLa = "Select CP14 FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " AND (CP10='107' OR CP10='424' OR CP10='126' OR CP10='805' OR CP10='438') ORDER BY CP27 DESC"
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  txtCaseField(12).Text = "" & rsA.Fields(0).Value
                  lblPromoter.Caption = GetStaffName(txtCaseField(12).Text)
                   'Add By Cheng 2002/11/22
                   '若承辦人離職, 則將承辦人代號清空
                   If lblPromoter.Caption = "" Then txtCaseField(12).Text = ""
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               If Me.txtCaseField(12).Text = "" Then
                  StrSQLa = "Select CP14 FROM CASEPROGRESS,STAFF WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " AND CP14=ST01(+) AND ST03<>'P12' AND CP09<'C' ORDER BY CP05 DESC, CP09 DESC "
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     txtCaseField(12).Text = "" & rsA.Fields(0).Value
                     lblPromoter.Caption = GetStaffName(txtCaseField(12).Text)
                     'Add By Cheng 2002/11/22
                     '若承辦人離職, 則將承辦人代號清空
                     If lblPromoter.Caption = "" Then txtCaseField(12).Text = ""
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
               End If
            End If
            'add by sonia 2018/10/5 美國最終核駁時,若有已發文之請求繼續審查則提醒程序人員CFP-029197
            If txtCaseField(0) = "1006" And field(9) = 美國國家代號 Then
               StrSQLa = "Select CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " AND CP10='424' AND CP27>0"
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  MsgBox "本案已提過請求繼續審查，請注意規費之報價！"
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               
               'Added by Morgan 2025/5/12
               strExc(0) = "select * from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06 is null and np07='214'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "本案下一程序有IDS尚未收文，RCE報價時請檢視是否需要加上差額！", vbExclamation
               End If
               
               strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='214' and cp158=0 and cp159=0"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "本案已收文IDS尚未發文，RCE報價時請檢視是否需要加上差額！", vbExclamation
               End If
               'end 2025/5/12
            End If
            'end 2018/10/5
         End If
      Case Else:
   End Select
   If Cancel Then
      txtCaseField_GotFocus Index
      txtCaseField(Index).SetFocus 'Added by Morgan 2021/12/22
   End If
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Function CheckKeyIn(intIndex As Integer) As Integer
' 91.03.25 modify by louis
' Input : intIndex
'         bComputeCP48 == 是否檢查承辦限有輸入(沒資料則自動計算)
'                         TRUE ==> 檢查並自動計算
'                         FALSE ==> 不檢查且不會自動計算
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CheckKeyIn(intIndex As Integer, Optional ByVal bComputeCP48 As Boolean = True) As Integer
Dim strTemp As String, strTemp1 As String, bolIsChina As Boolean, strCusTemp As String

   CheckKeyIn = -1
   Select Case intIndex
             Case 0 '來函性質
                        'Add By Cheng 2002/01/04
                        If Len(Me.txtCaseField(0).Text) > 0 Then
                           If Len(Me.txtCaseField(0).Text) <> 4 Then
                              MsgBox "來函性質欄位值必須為四碼 !", vbExclamation
                              Exit Function
                           End If
                        End If
                        'edit by nickc 2007/02/02 不用 dll 了
                        'If objPublicData.GetCaseProperty(cp(1), txtCaseField(intIndex), strTemp) Then
                        If ClsPDGetCaseProperty(cp(1), txtCaseField(intIndex), strTemp) Then
                           lblProperty = strTemp
                           CheckKeyIn = 1
'                           If objPublicData.GetCaseWorkDays(cp(1), lblCaseField(9), txtCaseField(intIndex), strTemp) Then
'                              If strTemp <> "" Then
'                                 strTemp1 = CompDate(2, Val(strTemp), strSrvDate(1))
'                                 If bComputeCP48 Then
'                                    txtCaseField(13) = TransDate(strTemp1, 1)
'                                 End If
'                              End If
'                           End If
                        End If
             Case 1 '下一程序
                        'Add By Cheng 2002/01/04
                        If Len(Me.txtCaseField(1).Text) > 0 Then
                           If Len(Me.txtCaseField(1).Text) <> 3 Then
                              MsgBox "下一程序欄位值必須為三碼 !", vbExclamation
                              Exit Function
                           End If
                        End If
                       'Modified by Lydia 2016/12/26 英國核准先行通知(1005),不輸入下一程序,法定期限及本所期限欄不可空白
                       'If txtCaseField(intIndex) = "" Then
                       If lblCaseField(9) = "201" And txtCaseField(0) = "1005" Then
                          If txtCaseField(1) <> "" Then
                             MsgBox "英國核准先行通知，不輸入下一程序 !", vbExclamation
                             txtCaseField(1) = ""
                             lblNextCaseProperty = ""
                             Exit Function
                          Else
                             CheckKeyIn = 1
                          End If
                       ElseIf txtCaseField(intIndex) = "" Then
                       'end 2016/12/26
                          txtCaseField(2) = ""
                          txtCaseField(3) = ""
                          CheckKeyIn = 1
                       Else
                          If lblCaseField(9) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
                          'edit by nickc 2007/02/02 不用 dll 了
                          'If objPublicData.GetCaseProperty(cp(1), txtCaseField(intIndex), strTemp, bolIsChina) Then
                          If ClsPDGetCaseProperty(cp(1), txtCaseField(intIndex), strTemp, bolIsChina) Then
                              lblNextCaseProperty = strTemp
                              CheckKeyIn = 1
                           End If
                        End If
             Case 2 '本所期限
                        '若有輸入本所期限
                        If IsEmptyText(txtCaseField(2)) = False Then
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              '91.11.27 MODIFY BY SONIA
                              'If CheckReKey(txtCaseField(intIndex)) Then
                              '   CheckKeyIn = 1
                              'End If
                              If Val(txtCaseField(2)) <= Val(txtCaseField(3)) Then
                                 If CheckReKey(txtCaseField(intIndex)) Then
                                    CheckKeyIn = 1
                                 End If
                              Else
                                 ShowMsg MsgText(1033)
                              End If
                              '91.11.27 END
                           End If
                           'Add By Cheng 2002/03/11
                           If Val(Me.txtCaseField(intIndex).Text) + 19110000 < strSrvDate(1) Then
                              MsgBox "本所期限不可小於系統日期!!!", vbExclamation
                              CheckKeyIn = -1
                            'Add By Cheng 2003/12/08
                            '若本所期限非工作天則直接調整至最近的工作天
                            Else
                                Me.txtCaseField(intIndex).Text = TransDate(PUB_GetWorkDay1(Me.txtCaseField(intIndex).Text, True), 1)
                           End If
                           '2008/11/27 add by sonia 約定期限未輸入時預設為本所期限-28天且須為工作天
                           'Modified by Morgan 2024/12/6 改控制enable，這樣駐點才不會亂跳
                           'If txtCaseField(30).Visible = True And txtCaseField(30) = "" Then
                           If txtCaseField(30).Enabled And txtCaseField(30) = "" Then
                           'end 2024/12/6
                              txtCaseField(30) = TransDate(PUB_GetWorkDay1(CompDate(2, -28, Me.txtCaseField(intIndex).Text), True), 1)
                           End If
                           '2008/11/27 END
                        Else
                           CheckKeyIn = 1
                        End If
             Case 3 '法定期限
                        If IsEmptyText(txtCaseField(3)) = False Then
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                           
                              'Add by Morgan 2003/12/01
                              'Modify by Morgan 2006/3/23 加最終核駁
                              'Modify by Morgan 2010/3/10 +建議性處分書 1220
                              If m_PA09 = 美國國家代號 And (txtCaseField(0) = 核駁 Or txtCaseField(0) = 通知要求選取 Or txtCaseField(0) = "1006" Or txtCaseField(0) = "1220") Then
                                 If IsEmptyText(txtCaseField(4)) = False Then
                                    strTemp = ChangeWStringToWDateString(DBDATE(txtCaseField(4)))
                                    strTemp1 = DateAdd("M", 3, strTemp)
                                    If (txtCaseField(3) <> ChangeWDateStringToTString(strTemp1)) Then
                                       If (MsgBox("法定期限與准駁通知日相差不為3個月，是否要繼續？", vbYesNo)) <> 6 Then
                                          Exit Function
                                       End If
                                    End If
                                 End If
                              End If
                              'End 2003/12/01
                              
                              If txtCaseField(2).Text = "" Then
                                 txtCaseField(2).Text = TransDate(CompDate(2, -14, TransDate(txtCaseField(3), 2)), 1)
                                 'add by sonia 2021/4/16 加拿大核駁期限延期很嚴格,所以本所期限再往前1個月
                                 If m_PA09 = "102" And (txtCaseField(0) = 核駁 Or txtCaseField(0) = "1006") Then
                                    txtCaseField(2).Text = TransDate(CompDate(1, -1, TransDate(txtCaseField(2), 2)), 1)
                                 End If
                                 'end 2021/4/16
                                 'Add By Cheng 2003/12/08
                                 '本所期限若非工作天則抓最近工作天
                                 Me.txtCaseField(2).Text = TransDate(PUB_GetWorkDay1(Me.txtCaseField(2).Text, True), 1)
                              End If
                              '91.11.27 CANCEL BY SONIA 移至本所期限檢查
                              'If txtCaseField(2) <= txtCaseField(3) Then
                              '   If CheckReKey(txtCaseField(intIndex)) Then
                              '      CheckKeyIn = 1
                              '   End If
                              'Else
                              '   ShowMsg MsgText(1033)
                              'End If
                              If CheckReKey(txtCaseField(intIndex)) Then
                                 CheckKeyIn = 1
                              End If
                              '91.11.27 END
                           End If
                        Else
                           CheckKeyIn = 1
                        End If
            Case 4
                        If cp(10) <> 讓與 Then
                           'Modify by Morgan2006/3/23 加最終核駁
                           'Modify by Morgan 2010/3/10 +建議性處分書 1220
                           If txtCaseField(intIndex) = "" And txtCaseField(0) <> 核准 And txtCaseField(0) <> 核駁 And txtCaseField(0) <> "1006" And txtCaseField(0) <> "1220" Then
                              CheckKeyIn = 1
                           ElseIf CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              If Val(txtCaseField(intIndex)) <= Val(strSrvDate(2)) Then
                                 CheckKeyIn = 1
                              Else
                                 ShowMsg MsgText(1052)
                              End If
                           End If
                        Else
                           If txtCaseField(intIndex) = "" Then
                              If m_PA09 <> "101" Then
                                 CheckKeyIn = 1
                              Else
                                 ShowMsg MsgText(9003)
                              End If
                           ElseIf CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              If Val(txtCaseField(intIndex)) <= Val(strSrvDate(2)) Then
                                 CheckKeyIn = 1
                              Else
                                 ShowMsg MsgText(1052)
                              End If
                           End If
                        End If
             Case 5, 14
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 9
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Or txtCaseField(intIndex) = "Y" Then
                           If txtCaseField(intIndex) = "" And txtCaseField(0) = 核准 Then
                              ShowMsg MsgText(1053)
                           Else
                              CheckKeyIn = 1
                           End If
                        Else
                           ShowMsg MsgText(9177)
                        End If
             Case 12 '承辦人
                        m_CP14ST06 = "1" '2010/1/20 add by sonia
                        If txtCaseField(intIndex) = "" Then
                           lblPromoter = "" 'Add by Morgan 2004/12/2
                           CheckKeyIn = 1
                        Else
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetStaff(txtCaseField(intIndex), strTemp) Then
                           If ClsPDGetStaff(txtCaseField(intIndex), strTemp) Then
                              lblPromoter = strTemp
                              CheckKeyIn = 1
                              m_CP14ST06 = PUB_GetST06(txtCaseField(intIndex))  '2010/1/20 add by sonia
                              '92.5.8 ADD BY SONIA
                              strExc(0) = "SELECT ST03 FROM STAFF WHERE ST01='" & txtCaseField(intIndex) & "'"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 
                                 If Not IsNull(RsTemp.Fields("ST03")) And RsTemp.Fields("ST03") = "P12" Then
                                    txtCaseField(14) = "N"
                                 'Added by Morgan 2019/2/26
                                 ElseIf Left("" & RsTemp.Fields("ST03"), 2) <> "P1" Then
                                    MsgBox "承辦人非專利處人員，請重新分案！", vbExclamation
                                    CheckKeyIn = -1
                                 'end 2019/2/26
                                 End If
                              End If
                              '92.5.8 END
                           Else
                              lblPromoter = ""
                           End If
                        End If
                     
                        '2010/1/20 add by sonia 重新依承辦人所別以系統日或下一個工作天計算承辦期限
                        If txtCaseField(12).Tag <> txtCaseField(12) Then 'Add by Morgan 2010/5/25
                           If txtCaseField(13).Enabled Then 'Add by Morgan 2010/10/1
                              If m_CP14ST06 <> "1" Then
                                 txtCaseField(13) = TransDate(Pub_GetHandleDay(m_PA01, m_PA09, txtCaseField(0), CompWorkDay(2, strSrvDate(1), 0), IIf(txtCaseField(2) = "", "", TransDate(txtCaseField(2), 2))), 1)
                              Else
                                 txtCaseField(13) = TransDate(Pub_GetHandleDay(m_PA01, m_PA09, txtCaseField(0), , IIf(txtCaseField(2) = "", "", TransDate(txtCaseField(2), 2))), 1)
                              End If
                           End If 'Add by Morgan 2010/10/1
                           txtCaseField(12).Tag = txtCaseField(12) 'Add by Morgan 2010/5/25
                        End If
                        '2010/1/20 end
            'Mark by Amy 2014/09/17 承辦期限欄位隱藏
'            Case 13 '承辦期限
'                        If txtCaseField(intIndex) = "" Then
'                           ' 91.03.25 modify by louis (忽略承辦期限空白的情形)
'                           If bComputeCP48 Then
'                              'edit by nickc 2007/02/02 不用 dll 了
'                              'If objPublicData.GetCaseWorkDays(cp(1), lblCaseField(9), txtCaseField(0), strTemp) Then
'                              'Modify by Morgan 2007/10/16 工作天函數統一
'                              'If ClsPDGetCaseWorkDays(cp(1), lblCaseField(9), txtCaseField(0), strTemp) Then
'                              strTemp = GetWorkDays(cp(1), lblCaseField(9), txtCaseField(0))
'                              'end 2007/10/16
'                                 If strTemp <> "" And txtCaseField(1) <> "" Then
'                                    ShowMsg MsgText(1049)
'                                 Else
'                                    CheckKeyIn = 1
'                                 End If
'                              'End If
'                           Else
'                              CheckKeyIn = 1
'                           End If
'                        ElseIf CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
'                           CheckKeyIn = 1
'                        End If
'                        If txtCaseField(13).Enabled Then 'Add by Morgan 2010/10/1
'                           'Add By Cheng 2002/05/06
'                           '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
'                           If Len(Me.txtCaseField(2).Text) > 0 And Len(Me.txtCaseField(13).Text) > 0 Then
'                              If Val(Me.txtCaseField(2).Text) < Val(Me.txtCaseField(13).Text) Then
'                                 MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
'                                 CheckKeyIn = -1
'                                 Exit Function
'                              End If
'                           End If
'                        End If
            'end 2014/09/17
            Case 24 '公開費
                        'Modify By Cheng 2002/07/24
                        '若有顯示公開費欄位, 則不可空白
                        If Me.txtCaseField(intIndex).Visible Then
'                        'Add By Cheng 2002/03/07
                           '若有輸入公開費
                           If Len(Me.txtCaseField(intIndex).Text) > 0 Then
                              If IsNumeric(Me.txtCaseField(intIndex).Text) = False Then
                                 ShowMsg MsgText(9002)
                              Else
                                 CheckKeyIn = 1
                              End If
                           'Modify by Morgan 2007/3/27 不再限制一定要輸(96/4/1以後)--郭
                           '若沒輸入公開費
                           ElseIf strSrvDate(1) < 20070401 Then
                              MsgBox "請輸入公開費!!!", vbExclamation + vbOKOnly
                           Else
                              CheckKeyIn = 1
                           'End 2007/3/27
                           End If
                         Else
                           CheckKeyIn = 1
                         End If
                        
             'Add By Cheng 2002/07/25
             Case 27 '修圖本所期限
               If Me.txtCaseField(27).Text <> "" Then
                  If CheckIsTaiwanDate(Me.txtCaseField(27).Text) = False Then
                     Me.SSTab1.Tab = 2
                  
                  'Modify by Morgan 2010/8/11 百年蟲
                  'ElseIf Me.txtCaseField(27).Text > Me.txtCaseField(2).Text Then
                  ElseIf Val(txtCaseField(27)) > Val(txtCaseField(2)) Then
                     MsgBox "修圖本所期限不可大於基本資料的本所期限!!!", vbExclamation + vbOKOnly
                    'Add By Cheng 2003/12/08
                  '若修改本所期限非工作天
                  ElseIf ChkWorkDay(DBDATE(Me.txtCaseField(27).Text)) = False Then
                     MsgBox "修圖本所期限非工入天, 請重新輸入!!!", vbExclamation + vbOKOnly
                  Else
                     CheckKeyIn = 1
                  End If
               Else
                  CheckKeyIn = 1
               End If
             Case 28 '修圖法定期限
               If Me.txtCaseField(28).Text <> "" Then
                  If CheckIsTaiwanDate(Me.txtCaseField(28).Text) = False Then
                     Me.SSTab1.Tab = 2
                  
                  'Modify by Morgan 2010/8/11 百年蟲
                  'ElseIf Me.txtCaseField(28).Text > Me.txtCaseField(2).Text Then
                  ElseIf Val(txtCaseField(28)) > Val(txtCaseField(2)) Then
                     MsgBox "修圖法定期限不可大於基本資料的本所期限!!!", vbExclamation + vbOKOnly
                  Else
                     CheckKeyIn = 1
                  End If
               Else
                  CheckKeyIn = 1
               End If
             Case 11 '美國讓與登記號
                  If CheckLengthIsOK(txtCaseField(11), 20) = True Then
                     CheckKeyIn = 1
                  End If
             'Add by Morgan 2005/5/20
             Case 30 '約定期限
               If IsEmptyText(txtCaseField(intIndex)) = False Then
                  If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                     CheckKeyIn = 1
                  End If
                  If Val(Me.txtCaseField(intIndex).Text) + 19110000 < strSrvDate(1) Then
                     MsgBox "約定期限不可小於系統日期!!!", vbExclamation
                     CheckKeyIn = -1
                  'Added by Morgan 2021/10/14
                  ElseIf Val(Me.txtCaseField(intIndex).Text) > Val(Me.txtCaseField(2).Text) And Val(Me.txtCaseField(2).Text) > 0 Then
                     MsgBox "約定期限不可晚於本所期限!!!", vbExclamation
                     CheckKeyIn = -1
                  'end 2021/10/14
                  '2008/11/27 ADD BY SONIA
                  Else
                     Me.txtCaseField(intIndex).Text = TransDate(PUB_GetWorkDay1(Me.txtCaseField(intIndex).Text, True), 1)
                  '2008/11/27 END
                  End If
               Else
                  CheckKeyIn = 1
               End If
             '2011/6/2 add by sonia
             Case 17, 39
                  If CheckLengthIsOK(txtCaseField(intIndex), txtCaseField(intIndex).MaxLength) = True Then
                     CheckKeyIn = 1
                  End If
             '2011/6/2 end
             Case Else
                  CheckKeyIn = 1

   End Select
End Function

Private Sub txtCaseField_GotFocus(Index As Integer)
   txtCaseField(Index).SelStart = 0
   txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
   '儲存未修改前之值至Tag中,供再確認時使用
   txtCaseField(Index).Tag = txtCaseField(Index)
   '91.11.13 ADD BY SONIA
   If Index = 0 Then
      bolDo = True
   Else
      bolDo = False
   End If
   
   'Added by Morgan 2021/12/8
   'visible從False變True時駐點也會跳來,要控制再跳到原來的下一個駐點
   If Index = m_iNoStopIdx Then
      If m_iNextIndex <> -1 Then
         txtCaseField(m_iNextIndex).SetFocus
         m_iNextIndex = -1
      End If
      m_iNoStopIdx = -1
   'ElseIf Index = m_iNextIndex Then
   '   txtCaseField(Index).SetFocus 'Form2.0關閉對話框回來游標會不見
   End If
   'end 2021/12/9
End Sub

' 90.07.12 modify by louis (檢查資料是否輸入完整)
Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse As Long
Dim strDate As String
   
   CheckDataValid = False
   
   ' 准駁通知日
   If IsEmptyText(txtCaseField(4)) = False Then
      strDate = txtCaseField(4)
      If CheckIsTaiwanDate(strDate, False) = False Then
         strTit = "檢核資料"
         strMsg = "准駁通知日日期格式不正確!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtCaseField(4).SetFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 本所期限
   If IsEmptyText(txtCaseField(2)) = False Then
      strDate = txtCaseField(2)
      If CheckIsTaiwanDate(strDate, False) = False Then
         strTit = "檢核資料"
         strMsg = "本所期限日期格式不正確!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtCaseField(2).SetFocus
         GoTo EXITSUB
      End If
      If Val(txtCaseField(2)) < Val(strSrvDate(2)) Then
         strTit = "檢核資料"
         strMsg = "本所期限不可小於系統日!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtCaseField(2).SetFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 法定期限
   If IsEmptyText(txtCaseField(3)) = False Then
      strDate = txtCaseField(3)
      If CheckIsTaiwanDate(strDate, False) = False Then
         strTit = "檢核資料"
         strMsg = "法定期限日期格式不正確!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtCaseField(3).SetFocus
         GoTo EXITSUB
      End If
   End If
   
   If txtCaseField(13).Enabled Then 'Add by Morgan 2010/10/1
      ' 承辦期限
      If IsEmptyText(txtCaseField(13)) = False Then
         strDate = txtCaseField(13)
         If CheckIsTaiwanDate(strDate, False) = False Then
            strTit = "檢核資料"
            strMsg = "承辦期限日期格式不正確!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            txtCaseField(13).SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   ' 有輸入下一程序時才需要輸入本所期限及法定期限
   If IsEmptyText(txtCaseField(1)) = False Then
      
      If IsEmptyText(txtCaseField(3)) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入法定期限!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtCaseField(3).SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(txtCaseField(2)) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入本所期限!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtCaseField(2).SetFocus
         GoTo EXITSUB
      End If
   End If
   
   '檢查是否有按指定國家的按鈕，並且又指定國家
   If (strMoneyCountry = "" Or strMoney = "") And cmdCountry.Enabled = True And txtCaseField(0) = "1001" Then
      strTit = "檢核資料"
      strMsg = "請輸入指定國家!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      cmdCountry.SetFocus
      GoTo EXITSUB
   End If
   'Add By Cheng 2002/07/25
   'Modify By Cheng 2002/07/31
'   If m_PA09 = 美國國家代號 And IsNumeric(Me.txtCaseField(7).Text) Then
   If m_PA09 = 美國國家代號 Then
      If txtCaseField(7) <> "" Then
         If Me.txtCaseField(27).Text = "" Then
            strTit = "檢核資料"
            strMsg = "請輸入修圖本所期限!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Me.SSTab1.Tab = 2
            Me.txtCaseField(27).SetFocus
            GoTo EXITSUB
         End If
         If Me.txtCaseField(28).Text = "" Then
            strTit = "檢核資料"
            strMsg = "請輸入修圖法定期限!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Me.SSTab1.Tab = 2
            Me.txtCaseField(28).SetFocus
            GoTo EXITSUB
         End If
      End If
      If Val(txtCaseField(27).Text) > Val(txtCaseField(28).Text) Then
         MsgBox "修圖本所期限不可大於修圖法定期限!!!", vbExclamation + vbOKOnly
         Me.SSTab1.Tab = 2
         Me.txtCaseField(27).SetFocus
         GoTo EXITSUB
      End If
   Else
      Me.txtCaseField(27).Text = Empty
      Me.txtCaseField(28).Text = Empty
   End If
   
   'Added by Morgan 2020/12/24
   '美國發明核准若有IDS期限時需報價
   m_bolIDSPrice = False
   If field(9) = 美國國家代號 And field(8) = "1" And txtCaseField(0) = "1001" And txtCaseField(1) = "601" Then
      'Modified by Morgan 2022/11/24 排除IDS不續辦者--郭 Ex:CFP-31552
      'Modified by Morgan 2022/11/24 不續辦還是要檢查(維持請作單條件,由工程師判斷)--郭
      strExc(0) = "select np09,np15,np06 from nextprogress where np02='" & field(1) & "' and np03='" & field(2) & "' and np04='" & field(3) & "'" & _
         " and np05='" & field(4) & "' and np07='214' and NVL(np06,'N')='N' order by np09 asc"
      'end 2022/11/24
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         intI = 0
         '有超過法限，建議客戶提RCE(要報價)
         If RsTemp("np09") < strSrvDate(1) Then
            intI = 1
            m_bolRCE = True
            If txtCaseField(42) = "" Or txtCaseField(43) = "" Then
               MsgBox "本案有 IDS 期限已過法限，將會建議客戶提 RCE，請輸入報價！", vbExclamation, "RCE報價提醒"
               If txtCaseField(42) = "" Then
                  txtCaseField(42).SetFocus
               Else
                  txtCaseField(43).SetFocus
               End If
               GoTo EXITSUB
            End If
            
         '法限>=系統日+1個月，提醒客戶處理IDS(要報價)
         ElseIf RsTemp("np09") >= CompDate(1, 1, strSrvDate(1)) Then
            intI = 1
         End If
         If intI = 1 Then
            strExc(1) = "IDS相關案號及期限："
            With RsTemp
            Do While Not .EOF
               strExc(1) = strExc(1) & vbCrLf & vbCrLf & .Fields("NP15") & vbCrLf & "　　" & "法限：" & ChangeWStringToTDateString(.Fields("np09")) & IIf(.Fields("np06") = "N", "(不續辦)", "")
               .MoveNext
            Loop
            End With
            If MsgBox("本案需ＩＤＳ報價，請確認輸入金額是否包含下列案件？" & vbCrLf & vbCrLf & strExc(1), vbExclamation + vbYesNo + vbDefaultButton2, "IDS報價提醒") = vbNo Then
               GoTo EXITSUB
            End If
            If txtCaseField(40) = "" Or txtCaseField(41) = "" Then
               MsgBox "請輸入ＩＤＳ報價！", vbExclamation, "IDS報價提醒"
               If txtCaseField(40) = "" Then
                  txtCaseField(40).SetFocus
               Else
                  txtCaseField(41).SetFocus
               End If
               GoTo EXITSUB
            End If
            m_bolIDSPrice = True
         End If
      End If
   End If
   'end 2020/12/24
   
   'Added by Morgan 2023/12/11
   'CFP其他國家有實體審查的案件(美國及有管制實審期限國家)，輸入核准時，若有美國案且沒有收到OA，詢問USER是否要報價IDS
   m_USCaseNo = ""
   If txtCaseField(0) = "1001" And InStr("101,102,103,301,302,303,107,307", cp(10)) > 0 Then
      intI = 0
      '美國
      If m_PA09 = 美國國家代號 Then
         intI = 1
         
      '有管制實審期限國家
      Else
         strExc(0) = "select na26,na28,na30 from nation where na01='" & field(9) & "' and decode('" & field(8) & "','1',na26,'2',na28,'3',na30) is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      End If
      If intI = 1 Then
         nResponse = 0
         strExc(0) = "select a.cp10,b.cp10 from caseprogress a,caseprogress b" & _
            " where a.cp01='" & field(1) & "' and a.cp02='" & field(2) & "' and a.cp03='" & field(3) & "' and a.cp04='" & field(4) & "'" & _
            " and substr(a.cp09,1,1)='C' and b.cp09(+)=a.cp43"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         'OA檢查
         If intI = 1 Then
            With RsTemp
            Do While Not .EOF
               If PUB_CheckIDSOA(field(1), field(9), .Fields(0), "" & .Fields(1)) Then
                  nResponse = 1
                  Exit Do
               End If
               .MoveNext
            Loop
            End With
         End If
         If nResponse = 0 Then
            m_USCaseNo = PUB_GetUSCaseNo(field(1), field(2), field(3), field(4))
         End If
      End If
   End If
   If m_USCaseNo <> "" Then
      If txtIDSFee(1) = "" Or txtIDSFee(2) = "" Or txtIDSPt(1) = "" Or txtIDSPt(2) = "" Then
         If MsgBox("本案未曾收到OA且有相關美國案(" & m_USCaseNo & ")，是否要報價ＩＤＳ？", vbYesNo + vbDefaultButton2 + vbExclamation, "ＩＤＳ報價") = vbYes Then
            SSTab1.Tab = 1
            If txtIDSFee(1) = "" Then
               txtIDSFee(1).SetFocus
            ElseIf txtIDSPt(1) = "" Then
               txtIDSPt(1).SetFocus
            ElseIf txtIDSFee(2) = "" Then
               txtIDSFee(2).SetFocus
            ElseIf txtIDSPt(2) = "" Then
               txtIDSPt(2).SetFocus
            End If
            Exit Function
         Else
            m_USCaseNo = ""
         End If
      End If
   
   'Added by Morgan 2025/4/30
   '有關CFP其他國家在輸入OA時，除了原有IDS控管不變外，請增加跳提醒程序人員有報IDS，提醒畫面為"本案有相關美國案，請一併報IDS。"
   '，程序人員就一定要輸IDS金額，並直接於IDS報價欄位輸入金額，執行後將輸入的金額帶入報價備註"IDS第一階段12,000(6P)第二階段16,000(6P)"
   '若是相關美國案有已收文未發文之IDS，則請跳訊息"本案有相關美國案已收文IDS尚未發文，請確認加上本案的前案是否有超過25件，若有，則請報差價給工程師。"
   ElseIf PUB_CheckIDSOA(field(1), field(9), txtCaseField(0), cp(10)) = True Then
      strExc(1) = PUB_GetUSCaseNo(field(1), field(2), field(3), field(4))
      If strExc(1) <> "" Then
         If txtIDSFee(1) = "" Or txtIDSFee(2) = "" Or txtIDSPt(1) = "" Or txtIDSPt(2) = "" Then
            If ChkIDS(strExc(1)) = True Then
               SSTab1.Tab = 1
               If txtIDSFee(1) = "" Then
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
   'end 2025/4/30
   End If
   'end 2023/12/11
   
   '91.12.26 ADD BY SONIA
   'Modify by Morgan 2007/8/9 加繼承
   'If m_PA09 = 美國國家代號 And cp(10) = 讓與 And Me.txtCaseField(0).Text = 核准 And txtCaseField(4) = "" Then
   'Modify by Morgan 2007/10/3 加授權
   'Modify by Morgan 2008/1/14 加變更
   If m_PA09 = 美國國家代號 And (cp(10) = 讓與 Or cp(10) = 繼承 Or cp(10) = 授權 Or cp(10) = 變更) And Me.txtCaseField(0).Text = 核准 And txtCaseField(4) = "" Then
      MsgBox "美國讓與/繼承/授權/變更核准請輸入准駁通知日!!!", vbExclamation + vbOKOnly
      Me.txtCaseField(4).SetFocus
      GoTo EXITSUB
   End If
   'Modify by Morgan 2007/8/9 加繼承
   'If m_PA09 = 美國國家代號 And cp(10) = 讓與 And Me.txtCaseField(0).Text = 核准 And (Me.txtCaseField(11).Text = "" Or Me.txtCaseField(11).Text = "第號/第格") Then
   'Modify by Morgan 2007/10/3 加授權
   'Modify by Morgan 2008/1/14 加變更
   If m_PA09 = 美國國家代號 And (cp(10) = 讓與 Or cp(10) = 繼承 Or cp(10) = 授權 Or cp(10) = 變更) And Me.txtCaseField(0).Text = 核准 And (Me.txtCaseField(11).Text = "" Or Me.txtCaseField(11).Text = "第號/第格") Then
      MsgBox "美國讓與/繼承核准請輸入讓渡登記號，此欄記錄在案件進度檔的'大陸申請案號'欄!!!", vbExclamation + vbOKOnly
      'Added by Morgan 2017/5/8 變更不一定會有讓渡登記號(有同時辦讓渡才有) --玫音 Ex.CFP-18957
      If cp(10) = 變更 Then
         If MsgBox("是否有讓渡登記號?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
            Me.txtCaseField(11).SetFocus
            GoTo EXITSUB
         Else
            txtCaseField(11).Text = ""
         End If
      Else
      'end 2017/5/8
         MsgBox "美國讓與/繼承核准請輸入讓渡登記號，此欄記錄在案件進度檔的'大陸申請案號'欄!!!", vbExclamation + vbOKOnly
         Me.txtCaseField(11).SetFocus
         GoTo EXITSUB
      End If 'Added by Morgan 2017/5/8
   End If
   '91.12.26 END
   
   'Added by Morgan 2018/10/4 核准有領證期限一定要出客戶函--玫音
   If txtCaseField(0) = 核准 And txtCaseField(1) = "601" Then
      If txtCaseField(5) = "N" Then
         MsgBox "核准且有領證是否列印通知函不可為 N！", vbExclamation
         txtCaseField(5).SetFocus
         GoTo EXITSUB
      End If
   End If
   'end 2018/10/4
   
   'Added by Lydia 2020/11/19 CFP英國脫歐案管制：若有英國再註冊來函時，除非英國新案已收文否則再註冊仍輸在歐盟案
   'Modified by Lydia 2020/12/01 改成歐盟案或英國案
   'If field(1) = "CFP" And field(9) = "239" And txtCaseField(0) = "1608" And txtCaseField(26).Visible = True Then
   If field(1) = "CFP" And (field(9) = "239" Or field(9) = "201") And txtCaseField(0) = "1608" And txtCaseField(26).Visible = True Then
       If Trim(txtCaseField(26)) = "" Then
            'Modified by Lydia 2020/12/01
            'MsgBox "歐盟案之" & lblProperty.Caption & " 需要輸入英國脫歐案專利號數!!!", vbExclamation + vbOKOnly
            MsgBox " 需要輸入英國脫歐案專利號數!!!", vbExclamation + vbOKOnly
            txtCaseField(26).SetFocus
            txtCaseField_GotFocus 26
            GoTo EXITSUB
       Else '檢查:是否符合編碼原則=> 9+歐盟設計專利號(拿掉-符號)
            strExc(1) = "9" & Replace(field(22), "-", "")
            If Trim(txtCaseField(26)) <> strExc(1) Then
                strExc(2) = "歐盟設計專利號數：" & field(22) & vbCrLf & _
                                 "英國專利號數編碼：9+歐盟設計專利號(拿掉-符號)" & vbCrLf & _
                                 "目前欄位輸入號數〔" & txtCaseField(26) & "〕" & vbCrLf & vbCrLf & "請問是否繼續存檔？"
                If MsgBox(strExc(2), vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                    txtCaseField(26).SetFocus
                    txtCaseField_GotFocus 26
                    GoTo EXITSUB
                End If
            End If
       End If
   End If
   'end 2020/1/19
   
   CheckDataValid = True
EXITSUB:
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim bPaper As Boolean

   bolDo = False
   TxtValidate = False
   
   'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   'Added by Morgan 2013/1/14
   '檢查移轉或讓與的受讓人(5個)與基本檔是否相同
   If txtCaseField(0) = "1001" And InStr("701,702,703,708", cp(10)) > 0 Then
      If PUB_ChkAsignCaseCustNo(cp(9)) = False Then
         Exit Function
      End If
   End If
   'end 2013/1/14
            
   For Each objTxt In Me.txtCaseField
      If objTxt.Enabled = True Then
         Cancel = False
         txtCaseField_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   txtPA14_Validate Cancel
   If Cancel = True Then
      txtPA14.SetFocus
      Exit Function
   End If
   
   'Added by Morgan 2024/4/18
   '日本核駁答辯期限若有提超級加速審查時提醒設定為2個月而非3個月--禧佩
   If field(9) = "011" And txtCaseField(0) = "1002" And txtCaseField(1) = "107" Then
      If PUB_ChkCPExist(cp(), "422", 2) = True Then
         If MsgBox("本案若提出為「超級」加速審查，答辯期限應設定為2個月而非3個月，否則案件將退回一般審查。" & vbCrLf & vbCrLf & "是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion, "「超級」加速審查提醒") = vbNo Then
            Exit Function
         End If
      End If
   End If
   'end 2024/4/18
   
   'Add by Morgan 2010/2/8
   '來函性質為1801或1802時控制1.對造號數不可空白 2.對造名稱(中,英,日)不可全部空白
   If txtCaseField(0) = "1801" Or txtCaseField(0) = "1802" Then
      If RTrim(txtCaseField(15)) = "" Then
         SSTab1.Tab = 0
         MsgBox "對造號數不可空白！", vbExclamation
         txtCaseField(15).SetFocus
         Exit Function
      ElseIf RTrim(txtCaseField(21) & txtCaseField(22) & txtCaseField(23)) = "" Then
         SSTab1.Tab = 1
         MsgBox "對造名稱不可空白！", vbExclamation
         txtCaseField(21).SetFocus
         Exit Function
      Else
         PUB_ChkCustNameExist txtCaseField(21), txtCaseField(22), txtCaseField(23)
      End If
   End If
   
   'Add by Morgan 2011/4/20
   'EPC收到檢索報告時要更新實審及指定費的期限(准駁通知日+6個月)
   If field(9) = "221" And txtCaseField(0) = "1209" And txtCaseField(4) = "" Then
      MsgBox "請輸入准駁通知日以便計算實審及指定費的期限！"
      txtCaseField(4).SetFocus
      Exit Function
   End If
   
   'Added by Morgan 2015/6/23
   '1001,1002,1202,1209,1802,1807,1809,1810 E化提醒
   If InStr("1001,1002,1202,1221,1209,1802,1807,1809,1810", txtCaseField(0)) > 0 Then
      If PUB_GetEMailFlag(field(1) & field(2) & field(3) & field(4), , , bPaper) = True And bPaper = False Then
         MsgBox "E化案件，不印前案!!", vbExclamation
      End If
   End If
   'end 2015/6/23
   
   'Added by Lydia 2016/12/26 英國核准先行通知(1005),法定期限及本所期限欄不可空白
   If field(9) = "201" And txtCaseField(0) = "1005" And (txtCaseField(3) = "" Or txtCaseField(2) = "") Then
      MsgBox "法定期限及本所期限欄不可空白!!", vbExclamation
      If txtCaseField(3) = "" Then
         txtCaseField(3).SetFocus
      Else
         txtCaseField(2).SetFocus
      End If
      Exit Function
   End If
   'end 2016/12/26
   
   'Added by Morgan 2020/3/9
   '美國通知要求選取1206提醒報價
   If field(9) = "101" And txtCaseField(0) = "1206" Then
      'Modified by Morgan 2022/9/27 刪掉IDS--陳玫音,郭雅娟
      If PUB_ChkCPExist(cp(), "106") = True Then
         'strExc(0) = "分割、IDS及主張優先權費用"
         strExc(0) = "分割及主張優先權費用"
      Else
         'strExc(0) = "分割及IDS費用"
         strExc(0) = "分割費用"
      End If
      
      SSTab1.Tab = 1
      '報價備註
      If txtCaseField(39) = "" Then
         txtCaseField(39).SetFocus
         MsgBox "請報" & strExc(0), vbExclamation + vbCritical, "報價備註檢查"
         Exit Function
      Else
         If MsgBox("是否已輸入" & strExc(0), vbYesNo + vbExclamation + vbDefaultButton2, "報價備註檢查") = vbNo Then
            txtCaseField(39).SetFocus
            Exit Function
         End If
      End If
   End If
   'end 2020/3/9
               
   'Added by Morgan 2018/7/10 CFP電子化
   If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) Then
      'Modified by Morgan 2021/2/26 +審查意見通知函 1202
      If txtFiles = "" And (txtCaseField(0) = "1002" Or txtCaseField(0) = "1006" Or txtCaseField(0) = "1202" Or txtCaseField(0) = "1209") Then
         Me.SSTab1.Tab = 0 'Added by Morgan 2024/6/11
         MsgBox "請輸入引證前案檔案數量!!", vbExclamation
         txtFiles.SetFocus
         Exit Function
      End If
   End If
   'end 2018/7/10
   
   'Added by Morgan 2020/7/30
   '若為來函期限2次確認退回時需檢查法限是否一致
   If m_strIR01 <> "" Then
      If PUB_ChkReKeyInOk(m_strIR01, m_strIR02, m_strIR03, m_strIR04, txtCaseField(3).Text, m_bolReKeyInOK) = False Then
         txtCaseField(3).SetFocus
         Exit Function
      End If
   End If
   'end 2020/7/30
   
'Removed by Morgan 2019/6/24 取消,因報價定稿一般都非當日列印,改CFP程序可自行點選王副總案件判發--郭
'   'Added by Morgan 2018/11/12
'   If m_bolJudgerAlert And Text37.Enabled Then
'      If Text37 = "71011" Then
'         MsgBox "王副總及郭經理同時請假，請改判發人為職代！", vbExclamation
'         Text37.SetFocus
'         Exit Function
'      ElseIf Text37 = strUserNum Then
'         MsgBox "判發人不可為自己！", vbExclamation
'         Text37.SetFocus
'         Exit Function
'      End If
'   End If
'   'end 2018/11/12
'end 2019/6/24
   
   'Added by Morgan 2021/6/3
   m_bolAdd217BCP = False
   If txtCaseField(0) = "1236" And field(9) = "019" And txtCaseField(1) = "217" Then
      If cp(10) = "101" And DBDATE(cp(5)) > "19221111" And DBDATE(cp(5)) < "20210519" Then
         If MsgBox("本案發明申請為 2021/5/19 以前收文，申請費內含公開費，將設定【不通知】客戶並自動內部收文【公開費】！" & vbCrLf & vbCrLf & "是否確定要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            txtCaseField(5) = "N"
            m_bolAdd217BCP = True
         Else
            Exit Function
         End If
         
      ElseIf txtCaseField(5) = "" Then
         If Trim(txtCaseField(6)) = "" Then
            MsgBox "【公開費】欄位不可空白！"
            txtCaseField(6).SetFocus
            Exit Function
         ElseIf Trim(txtCaseField(8)) = "" Then
            MsgBox "【點數】欄位不可空白！"
            txtCaseField(8).SetFocus
            Exit Function
         End If
      End If
   End If
   'end 2021/6/3
   
   'Added by Morgan 2021/11/4
   '印尼發明/新型核准
   If m_IDNGrant Then
      If txtCaseField(6).Enabled And Trim(txtCaseField(6)) = "" Then
         MsgBox "請輸入年費金額！", vbCritical
         txtCaseField(6).SetFocus
         Exit Function
      End If
      If txtCaseField(8).Enabled And Trim(txtCaseField(8)) = "" Then
         MsgBox "請輸入年費點數！", vbCritical
         txtCaseField(8).SetFocus
         Exit Function
      End If
      strExc(0) = "select * from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06 is null and np07='605'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "下一程序已有年費期限，請確認是否重複輸入！", vbCritical
         Exit Function
      End If
   End If
   'end 2021/11/4
   
   bolDo = True
   TxtValidate = True
End Function

'Add By Cheng 2003/09/05
Private Function blnUpdateNP(strCaseNo As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strPA08 As String '專利種類
Dim strPA09 As String '申請國家
Dim strPA16 As String '目前准駁
Dim strPA20 As String '准駁通知日

   blnUpdateNP = False
   StrSQLa = "Select * From Patent Where " & ChgPatent(strCaseNo)
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       strPA08 = "" & rsA("PA08").Value
       strPA09 = "" & rsA("PA09").Value
       strPA16 = "" & rsA("PA16").Value
       strPA20 = "" & rsA("PA20").Value
       If rsA.State <> adStateClosed Then rsA.Close
       Set rsA = Nothing
       StrSQLa = "Select * From Nation Where NA01='" & strPA09 & "' "
       rsA.CursorLocation = adUseClient
       rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
       If rsA.RecordCount > 0 Then
         'Modify by Morgan 2006/9/19 核准日起算或准後繳的要新增下一程序
           Select Case strPA08
           Case "1" '發明
               '若核准日起算或准後繳的
               If "" & rsA("NA06").Value = "4" Or "" & rsA("NA56").Value = "Y" Then
                   '新增下一程序
                   blnUpdateNP = True
                   If rsA.State <> adStateClosed Then rsA.Close
                   Set rsA = Nothing
                   Exit Function
               Else
                   If rsA.State <> adStateClosed Then rsA.Close
                   Set rsA = Nothing
                   GoTo CompNextDate
               End If
           Case "2" '新型
               '若核准日起算或准後繳的
               If "" & rsA("NA08").Value = "4" Or "" & rsA("NA57").Value = "Y" Then
                   '新增下一程序
                   blnUpdateNP = True
                   If rsA.State <> adStateClosed Then rsA.Close
                   Set rsA = Nothing
                   Exit Function
               Else
                   If rsA.State <> adStateClosed Then rsA.Close
                   Set rsA = Nothing
                   GoTo CompNextDate
               End If
           Case "3" '設計
               '若核准日起算或准後繳的
               If "" & rsA("NA10").Value = "4" Or "" & rsA("NA58").Value = "Y" Then
                   '新增下一程序
                   blnUpdateNP = True
                   If rsA.State <> adStateClosed Then rsA.Close
                   Set rsA = Nothing
                   Exit Function
               Else
                   If rsA.State <> adStateClosed Then rsA.Close
                   Set rsA = Nothing
                   GoTo CompNextDate
               End If
           End Select
       End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   Exit Function
CompNextDate:
   
   If m_blnCompNextDate = True Then
       blnUpdateNP = False
   
ReDo:
       m_i = m_i + 1
       If m_i > UBound(m_varTemp) Then Exit Function
            
       m_dobDateAdd = m_varTemp(m_i - 1)
       
       If m_NP07 = "605" Then m_dobDateAdd = m_dobDateAdd - 1 'Add by Morgan 2013/6/20 年費要減一年
            
       '法定期限
       m_strDate = CompDate(0, m_dobDateAdd, m_strStartDate)
       
       'Modify by Morgan 2004/9/29
       '法定期限不必減一天,
       'm_strDate = CompDate(2, -1, m_strDate)
       '本所期限
       'm_strDate1 = CompDate(1, -1, m_strDate)
       '本所期限改抓共用
       strExc(1) = field(1)
       strExc(2) = field(9)
       strExc(3) = TransDate(m_strDate, 2)
       GetCtrlDT strExc
       m_strDate1 = DBDATE(strExc(0))
       '2004/9/29 end
       
       '若法定期限年小於系統年
       If Left(m_strDate, 4) < Left(strSrvDate(1), 4) Then
           '93.12.22 MODIFY BY SONIA 改掛畫面上之期限
           '重算
           'GoTo ReDo
           m_strDate1 = TransDate(txtCaseField(2), 2)
           m_strDate = TransDate(txtCaseField(3), 2)
           blnUpdateNP = True
           Exit Function
           '93.12.22 END
       '若法定期限年大於等於系統年
       Else
           blnUpdateNP = True
           Exit Function
       End If
   End If

End Function

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


Private Sub txtPA14_GotFocus()
   TextInverse txtPA14
End Sub

Private Sub txtPA14_Validate(Cancel As Boolean)
   If IsEmptyText(txtPA14) = False Then
      If CheckIsTaiwanDate(txtPA14) = False Then
         SSTab1.Tab = 1
         Cancel = True
      End If
   End If
End Sub

Private Sub txtPA15_GotFocus()
   TextInverse txtPA15
End Sub

Private Sub txtPA22_GotFocus()
   TextInverse txtPA22
End Sub

Private Function AutoIssue() As Boolean
   '2005/4/20 MODIFY BY SONIA
   'If txtCaseField(0) = 核准 And txtCaseField(1) = "" And InStr(發明申請 & "," & 新型申請 & "," & 設計申請 & "," & 追加申請 & "," & 答辯 & "," & CIP申請 & "," & CPA申請 & "," & 再發行, cp(10)) > 0 Then
   '2005/5/19 MODIFY BY SONIA 加分割
   '2007/8/3 MODIFY BY SONA 加424請求繼續審查
   'Modify by Sonia  2013/11/15 加復審805(CFP-023821)
   'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
   'Modified by Lydia 2016/08/27 +438 再考量試行計畫(AFCP2.0)
   If txtCaseField(0) = 核准 And txtCaseField(1) = "" And InStr(發明申請 & "," & 新型申請 & "," & 設計申請 & "," & 追加申請 & "," & 答辯 & "," & CIP申請 & "," & CPA申請 & "," & 訴願 & "," & 分割 & "," & 再發行 & ",424,805,126,438", cp(10)) > 0 Then
      AutoIssue = True
   Else
      AutoIssue = False
   End If
End Function

'Removed by Morgan 2020/9/28 取消(不會通知核准,直接輸證書)--玫音
'Private Sub txtPublic_GotFocus()
'   TextInverse txtPublic
'   'edit by nickc 2007/07/11 切換輸入法改用API
'   'txtPublic.IMEMode = 2
'   CloseIme
'End Sub

'Private Sub txtPublic_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'      Beep
'   End If
'End Sub
'end 2020/9/28

'Add by Morgan 2005/5/20 檢查是否下一程序有答辯不續辦
Private Function IsNP107N() As Boolean
On Error GoTo ErrHnd
   m_strNP07 = "" 'Added by Morgan 2013/4/18
   'Modify by Morgan 2006/6/2
   'Modify by Morgan 2006/11/10 加修正204不續辦
   '2007/8/3 MODIFY BY SONIA 加請求繼續審查424
   '2008/11/28 MODIFY BY SONIA 改判斷因107,204,424而閉卷的案件
   'strSQL = "select 1 from nextprogress where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06='N' and np07 in ('107','204')" & _
      " union select 1 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('107','204','424') and cp57>0"
   'Modified by Morgan 2013/4/18 +回覆檢索報告218
   'Modify by Sonia  2013/11/15 加復審805(CFP-023821)
   'Modified by Morgan 2016/2/16 +期末拋棄126(CFP-026475)
   'Modified by Lydia 2016/08/24 +438 再考量試行計畫(AFCP2.0)
   'Modified by Morgan 2016/10/6 +訴願501(CFP-027892)
   strSql = "select np07 from nextprogress,CASEPROGRESS C2 where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' and np06='N' and np07 in ('107','204','218','438','501') AND NP01=C2.CP43 AND NP02=C2.CP01 AND NP03=C2.CP02 AND NP04=C2.CP03 AND NP05=C2.CP04 AND '913'=C2.CP10 AND NP11=C2.CP05" & _
      " union select c1.cp10 from caseprogress C1,CASEPROGRESS C2 where C1.cp01='" & cp(1) & "' and C1.cp02='" & cp(2) & "' and C1.cp03='" & cp(3) & "' and C1.cp04='" & cp(4) & "' and C1.cp10 in ('107','204','126','218','424','805','438','501') and C1.cp57>0 AND C1.CP09=C2.CP43 AND C1.CP01=C2.CP01 AND C1.CP02=C2.CP02 AND C1.CP03=C2.CP03 AND C1.CP04=C2.CP04 AND '913'=C2.CP10 AND C1.CP57=C2.CP05"
   '2008/11/28 END
   'end 2006/6/2
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         IsNP107N = True
         m_strNP07 = .Fields(0) 'Added by Morgan 2013/4/18
      End If
   End With
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Add by Morgan 2008/5/12
Private Sub StartLetter1(ByVal ET03 As String, Optional strNP01 As String, Optional strNP22 As String)
Dim strTxt(1 To 99) As String, iStep As Integer, strTmp As Variant
Dim strTemp1 As String, strStartDate As String, strTemp As Variant
Dim bolTmp As Boolean, StrExt1 As String, StrExt2 As String, i As Integer
Dim iEPC As Integer 'EPC 指定國家順序
Dim iPos As Integer '字元搜尋位置
Dim Jjj As Integer
   
   Jjj = 1
   
   'Added by Morgan 2023/12/11
   If m_USCaseNo <> "" Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'美國案本所案號','" & m_USCaseNo & "')"
      Jjj = Jjj + 1
      
      If Val(txtIDSFee(1)) > 0 Then
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV08) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'IDS報價1','" & txtIDSFee(1) & "','Y','N')"
         Jjj = Jjj + 1
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'IDS報價1點數','" & txtIDSPt(1) & "')"
         Jjj = Jjj + 1
      End If
      
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV08) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'IDS報價2','" & txtIDSFee(2) & "','Y','N')"
      Jjj = Jjj + 1
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'IDS報價2點數','" & txtIDSPt(1) & "')"
      Jjj = Jjj + 1
   End If
   'end 2023/12/11
   
   'Added by Morgan 2023/6/12
   If field(9) = "023" And txtCaseField(1) = "601" Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'俄羅斯才印','♀','')"
         Jjj = Jjj + 1
   End If
   'end 2023/6/12
   
   If Val(txtCaseField(6)) > 0 Then
      'Added by Morgan 2021/11/4
      '印尼發明/新型核准
      If m_IDNGrant Then
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'印尼案要印','♀','')"
         Jjj = Jjj + 1
         
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'下一程序','605')"
         Jjj = Jjj + 1
      
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'第幾年','" & Text5(0) & "-" & Text5(1) & "','')"
         Jjj = Jjj + 1
         
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'年費法定期限','" & m_strDate & "','')"
         Jjj = Jjj + 1
         
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'年費本所期限','" & m_strDate1 & "','')"
         Jjj = Jjj + 1
         
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'年費','" & Val(txtCaseField(6)) & "','Y')"
         Jjj = Jjj + 1
         
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'年費點數','" & Val(txtCaseField(8)) & "','')"
         Jjj = Jjj + 1
      Else
      'end 2021/11/4
      
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'領證費','" & Val(txtCaseField(6)) & "','Y')"
         Jjj = Jjj + 1
         
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'領證費點數','" & Val(txtCaseField(8)) & "','')"
         Jjj = Jjj + 1
      End If
   End If
   
   
   If Val(txtCaseField(7)) > 0 Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV08) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'製圖費','" & Val(txtCaseField(7)) & "','Y','N')"
      Jjj = Jjj + 1
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'製圖費點數','" & Val(txtCaseField(7)) / 1000 & "','')"
      Jjj = Jjj + 1
   End If
   'Add by Morgan 2008/11/17
   If Val(txtCaseField(31)) > 0 Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV08) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'面詢費','" & Val(txtCaseField(31)) & "','Y','N')"
      Jjj = Jjj + 1
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'面詢費點數','" & Val(txtCaseField(32)) & "','')"
      Jjj = Jjj + 1
   End If
   
   If Val(txtCaseField(33)) > 0 Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV08) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'修正費','" & Val(txtCaseField(33)) & "','Y','N')"
      Jjj = Jjj + 1
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'修正費點數','" & Val(txtCaseField(34)) & "','')"
      Jjj = Jjj + 1
   End If
   
   If Val(txtCaseField(31)) + Val(txtCaseField(33)) > 0 Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'修正費+面詢費','" & Val(txtCaseField(31)) + Val(txtCaseField(33)) & "','')"
      Jjj = Jjj + 1
   End If
   
   'Memo by Morgan 2021/8/17 例外欄位名稱不可改，接洽單會用此名稱來檢查是否有跟客戶收其他費用(領證費<>費用總計)
   strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
      "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用總計','" & Val(txtCaseField(6)) + Val(txtCaseField(7)) + Val(txtCaseField(31)) + Val(txtCaseField(33)) & "','')"
   Jjj = Jjj + 1
   
   strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
      "VALUES ('" & strNP01 & "'," & strNP22 & ",'點數合計','" & Val(txtCaseField(8)) + Val(txtCaseField(7)) / 1000 + Val(txtCaseField(32)) + Val(txtCaseField(34)) & "','')"
   Jjj = Jjj + 1
   'end 2008/11/17
   
   If Val(Me.txtCaseField(10)) > 0 Then
      'Modify by Morgan 2008/11/17 智權人員報價畫面也要帶讓渡費,點數固定為5
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV08) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'讓渡費','" & Val(txtCaseField(10)) & "','Y','N')"
      Jjj = Jjj + 1
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'讓渡費點數','5','')"
      Jjj = Jjj + 1
   End If
   
   'Added by Morgan 2020/12/24
   If m_bolIDSPrice = True Then
      If m_bolRCE = True Then
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV08) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'RCE報價','" & Val(txtCaseField(42)) & "','Y','N')"
         Jjj = Jjj + 1
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'RCE報價點數','" & Val(txtCaseField(43)) & "','')"
         Jjj = Jjj + 1
      
         strExc(1) = "RCE提醒"
         
      Else
         strExc(1) = "IDS提醒"
      End If
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
        "VALUES ('" & strNP01 & "'," & strNP22 & ",'" & strExc(1) & "','♀')"
      Jjj = Jjj + 1
             
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05,LCV08) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'IDS報價','" & Val(txtCaseField(40)) & "','Y','N')"
      Jjj = Jjj + 1
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'IDS報價點數','" & Val(txtCaseField(41)) & "','')"
      Jjj = Jjj + 1
   End If
   'end 2020/12/24
               
   Select Case lblCaseField(9)
      Case "221" 'EPC
         Select Case txtCaseField(0)
            Case 核准
               '加入epc 核准國家領證費用
               If Trim(strMoneyCountry) <> "" And Trim(strMoney) <> "" Then
                  strTmp = Split(strMoneyCountry, ",")
                  strTemp = Split(strMoney, ",")
                  StrExt1 = ""
                  StrExt2 = ""
                  For i = 0 To UBound(strTmp)
                     If Val(strTemp(i)) <> 0 Then
                        '指定國家抓中文
                        strExc(1) = GetNationName(strTmp(i), 0)
                        StrExt2 = str(Val(StrExt2) + Val(strTemp(i)))
                        '註冊費
                        strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) VALUES ('" & strNP01 & "'," & strNP22 & ",'" & strExc(1) & "註冊費','" & strTemp(i) & "','Y')"
                        Jjj = Jjj + 1
                        '註冊費點數
                        'Modified by Morgan 2012/12/18 改預設8點
                        strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04)  select '" & strNP01 & "'," & strNP22 & ",'" & strExc(1) & "註冊費點數',nvl(max(YF06/1000),8) from patentyearfee where yf01='" & strTmp(i) & "' and yf02='1' and yf03='Y00000000' and yf04='224' and yf05='1'"
                        Jjj = Jjj + 1
                     End If
                  Next i
                  
                  strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                     "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用合計','" & Val(txtCaseField(6)) + Val(StrExt2) & "','')"
                  Jjj = Jjj + 1
                  
                  'EPC其他成員國
                  If m_bolEPC7Up = True Then
                     StrExt2 = PUB_GetNationName(m_strRestEPCMember)
                     StrExt2 = Replace(StrExt2, ",", "、")
                     i = 0: iPos = 0
                     Do
                        iPos = i
                        i = i + 1
                        i = InStr(i, StrExt2, "、")
                     Loop While i > 0
                     If iPos > 0 Then StrExt2 = Left(StrExt2, iPos - 1) & "及" & Mid(StrExt2, iPos + 1)
                     
                     strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                        "VALUES ('" & strNP01 & "'," & strNP22 & ",'EPC其他成員國','" & StrExt2 & "')"
                     Jjj = Jjj + 1
                     
                  End If
               End If
         End Select
         
      Case "101" '美國
          'Add by Morgan 2005/11/11
          If CheckStr(txtCaseField(4)) <> "" Then
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
               "VALUES ('" & strNP01 & "'," & strNP22 & ",'准駁日','" & DBDATE(txtCaseField(4)) & "')"
             Jjj = Jjj + 1
          End If
                    
          If CheckStr(txtCaseField(6)) <> "" Or CheckStr(txtCaseField(7)) <> "" Then
               strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用合計','" & Val(txtCaseField(6)) + Val(txtCaseField(7)) & "','')"
               Jjj = Jjj + 1
          End If
         
      Case Else '其他國家
      
         Select Case ET03
            Case "14"
               '形式審查合格通知
               Select Case lblCaseField(9)
                  Case "042"
                     If m_PA08 = "3" Then  '越南設計
                        strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                           "VALUES ('" & strNP01 & "'," & strNP22 & ",'列印備註','公開後6個月')"
                     Else
                        strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                           "VALUES ('" & strNP01 & "'," & strNP22 & ",'列印備註','之後')"
                     End If
                     Jjj = Jjj + 1
                  Case Else
                     strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                        "VALUES ('" & strNP01 & "'," & strNP22 & ",'列印備註','之後')"
                     Jjj = Jjj + 1
              End Select

            Case Else
            
               'Add by Morgan 2005/2/17 歐盟設計用
               'Removed by Morgan 2020/9/28 取消(不會通知核准,直接輸證書)--玫音
               'If txtPublic.Text <> "" Then
               '   strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
               '         "VALUES ('" & strNP01 & "'," & strNP22 & ",'公告與否','" & IIf(txtPublic.Text = "1", "已", "將") & "')"
               '   Jjj = Jjj + 1
               'End If
               'end 2020/9/28
               
         End Select
         
         '一案兩請
         If m_bolIsDualApp = True Then
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
               "VALUES ('" & strNP01 & "'," & strNP22 & ",'放棄專利權費','" & Format(txtAbandonFee) & "','')"
            Jjj = Jjj + 1
            
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04,LCV05) " & _
               "VALUES ('" & strNP01 & "'," & strNP22 & ",'費用合計','" & Format(Val(txtAbandonFee) + Val(txtCaseField(6))) & "','')"
            Jjj = Jjj + 1
         
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'新型專利號數','" & m_stCertNo & "')"
            Jjj = Jjj + 1
            
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'新型本所號','" & m_stCaseNo & "')"
            Jjj = Jjj + 1
            
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'新型申請案號','" & m_stAppNo & "')"
            Jjj = Jjj + 1
            
            strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
                  "VALUES ('" & strNP01 & "'," & strNP22 & ",'新型案件名稱','" & m_stCaseName & "')"
            Jjj = Jjj + 1
         End If
     End Select
  
   If txtCaseField(30) <> "" Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'約定期限','" & DBDATE(txtCaseField(30)) & "')"
      Jjj = Jjj + 1
   End If
   
   If CheckStr(txtCaseField(2)) <> "" Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'本所期限','" & DBDATE(txtCaseField(2)) & "')"
      Jjj = Jjj + 1
   End If
   
   If CheckStr(txtCaseField(3)) <> "" Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'法定期限','" & DBDATE(txtCaseField(3)) & "')"
      Jjj = Jjj + 1
   End If
         
   If txtCaseField(1) <> "" Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'下一程序','" & txtCaseField(1) & "')"
      Jjj = Jjj + 1
   End If
   
   If m_strRetSheet2NP07 <> "" Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'下一程序2','" & m_strRetSheet2NP07 & "')"
      Jjj = Jjj + 1
   End If
   
   If lblNextCaseProperty <> "" Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'下一程序名稱','" & lblNextCaseProperty & "')"
      Jjj = Jjj + 1
   End If
                  
   '美國讓渡/繼承登記號
   If txtCaseField(11) <> "" Then
      i = InStr(1, txtCaseField(11), "/", 1)
      If i > 0 Then
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'補文件 V 1','" & Mid(txtCaseField(11), 1, i - 1) & "')"
         Jjj = Jjj + 1
         
         strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
            "VALUES ('" & strNP01 & "'," & strNP22 & ",'補文件 V 2','" & Right(txtCaseField(11), Val(Len(txtCaseField(11)) - i)) & "')"
         Jjj = Jjj + 1
      End If
   End If
   
   If m_PA09 = 美國國家代號 And (cp(10) = 繼承 Or cp(10) = 授權 Or cp(10) = 變更) And txtCaseField(0).Text = 核准 Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'相關收文性質','" & GetCaseTypeName(cp(1), cp(10)) & "')"
      Jjj = Jjj + 1
   End If
   
   If Text5(0).Enabled = True And Text5(0) <> "" Then
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'含年費','(含第" & Text5(0) & "至" & Text5(1) & "年年費)')"
      Jjj = Jjj + 1
      'Add by Morgan 2009/2/3
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'年費起年','" & Text5(0) & "')"
      Jjj = Jjj + 1
      
      strTxt(Jjj) = "INSERT INTO LetterCacheVar (LCV01,LCV02,LCV03,LCV04) " & _
         "VALUES ('" & strNP01 & "'," & strNP22 & ",'年費迄年','" & Text5(1) & "')"
      Jjj = Jjj + 1
      
   End If
   
   If Not ClsLawExecSQL(Jjj - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Sub

'2013/8/12 add by sonia 檢查申請人是否為特殊客戶
Private Function PUB_CheckspecialCust(s_CP09 As String, s_CP10 As String) As Boolean
Dim stSQL As String, adoRst As ADODB.Recordset, intR As Integer

On Error GoTo ErrHnd

   PUB_CheckspecialCust = False
   
   'Modified by Morgan 2021/11/5 來函性質改用常數判斷
   'If InStr("1002,1006,1206,1209,1220", s_CP10) > 0 Then
   If InStr(PatentOAPtyList, s_CP10) > 0 Then
   'end 2021/11/5
   Else
      Exit Function
   End If
   
   stSQL = "select pa26,cp14,pa09,pa01,pa02,pa03,pa04 from patent,caseprogress where cp09='" & s_CP09 & "' " & _
           "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      '林特助之客戶華碩電腦X6901101
      If Left(adoRst("pa26"), 8) = "X6901101" Then PUB_CheckspecialCust = True
      '2013/9/12 雅娟要求加入義隆X2906604,慧汶要求加入新日興X1484304
      If Left(adoRst("pa26"), 8) = "X2906604" Then PUB_CheckspecialCust = True
      If Left(adoRst("pa26"), 8) = "X1484304" Then PUB_CheckspecialCust = True
      'Added by Morgan 2018/1/11 +義晶X2906606--禧佩,偉城
      If Left(adoRst("pa26"), 8) = "X2906606" Then PUB_CheckspecialCust = True
      '2015/12/22 雅娟要求加入和碩X70017
      If Left(adoRst("pa26"), 8) = "X7001700" Then PUB_CheckspecialCust = True
      'Added by Morgan 2019/1/9 +X79217 南京嵐煜生物科技有限公司---郭雅娟
      If Left(adoRst("pa26"), 8) = "X7921700" Then PUB_CheckspecialCust = True
      'Added by Morgan 2021/11/5 +X3880503 財團法人資訊工業策進會---禧佩
      If Left(adoRst("pa26"), 8) = "X3880503" Then PUB_CheckspecialCust = True
      'Added by Morgan 2021/11/5 +X53094010強茂公司---郭雅娟/吳邑君
      If Left(adoRst("pa26"), 8) = "X5309401" Then PUB_CheckspecialCust = True
   End If
   Exit Function
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function
'2013/8/12 end

'Added by Lydia 2015/04/10 呼叫~申請人國外ID資料維護
Private Sub CmdAFID03_Click(Index As Integer)

   Set frm880021.m_PrevF = Me
   frm880021.strXNo = field(26 + Index)
   frm880021.strXNation = m_PA09
   frm880021.lblTitle.Caption = "申請人" & str(Index + 1)
   frm880021.lblCust.Caption = lblAgent.Caption '客戶名稱
   frm880021.lblNation.Caption = lblNation.Caption  '國名
   frm880021.Show
End Sub
'end 2015/04/10
'Added by Morgan 2020/7/16 計算期限
Private Sub GetTime()
   Dim i As Integer
   Dim strFromDate As String '期限起算日
   Dim iDays1 As Integer, iDays2 As Integer, iDays3 As Integer, iDays4 As Integer 'Add by Morgan 2009/12/1
   
   If txtCaseField(4) <> "" Then
      '起算日=准駁通知日
      strFromDate = txtCaseField(4)
         
      '文到天數
      If Option4(0).Value = True Then
         txtCaseField(3) = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
         
      '文到月數
      ElseIf Option4(1).Value = True Then
         txtCaseField(3) = TransDate(AddMonth(strFromDate, Val(Text11)), 1)
         
      End If
      CheckKeyIn 3
   End If
End Sub

'Added by Morgan 2021/11/4
'設定印尼核准繳年費年度
Private Sub SetINDYear()
   If m_IDNGrant Then
      If field(10) <> "" Then
         If ChkDate(txtCaseField(4), False) Then
            '核准日6個月內須繳交自申請日起算累計至核准日次年之年費，之後則逐年提前於屆滿前1個月(申請日)繳交
             Text5(1) = PUB_GetIDNToYear(field(10), txtCaseField(4))
             
             'Added by Morgan 2022/11/1 若核准日次年的年費期限與再次年的期限相差不超過6個月時合併通知
             '核准日次年的年費期限 strexc(1)
             strExc(1) = GetIDN1st605FeeDate(txtCaseField(4))
             '年費起算日 strexc(2)
             GetMoneyDate Val(field(8)), field(9), field(), strExc(2)
             '再次年法限 strexc(3)
             strExc(3) = GetFeeNextDate(strExc(2), Text5(1), field(9), field(8))
             '核准日次年的年費期限+6個月
             strExc(4) = CompDate(1, 6, strExc(1))
             If strExc(4) >= strExc(3) Then
               Text5(1) = Val(Text5(1)) + 1
             End If
             'end 2022/11/1
         End If
      End If
   End If
End Sub

'Added by Morgan 2025/5/13
Private Function ChkIDS(pUSList As String) As Boolean
   Dim arrUSNO() As String, arrNo() As String, pa(4) As String
   Dim ii As Integer, jj As Integer
   Dim intQ As Integer, stSQL As String, stMsg As String
   Dim rsQuery As ADODB.Recordset
   Dim bolA As Boolean, bolB As Boolean
   
   arrUSNO = Split(pUSList, "、")
   For jj = LBound(arrUSNO) To UBound(arrUSNO)
      stMsg = stMsg & "本案有相關美國案(" & arrUSNO(jj) & ")"
      arrNo = Split(arrUSNO(jj), "-")
      
      Erase pa()
      intQ = 1
      For ii = LBound(arrNo) To UBound(arrNo)
         pa(intQ) = arrNo(ii)
         intQ = intQ + 1
      Next
      If pa(3) = "" Then pa(3) = "0"
      If pa(4) = "" Then pa(4) = "00"
      
      
      stSQL = "select cp09,cp14,cp07,pa05,cu04 from caseprogress,patent,customer" & _
         " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
         " and cp10='214' and cp27||cp57 is null" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         stMsg = stMsg & "已收文IDS尚未發文，請確認加上本案的前案是否有超過25件，若有，則請報差價給工程師！"
         bolA = True
      Else
         stMsg = stMsg & "，請一併報IDS！"
         bolB = True
      End If
      stMsg = stMsg & vbCrLf & vbCrLf
   Next
   If bolA And Not bolB Then
      If MsgBox(stMsg & "是否要報差價給工程師？", vbQuestion + vbYesNo + vbDefaultButton1, "IDS報價提醒") = vbYes Then
         ChkIDS = True
      End If
   Else
      MsgBox stMsg, vbExclamation, "IDS報價提醒"
      ChkIDS = True
   End If
End Function
