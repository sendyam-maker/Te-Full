VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030207 
   BorderStyle     =   1  '單線固定
   Caption         =   "委任狀中譯文"
   ClientHeight    =   7125
   ClientLeft      =   825
   ClientTop       =   975
   ClientWidth     =   8730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8730
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   4
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   4
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "申請人編號："
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Value           =   -1  'True
      Width           =   1275
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5680
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10028
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "案件資料"
      TabPicture(0)   =   "frm030207.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3(4)"
      Tab(0).Control(1)=   "Label3(3)"
      Tab(0).Control(2)=   "Label3(2)"
      Tab(0).Control(3)=   "Label3(1)"
      Tab(0).Control(4)=   "Label4(4)"
      Tab(0).Control(5)=   "Label4(3)"
      Tab(0).Control(6)=   "Label4(2)"
      Tab(0).Control(7)=   "Label4(1)"
      Tab(0).Control(8)=   "Label3(0)"
      Tab(0).Control(9)=   "Label4(0)"
      Tab(0).Control(10)=   "Label4(5)"
      Tab(0).Control(11)=   "Label4(6)"
      Tab(0).Control(12)=   "Label4(7)"
      Tab(0).Control(13)=   "Label4(8)"
      Tab(0).Control(14)=   "Text1(5)"
      Tab(0).Control(15)=   "Text1(6)"
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "代表人１"
      TabPicture(1)   =   "frm030207.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo2(0)"
      Tab(1).Control(1)=   "txtCaseField(0)"
      Tab(1).Control(2)=   "txtCaseField(1)"
      Tab(1).Control(3)=   "txtCaseField(2)"
      Tab(1).Control(4)=   "Combo2(1)"
      Tab(1).Control(5)=   "txtCaseField(3)"
      Tab(1).Control(6)=   "txtCaseField(4)"
      Tab(1).Control(7)=   "txtCaseField(5)"
      Tab(1).Control(8)=   "Combo2(2)"
      Tab(1).Control(9)=   "txtCaseField(6)"
      Tab(1).Control(10)=   "txtCaseField(7)"
      Tab(1).Control(11)=   "txtCaseField(8)"
      Tab(1).Control(12)=   "Combo2(3)"
      Tab(1).Control(13)=   "txtCaseField(9)"
      Tab(1).Control(14)=   "txtCaseField(10)"
      Tab(1).Control(15)=   "txtCaseField(11)"
      Tab(1).Control(16)=   "Combo2(4)"
      Tab(1).Control(17)=   "txtCaseField(12)"
      Tab(1).Control(18)=   "txtCaseField(13)"
      Tab(1).Control(19)=   "txtCaseField(14)"
      Tab(1).Control(20)=   "Label5(8)"
      Tab(1).Control(21)=   "Label5(7)"
      Tab(1).Control(22)=   "Label5(6)"
      Tab(1).Control(23)=   "Label5(5)"
      Tab(1).Control(24)=   "Label5(4)"
      Tab(1).Control(25)=   "Label5(3)"
      Tab(1).Control(26)=   "Label14(1)"
      Tab(1).Control(27)=   "Label18(2)"
      Tab(1).Control(28)=   "Label5(24)"
      Tab(1).Control(29)=   "Label5(25)"
      Tab(1).Control(30)=   "Label5(26)"
      Tab(1).Control(31)=   "Label5(27)"
      Tab(1).Control(32)=   "Label5(28)"
      Tab(1).Control(33)=   "Label5(29)"
      Tab(1).Control(34)=   "Label14(2)"
      Tab(1).Control(35)=   "Label18(1)"
      Tab(1).Control(36)=   "Label5(33)"
      Tab(1).Control(37)=   "Label5(34)"
      Tab(1).Control(38)=   "Label5(35)"
      Tab(1).Control(39)=   "Label14(3)"
      Tab(1).ControlCount=   40
      TabCaption(2)   =   "代表人２"
      TabPicture(2)   =   "frm030207.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label14(4)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label5(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label5(9)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label18(3)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label14(5)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label5(10)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label5(11)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label5(12)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label5(13)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label5(14)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label5(15)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label18(4)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label14(6)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label5(16)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label5(17)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Label5(18)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Label5(19)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Label5(20)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Label5(21)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "txtCaseField(29)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txtCaseField(28)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "txtCaseField(27)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Combo2(9)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "txtCaseField(26)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "txtCaseField(25)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "txtCaseField(24)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Combo2(8)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "txtCaseField(23)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "txtCaseField(22)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "txtCaseField(21)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Combo2(7)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "txtCaseField(20)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "txtCaseField(19)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "txtCaseField(18)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Combo2(6)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "txtCaseField(17)"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "txtCaseField(16)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "txtCaseField(15)"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Combo2(5)"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).ControlCount=   40
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   6
         Left            =   -73470
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2700
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   5
         Left            =   -73470
         MaxLength       =   8
         TabIndex        =   6
         Top             =   3120
         Width           =   1455
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   0
         Left            =   -73425
         TabIndex        =   14
         Top             =   420
         Width           =   6135
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   0
         Left            =   -73425
         TabIndex        =   15
         Top             =   689
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   1
         Left            =   -73425
         TabIndex        =   16
         Top             =   943
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   2
         Left            =   -73425
         TabIndex        =   17
         Top             =   1197
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   1
         Left            =   -73425
         TabIndex        =   18
         Top             =   1451
         Width           =   6135
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   3
         Left            =   -73425
         TabIndex        =   19
         Top             =   1720
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   4
         Left            =   -73425
         TabIndex        =   20
         Top             =   1974
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   5
         Left            =   -73425
         TabIndex        =   21
         Top             =   2228
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   2
         Left            =   -73425
         TabIndex        =   22
         Top             =   2482
         Width           =   6135
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   6
         Left            =   -73425
         TabIndex        =   23
         Top             =   2751
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   7
         Left            =   -73425
         TabIndex        =   24
         Top             =   3005
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   8
         Left            =   -73425
         TabIndex        =   25
         Top             =   3259
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   3
         Left            =   -73425
         TabIndex        =   26
         Top             =   3513
         Width           =   6135
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   9
         Left            =   -73425
         TabIndex        =   27
         Top             =   3782
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   10
         Left            =   -73425
         TabIndex        =   28
         Top             =   4036
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   11
         Left            =   -73425
         TabIndex        =   29
         Top             =   4290
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   4
         Left            =   -73425
         TabIndex        =   30
         Top             =   4544
         Width           =   6135
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   12
         Left            =   -73425
         TabIndex        =   31
         Top             =   4813
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   13
         Left            =   -73425
         TabIndex        =   32
         Top             =   5067
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   14
         Left            =   -73425
         TabIndex        =   33
         Top             =   5310
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   285
         Index           =   5
         Left            =   1575
         TabIndex        =   54
         Top             =   420
         Width           =   6135
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10821;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   15
         Left            =   1575
         TabIndex        =   55
         Top             =   679
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   16
         Left            =   1575
         TabIndex        =   56
         Top             =   938
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   17
         Left            =   1575
         TabIndex        =   57
         Top             =   1197
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   285
         Index           =   6
         Left            =   1575
         TabIndex        =   58
         Top             =   1456
         Width           =   6135
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10821;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   18
         Left            =   1575
         TabIndex        =   59
         Top             =   1715
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   19
         Left            =   1575
         TabIndex        =   60
         Top             =   1974
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   20
         Left            =   1575
         TabIndex        =   61
         Top             =   2233
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   285
         Index           =   7
         Left            =   1575
         TabIndex        =   62
         Top             =   2492
         Width           =   6135
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10821;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   21
         Left            =   1575
         TabIndex        =   63
         Top             =   2751
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   22
         Left            =   1575
         TabIndex        =   64
         Top             =   3010
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   23
         Left            =   1575
         TabIndex        =   65
         Top             =   3269
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   285
         Index           =   8
         Left            =   1575
         TabIndex        =   66
         Top             =   3528
         Width           =   6135
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10821;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   24
         Left            =   1575
         TabIndex        =   67
         Top             =   3787
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   25
         Left            =   1575
         TabIndex        =   68
         Top             =   4046
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   26
         Left            =   1575
         TabIndex        =   69
         Top             =   4305
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   285
         Index           =   9
         Left            =   1575
         TabIndex        =   70
         Top             =   4564
         Width           =   6135
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "10821;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   27
         Left            =   1575
         TabIndex        =   71
         Top             =   4823
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   28
         Left            =   1575
         TabIndex        =   72
         Top             =   5082
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   285
         Index           =   29
         Left            =   1575
         TabIndex        =   73
         Top             =   5340
         Width           =   6135
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "10821;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "(可輸入10個中文字)"
         Height          =   255
         Index           =   8
         Left            =   -69870
         TabIndex        =   107
         Top             =   2708
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "(西元年月日)"
         Height          =   255
         Index           =   7
         Left            =   -71850
         TabIndex        =   106
         Top             =   3128
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "地　　點："
         Height          =   255
         Index           =   6
         Left            =   -74400
         TabIndex        =   105
         Top             =   2708
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "簽署日期："
         Height          =   255
         Index           =   5
         Left            =   -74400
         TabIndex        =   104
         Top             =   3128
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "申請人1："
         Height          =   255
         Index           =   0
         Left            =   -74400
         TabIndex        =   103
         Top             =   480
         Width           =   855
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Index           =   0
         Left            =   -73545
         TabIndex        =   102
         Top             =   480
         Width           =   6500
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "11465;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "申請人2："
         Height          =   255
         Index           =   1
         Left            =   -74400
         TabIndex        =   101
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "申請人3："
         Height          =   255
         Index           =   2
         Left            =   -74400
         TabIndex        =   100
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "申請人4："
         Height          =   255
         Index           =   3
         Left            =   -74400
         TabIndex        =   99
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "申請人5："
         Height          =   255
         Index           =   4
         Left            =   -74400
         TabIndex        =   98
         Top             =   2160
         Width           =   855
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Index           =   1
         Left            =   -73545
         TabIndex        =   97
         Top             =   900
         Width           =   6500
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "11465;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Index           =   2
         Left            =   -73560
         TabIndex        =   96
         Top             =   1320
         Width           =   6500
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "11465;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Index           =   3
         Left            =   -73545
         TabIndex        =   95
         Top             =   1740
         Width           =   6500
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "11465;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Index           =   4
         Left            =   -73545
         TabIndex        =   94
         Top             =   2160
         Width           =   6500
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "11465;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   8
         Left            =   -73920
         TabIndex        =   53
         Top             =   2228
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   7
         Left            =   -73920
         TabIndex        =   52
         Top             =   1974
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   6
         Left            =   -73920
         TabIndex        =   51
         Top             =   1720
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   5
         Left            =   -73920
         TabIndex        =   50
         Top             =   1197
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   4
         Left            =   -73920
         TabIndex        =   49
         Top             =   943
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   3
         Left            =   -73920
         TabIndex        =   48
         Top             =   689
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人1"
         Height          =   180
         Index           =   1
         Left            =   -74280
         TabIndex        =   47
         Top             =   420
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人2"
         Height          =   180
         Index           =   2
         Left            =   -74280
         TabIndex        =   46
         Top             =   1451
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   24
         Left            =   -73920
         TabIndex        =   45
         Top             =   4290
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   25
         Left            =   -73920
         TabIndex        =   44
         Top             =   4036
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   26
         Left            =   -73920
         TabIndex        =   43
         Top             =   3782
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   27
         Left            =   -73920
         TabIndex        =   42
         Top             =   3259
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   28
         Left            =   -73920
         TabIndex        =   41
         Top             =   3005
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   29
         Left            =   -73920
         TabIndex        =   40
         Top             =   2751
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人3"
         Height          =   180
         Index           =   2
         Left            =   -74280
         TabIndex        =   39
         Top             =   2482
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人4"
         Height          =   180
         Index           =   1
         Left            =   -74280
         TabIndex        =   38
         Top             =   3513
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   33
         Left            =   -73920
         TabIndex        =   37
         Top             =   5310
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   34
         Left            =   -73920
         TabIndex        =   36
         Top             =   5067
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   35
         Left            =   -73920
         TabIndex        =   35
         Top             =   4813
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人5"
         Height          =   180
         Index           =   3
         Left            =   -74280
         TabIndex        =   34
         Top             =   4544
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   21
         Left            =   1080
         TabIndex        =   93
         Top             =   2233
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   20
         Left            =   1080
         TabIndex        =   92
         Top             =   1950
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   19
         Left            =   1080
         TabIndex        =   91
         Top             =   1715
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   18
         Left            =   1080
         TabIndex        =   90
         Top             =   1197
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   17
         Left            =   1080
         TabIndex        =   89
         Top             =   938
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   16
         Left            =   1080
         TabIndex        =   88
         Top             =   679
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人6"
         Height          =   180
         Index           =   6
         Left            =   720
         TabIndex        =   87
         Top             =   420
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人7"
         Height          =   180
         Index           =   4
         Left            =   720
         TabIndex        =   86
         Top             =   1456
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   15
         Left            =   1080
         TabIndex        =   85
         Top             =   4357
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   14
         Left            =   1080
         TabIndex        =   84
         Top             =   4098
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   13
         Left            =   1080
         TabIndex        =   83
         Top             =   3839
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   12
         Left            =   1080
         TabIndex        =   82
         Top             =   3321
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   11
         Left            =   1080
         TabIndex        =   81
         Top             =   3062
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   10
         Left            =   1080
         TabIndex        =   80
         Top             =   2751
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人8"
         Height          =   180
         Index           =   5
         Left            =   720
         TabIndex        =   79
         Top             =   2492
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人9"
         Height          =   180
         Index           =   3
         Left            =   720
         TabIndex        =   78
         Top             =   3580
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   9
         Left            =   1080
         TabIndex        =   77
         Top             =   5392
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   2
         Left            =   1080
         TabIndex        =   76
         Top             =   5134
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   1
         Left            =   1080
         TabIndex        =   75
         Top             =   4875
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人10"
         Height          =   180
         Index           =   4
         Left            =   720
         TabIndex        =   74
         Top             =   4616
         Width           =   720
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3630
      TabIndex        =   7
      Top             =   150
      Width           =   795
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7800
      TabIndex        =   9
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Word 編輯(&W)"
      Enabled         =   0   'False
      Height          =   400
      Index           =   0
      Left            =   5715
      TabIndex        =   8
      Top             =   60
      Width           =   1920
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   3120
      MaxLength       =   2
      TabIndex        =   3
      Top             =   210
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   2
      Top             =   210
      Width           =   300
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   1
      Top             =   210
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   0
      Top             =   210
      Width           =   495
   End
   Begin MSForms.Label Label2 
      Height          =   285
      Left            =   1560
      TabIndex        =   108
      Top             =   570
      Width           =   7065
      VariousPropertyBits=   27
      Size            =   "12462;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   525
      TabIndex        =   10
      Top             =   570
      Width           =   900
   End
End
Attribute VB_Name = "frm030207"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/01 改成Form2.0; Label2、Label3(index)、txtCaseField(index)、Combo2(index)
'Create by Amy 2013/07/11
Option Explicit
Dim i As Integer, j As Integer, intA As Integer
Dim rsA As New ADODB.Recordset
Dim strSql As String
Dim IsTM As Boolean
Dim strText(0 To 3) As String
Dim bolRetry As Boolean '是否已發生錯誤且重試
Dim CountNPerson As Integer 'count 申請人非個人( CU15<>0)
Dim strSysKind As String 'Add by Amy 2013/09/05 +內商使用預帶系統別

Private Sub cmdOK_Click(Index As Integer)
    Select Case Index
        Case 0
            If TxtValidate = False Then Exit Sub
            If IsTM Then
                Call FormSave
            End If
            Call RunWordOpen
        Case 1
            Unload Me
    End Select
End Sub

Private Sub Combo2_Click(Index As Integer)
Dim strTmp As String
  If Combo2(Index) = "" Then
      For i = 0 To 2
         txtCaseField(Index * 3 + i) = ""
      Next
      Exit Sub
   End If
   
   strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
   strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
   strExc(0) = "Select " & strExc(1) & " From Customer Where " & ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      For i = 0 To 2
         
         If Not IsNull(RsTemp.Fields(i)) Then
            txtCaseField(Index * 3 + i) = RsTemp.Fields(i)
         Else
            txtCaseField(Index * 3 + i) = ""
         End If
      Next
   End If
End Sub

Private Sub Command1_Click()
Dim Cancel As Boolean
Dim strNo As String, strName As String, strCName As String
Dim CountApp As Integer '申請人個數
Dim k As Integer 'Add by Amy 2015/08/19

    SSTab1.Tab = 0
    If Option1(0).Value Then
        If Len(Trim(Text1(0))) = 0 Or Len(Text1(1)) <> 6 Then
            MsgBox "本所案號輸入錯誤，請重新輸入 !", vbCritical
            Text1(0).SetFocus
            Exit Sub
        End If
        'Modify by Amy 2013/09/05 +內商使用系統別判斷
        If Text1(0) = "FCT" Or Text1(0) = "T" Then
            If Text1(2) = "" Then Text1(2) = "0"
            If Text1(3) = "" Then Text1(3) = "00"
        Else
            MsgBox "本所案號輸入錯誤，請重新輸入 !", vbCritical
           Text1(0).SetFocus
            Exit Sub
        End If
        
    Else
        If Len(Trim(Text1(4))) = 0 Then
            MsgBox "申請人編號不可為空!", vbCritical
            Text1(0).SetFocus
            Exit Sub
        End If
    End If
   
    IsTM = False: CountApp = 0: CountNPerson = 0
    If Option1(0).Value Then '本所案號查
        IsTM = True
        strText(0) = "TM01 = '" & Text1(0) & "' And TM02='" & Text1(1) & "' And TM03='" & Text1(2) & "' And TM04='" & Text1(3) & "'"
        strSql = "Select Decode(TM05,null,Decode(TM06,null,TM07,TM06),TM05) CaseName,Nvl(TM23,'Null'),Nvl(TM78,'Null'),Nvl(TM79,'Null'),Nvl(TM80,'Null'),Nvl(TM81,'Null')," & _
                     "TM47,TM48,TM49,TM50,TM51,TM52,TM94,TM95,TM96,TM97,TM98,TM99,TM100,TM101,TM102,TM103,TM104,TM105,TM106,TM107,TM108,TM109,TM110,TM111,TM112,TM113,TM114,TM115,TM116,TM117, " & _
                     "TM24,TM82,TM83,TM84,TM85 From TradeMark Where  " & strText(0)
           
    Else '申請人編號查
        'Modified by Lydia 2019/08/15 +國別
        'strSql = "Select ' ' CaseName,Decode(CU05,null,Decode(CU04,null,CU06,CU04),CU05||CU88||CU89||CU90) CusName,Decode(CU40,null,Decode(CU39,null,CU41,CU39),CU40),Decode(CU43,null,Decode(CU42,null,CU44,CU42),CU43)," & _
                                      "Decode(CU46,null,Decode(CU45,null,CU47,CU45),CU46),Decode(CU49,null,Decode(CU48,null,CU50,CU48),CU49)," & _
                                      "Decode(CU52,null,Decode(CU51,null,CU53,CU51),CU52),Decode(CU55,null,Decode(CU54,null,CU56,CU54),CU55),CU23 Addr,CU04,CU15 From Customer " & _
                    "Where " & ChgCustomer(Text1(4))
        strSql = "Select ' ' CaseName,Decode(CU05,null,Decode(CU04,null,CU06,CU04),CU05||CU88||CU89||CU90) CusName,Decode(CU40,null,Decode(CU39,null,CU41,CU39),CU40),Decode(CU43,null,Decode(CU42,null,CU44,CU42),CU43)," & _
                                      "Decode(CU46,null,Decode(CU45,null,CU47,CU45),CU46),Decode(CU49,null,Decode(CU48,null,CU50,CU48),CU49)," & _
                                      "Decode(CU52,null,Decode(CU51,null,CU53,CU51),CU52),Decode(CU55,null,Decode(CU54,null,CU56,CU54),CU55),CU23 Addr,CU04,CU15,NA81 From Customer,Nation " & _
                    "Where " & ChgCustomer(Text1(4)) & " AND CU10=NA01(+) "
    End If
      
    strText(1) = "": strText(2) = ""
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, strSql)
     
     If intA = 1 Then
        FormClear False
         For i = 0 To 9
            Combo2(i).Clear
            Combo2(i).AddItem ""
         Next
        Label2.Caption = "" & rsA.Fields("CaseName")
        
        If IsTM Then
            '本所案號查設定顯示資料
            For i = 1 To 5
                If rsA.Fields(i) <> "Null" Then
                    CountApp = CountApp + 1
                    strNo = rsA.Fields(i)
                    strName = GetCusName(strNo, strCName)
                    'Modify by Amy 2014/08/11 解申請人名稱有全型空白造成簽署人切割有問題
                    strText(1) = strText(1) & i & "." & strCName & "&nbsp;" 'For Word申請人名稱
                    '申請人
                    Label3(i - 1) = strNo & " - " & strName
           
                    '設定代表人選項 (英->中->日)
                    strExc(0) = "Select Decode(CU40,null,Decode(CU39,null,CU41,CU39),CU40),Decode(CU43,null,Decode(CU42,null,CU44,CU42),CU43)," & _
                                      "Decode(CU46,null,Decode(CU45,null,CU47,CU45),CU46),Decode(CU49,null,Decode(CU48,null,CU50,CU48),CU49)," & _
                                      "Decode(CU52,null,Decode(CU51,null,CU53,CU51),CU52),Decode(CU55,null,Decode(CU54,null,CU56,CU54),CU55) FROM Customer Where " & ChgCustomer(strNo)
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    If intI = 1 Then
                        For j = 1 To 6
                            If IsNull(RsTemp.Fields(j - 1)) Then
                                strExc(0) = ""
                            Else
                                strExc(0) = "-" & RsTemp.Fields(j - 1)
                            End If
                            'Modify by Amy 2015/08/19 改所有下拉帶全部申請人的所有代表人-湘蘭 FCT-037789 申請人1的代表人3個
                            '原:代表人1/2下拉,帶申請人1的代表人；代表人3/4下拉,帶申請人2的代表人…
    '                        Combo2((i - 1) * 2).AddItem strNo & "-" & j & strExc(0)
    '                        Combo2((i - 1) * 2 + 1).AddItem strNo & "-" & j & strExc(0)
                            If strExc(0) <> MsgText(601) Then
                                For k = 0 To 9
                                    Combo2(k).AddItem strNo & "-" & j & strExc(0)
                                Next k
                            End If
                        Next j
                    End If
                               
                Else
                    strNo = ""
                    Label3(i - 1) = ""
                End If
            Next i
            'Modify by Amy 2014/08/11 原:Right(strText(1), Len(strText(1)) - 2)
            If CountApp = 1 Then strText(1) = Mid(strText(1), 3) 'For Word申請人名稱(只有一個申請人不需+序號)
        
           '代表人
           For i = 0 To 29
             If Not IsNull(rsA.Fields(i + 6)) Then
                  txtCaseField(i) = rsA.Fields(i + 6)
             Else
                  txtCaseField(i) = ""
             End If
          Next i
          
          'For Word 申請人地址
          CountApp = CountApp + 35
          For i = 36 To CountApp
            If CountApp = 36 Then
                strText(2) = strText(2) & IIf(IsNull(rsA.Fields(i)), "", rsA.Fields(i)) & "　"
            Else
                strText(2) = strText(2) & i - 35 & "." & IIf(IsNull(rsA.Fields(i)), "", rsA.Fields(i)) & "　"
            End If
          Next i
        
        Else '申請人編號查
            '2013/07/22 +個人判斷
            If rsA.Fields("CU15") <> "0" Then CountNPerson = CountNPerson + 1
            
            'For Word申請人名稱
            'Modify by Amy 2014/08/11 全型空白改為&nbsp;
            'strText(1) = IIf(IsNull(rsA.Fields("CU04")), "", rsA.Fields("CU04")) & "　"
            'Added by Lydia 2019/08/15 +XX商
            If "" & rsA.Fields("CU15") = "1" Then
                strText(1) = IIf(IsNull(rsA.Fields("CU04")), "", rsA.Fields("NA81") & rsA.Fields("CU04")) & "&nbsp;"
            'Added by Lydia 2019/08/28 FCT和T的自然人+XX籍(紙本)
            ElseIf "" & rsA.Fields("CU15") = "0" Then
                strText(1) = IIf(IsNull(rsA.Fields("CU04")), "", Replace(rsA.Fields("NA81"), "商", "籍") & rsA.Fields("CU04")) & "&nbsp;"
            Else
            'end 2019/08/15
                strText(1) = IIf(IsNull(rsA.Fields("CU04")), "", rsA.Fields("CU04")) & "&nbsp;"
            End If  'end 2019/08/15
            
            '申請人1
            Label3(0) = Text1(4) & " - " & rsA.Fields("CusName")
            
            '設定代表人
            For i = 0 To 4
                For j = 2 To 7
                   strNo = IIf(IsNull(rsA.Fields(j)), "", rsA.Fields(j))
                   Combo2(i * 2).AddItem Text1(4) & "-" & j - 1 & strNo
                   Combo2(i * 2 + 1).AddItem Text1(4) & "-" & j - 1 & strNo
                 Next j
            Next i
            
            'For Word申請人地址
            strText(2) = rsA.Fields("Addr") & "　"
        End If
        'Mark by Amy 2014/08/11
        'strText(1) = Left(strText(1), Len(strText(1)) - 1)
        strText(1) = Left(strText(1), Len(strText(1)) - 6)
        strText(2) = Left(strText(2), Len(strText(2)) - 1)
        cmdOK(0).Caption = IIf(IsTM = True, "Word 編輯及存檔(&W)", "Word 編輯(&W)")
        cmdOK(0).Enabled = True
     Else
        MsgBox "查無資料"
     End If
     Set rsA = Nothing
End Sub

Private Sub Form_Activate()
    Text1(1).SetFocus
End Sub

Private Sub Form_Load()
    'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    txtCaseField(0).MaxLength = Pub_MaxCEL10
    txtCaseField(1).MaxLength = Pub_MaxCEL11
    txtCaseField(3).MaxLength = Pub_MaxCEL10
    txtCaseField(4).MaxLength = Pub_MaxCEL11
    txtCaseField(6).MaxLength = Pub_MaxCEL10
    txtCaseField(7).MaxLength = Pub_MaxCEL11
    txtCaseField(9).MaxLength = Pub_MaxCEL10
    txtCaseField(10).MaxLength = Pub_MaxCEL11
    txtCaseField(12).MaxLength = Pub_MaxCEL10
    txtCaseField(13).MaxLength = Pub_MaxCEL11
    txtCaseField(15).MaxLength = Pub_MaxCEL10
    txtCaseField(16).MaxLength = Pub_MaxCEL11
    txtCaseField(18).MaxLength = Pub_MaxCEL10
    txtCaseField(19).MaxLength = Pub_MaxCEL11
    txtCaseField(21).MaxLength = Pub_MaxCEL10
    txtCaseField(22).MaxLength = Pub_MaxCEL11
    txtCaseField(24).MaxLength = Pub_MaxCEL10
    txtCaseField(25).MaxLength = Pub_MaxCEL11
    txtCaseField(27).MaxLength = Pub_MaxCEL10
    txtCaseField(28).MaxLength = Pub_MaxCEL11
    'end 2016/09/10

   MoveFormToCenter Me
   SSTab1.Tab = 0
   'Add by Amy 2013/09/05 +內商使用預帶系統別
   If Left(Pub_StrUserSt03, 1) = "F" Then
        strSysKind = "FCT"
   ElseIf Left(Pub_StrUserSt03, 2) = "P2" Then
        strSysKind = "T"
   Else
        strSysKind = "FCT"
        Text1(0).Locked = False
   End If
   FormClear True
End Sub

Private Sub FormClear(ByVal bolAll As Boolean)
 'Modified by Lydia 2021/09/01
 'Dim txt As TextBox, Lbl As LABEL
 'Dim Cmb2 As ComboBox
 Dim oObject As Control
 
  If bolAll Then
    'Modified by Lydia 2021/09/01
'    For Each txt In Text1
'      If txt.Index = 0 Then
'        txt.Text = strSysKind 'Modify by Amy 2013/09/05 +內商使用預帶系統別 原外商用帶"FCT"
'      Else
'        txt.Text = ""
'      End If
'    Next
    For Each oObject In Text1
      If oObject.Index = 0 Then
        oObject.Text = strSysKind 'Modify by Amy 2013/09/05 +內商使用預帶系統別 原外商用帶"FCT"
      Else
        oObject.Text = ""
      End If
    Next
    'end 2021/09/01
  Else
   Text1(5) = ""
   Text1(6) = ""
  End If
  
   Label2 = ""
   'Modified by Lydia 2021/09/01
'   For Each Lbl In Label3
'      Lbl.Caption = ""
'   Next
'   For Each Cmb2 In Combo2
'     Cmb2.Clear
'     Cmb2.AddItem ""
'    Next
'    For Each txt In txtCaseField
'        txt.Text = ""
'   Next
   For Each oObject In Label3
       oObject.Caption = ""
   Next
   For Each oObject In Combo2
       oObject.Clear
       oObject.AddItem ""
   Next
   For Each oObject In txtCaseField
       oObject.Text = ""
   Next
   'end 2021/09/01
End Sub

Private Sub Form_Unload(Cancel As Integer)
     Set frm030207 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0
            Text1(1).SetFocus
            Text1_GotFocus (1)
            
        Case 1
            Text1(4).SetFocus
            Text1_GotFocus (4)
            
    End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Select Case Index
        Case 0, 1, 2, 3
            Option1(0).Value = True
            Text1(4) = ""
            CloseIme
        Case 4
            Option1(1).Value = True
            Text1(0) = strSysKind 'Modify by Amy 2013/09/05 +內商使用預帶系統別 原外商用帶"FCT"
            For i = 1 To 3
                Text1(i) = ""
            Next i
            CloseIme
        Case 6
            OpenIme
        Case Else
           CloseIme
    End Select
    TextInverse Text1(Index)
End Sub

Private Function TxtValidate() As Boolean
    Dim bCancel As Boolean
    Dim CountP As Integer '代表人個數
    
   TxtValidate = False
    If Len(Trim(Text1(5))) = 0 Then
        MsgBox "簽署日期不可為空!", vbCritical
        Text1(5).SetFocus
         Exit Function
     Else
         Text1_Validate 5, bCancel
         If bCancel = True Then
               Text1(5).SetFocus
               TextInverse Text1(5)
               Exit Function
          End If
     End If
            
      If Len(Trim(Text1(6))) = 0 Then
           MsgBox "地點不可為空!", vbCritical
           Text1(6).SetFocus
           Exit Function
      Else
          Text1_Validate 6, bCancel
          If bCancel = True Then
               Text1(6).SetFocus
               TextInverse Text1(6)
               Exit Function
          End If
      End If

     strText(3) = "": CountP = 0
     '2013/07/22 +判斷若有一個申請人非個人則代表人至少有一個值
     If CountNPerson > 0 Then
        For i = 0 To 9 '取畫面上代表人名稱(中)，判斷至少一個有值
             'Modify by Amy 2014/08/11 全型改為&nbsp;
             If Len(Trim(txtCaseField(i * 3))) > 0 Then strText(3) = strText(3) & i + 1 & "." & txtCaseField(i * 3) & "&nbsp;": CountP = CountP + 1
        Next i
     
        If strText(3) = "" Then
            MsgBox "代表人中文名稱不可為空!", vbCritical
            txtCaseField(0).SetFocus
            Exit Function
        Else
            'Modiby by Amy 2014/08/11
            strText(3) = Left(strText(3), Len(strText(3)) - 6)
            If CountP = 1 Then strText(3) = Mid(strText(3), 3) '原:Right(strText(3), Len(strText(3)) - 2)
            'end 2014/08/11
        End If
     Else
        strText(3) = strText(1)
     End If
         
    TxtValidate = True
End Function

'反白
Public Sub TextInverse(ByRef txtTemp As TextBox)
txtTemp.SelStart = 0
txtTemp.SelLength = Len(txtTemp.Text)
End Sub

Private Sub SetCombo2Data(ByVal nIndex As Integer, ByVal strData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To Combo2(nIndex).ListCount - 1
      If Combo2(nIndex).List(nPos) = strData Then
         bFind = True
         Exit For
      End If
   Next nPos
   If Not bFind Then
      Combo2(nIndex).AddItem strData
      Combo2(nIndex).Refresh
      Combo2(nIndex).ListIndex = Combo2(nIndex).ListCount - 1
   Else
      Combo2(nIndex).ListIndex = nPos
   End If
End Sub

'2013/07/22 +回傳個人判斷
'抓取英文名稱，無英文抓中文，無中文抓日文，並回傳中文名稱
Private Function GetCusName(ByVal strNo, ByRef txtTemp As String) As String
    GetCusName = ""
    If strNo = "" Or strNo = "Null" Then
        Exit Function
    Else
         'Modified by Lydia 2019/08/08 +外商國名
         'strExc(0) = "Select Decode(CU05,null,Decode(CU04,null,CU06,CU04),CU05||CU88||CU89||CU90) CusName,CU04,CU15 From Customer " & _
                         "Where " & ChgCustomer(strNo)
         'Modified by Lydia 2019/08/28 FCT和T的自然人+XX籍(紙本)
         'strExc(0) = "Select DECODE(CU05,NULL,DECODE(CU04,NULL,CU06,DECODE(CU15,'1',NA81,'')||CU04), CU05||CU88||CU89||CU90) CUSNAME,CU04,CU15,NA81 " & _
                          "FROM CUSTOMER,NATION WHERE " & ChgCustomer(strNo) & " and cu10=na01(+) "
         strExc(0) = "Select DECODE(CU05,NULL,DECODE(CU04,NULL,CU06,DECODE(CU15,0,SUBSTR(NA81,1,LENGTH(NA81)-1)||'籍',1,NA81,NULL)||CU04), CU05||CU88||CU89||CU90) CUSNAME, " & _
                          "CU04,CU15,NA81 FROM CUSTOMER,NATION WHERE " & ChgCustomer(strNo) & " and cu10=na01(+) "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            GetCusName = "" & RsTemp.Fields("CusName")
            txtTemp = "" & RsTemp.Fields("CU04") 'Modify by Amy 2015/01/16 未輸中文名稱會錯
            If "" & RsTemp.Fields("CU15") = "1" And txtTemp <> "" Then txtTemp = RsTemp.Fields("na81") & txtTemp   'Added by Lydia 2019/08/08 +外商國名
            If "" & RsTemp.Fields("CU15") = "0" And txtTemp <> "" Then txtTemp = Replace(RsTemp.Fields("na81"), "商", "籍") & txtTemp 'Added by Lydia 2019/08/28 FCT和T的自然人+XX籍(紙本)
            If "" & RsTemp.Fields("CU15") <> "0" Then CountNPerson = CountNPerson + 1 '2013/07/22 Add
         End If
    End If
End Function

Private Sub FormSave()
    '以本所案號查詢時按下Word編輯及存檔鈕需同時更新商標基本檔之代表人欄位
    Dim strSqlU As String, stUpdates As String
   
    On Error GoTo CheckingErr
    
      For i = 0 To 29
        If i <= 5 Then
            stUpdates = stUpdates & ",TM" & i + 47 & "=" & CNULL(ChgSQL(txtCaseField(i)))
        Else
            stUpdates = stUpdates & ",TM" & i + 88 & "=" & CNULL(ChgSQL(txtCaseField(i)))
        End If
    Next i
    stUpdates = Right(stUpdates, Len(stUpdates) - 1)
    If stUpdates <> "" Then
        strSqlU = "Update TradeMark Set " & stUpdates & " Where " & strText(0)
        Pub_SeekTbLog strSqlU
        cnnConnection.Execute strSqlU
    End If
    
CheckingErr:

   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub RunWordOpen()
   Dim strTp() As String
   Dim intAddr As Integer '2013/8/16 +地點置中用
   
   bolRetry = False
    On Error GoTo ErrHand
    
    If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
    g_WordAp.Documents.add
  
    With g_WordAp
      .Selection.Font.Name = "新細明體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(2.5) 'Modify by Amy 2014/07/14 +一行文字 原:4.1
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2.5)
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      
      'Add by Amy 2014/07/14
      .Selection.Font.Size = 8
      .Selection.Font.Bold = False
      .Selection.TypeText "申請用委任書"
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '靠右
      .Selection.TypeParagraph
      'end 2014/07/14

      .Selection.Font.Size = 22
      .Selection.Font.Bold = True
      .Selection.TypeText "委 任 狀 中 譯 文"
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
      .Selection.TypeParagraph

      .Selection.Font.Size = 16
      .Selection.Font.Bold = False
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify '左右對齊
      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5  '1.5倍行高
      .Selection.TypeParagraph

      .Selection.TypeText "　　本人／公司係"
      .Selection.Font.Underline = wdUnderlineSingle
      'Modify by Amy 2014/08/11 +Replace
      'Modified by Lydia 2019/08/16 阿蓮: 內文中有底線的文字段落的前後加全形空白
      .Selection.TypeText "　" & Replace(strText(1), "&nbsp;", "　") & "　"
      .Selection.Font.Underline = wdUnderlineNone
      .Selection.TypeText "，址設於"
      .Selection.Font.Underline = wdUnderlineSingle
      'Modified by Lydia 2019/08/16 阿蓮: 內文中有底線的文字段落的前後加全形空白
      .Selection.TypeText "　" & strText(2) & "　"
      .Selection.Font.Underline = wdUnderlineNone
      
      '2013/07/22 判斷若申請人有一個非個人則代表人就顯示
      If CountNPerson > 0 Then
        .Selection.TypeText "，代表人"
        .Selection.Font.Underline = wdUnderlineSingle
        'Modify by Amy 2014/08/11 +Replace
        'Modified by Lydia 2019/08/16 阿蓮: 內文中有底線的文字段落的前後加全形空白
        .Selection.TypeText "　" & Replace(strText(3), "&nbsp;", "　") & "　"
        .Selection.Font.Underline = wdUnderlineNone
      End If
      'Modify by Amy 2014/7/14 修改內容
      'Modify by Amy 2018/04/12 換內容
'      .Selection.TypeText "，茲委任台一國際專利法律事務所．閻啟泰先生、林景郁先生，為本人／公司在中華民國之代理人，" & _
'                                    "有代理提出商標註冊之申請，申請修正及變更；申請註冊延展；讓與登記及授權或強制授權登記；質權之設立、" & _
'                                    "變更及消滅登記；申請減縮商品／服務；申請補發證書及申請證明書；向主管機關請求影印查閱、抄錄相關文件；" & _
'                                    "撤回上述申請，收受有關商標文件，為評定、異議、廢止、訴願、答辯或撤回之提出，訴願程序中之口頭陳述及言詞辯論，" & _
'                                    "及處理所有有關商標註冊之建立及保護之權，並有選任及解任複代理人之權。"
    'Added by Lydia 2022/03/24 取得人員名稱轉Unicode
    strSql = "select st02||decode(st22,'M','先生','小姐') as sname from staff where st01 in ('81040','94007') and st04='1' order by st01"
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strSql)
    If intI = 1 Then
       strExc(1) = RsTemp.GetString(adClipString, , , "、")
       strExc(1) = Mid(strExc(1), 1, Len(strExc(1)) - 1)
       strExc(1) = PUB_Big5toUnicode(strExc(1))
    End If
    'end 2022/03/24
    'Modify by Amy 2020/03/31 拿掉 台一國際專利法律事務所‧
    'Modified by Lydia 2022/03/24 改用Unicode ; 閻啟泰先生、林景郁先生=> strExc(1)
    .Selection.TypeText "，玆委任" & strExc(1) & "，為本人/公司在中華民國之代理人，" & _
                                    "有分別及共同代理向主管機關提出商標註冊之申請，申請修正及變更；申請註冊延展；讓與登記及授權或強制授權登記，質權之設立、" & _
                                    "變更及消滅登記；申請減縮商品/服務；申請補發證書及申請證明書；請求影印查閱、抄錄相關文件；" & _
                                    "撤回上述申請，收受有關商標文件，為評定、異議、廢止、訴願、答辯或撤回之提出，訴願程序中之口頭陳述及言詞辯論，" & _
                                    "及處理所有有關商標註冊之建立及保護之權，並有選任及解任複代理人之權。"

    .Selection.TypeParagraph
    .Selection.TypeParagraph
    

    strExc(0) = Text1(5)
    For i = 0 To 2
        .Selection.Font.Underline = wdUnderlineSingle
        If i = 0 Then
            .Selection.TypeText Space(2) & Left(strExc(0), 4) & Space(2)
            strExc(0) = Right(strExc(0), Len(strExc(0)) - 4)
        Else
            .Selection.TypeText Space(2) & Left(strExc(0), 2) & Space(2)
            strExc(0) = Right(strExc(0), Len(strExc(0)) - 2)
        End If
        .Selection.Font.Underline = wdUnderlineNone
        Select Case i
            Case 0
                .Selection.TypeText "年"
            Case 1
                .Selection.TypeText "月"
            Case 2
                .Selection.TypeText "日　簽於 "
        End Select
    Next
    'Modify by Amy 2013/08/16 +地點置中
    intAddr = (20 - GetTextLength(CheckStr(Text1(6)))) \ 2
    strExc(0) = Space(4 + intAddr) & CheckStr(Text1(6)) & Space(4 + intAddr)
    .Selection.Font.Underline = wdUnderlineSingle
    '.Selection.TypeText convForm(CheckStr(Text1(6)), 28)
    .Selection.TypeText convForm(strExc(0), 28)
    .Selection.Font.Underline = wdUnderlineNone
    
    .Selection.TypeParagraph
    
    'Modify by Amy 2014/08/11 全型空白改&nbsp;
    'strTp = Split(strText(3), "　")
    strTp = Split(strText(3), "&nbsp;")
    '簽署人 -代表人
    For i = 0 To UBound(strTp)
        'Modify by Amy 2013/08/16 簽署人2個以上顯示序號 -外商阿蓮
        'If UBound(strTp) <> 0 Then strTp(i) = Right(strTp(i), Len(strTp(i)) - 2)
        If i = 0 Then
            If GetTextLength(strTp(i)) > 30 Then
                .Selection.TypeText "簽 署 人："
                .Selection.Font.Underline = wdUnderlineSingle
                'Modified by Lydia 2019/08/16 阿蓮: 簽 署 人的前後加全形空白
                .Selection.TypeText "　" & strTp(i) & "　"
            Else
                .Selection.TypeText Space(20) & "簽 署 人："
                .Selection.Font.Underline = wdUnderlineSingle
                 'Modified by Lydia 2019/08/16
                .Selection.TypeText "　" & convForm(CheckStr(strTp(i)), 30) & "　"
            End If
            .Selection.Font.Underline = wdUnderlineNone
            .Selection.TypeParagraph
        Else
            If GetTextLength(strTp(i)) > 30 Then
                .Selection.Font.Underline = wdUnderlineSingle
                 'Modified by Lydia 2019/08/16
                .Selection.TypeText "　" & strTp(i) & "　"
            Else
                .Selection.TypeText Space(30)
                .Selection.Font.Underline = wdUnderlineSingle
                 'Modified by Lydia 2019/08/16
                .Selection.TypeText "　" & convForm(CheckStr(strTp(i)), 30) & "　"
            End If
            .Selection.Font.Underline = wdUnderlineNone
            .Selection.TypeParagraph
        End If
    Next i
    'Mark by Amy 2014/07/21 不需顯示-外商阿蓮 Add 2014/07/14 +職稱空著User自填-外商陳經理
'    .Selection.TypeText Space(20) & "職    稱："
'    .Selection.Font.Underline = wdUnderlineSingle
    'end 2014/0714
   End With

   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize
   
ErrHand:
    If Err.Number <> 0 Then
       Select Case Err.Number
         Case 91, 462:
            If bolRetry = True Then
               MsgBox Err.Description, vbCritical
            Else
               Set g_WordAp = New Word.Application
               g_WordAp.Documents.add
               bolRetry = True
               Resume Next
            End If
         Case Else:
            MsgBox Err.Description, vbCritical
            Resume
      End Select
  End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0, 4
            KeyAscii = UpperCase(KeyAscii)
    End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 5
         If Len(Text1(5)) > 0 Then
            'Add by Amy 2013/08/08 +數值判斷
            If IsNumeric(Text1(5)) = False Then
                MsgBox "簽署日期不可輸入文字!"
                Text1(5).SetFocus
                TextInverse Text1(5)
                Cancel = True
                Exit Sub
            End If
            If CheckIsTaiwanDate(Text1(5) - 19110000) = False Then
                Text1(5).SetFocus
                TextInverse Text1(5)
                Cancel = True
                Exit Sub
             End If
            If Text1(5) > strSrvDate(1) Then
                MsgBox "簽署日期不可大於系統日 !"
                Text1(5).SetFocus
                TextInverse Text1(5)
                 Cancel = True
                Exit Sub
            End If
          End If
     
        Case 6
            If Len(Text1(6)) > 0 Then
                If GetTextLength(Text1(6)) > 20 Then
                    MsgBox "地點超過10個中文字!", vbCritical
                    Text1(6).SetFocus
                    TextInverse Text1(6)
                    Cancel = True
                    Exit Sub
                End If
            End If
    End Select
End Sub

Private Sub txtCaseField_GotFocus(Index As Integer)
    If Index Mod 3 = 1 Then
        CloseIme
    Else
        OpenIme
    End If
End Sub
