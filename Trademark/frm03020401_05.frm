VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020401_05 
   BorderStyle     =   1  '單線固定
   Caption         =   "變更事項"
   ClientHeight    =   5655
   ClientLeft      =   3765
   ClientTop       =   1905
   ClientWidth     =   8550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8550
   Begin VB.CommandButton cmdApprove 
      Caption         =   "全部核准(&P)"
      Height          =   400
      Left            =   3960
      TabIndex        =   69
      Top             =   60
      Width           =   1212
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6240
      TabIndex        =   3
      Top             =   60
      Width           =   1212
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5220
      TabIndex        =   2
      Top             =   60
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7500
      TabIndex        =   4
      Top             =   60
      Width           =   912
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin VB.TextBox textCE01 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   600
      Width           =   2532
   End
   Begin TabDlg.SSTab tabCtrl 
      Height          =   4635
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8176
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm03020401_05.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label22"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label21"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label20"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label19"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label18"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label17"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label16"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label15"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label14"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label13"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label10"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label8"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label6"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label5"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCE04"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCE02"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCE10"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCE11"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCE12"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCE13"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCE14"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCE15"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCE55"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCE53"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCE51"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textCE17"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCE63"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textCE64"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textCE65"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textCE22"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textCE52"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textCE54"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCE56"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textCE16"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCE03"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCE09"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm03020401_05.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label44"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label43"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label42"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label41"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label40"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label39"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label38"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label37"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label36"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label35"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label34"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label33"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label32"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label31"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label30"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label29"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label28"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label27"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label26"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label25"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label24"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label23"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label45"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "textCE41_1"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "textCE23"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "textCE24"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "textCE25"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "textCE57"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "textCE39"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "textCE41"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "textCE42"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "textCE43"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "textCE45"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "textCE47"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "textCE49"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "textCE61"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "textCE59"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "textCE62"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "textCE50"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "textCE48"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "textCE46"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "textCE44"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "textCE60"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "textCE40"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "textCE58"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "textCE38"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).ControlCount=   46
      Begin VB.TextBox textCE09 
         Height          =   285
         Left            =   -73680
         TabIndex        =   22
         Top             =   330
         Width           =   372
      End
      Begin VB.TextBox textCE03 
         Height          =   285
         Left            =   -73680
         TabIndex        =   21
         Top             =   630
         Width           =   372
      End
      Begin VB.TextBox textCE16 
         Height          =   285
         Left            =   -73680
         TabIndex        =   20
         Top             =   930
         Width           =   372
      End
      Begin VB.TextBox textCE56 
         Height          =   285
         Left            =   -73680
         TabIndex        =   19
         Top             =   2730
         Width           =   372
      End
      Begin VB.TextBox textCE54 
         Height          =   285
         Left            =   -73680
         TabIndex        =   18
         Top             =   3030
         Width           =   372
      End
      Begin VB.TextBox textCE52 
         Height          =   285
         Left            =   -73680
         TabIndex        =   17
         Top             =   3330
         Width           =   372
      End
      Begin VB.TextBox textCE22 
         Height          =   285
         Left            =   -73680
         TabIndex        =   16
         Top             =   3630
         Width           =   372
      End
      Begin VB.TextBox textCE65 
         Height          =   285
         Left            =   -73680
         TabIndex        =   15
         Top             =   3930
         Width           =   372
      End
      Begin VB.TextBox textCE38 
         Height          =   285
         Left            =   1320
         TabIndex        =   14
         Top             =   330
         Width           =   372
      End
      Begin VB.TextBox textCE58 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   1230
         Width           =   372
      End
      Begin VB.TextBox textCE40 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   1530
         Width           =   372
      End
      Begin VB.TextBox textCE60 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   1830
         Width           =   372
      End
      Begin VB.TextBox textCE44 
         Height          =   285
         Left            =   1320
         TabIndex        =   10
         Top             =   2126
         Width           =   372
      End
      Begin VB.TextBox textCE46 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         Top             =   3030
         Width           =   372
      End
      Begin VB.TextBox textCE48 
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   3330
         Width           =   372
      End
      Begin VB.TextBox textCE50 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   3630
         Width           =   372
      End
      Begin VB.TextBox textCE62 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   3930
         Width           =   372
      End
      Begin MSForms.TextBox textCE59 
         Height          =   285
         Left            =   3240
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   1830
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE61 
         Height          =   525
         Left            =   2520
         TabIndex        =   97
         Top             =   3930
         Width           =   5535
         VariousPropertyBits=   -1467989989
         Size            =   "9763;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE49 
         Height          =   285
         Left            =   3240
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   3630
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE47 
         Height          =   285
         Left            =   3240
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   3330
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE45 
         Height          =   285
         Left            =   3240
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   3030
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE43 
         Height          =   285
         Left            =   3240
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   2730
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE42 
         Height          =   285
         Left            =   3240
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   2430
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41 
         Height          =   285
         Left            =   3240
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   2130
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE39 
         Height          =   285
         Left            =   3240
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1530
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE57 
         Height          =   285
         Left            =   3240
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1230
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE25 
         Height          =   285
         Left            =   3240
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   930
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE24 
         Height          =   285
         Left            =   3240
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   630
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE23 
         Height          =   285
         Left            =   3240
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   330
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8488;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41_1 
         Height          =   885
         Left            =   3240
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   2130
         Width           =   4815
         VariousPropertyBits=   671105055
         Size            =   "8493;1561"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE64 
         Height          =   285
         Left            =   -71700
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   4230
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE63 
         Height          =   285
         Left            =   -71700
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   3930
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE17 
         Height          =   285
         Left            =   -71700
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   3630
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE51 
         Height          =   285
         Left            =   -71700
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   3330
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE53 
         Height          =   285
         Left            =   -71700
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   3030
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE55 
         Height          =   285
         Left            =   -71700
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   2730
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE15 
         Height          =   285
         Left            =   -71700
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   2430
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE14 
         Height          =   285
         Left            =   -71700
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2130
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE13 
         Height          =   285
         Left            =   -71700
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1830
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE12 
         Height          =   285
         Left            =   -71700
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   1530
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE11 
         Height          =   285
         Left            =   -71700
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1230
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE10 
         Height          =   285
         Left            =   -71700
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   930
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE02 
         Height          =   285
         Left            =   -71700
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   630
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE04 
         Height          =   285
         Left            =   -71700
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   330
         Width           =   4695
         VariousPropertyBits=   671105055
         Size            =   "8276;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label45 
         Caption         =   "案件名稱 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   70
         Top             =   2142
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Index           =   1
         Left            =   -74880
         TabIndex        =   66
         Top             =   346
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "申請人 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   65
         Top             =   346
         Width           =   732
      End
      Begin VB.Label Label3 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   64
         Top             =   646
         Width           =   1092
      End
      Begin VB.Label Label4 
         Caption         =   "申請日 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   63
         Top             =   646
         Width           =   732
      End
      Begin VB.Label Label5 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   62
         Top             =   946
         Width           =   1092
      End
      Begin VB.Label Label6 
         Caption         =   "代表人1(中) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   61
         Top             =   946
         Width           =   1092
      End
      Begin VB.Label Label7 
         Caption         =   "代表人1(英) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   60
         Top             =   1254
         Width           =   1092
      End
      Begin VB.Label Label8 
         Caption         =   "代表人1(日) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   59
         Top             =   1552
         Width           =   1092
      End
      Begin VB.Label Label9 
         Caption         =   "代表人2(中) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   58
         Top             =   1850
         Width           =   1092
      End
      Begin VB.Label Label10 
         Caption         =   "代表人2(英) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   57
         Top             =   2148
         Width           =   1092
      End
      Begin VB.Label Label11 
         Caption         =   "代表人2(日) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   56
         Top             =   2446
         Width           =   1092
      End
      Begin VB.Label Label12 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   55
         Top             =   2746
         Width           =   1092
      End
      Begin VB.Label Label13 
         Caption         =   "代理人 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   54
         Top             =   2746
         Width           =   732
      End
      Begin VB.Label Label14 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   53
         Top             =   3046
         Width           =   1092
      End
      Begin VB.Label Label15 
         Caption         =   "代表人印鑑 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   52
         Top             =   3042
         Width           =   1092
      End
      Begin VB.Label Label16 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   51
         Top             =   3346
         Width           =   1092
      End
      Begin VB.Label Label17 
         Caption         =   "申請人印鑑 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   50
         Top             =   3340
         Width           =   1092
      End
      Begin VB.Label Label18 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   49
         Top             =   3646
         Width           =   1092
      End
      Begin VB.Label Label19 
         Caption         =   "申請人中譯文 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   48
         Top             =   3638
         Width           =   1332
      End
      Begin VB.Label Label20 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   47
         Top             =   3946
         Width           =   1092
      End
      Begin VB.Label Label21 
         Caption         =   "代表人1中譯文 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   46
         Top             =   3936
         Width           =   1332
      End
      Begin VB.Label Label22 
         Caption         =   "代表人2中譯文 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   45
         Top             =   4246
         Width           =   1332
      End
      Begin VB.Label Label23 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   44
         Top             =   346
         Width           =   1092
      End
      Begin VB.Label Label24 
         Caption         =   "申請地址(中) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   43
         Top             =   346
         Width           =   1212
      End
      Begin VB.Label Label25 
         Caption         =   "申請地址(英) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   42
         Top             =   657
         Width           =   1212
      End
      Begin VB.Label Label26 
         Caption         =   "申請地址(日) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   41
         Top             =   954
         Width           =   1212
      End
      Begin VB.Label Label27 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   40
         Top             =   1246
         Width           =   1092
      End
      Begin VB.Label Label28 
         Caption         =   "正商標號數 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   39
         Top             =   1246
         Width           =   1212
      End
      Begin VB.Label Label29 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   38
         Top             =   1546
         Width           =   1092
      End
      Begin VB.Label Label30 
         Caption         =   "商標種類 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   37
         Top             =   1546
         Width           =   1212
      End
      Begin VB.Label Label31 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   36
         Top             =   1846
         Width           =   1092
      End
      Begin VB.Label Label32 
         Caption         =   "圖樣 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   35
         Top             =   1846
         Width           =   612
      End
      Begin VB.Label Label33 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   34
         Top             =   2142
         Width           =   1092
      End
      Begin VB.Label Label34 
         Caption         =   "案件名稱(中) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   33
         Top             =   2160
         Width           =   1212
      End
      Begin VB.Label Label35 
         Caption         =   "案件名稱(英) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   32
         Top             =   2439
         Width           =   1212
      End
      Begin VB.Label Label36 
         Caption         =   "案件名稱(日) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   31
         Top             =   2736
         Width           =   1212
      End
      Begin VB.Label Label37 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   30
         Top             =   3046
         Width           =   1092
      End
      Begin VB.Label Label38 
         Caption         =   "縮減商品 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   29
         Top             =   3046
         Width           =   1212
      End
      Begin VB.Label Label39 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   28
         Top             =   3346
         Width           =   1092
      End
      Begin VB.Label Label40 
         Caption         =   "商品類別 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   27
         Top             =   3346
         Width           =   1212
      End
      Begin VB.Label Label41 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   26
         Top             =   3646
         Width           =   1092
      End
      Begin VB.Label Label42 
         Caption         =   "商品群組 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   25
         Top             =   3646
         Width           =   1212
      End
      Begin VB.Label Label43 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   24
         Top             =   3930
         Width           =   1092
      End
      Begin VB.Label Label44 
         Caption         =   "其它 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   23
         Top             =   3930
         Width           =   612
      End
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   68
      Top             =   600
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   2
      Left            =   4440
      TabIndex        =   67
      Top             =   600
      Width           =   732
   End
End
Attribute VB_Name = "frm03020401_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/12 改成Form2.0 ; 全部textCExx.textbox
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 收文號
Dim m_CE01 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
'Dim m_FieldList() As FIELDITEM
'Dim m_FieldCount As Integer
Dim m_CEList() As FIELDITEM
Dim m_CECount As Integer
' 針對商標基本檔欄位所用的暫存陣列
Dim m_TMList() As FIELDITEM
Dim m_TMCount As Integer

' 更新商標基本檔時所使用的變更事項檔欄位的暫存資料
' 申請日
Dim m_CE02 As String
' 申請人
Dim m_CE04 As String
' 商品種類代碼
Dim m_CE39 As String
Dim bolErr As Boolean 'Added by Lydia 2016/07/19
'Added by Lydia 2016/07/19 延展核准在存檔時,直接將變更事項確定全部核准
Public Function Get102_Approve() As Boolean 'Memo by Lydia 2017/07/28 預設全部核准 (+301變更)
    Call cmdApprove_Click
    Call cmdOK_Click
    If bolErr = False Then
       Get102_Approve = True
    Else
       Get102_Approve = False
    End If
End Function
' 檢查該欄位是否存在
Private Function IsCEFieldExist(ByVal strField As String) As Boolean
   Dim nIndex As Integer
   IsCEFieldExist = False
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         IsCEFieldExist = True
         Exit For
      End If
   Next nIndex
End Function
' 新增一個欄位
Private Sub AddCEField(ByVal strField As String, ByVal strOldData As String, ByVal nType As Integer)
   If IsCEFieldExist(strField) = True Then
      GoTo EXITSUB
   End If
   ReDim Preserve m_CEList(m_CECount + 1)
   m_CEList(m_CECount).fiName = strField
   m_CEList(m_CECount).fiOldData = strOldData
   m_CEList(m_CECount).fiNewData = strOldData
   m_CEList(m_CECount).fiType = nType
   m_CECount = m_CECount + 1
EXITSUB:
End Sub
' 設定欄位新值
Private Sub SetCEFieldNewData(ByVal strField As String, ByVal strNewData As String)
   Dim nIndex As Integer
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         m_CEList(nIndex).fiNewData = strNewData
         Exit For
      End If
   Next nIndex
End Sub
' 清除欄位串列
Private Sub ClearCEFields()
   Erase m_CEList
   m_CECount = 0
End Sub

' 檢查該商標基本檔的欄位是否存在
Private Function IsTMFieldExist(ByVal strField As String) As Boolean
   Dim nIndex As Integer
   IsTMFieldExist = False
   For nIndex = 0 To m_TMCount - 1
      If m_TMList(nIndex).fiName = strField Then
         IsTMFieldExist = True
         Exit For
      End If
   Next nIndex
End Function
' 新增一個欄位
Private Sub AddTMField(ByVal strField As String, ByVal strOldData As String, ByVal nType As Integer)
   If IsTMFieldExist(strField) = True Then
      GoTo EXITSUB
   End If
   ReDim Preserve m_TMList(m_TMCount + 1)
   m_TMList(m_TMCount).fiName = strField
   m_TMList(m_TMCount).fiOldData = strOldData
   m_TMList(m_TMCount).fiNewData = strOldData
   m_TMList(m_TMCount).fiType = nType
   m_TMCount = m_TMCount + 1
EXITSUB:
End Sub
' 設定欄位新值
Private Sub SetTMFieldNewData(ByVal strField As String, ByVal strNewData As String)
   Dim nIndex As Integer
   For nIndex = 0 To m_TMCount - 1
      If m_TMList(nIndex).fiName = strField Then
         m_TMList(nIndex).fiNewData = strNewData
         Exit For
      End If
   Next nIndex
End Sub
' 清除欄位串列
Private Sub ClearTMFields()
   Erase m_TMList
   m_TMCount = 0
End Sub

' 更新欄位內容
Private Sub UpdateFieldNewData()
   SetCEFieldNewData "CE03", textCE03: SetCEFieldNewData "CE09", textCE09: SetCEFieldNewData "CE16", textCE16: SetCEFieldNewData "CE22", textCE22: SetCEFieldNewData "CE38", textCE38
   SetCEFieldNewData "CE40", textCE40: SetCEFieldNewData "CE44", textCE44: SetCEFieldNewData "CE46", textCE46: SetCEFieldNewData "CE48", textCE48: SetCEFieldNewData "CE50", textCE50
   SetCEFieldNewData "CE52", textCE52: SetCEFieldNewData "CE54", textCE54: SetCEFieldNewData "CE56", textCE56: SetCEFieldNewData "CE58", textCE58: SetCEFieldNewData "CE60", textCE60
   SetCEFieldNewData "CE62", textCE62: SetCEFieldNewData "CE65", textCE65
End Sub

Private Sub cmdApprove_Click()
    'Add By Cheng 2002/12/11
    '第一頁
    If Me.textCE04.Text <> "" Then Me.textCE09.Text = "1"
    If Me.textCE02.Text <> "" Then Me.textCE03.Text = "1"
    If Me.textCE10.Text <> "" Or Me.textCE11.Text <> "" Or _
        Me.textCE12.Text <> "" Or Me.textCE13.Text <> "" Or _
        Me.textCE14.Text <> "" Or Me.textCE15.Text <> "" Then Me.textCE16.Text = "1"
    If Me.textCE55.Text <> "" Then Me.textCE56.Text = "1"
    If Me.textCE53.Text <> "" Then Me.textCE54.Text = "1"
    If Me.textCE51.Text <> "" Then Me.textCE52.Text = "1"
    If Me.textCE17.Text <> "" Then Me.textCE22.Text = "1"
    If Me.textCE63.Text <> "" Or Me.textCE64.Text <> "" Then Me.textCE65.Text = "1"
    If Me.textCE17.Text <> "" Then Me.textCE22.Text = "1"
    '第二頁
    If Me.textCE23.Text <> "" Or Me.textCE24.Text <> "" Or Me.textCE25.Text <> "" Then Me.textCE38.Text = "1"
    If Me.textCE57.Text <> "" Then Me.textCE58.Text = "1"
    If Me.textCE39.Text <> "" Then Me.textCE40.Text = "1"
    If Me.textCE59.Text <> "" Then Me.textCE60.Text = "1"
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        If Me.textCE41_1.Text <> "" Then Me.textCE44.Text = "1"
    Case Else
        If Me.textCE41.Text <> "" Or Me.textCE42.Text <> "" Or Me.textCE43.Text <> "" Then Me.textCE44.Text = "1"
    End Select
    If Me.textCE45.Text <> "" Then Me.textCE46.Text = "1"
    If Me.textCE47.Text <> "" Then Me.textCE48.Text = "1"
    If Me.textCE49.Text <> "" Then Me.textCE50.Text = "1"
    If Me.textCE61.Text <> "" Then Me.textCE62.Text = "1"
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm03020401_04.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm03020401_04
   Unload frm03020401_03
   Unload frm03020401_02
   Unload frm03020401_01
   Unload Me
End Sub

Private Sub cmdOK_Click()

    'Added by Lydia 2021/09/13 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         Exit Sub
    End If

   bolErr = True 'Added by Lydia 2016/07/19
   UpdateFieldNewData
      'edit by  nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
   bolErr = False 'Added by Lydia 2016/07/19
   Unload Me
   ' 90.07.19 modify (回到前畫面)
   'Unload frm03020401_04
   'Unload frm03020401_03
   'Unload frm03020401_02
   'frm03020401_01.Show
    'Modify By Cheng 2002/12/11
'   frm03020404_01.Show
    frm03020401_04.Show
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textCE01.BackColor = &H8000000F
   textTMKey.BackColor = &H8000000F
   textCE02.BackColor = &H8000000F
   textCE04.BackColor = &H8000000F
   textCE10.BackColor = &H8000000F
   textCE11.BackColor = &H8000000F
   textCE12.BackColor = &H8000000F
   textCE13.BackColor = &H8000000F
   textCE14.BackColor = &H8000000F
   textCE15.BackColor = &H8000000F
   textCE17.BackColor = &H8000000F
   textCE23.BackColor = &H8000000F
   textCE24.BackColor = &H8000000F
   textCE25.BackColor = &H8000000F
   textCE39.BackColor = &H8000000F
   textCE41.BackColor = &H8000000F
   textCE41_1.BackColor = &H8000000F
   textCE42.BackColor = &H8000000F
   textCE43.BackColor = &H8000000F
   textCE45.BackColor = &H8000000F
   textCE47.BackColor = &H8000000F
   textCE49.BackColor = &H8000000F
   textCE51.BackColor = &H8000000F
   textCE53.BackColor = &H8000000F
   textCE55.BackColor = &H8000000F
   textCE57.BackColor = &H8000000F
   textCE59.BackColor = &H8000000F
   textCE61.BackColor = &H8000000F
   textCE63.BackColor = &H8000000F
   textCE64.BackColor = &H8000000F
   
    'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    textCE10.MaxLength = Pub_MaxCEL10
    textCE11.MaxLength = Pub_MaxCEL11
    textCE13.MaxLength = Pub_MaxCEL10
    textCE14.MaxLength = Pub_MaxCEL11
    'end 2016/09/10
    
   tabCtrl.Tab = 0
  
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ClearCEFields
   'Add By Cheng 2002/07/19
   Set frm03020401_05 = Nothing
End Sub

' 由客戶代號取得地址
' Input : strData ==> 客戶代號
'         nType ==> 種類
'                   0 : 表要取得的是中文地址
'                   1 : 表要取得的是英文地址
'                   2 : 表要取得的是日文地址
Private Function GetAddress(ByVal strData As String, ByVal nType As Integer) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetAddress = Empty
   If IsEmptyText(strData) = False Then
      If Len(strData) > 8 Then
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                        "CU02 = '" & Mid(strData, 9, 1) & "'"
      Else
         strSql = "SELECT * FROM Customer " & _
                  "WHERE CU01 = '" & Mid(strData, 1, 8) & "'"
      End If
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Select Case nType
            Case 0:
               If IsNull(rsTmp.Fields("CU23")) = False Then
                  GetAddress = rsTmp.Fields("CU23")
               End If
            Case 1:
               If IsNull(rsTmp.Fields("CU24")) = False Then
                  GetAddress = rsTmp.Fields("CU24")
               End If
            Case 2:
               If IsNull(rsTmp.Fields("CU29")) = False Then
                  GetAddress = rsTmp.Fields("CU29")
               End If
         End Select
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Function

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CE01 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
      ' 收文號
      Case 4: m_CE01 = strData
   End Select
End Sub

Public Sub QueryData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   ' 清除欄位串列
   ClearCEFields
   
   ' 清除暫存變數
   m_CE02 = Empty
   m_CE04 = Empty
    'Add By Cheng 2003/11/11
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        Me.Label34.Visible = False
        Me.textCE41.Visible = False
        Me.textCE41.Enabled = False
        Me.Label35.Visible = False
        Me.textCE42.Visible = False
        Me.textCE42.Enabled = False
        Me.Label36.Visible = False
        Me.textCE43.Visible = False
        Me.textCE43.Enabled = False
    Case Else
        Me.Label45.Visible = False
        Me.textCE41_1.Visible = False
        Me.textCE41_1.Enabled = False
    End Select
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textCE01 = m_CE01
   
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請日
      If IsNull(rsTmp.Fields("CE02")) = False Then
         m_CE02 = rsTmp.Fields("CE02")
         textCE02 = TAIWANDATE(rsTmp.Fields("CE02"))
      End If
      If IsNull(rsTmp.Fields("CE03")) = False Then
         textCE03 = rsTmp.Fields("CE03")
      End If
      AddCEField "CE03", textCE03, 0
      ' 申請人
      If IsNull(rsTmp.Fields("CE04")) = False Then
         m_CE04 = rsTmp.Fields("CE04")
         textCE04 = GetCustomerName(rsTmp.Fields("CE04"), 0)
      End If
      If IsNull(rsTmp.Fields("CE09")) = False Then
         textCE09 = rsTmp.Fields("CE09")
      End If
      AddCEField "CE09", textCE09, 0
      ' 代表人
      If IsNull(rsTmp.Fields("CE10")) = False Then
         textCE10 = rsTmp.Fields("CE10")
      End If
      If IsNull(rsTmp.Fields("CE11")) = False Then
         textCE10 = rsTmp.Fields("CE11")
      End If
      If IsNull(rsTmp.Fields("CE12")) = False Then
         textCE12 = rsTmp.Fields("CE12")
      End If
      If IsNull(rsTmp.Fields("CE13")) = False Then
         textCE13 = rsTmp.Fields("CE13")
      End If
      If IsNull(rsTmp.Fields("CE14")) = False Then
         textCE14 = rsTmp.Fields("CE14")
      End If
      If IsNull(rsTmp.Fields("CE15")) = False Then
         textCE15 = rsTmp.Fields("CE15")
      End If
      If IsNull(rsTmp.Fields("CE16")) = False Then
         textCE16 = rsTmp.Fields("CE16")
      End If
      AddCEField "CE16", textCE16, 0
      ' 申請人中譯文
      If IsNull(rsTmp.Fields("CE17")) = False Then
         textCE17 = rsTmp.Fields("CE17")
      End If
      If IsNull(rsTmp.Fields("CE22")) = False Then
         textCE22 = rsTmp.Fields("CE22")
      End If
      AddCEField "CE22", textCE22, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("CE23")) = False Then
         textCE23 = rsTmp.Fields("CE23")
      End If
      If IsNull(rsTmp.Fields("CE24")) = False Then
         textCE24 = rsTmp.Fields("CE24")
      End If
      If IsNull(rsTmp.Fields("CE25")) = False Then
         textCE25 = rsTmp.Fields("CE25")
      End If
      If IsNull(rsTmp.Fields("CE38")) = False Then
         textCE38 = rsTmp.Fields("CE38")
      End If
      AddCEField "CE38", textCE38, 0
      ' 專利商標種類代號
      If IsNull(rsTmp.Fields("CE39")) = False Then
         m_CE39 = rsTmp.Fields("CE39")
         textCE39 = rsTmp.Fields("CE39")
      End If
      If IsNull(rsTmp.Fields("CE40")) = False Then
         textCE40 = rsTmp.Fields("CE40")
      End If
      AddCEField "CE40", textCE40, 0
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            ' 案件名稱
            If IsNull(rsTmp.Fields("CE41")) = False Then
               textCE41_1 = rsTmp.Fields("CE41")
            End If
        Case Else
            ' 案件名稱
            If IsNull(rsTmp.Fields("CE41")) = False Then
               textCE41 = rsTmp.Fields("CE41")
            End If
            If IsNull(rsTmp.Fields("CE42")) = False Then
               textCE42 = rsTmp.Fields("CE42")
            End If
            If IsNull(rsTmp.Fields("CE43")) = False Then
               textCE43 = rsTmp.Fields("CE43")
            End If
        End Select
      If IsNull(rsTmp.Fields("CE44")) = False Then
         textCE44 = rsTmp.Fields("CE44")
      End If
      AddCEField "CE44", textCE44, 0
      ' 縮減商品
      If IsNull(rsTmp.Fields("CE45")) = False Then
         textCE45 = rsTmp.Fields("CE45")
      End If
      If IsNull(rsTmp.Fields("CE46")) = False Then
         textCE46 = rsTmp.Fields("CE46")
      End If
      AddCEField "CE46", textCE46, 0
      ' 商品類別
      If IsNull(rsTmp.Fields("CE47")) = False Then
         textCE47 = rsTmp.Fields("CE47")
      End If
      If IsNull(rsTmp.Fields("CE48")) = False Then
         textCE48 = rsTmp.Fields("CE48")
      End If
      AddCEField "CE48", textCE48, 0
      ' 商品群組
      If IsNull(rsTmp.Fields("CE49")) = False Then
         textCE49 = rsTmp.Fields("CE49")
      End If
      If IsNull(rsTmp.Fields("CE50")) = False Then
         textCE50 = rsTmp.Fields("CE50")
      End If
      AddCEField "CE50", textCE50, 0
      ' 申請人印鑑
      If IsNull(rsTmp.Fields("CE51")) = False Then
         textCE51 = rsTmp.Fields("CE51")
      End If
      If IsNull(rsTmp.Fields("CE52")) = False Then
         textCE52 = rsTmp.Fields("CE52")
      End If
      AddCEField "CE52", textCE52, 0
      ' 代表人印鑑
      If IsNull(rsTmp.Fields("CE53")) = False Then
         textCE53 = rsTmp.Fields("CE53")
      End If
      If IsNull(rsTmp.Fields("CE54")) = False Then
         textCE54 = rsTmp.Fields("CE54")
      End If
      AddCEField "CE54", textCE54, 0
      ' 代理人
      If IsNull(rsTmp.Fields("CE55")) = False Then
         textCE55 = rsTmp.Fields("CE55")
      End If
      If IsNull(rsTmp.Fields("CE56")) = False Then
         textCE56 = rsTmp.Fields("CE56")
      End If
      AddCEField "CE56", textCE56, 0
      ' 正商標號數
      If IsNull(rsTmp.Fields("CE57")) = False Then
         textCE57 = rsTmp.Fields("CE57")
      End If
      If IsNull(rsTmp.Fields("CE58")) = False Then
         textCE58 = rsTmp.Fields("CE58")
      End If
      AddCEField "CE58", textCE58, 0
      ' 圖樣
      If IsNull(rsTmp.Fields("CE59")) = False Then
         textCE59 = rsTmp.Fields("CE59")
      End If
      If IsNull(rsTmp.Fields("CE60")) = False Then
         textCE60 = rsTmp.Fields("CE60")
      End If
      AddCEField "CE60", textCE60, 0
      ' 其它
      If IsNull(rsTmp.Fields("CE61")) = False Then
         textCE61 = rsTmp.Fields("CE61")
      End If
      If IsNull(rsTmp.Fields("CE62")) = False Then
         textCE62 = rsTmp.Fields("CE62")
      End If
      AddCEField "CE62", textCE62, 0
      ' 代表人譯文
      If IsNull(rsTmp.Fields("CE63")) = False Then
         textCE63 = rsTmp.Fields("CE63")
      End If
      If IsNull(rsTmp.Fields("CE64")) = False Then
         textCE64 = rsTmp.Fields("CE64")
      End If
      If IsNull(rsTmp.Fields("CE65")) = False Then
         textCE65 = rsTmp.Fields("CE65")
      End If
      AddCEField "CE65", textCE65, 0
      
      OnUpdateCtrlState rsTmp
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Sub OnUpdateCtrlState(ByRef rsTmp As ADODB.Recordset)
   '
   EnableTextBox textCE09, False
   If IsNull(rsTmp.Fields("CE04")) = False Then
      If IsEmptyText(rsTmp.Fields("CE04")) = False Then
         EnableTextBox textCE09, True
      End If
   End If
   '
   EnableTextBox textCE03, False
   If IsNull(rsTmp.Fields("CE02")) = False Then
      If IsEmptyText(rsTmp.Fields("CE02")) = False Then
         EnableTextBox textCE03, True
      End If
   End If
   '
   EnableTextBox textCE16, False
   If IsNull(rsTmp.Fields("CE10")) = False Then
      If IsEmptyText(rsTmp.Fields("CE10")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE11")) = False Then
      If IsEmptyText(rsTmp.Fields("CE11")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE12")) = False Then
      If IsEmptyText(rsTmp.Fields("CE12")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE13")) = False Then
      If IsEmptyText(rsTmp.Fields("CE13")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE14")) = False Then
      If IsEmptyText(rsTmp.Fields("CE14")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE15")) = False Then
      If IsEmptyText(rsTmp.Fields("CE15")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   '
   EnableTextBox textCE56, False
   If IsNull(rsTmp.Fields("CE55")) = False Then
      If IsEmptyText(rsTmp.Fields("CE55")) = False Then
         EnableTextBox textCE56, True
      End If
   End If
   '
   EnableTextBox textCE54, False
   If IsNull(rsTmp.Fields("CE53")) = False Then
      If IsEmptyText(rsTmp.Fields("CE53")) = False Then
         EnableTextBox textCE54, True
      End If
   End If
   '
   EnableTextBox textCE52, False
   If IsNull(rsTmp.Fields("CE51")) = False Then
      If IsEmptyText(rsTmp.Fields("CE51")) = False Then
         EnableTextBox textCE52, True
      End If
   End If
   '
   EnableTextBox textCE22, False
   If IsNull(rsTmp.Fields("CE17")) = False Then
      If IsEmptyText(rsTmp.Fields("CE17")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE18")) = False Then
      If IsEmptyText(rsTmp.Fields("CE18")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE19")) = False Then
      If IsEmptyText(rsTmp.Fields("CE19")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE20")) = False Then
      If IsEmptyText(rsTmp.Fields("CE20")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE21")) = False Then
      If IsEmptyText(rsTmp.Fields("CE21")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   '
   EnableTextBox textCE65, False
   If IsNull(rsTmp.Fields("CE63")) = False Then
      If IsEmptyText(rsTmp.Fields("CE63")) = False Then
         EnableTextBox textCE65, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE64")) = False Then
      If IsEmptyText(rsTmp.Fields("CE64")) = False Then
         EnableTextBox textCE65, True
      End If
   End If
   '
   EnableTextBox textCE38, False
   If IsNull(rsTmp.Fields("CE23")) = False Then
      If IsEmptyText(rsTmp.Fields("CE23")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE24")) = False Then
      If IsEmptyText(rsTmp.Fields("CE24")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE25")) = False Then
      If IsEmptyText(rsTmp.Fields("CE25")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE26")) = False Then
      If IsEmptyText(rsTmp.Fields("CE26")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE27")) = False Then
      If IsEmptyText(rsTmp.Fields("CE27")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE28")) = False Then
      If IsEmptyText(rsTmp.Fields("CE28")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE29")) = False Then
      If IsEmptyText(rsTmp.Fields("CE29")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE30")) = False Then
      If IsEmptyText(rsTmp.Fields("CE30")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE31")) = False Then
      If IsEmptyText(rsTmp.Fields("CE31")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE32")) = False Then
      If IsEmptyText(rsTmp.Fields("CE32")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE33")) = False Then
      If IsEmptyText(rsTmp.Fields("CE33")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE34")) = False Then
      If IsEmptyText(rsTmp.Fields("CE34")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE35")) = False Then
      If IsEmptyText(rsTmp.Fields("CE35")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE36")) = False Then
      If IsEmptyText(rsTmp.Fields("CE36")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE37")) = False Then
      If IsEmptyText(rsTmp.Fields("CE37")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   '
   EnableTextBox textCE58, False
   If IsNull(rsTmp.Fields("CE57")) = False Then
      If IsEmptyText(rsTmp.Fields("CE57")) = False Then
         EnableTextBox textCE58, True
      End If
   End If
   '
   EnableTextBox textCE40, False
   If IsNull(rsTmp.Fields("CE39")) = False Then
      If IsEmptyText(rsTmp.Fields("CE39")) = False Then
         EnableTextBox textCE40, True
      End If
   End If
   '
   EnableTextBox textCE60, False
   If IsNull(rsTmp.Fields("CE59")) = False Then
      If IsEmptyText(rsTmp.Fields("CE59")) = False Then
         EnableTextBox textCE60, True
      End If
   End If
   '
   EnableTextBox textCE44, False
   If IsNull(rsTmp.Fields("CE41")) = False Then
      If IsEmptyText(rsTmp.Fields("CE41")) = False Then
         EnableTextBox textCE44, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE42")) = False Then
      If IsEmptyText(rsTmp.Fields("CE42")) = False Then
         EnableTextBox textCE44, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE43")) = False Then
      If IsEmptyText(rsTmp.Fields("CE43")) = False Then
         EnableTextBox textCE44, True
      End If
   End If
   '
   EnableTextBox textCE46, False
   If IsNull(rsTmp.Fields("CE45")) = False Then
      If IsEmptyText(rsTmp.Fields("CE45")) = False Then
         EnableTextBox textCE46, True
      End If
   End If
   '
   EnableTextBox textCE48, False
   If IsNull(rsTmp.Fields("CE47")) = False Then
      If IsEmptyText(rsTmp.Fields("CE47")) = False Then
         EnableTextBox textCE48, True
      End If
   End If
   '
   EnableTextBox textCE50, False
   If IsNull(rsTmp.Fields("CE49")) = False Then
      If IsEmptyText(rsTmp.Fields("CE49")) = False Then
         EnableTextBox textCE50, True
      End If
   End If
   '
   EnableTextBox textCE62, False
   If IsNull(rsTmp.Fields("CE61")) = False Then
      If IsEmptyText(rsTmp.Fields("CE61")) = False Then
         EnableTextBox textCE62, True
      End If
   End If
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strSql As String
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans
   
   'Add By Sindy 2012/2/16 按確定時先檢查變更事項有值的欄位, 至少要輸入一項核准, 不可都不選
   If Me.textCE09.Text = "1" Or Me.textCE03.Text = "1" Or _
      Me.textCE16.Text = "1" Or Me.textCE56.Text = "1" Or _
      Me.textCE54.Text = "1" Or Me.textCE52.Text = "1" Or _
      Me.textCE22.Text = "1" Or Me.textCE65.Text = "1" Or _
      Me.textCE38.Text = "1" Or Me.textCE58.Text = "1" Or _
      Me.textCE40.Text = "1" Or Me.textCE60.Text = "1" Or _
      Me.textCE44.Text = "1" Or Me.textCE46.Text = "1" Or _
      Me.textCE48.Text = "1" Or Me.textCE50.Text = "1" Or _
      Me.textCE62.Text = "1" Then
      frm03020401_04.m_blnClkChgButton = True
   End If
   '2012/2/16 End
   
   strSql = "UPDATE ChangeEvent SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CECount - 1
      strTmp = Empty
      If m_CEList(nIndex).fiOldData <> m_CEList(nIndex).fiNewData Then
'         'Add By Sindy 2012/2/6 按確定時先檢查變更事項有值的欄位, 至少要輸入一項核准, 不可都不選
'         If m_CEList(nIndex).fiNewData = "1" Then
'            frm03020401_04.m_blnClkChgButton = True
'         End If
'         '2012/2/6 End
         If m_CEList(nIndex).fiType = 0 Then
            strTmp = m_CEList(nIndex).fiName & " = '" & m_CEList(nIndex).fiNewData & "'"
         Else
            If m_CEList(nIndex).fiNewData = Empty Then
               strTmp = m_CEList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_CEList(nIndex).fiName & " = " & m_CEList(nIndex).fiNewData
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
   
   strSql = strSql & " " & _
                  "WHERE CE01 = '" & m_CE01 & "'"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
'   'Add By Sindy 2012/2/6
'   Else
'      frm03020401_04.m_blnClkChgButton = False
'   '2012/2/6 End
   End If
   
    'Modify By Cheng 2003/04/11
    '不在此更新基本檔
'   ' 更新商標基本檔
'   OnUpdateTradeMark
   
   ' 90.07.19 modify 不替前畫面儲存
   ' 直接儲存前畫面的資料
   'frm03020401_04.OnSaveData
   
 '911107 nick transation
  cnnConnection.CommitTrans
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
     'edit by nick 2004/11/03
End Function

Public Sub OnUpdateTradeMark()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strPS As String
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim bModifyCE09 As Boolean
      
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      ' 申請日
      If textCE03 = "1" Then
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM11")) = False Then: strTmp = rsTmp.Fields("TM11")
         AddTMField "TM11", strTmp, 1
         SetTMFieldNewData "TM11", m_CE02
      End If
        'Modify By Cheng 2003/03/07
        '不在此更新申請人資料
'      ' 申請人
'      bModifyCE09 = False
'      If textCE09 = "1" Then
'         strTmp = Empty
'         If IsNull(rsTmp.Fields("TM23")) = False Then: strTmp = rsTmp.Fields("TM23")
'         AddTMField "TM23", strTmp, 0
'         SetTMFieldNewData "TM23", m_CE04
'         '連帶變更申請人的地址(中, 英, 日)
'         bModifyCE09 = True
'         strTmp = Empty
'         If IsNull(rsTmp.Fields("TM24")) = False Then: strTmp = rsTmp.Fields("TM24")
'         AddTMField "TM24", strTmp, 0
'         SetTMFieldNewData "TM24", GetAddress(m_CE04, 0)
'         strTmp = Empty
'         If IsNull(rsTmp.Fields("TM25")) = False Then: strTmp = rsTmp.Fields("TM25")
'         AddTMField "TM25", strTmp, 0
'         SetTMFieldNewData "TM25", GetAddress(m_CE04, 1)
'         strTmp = Empty
'         If IsNull(rsTmp.Fields("TM26")) = False Then: strTmp = rsTmp.Fields("TM26")
'         AddTMField "TM26", strTmp, 0
'         SetTMFieldNewData "TM26", GetAddress(m_CE04, 2)
'      End If
      ' 申請地址
      If textCE38 = "1" And bModifyCE09 = False Then
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM24")) = False Then: strTmp = rsTmp.Fields("TM24")
         AddTMField "TM24", strTmp, 0
         SetTMFieldNewData "TM24", textCE23
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM25")) = False Then: strTmp = rsTmp.Fields("TM25")
         AddTMField "TM25", strTmp, 0
         SetTMFieldNewData "TM25", textCE24
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM26")) = False Then: strTmp = rsTmp.Fields("TM26")
         AddTMField "TM26", strTmp, 0
         SetTMFieldNewData "TM26", textCE25
      End If
      ' 代表人
      If textCE16 = "1" Then
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM47")) = False Then: strTmp = rsTmp.Fields("TM47")
         AddTMField "TM47", strTmp, 0
         SetTMFieldNewData "TM47", textCE10
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM48")) = False Then: strTmp = rsTmp.Fields("TM48")
         AddTMField "TM48", strTmp, 0
         SetTMFieldNewData "TM48", textCE11
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM49")) = False Then: strTmp = rsTmp.Fields("TM49")
         AddTMField "TM49", strTmp, 0
         SetTMFieldNewData "TM49", textCE12
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM50")) = False Then: strTmp = rsTmp.Fields("TM50")
         AddTMField "TM50", strTmp, 0
         SetTMFieldNewData "TM50", textCE13
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM51")) = False Then: strTmp = rsTmp.Fields("TM51")
         AddTMField "TM51", strTmp, 0
         SetTMFieldNewData "TM51", textCE14
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM52")) = False Then: strTmp = rsTmp.Fields("TM52")
         AddTMField "TM52", strTmp, 0
         SetTMFieldNewData "TM52", textCE15
      End If
      ' 商標種類代號
      If textCE40 = "1" Then
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM08")) = False Then: strTmp = rsTmp.Fields("TM08")
         AddTMField "TM08", strTmp, 0
         SetTMFieldNewData "TM08", m_CE39
      End If
      ' 商品類別
      If textCE48 = "1" Then
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM09")) = False Then: strTmp = rsTmp.Fields("TM09")
         AddTMField "TM09", strTmp, 0
         SetTMFieldNewData "TM09", textCE47
      End If
      ' 正商標號數
      If textCE58 = "1" Then
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM27")) = False Then: strTmp = rsTmp.Fields("TM27")
         AddTMField "TM27", strTmp, 0
         SetTMFieldNewData "TM27", textCE57
      End If
      ' 商標名稱
      If textCE44 = "1" Then
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            strTmp = Empty
            If IsNull(rsTmp.Fields("TM05")) = False Then: strTmp = rsTmp.Fields("TM05")
            AddTMField "TM05", strTmp, 0
            SetTMFieldNewData "TM05", textCE41_1
        Case Else
            strTmp = Empty
            If IsNull(rsTmp.Fields("TM05")) = False Then: strTmp = rsTmp.Fields("TM05")
            AddTMField "TM05", strTmp, 0
            SetTMFieldNewData "TM05", textCE41
            strTmp = Empty
            If IsNull(rsTmp.Fields("TM06")) = False Then: strTmp = rsTmp.Fields("TM06")
            AddTMField "TM06", strTmp, 0
            SetTMFieldNewData "TM06", textCE42
            strTmp = Empty
            If IsNull(rsTmp.Fields("TM07")) = False Then: strTmp = rsTmp.Fields("TM07")
            AddTMField "TM07", strTmp, 0
            SetTMFieldNewData "TM07", textCE43
        End Select
      End If
      ' 商品群組
      If textCE50 = "1" Then
         strTmp = Empty
         If IsNull(rsTmp.Fields("TM32")) = False Then: strTmp = rsTmp.Fields("TM32")
         AddTMField "TM32", strTmp, 0
         SetTMFieldNewData "TM32", textCE49
      End If
   End If
   rsTmp.Close
   ' 更新商標基本檔
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMCount - 1
      strTmp = Empty
      If m_TMList(nIndex).fiOldData <> m_TMList(nIndex).fiNewData Then
         If m_TMList(nIndex).fiType = 0 Then
            strTmp = m_TMList(nIndex).fiName & " = '" & m_TMList(nIndex).fiNewData & "'"
         Else
            If m_TMList(nIndex).fiNewData = Empty Then
               strTmp = m_TMList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_TMList(nIndex).fiName & " = " & m_TMList(nIndex).fiNewData
            End If
         End If
         ' 將更新的項目以原資料儲存到 strPS 中
         If m_TMList(nIndex).fiOldData <> Empty Then
            If strPS <> Empty Then: strPS = strPS & ","
            strPS = strPS & m_TMList(nIndex).fiOldData
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
   ' 組成SQL語法
   strSql = strSql & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "'"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
   End If
   
   ' 取得案件進度檔原備註欄位的內容
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CE01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CP64")) = False Then
         If IsEmptyText(rsTmp.Fields("CP64")) = False Then
            strPS = rsTmp.Fields("CP64") & "," & strPS
         End If
      End If
   End If
   rsTmp.Close
   ' 更新案件進度檔的進度備註欄位
   strSql = "UPDATE CaseProgress SET CP64 = '" & strPS & "' " & _
            "WHERE CP09 = '" & m_CE01 & "' "
   cnnConnection.Execute strSql
   
   ' 清除欄位
   ClearTMFields
   
   Set rsTmp = Nothing
End Sub

Private Function CheckIs1Or2(ByVal strData As String) As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckIs1Or2 = True
   If IsEmptyText(strData) = False Then
      Select Case strData
         Case "1", "2":
         Case Else
            CheckIs1Or2 = False
            strTit = "資料檢核"
            strMsg = "只可輸入1或2"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End Select
   End If
End Function

Private Sub textCE03_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE03) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE09_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE09) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE16_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE16) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE22_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE22) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE38_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE38) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE40_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE40) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE44_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE44) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE46_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE46) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE48_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE48) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE50_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE50) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE52_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE52) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE54_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE54) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE56_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE56) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE58_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE58) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE60_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE60) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE62_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE62) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE65_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE65) = False Then
      Cancel = True
   End If
End Sub

Private Sub textCE03_GotFocus()
   InverseTextBox textCE03
End Sub

Private Sub textCE09_GotFocus()
   InverseTextBox textCE09
End Sub

Private Sub textCE16_GotFocus()
   InverseTextBox textCE16
End Sub

Private Sub textCE22_GotFocus()
   InverseTextBox textCE22
End Sub

Private Sub textCE38_GotFocus()
   InverseTextBox textCE38
End Sub

Private Sub textCE40_GotFocus()
   InverseTextBox textCE40
End Sub

Private Sub textCE44_GotFocus()
   InverseTextBox textCE44
End Sub

Private Sub textCE46_GotFocus()
   InverseTextBox textCE46
End Sub

Private Sub textCE48_GotFocus()
   InverseTextBox textCE48
End Sub

Private Sub textCE50_GotFocus()
   InverseTextBox textCE50
End Sub

Private Sub textCE52_GotFocus()
   InverseTextBox textCE52
End Sub

Private Sub textCE54_GotFocus()
   InverseTextBox textCE54
End Sub

Private Sub textCE56_GotFocus()
   InverseTextBox textCE56
End Sub

Private Sub textCE58_GotFocus()
   InverseTextBox textCE58
End Sub

Private Sub textCE60_GotFocus()
   InverseTextBox textCE60
End Sub

Private Sub textCE62_GotFocus()
   InverseTextBox textCE62
End Sub

Private Sub textCE65_GotFocus()
   InverseTextBox textCE65
End Sub



