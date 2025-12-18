VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030101_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(查名, 申請, 延展, 補換發註冊)"
   ClientHeight    =   5660
   ClientLeft      =   2180
   ClientTop       =   1540
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5660
   ScaleWidth      =   9150
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "案件進度(&C)"
      Height          =   350
      Left            =   1296
      TabIndex        =   38
      Top             =   0
      Width           =   1272
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   350
      Index           =   1
      Left            =   2592
      TabIndex        =   39
      Top             =   0
      Width           =   1152
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   350
      Left            =   3768
      TabIndex        =   40
      Top             =   0
      Width           =   1152
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H008080FF&
      Caption         =   "變更事項(&R)"
      Height          =   350
      Left            =   4944
      Style           =   1  '圖片外觀
      TabIndex        =   41
      Top             =   0
      Width           =   1152
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Left            =   7056
      TabIndex        =   42
      Top             =   0
      Width           =   1152
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6096
      TabIndex        =   37
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8208
      TabIndex        =   43
      Top             =   0
      Width           =   912
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3795
      Left            =   180
      TabIndex        =   63
      Top             =   1860
      Width           =   8955
      _ExtentX        =   15804
      _ExtentY        =   6703
      _Version        =   393216
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料1"
      TabPicture(0)   =   "frm030101_03.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label25"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label22"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label23"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label9"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(7)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(8)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(10)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label10"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label15"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label16"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(12)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label11"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label12"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label13"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label37"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblCP113(18)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCP44_2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textTM23_2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textTM07"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textTM06"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textTM05"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textTM05_1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCP27"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textUargeDate"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textTM22"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textTM21"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textPrint"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textTM09"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textTM08"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textTM32"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textTM27"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCP18"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textPetition"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCP26"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCF09"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textPriorityDoc"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textTM23"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textCP44"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textTM08_2"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtCP113"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "基本資料2"
      TabPicture(1)   =   "frm030101_03.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textCP09_S3"
      Tab(1).Control(1)=   "textCP09_S2"
      Tab(1).Control(2)=   "textCP09_S1"
      Tab(1).Control(3)=   "textCP09_S"
      Tab(1).Control(4)=   "cmdPriority"
      Tab(1).Control(5)=   "textPrintTNT"
      Tab(1).Control(6)=   "textPrintLetter"
      Tab(1).Control(7)=   "textTM67"
      Tab(1).Control(8)=   "textTM58"
      Tab(1).Control(9)=   "textCP64"
      Tab(1).Control(10)=   "Line2"
      Tab(1).Control(11)=   "Label29"
      Tab(1).Control(12)=   "Label28"
      Tab(1).Control(13)=   "Label27"
      Tab(1).Control(14)=   "Label26"
      Tab(1).Control(15)=   "Label21"
      Tab(1).Control(16)=   "Label20"
      Tab(1).Control(17)=   "Label19"
      Tab(1).Control(18)=   "Label18(0)"
      Tab(1).Control(19)=   "Label17"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "代表人"
      TabPicture(2)   =   "frm030101_03.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Combo2(1)"
      Tab(2).Control(1)=   "Combo2(0)"
      Tab(2).Control(2)=   "textTM50"
      Tab(2).Control(3)=   "textTM47"
      Tab(2).Control(4)=   "textTM48"
      Tab(2).Control(5)=   "textTM49"
      Tab(2).Control(6)=   "textTM51"
      Tab(2).Control(7)=   "textTM52"
      Tab(2).Control(8)=   "Label18(2)"
      Tab(2).Control(9)=   "Label14(1)"
      Tab(2).Control(10)=   "Label5(3)"
      Tab(2).Control(11)=   "Label5(4)"
      Tab(2).Control(12)=   "Label5(5)"
      Tab(2).Control(13)=   "Label5(6)"
      Tab(2).Control(14)=   "Label5(7)"
      Tab(2).Control(15)=   "Label5(8)"
      Tab(2).ControlCount=   16
      Begin VB.TextBox txtCP113 
         Height          =   270
         Left            =   3660
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1509
         Width           =   540
      End
      Begin VB.TextBox textCP09_S3 
         Height          =   270
         Left            =   -71490
         MaxLength       =   2
         TabIndex        =   28
         Top             =   1230
         Width           =   465
      End
      Begin VB.TextBox textCP09_S2 
         Height          =   270
         Left            =   -71940
         MaxLength       =   1
         TabIndex        =   27
         Top             =   1230
         Width           =   345
      End
      Begin VB.TextBox textCP09_S1 
         Height          =   270
         Left            =   -73020
         MaxLength       =   6
         TabIndex        =   26
         Top             =   1230
         Width           =   975
      End
      Begin VB.TextBox textCP09_S 
         Height          =   270
         Left            =   -73560
         MaxLength       =   1
         TabIndex        =   25
         Top             =   1230
         Width           =   465
      End
      Begin VB.TextBox textTM08_2 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   1704
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   2952
         Width           =   2052
      End
      Begin VB.ComboBox textCP44 
         Height          =   260
         Left            =   1200
         TabIndex        =   4
         Top             =   912
         Width           =   1620
      End
      Begin VB.TextBox textTM23 
         Height          =   270
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   11
         Top             =   1812
         Width           =   1092
      End
      Begin VB.CommandButton cmdPriority 
         Caption         =   "輸入(&V)"
         Height          =   252
         Left            =   -73560
         TabIndex        =   24
         Top             =   960
         Width           =   1032
      End
      Begin VB.TextBox textPrintTNT 
         Height          =   270
         Left            =   -69240
         MaxLength       =   1
         TabIndex        =   22
         Top             =   360
         Width           =   372
      End
      Begin VB.TextBox textPrintLetter 
         Height          =   270
         Left            =   -73200
         MaxLength       =   1
         TabIndex        =   21
         Top             =   360
         Width           =   372
      End
      Begin VB.TextBox textPriorityDoc 
         Height          =   270
         Left            =   6480
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1812
         Width           =   372
      End
      Begin VB.TextBox textCF09 
         Height          =   270
         Left            =   5160
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1512
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.TextBox textCP26 
         Height          =   270
         Left            =   5760
         MaxLength       =   1
         TabIndex        =   7
         Top             =   1212
         Width           =   372
      End
      Begin VB.TextBox textPetition 
         Height          =   270
         Left            =   4860
         MaxLength       =   8
         TabIndex        =   3
         Top             =   600
         Width           =   1092
      End
      Begin VB.TextBox textCP18 
         Height          =   270
         Left            =   4860
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         Top             =   312
         Width           =   1545
      End
      Begin VB.TextBox textTM27 
         Height          =   270
         Left            =   5472
         MaxLength       =   20
         TabIndex        =   18
         Top             =   2952
         Width           =   2532
      End
      Begin VB.TextBox textTM32 
         Height          =   270
         Left            =   1200
         MaxLength       =   699
         TabIndex        =   20
         Top             =   3444
         Width           =   7632
      End
      Begin VB.TextBox textTM08 
         Height          =   270
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   17
         Top             =   2952
         Width           =   372
      End
      Begin VB.TextBox textTM09 
         Height          =   270
         Left            =   1200
         MaxLength       =   395
         TabIndex        =   19
         Top             =   3216
         Width           =   7632
      End
      Begin VB.TextBox textPrint 
         Height          =   270
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1512
         Width           =   372
      End
      Begin VB.TextBox textTM21 
         Height          =   270
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1212
         Width           =   852
      End
      Begin VB.TextBox textTM22 
         Height          =   270
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1212
         Width           =   852
      End
      Begin VB.TextBox textUargeDate 
         Height          =   270
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   2
         Top             =   600
         Width           =   1092
      End
      Begin VB.TextBox textCP27 
         Height          =   270
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   0
         Top             =   312
         Width           =   1092
      End
      Begin MSForms.TextBox textTM67 
         Height          =   285
         Left            =   -73590
         TabIndex        =   23
         Top             =   660
         Width           =   7395
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "13039;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   1
         Left            =   -73590
         TabIndex        =   109
         Top             =   1800
         Width           =   7485
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "13203;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   0
         Left            =   -73590
         TabIndex        =   108
         Top             =   420
         Width           =   7485
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "13203;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05_1 
         Height          =   825
         Left            =   1440
         TabIndex        =   13
         Top             =   2100
         Width           =   7395
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "13039;1455"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM50 
         Height          =   300
         Left            =   -73575
         TabIndex        =   34
         Top             =   2130
         Width           =   7392
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM47 
         Height          =   300
         Left            =   -73575
         TabIndex        =   31
         Top             =   750
         Width           =   7392
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   50
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM48 
         Height          =   300
         Left            =   -73575
         TabIndex        =   32
         Top             =   1080
         Width           =   7392
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM58 
         Height          =   885
         Left            =   -73560
         TabIndex        =   30
         Top             =   2700
         Width           =   7395
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13044;1561"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   945
         Left            =   -73560
         TabIndex        =   29
         Top             =   1560
         Width           =   7395
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13044;1667"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05 
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   2112
         Width           =   7392
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "13039;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM06 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   2388
         Width           =   7392
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "13039;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM07 
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   2664
         Width           =   7392
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "13039;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_2 
         Height          =   285
         Left            =   2400
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1812
         Width           =   1332
         VariousPropertyBits=   671105055
         ForeColor       =   -2147483641
         Size            =   "4471;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   285
         Left            =   2880
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   912
         Width           =   5964
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "4471;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM49 
         Height          =   300
         Left            =   -73590
         TabIndex        =   33
         Top             =   1410
         Width           =   7395
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "13044;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM51 
         Height          =   300
         Left            =   -73575
         TabIndex        =   35
         Top             =   2460
         Width           =   7392
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   80
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM52 
         Height          =   300
         Left            =   -73575
         TabIndex        =   36
         Top             =   2760
         Width           =   7392
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "13039;529"
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
         Left            =   2760
         TabIndex        =   110
         Top             =   1554
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人2"
         Height          =   180
         Index           =   2
         Left            =   -74415
         TabIndex        =   107
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人1"
         Height          =   180
         Index           =   1
         Left            =   -74415
         TabIndex        =   106
         Top             =   525
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   3
         Left            =   -74055
         TabIndex        =   105
         Top             =   750
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   4
         Left            =   -74055
         TabIndex        =   104
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   5
         Left            =   -74055
         TabIndex        =   103
         Top             =   1410
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   6
         Left            =   -74055
         TabIndex        =   102
         Top             =   2130
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   7
         Left            =   -74055
         TabIndex        =   101
         Top             =   2460
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   8
         Left            =   -74055
         TabIndex        =   100
         Top             =   2760
         Width           =   345
      End
      Begin VB.Label Label37 
         Caption         =   "案件名稱 :"
         Height          =   252
         Left            =   120
         TabIndex        =   99
         Top             =   2112
         Width           =   1332
      End
      Begin VB.Line Line2 
         X1              =   -73290
         X2              =   -71220
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Label Label29 
         Caption         =   "案件備註 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   96
         Top             =   2700
         Width           =   972
      End
      Begin VB.Label Label28 
         Caption         =   "進度備註 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   95
         Top             =   1560
         Width           =   972
      End
      Begin VB.Label Label27 
         Caption         =   "查名本所案號 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   94
         Top             =   1260
         Width           =   1245
      End
      Begin VB.Label Label26 
         Caption         =   "優先權資料 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   93
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label21 
         Caption         =   "放棄專用權 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   92
         Top             =   660
         Width           =   1092
      End
      Begin VB.Label Label20 
         Caption         =   "是否列印TNT :"
         Height          =   252
         Left            =   -70560
         TabIndex        =   91
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label19 
         Caption         =   "(Y:印)"
         Height          =   252
         Left            =   -68760
         TabIndex        =   90
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label18 
         Caption         =   "是否列印指示信 :"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   89
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "(N:不印)"
         Height          =   255
         Left            =   -72720
         TabIndex        =   88
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "(Y / N)"
         Height          =   252
         Left            =   6960
         TabIndex        =   87
         Top             =   1812
         Width           =   612
      End
      Begin VB.Label Label12 
         Caption         =   "是否附優先權證明文件 :"
         Height          =   252
         Left            =   4440
         TabIndex        =   86
         Top             =   1812
         Width           =   1932
      End
      Begin VB.Label Label11 
         Caption         =   "可接獲回音"
         Height          =   255
         Left            =   5880
         TabIndex        =   85
         Top             =   1515
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "大約"
         Height          =   255
         Index           =   12
         Left            =   4680
         TabIndex        =   84
         Top             =   1515
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label16 
         Caption         =   "是否算案件數 :"
         Height          =   252
         Left            =   4440
         TabIndex        =   83
         Top             =   1212
         Width           =   1212
      End
      Begin VB.Label Label15 
         Caption         =   "(N:不算)"
         Height          =   252
         Left            =   6240
         TabIndex        =   82
         Top             =   1212
         Width           =   972
      End
      Begin VB.Label Label10 
         Caption         =   "提申期限 :"
         Height          =   255
         Left            =   3780
         TabIndex        =   81
         Top             =   605
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "點數 :"
         Height          =   255
         Index           =   10
         Left            =   3780
         TabIndex        =   80
         Top             =   317
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "正商標號數:"
         Height          =   252
         Index           =   8
         Left            =   4392
         TabIndex        =   79
         Top             =   2976
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "商品組群 :"
         Height          =   252
         Index           =   5
         Left            =   264
         TabIndex        =   78
         Top             =   3480
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "商標種類 :"
         Height          =   252
         Index           =   4
         Left            =   264
         TabIndex        =   77
         Top             =   2928
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "商品類別 :"
         Height          =   252
         Index           =   7
         Left            =   264
         TabIndex        =   76
         Top             =   3216
         Width           =   852
      End
      Begin VB.Label Label9 
         Caption         =   "案件中文名稱 :"
         Height          =   252
         Left            =   120
         TabIndex        =   75
         Top             =   2112
         Width           =   1332
      End
      Begin VB.Label Label8 
         Caption         =   "案件英文名稱 :"
         Height          =   252
         Left            =   120
         TabIndex        =   74
         Top             =   2412
         Width           =   1212
      End
      Begin VB.Label Label7 
         Caption         =   "案件日文名稱 :"
         Height          =   252
         Left            =   120
         TabIndex        =   73
         Top             =   2640
         Width           =   1452
      End
      Begin VB.Label Label6 
         Caption         =   "申請人 :"
         Height          =   252
         Left            =   120
         TabIndex        =   72
         Top             =   1812
         Width           =   852
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不印)"
         Height          =   252
         Left            =   1680
         TabIndex        =   70
         Top             =   1512
         Width           =   852
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   252
         Left            =   120
         TabIndex        =   69
         Top             =   1512
         Width           =   972
      End
      Begin VB.Label Label5 
         Caption         =   "延展後專用期限 :"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   68
         Top             =   1212
         Width           =   1452
      End
      Begin VB.Line Line1 
         X1              =   2640
         X2              =   2760
         Y1              =   1332
         Y2              =   1332
      End
      Begin VB.Label Label4 
         Caption         =   "代理人 :"
         Height          =   252
         Left            =   120
         TabIndex        =   66
         Top             =   912
         Width           =   972
      End
      Begin VB.Label Label14 
         Caption         =   "催審期限 :"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   65
         Top             =   612
         Width           =   972
      End
      Begin VB.Label Label25 
         Caption         =   "發文日 :"
         Height          =   252
         Left            =   120
         TabIndex        =   64
         Top             =   312
         Width           =   852
      End
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5424
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   372
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5424
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   672
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5424
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   972
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1344
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   672
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1344
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   372
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1344
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1272
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5424
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1272
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1344
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   972
      Width           =   2532
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1344
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   1572
      Width           =   2532
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5424
      TabIndex        =   111
      Top             =   1572
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label36 
      Caption         =   "S商品類別輸在""案件備註""欄!!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   8100
      TabIndex        =   98
      Top             =   510
      Width           =   930
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   264
      TabIndex        =   62
      Top             =   1572
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "業務區別 :"
      Height          =   252
      Index           =   2
      Left            =   4464
      TabIndex        =   60
      Top             =   372
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "發證日 :"
      Height          =   252
      Index           =   3
      Left            =   4464
      TabIndex        =   58
      Top             =   672
      Width           =   732
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   4464
      TabIndex        =   56
      Top             =   972
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   264
      TabIndex        =   54
      Top             =   672
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   264
      TabIndex        =   53
      Top             =   372
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   264
      TabIndex        =   52
      Top             =   1272
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   252
      Index           =   9
      Left            =   4464
      TabIndex        =   51
      Top             =   1272
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4464
      TabIndex        =   50
      Top             =   1572
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   252
      Left            =   264
      TabIndex        =   49
      Top             =   972
      Width           =   732
   End
End
Attribute VB_Name = "frm030101_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/10 改成Form2.0 ; textCP13、 textCP14、textCP44_2、textTM23_2、textCP64、textTM58
                                 '、textTM05、textTM06、textTM07、Combo2(Index)、textTM47、textTM48、textTM49、textTM50、textTM51、textTM52、textTM67(111/8/8)
'end 2021/08/10
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/10 日期欄已修改
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
' 申請國家
Dim m_TM10 As String
' 案件性質代號
Dim m_CP10 As String
' 承辦人代號
Dim m_CP14 As String
' 原專用期限起日
Dim m_TM21 As String
' 原專用期限止日
Dim m_TM22 As String
' 申請人1
Dim m_TM23 As String
Dim m_TM78 As String ' 申請人2
Dim m_TM79 As String ' 申請人3
Dim m_TM80 As String ' 申請人4
Dim m_TM81 As String ' 申請人5
' 申請國家的延展年度
Dim m_NA14 As String
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
' 優先權畫面所使用的變數
Dim m_Pa(1 To 4) As String '本所案號
Dim m_Priority(1 To 6) As String
'Add By Cheng 2002/08/23
Dim m_strCust1 As String '申請人1
'Add By Cheng 2003/10/06
Public m_blnClkChgButton As Boolean '是否按下變更事項按鈕
'Add By Cheng 2004/05/17
Dim m_blnOutGoingMsg107 As Boolean
'End
'add by nickc 2005/11/18
Dim m_TM24 As String
Dim m_tm25 As String
Dim m_tm26 As String
'add by nickc 2006/07/03 判斷用，因為原先已經被使用且更改值過了
' 原專用期限起日
Dim mm_TM21 As String
' 原專用期限止日
Dim mm_TM22 As String
'add by nickc 2006/09/07
Dim m_CP13 As String
'add by nickc 2008/02/22
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP07 As String 'Add By Sindy 2012/5/3
Dim m_990CP09 As String 'Add By Sindy 2016/12/20


Private Sub cmdCancel_Click()
   frm030101_01.Show
   Unload Me
End Sub

Private Sub cmdCaseProgress_Click()
   frm030101_04.SetData 0, m_TM01, True
   frm030101_04.SetData 1, m_TM02, False
   frm030101_04.SetData 2, m_TM03, False
   frm030101_04.SetData 3, m_TM04, False
   frm030101_04.SetData 4, m_CP09, False
   frm030101_04.SetParent Me
   Me.Hide
   frm030101_04.Show
   frm030101_04.QueryData
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm030101_01
   Unload Me
End Sub

Private Sub cmdMod_Click()
   frm030101_05.SetData 0, m_TM01, True
   frm030101_05.SetData 1, m_TM02, False
   frm030101_05.SetData 2, m_TM03, False
   frm030101_05.SetData 3, m_TM04, False
   frm030101_05.SetData 4, m_CP09, False
   frm030101_05.SetParent "frm030101_03"
   'Me.Hide
   frm030101_05.Show
   frm030101_05.QueryData
'    m_blnClkChgButton = True
End Sub

Private Sub cmdok_Click(Index As Integer)
'Add By Sindy 2011/1/26
Dim strApplID As String
Dim rsAddrNotAlike As New ADODB.Recordset
'2011/1/26 End
   
   'Modify By Sindy 2010/11/19 把「確定」及「同時發文」按鈕程式碼合併
   Select Case Index
      Case 0, 1
         If CheckDataValid = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            ' 設定滑鼠游標為等待狀態
            Screen.MousePointer = vbHourglass
            ' 更新欄位輸入的內容
            OnUpdateField
            ' 存檔
            'edit by nick 2004/11/03
            'OnSaveData
            If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
            ' 設定滑鼠游標為預設
            Screen.MousePointer = vbDefault
            
            'Added by Lydia 2016/12/19 CFT申請或分割案若抓不到審查時間或CF05=0時，仍可存檔但要發E-MAIL給特殊人員(V)外商陳經理
            If m_TM01 = "CFT" And (m_CP10 = "101" Or m_CP10 = "308") Then
               Call PUB_SetChkResultDateT(m_TM01, m_TM10, m_CP10, "19221111", strExc(5), m_TM02, m_TM03, m_TM04)
            End If
            
            'Add By Sindy 2011/1/26 檢查相同國家若有舊案申請地址與客戶目前申請地址不同者
            strApplID = ""
            If Trim(m_TM23) <> "" Then '申請人1
               If strApplID <> "" Then strApplID = strApplID & ","
               strApplID = strApplID & "'" & Trim(m_TM23) & "'"
            End If
            If Trim(m_TM78) <> "" Then '申請人2
               If strApplID <> "" Then strApplID = strApplID & ","
               strApplID = strApplID & "'" & Trim(m_TM78) & "'"
            End If
            If Trim(m_TM79) <> "" Then '申請人3
               If strApplID <> "" Then strApplID = strApplID & ","
               strApplID = strApplID & "'" & Trim(m_TM79) & "'"
            End If
            If Trim(m_TM80) <> "" Then '申請人4
               If strApplID <> "" Then strApplID = strApplID & ","
               strApplID = strApplID & "'" & Trim(m_TM80) & "'"
            End If
            If Trim(m_TM81) <> "" Then '申請人5
               If strApplID <> "" Then strApplID = strApplID & ","
               strApplID = strApplID & "'" & Trim(m_TM81) & "'"
            End If
            If ChkOCaseAndCAddrNotAlike(strApplID, m_TM10, m_TM01, m_CP10, rsAddrNotAlike, False) = True Then
               Set frm880018.fmParent = Me
               Set frm880018.RsTemp = rsAddrNotAlike
               frm880018.m_Appl1 = Trim(m_TM23)
               frm880018.m_Appl2 = Trim(m_TM78)
               frm880018.m_Appl3 = Trim(m_TM79)
               frm880018.m_Appl4 = Trim(m_TM80)
               frm880018.m_Appl5 = Trim(m_TM81)
               frm880018.Show vbModal
            End If
            '2011/1/26 End
            
            'Add By Sindy 2012/4/5 CFT,FCT所有案件性質發文時,檢查代表圖是否存在
            'Modify by Amy 2018/07/31 ChkIsExistImg不使用
            'Call ChkIsExistImg(m_TM01, m_TM02, m_TM03, m_TM04)
            If ChkImgByteFile(m_TM01, m_TM02, m_TM03, m_TM04) = False Then MsgBox "本案尚未放代表圖至系統！"
            
            'Add By Sindy 2024/8/19
            If frm030101_01.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2024/8/19 End
            If Index = 0 Then '確定鍵
               '*********** 90.11.23   nick  清畫面
               'frm030101_01.radio(0).Value = True
               'frm030101_01.textCP09.Enabled = True
               'frm030101_01.textCP09.Text = ""
               'frm030101_01.textTM01.Enabled = False
               'frm030101_01.textTM01.Text = ""
               'frm030101_01.textTM02.Enabled = False
               'frm030101_01.textTM02.Text = ""
               'frm030101_01.textTM02_2.Enabled = False
               'frm030101_01.textTM02_2.Text = ""
               'frm030101_01.textTM03.Enabled = False
               'frm030101_01.textTM03.Text = "'"
               'frm030101_01.textTM04.Enabled = False
               'frm030101_01.textTM04.Text = ""
               'frm030101_01.grdList.Clear
               'frm030101_01.grdList.Rows = 2
               'frm030101_01.RefreshData
               '***********************************
               'Add By Cheng 2002/04/30
               '若有未發文資料顯示警告
               If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = True Then
                  frm030101_01.Show
                  ' 90.12.07 modify by louis
                  'frm030101_01.Clear
                  'Add By Cheng 2002/01/10
                  frm030101_01.Clear1
               Else
                  'Add By Sindy 2024/8/19
                  If frm030101_01.bolIsEMPFlow = True Then
                     Unload frm030101_01
                     frm090202_4.Show
                  Else
                  '2024/8/19 End
                     frm030101_01.Show
                     frm030101_01.Clear1
                  End If
               End If
               Unload Me
            ElseIf Index = 1 Then '同時發文鍵
               ' 呼叫第一個畫面
               frm030101_01.SetData 0, m_TM01, True
               frm030101_01.SetData 1, m_TM02, False
               frm030101_01.SetData 2, m_TM03, False
               frm030101_01.SetData 3, m_TM04, False
               frm030101_01.SetQueryFromTM
               Unload Me
               frm030101_01.Show
               frm030101_01.radio(1).Value = True
               frm030101_01.radio_Click 1
               frm030101_01.QueryData
            End If
         End If
      Case Else
   End Select
End Sub

Private Sub cmdPriority_Click()
   ' 修改優先權資料
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify By Sindy 2017/10/12 + , m_Priority(6)
   'Modify by Sindy 2019/1/23 + m_TM01 & m_TM02 & m_TM03 & m_TM04
   ModifyPriority m_Priority(1), m_Priority(2), m_Priority(3), , , m_TM01 & m_TM02 & m_TM03 & m_TM04, , , m_Priority(4), m_Priority(5), m_Priority(6)
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

'Private Sub cmdTogether_Click()
'   If CheckDataValid = True Then
'      'Add By Cheng 2002/07/15
'      '重新檢查欄位有效性
'      If TxtValidate = False Then Exit Sub
'      ' 設定滑鼠游標為等待狀態
'      Screen.MousePointer = vbHourglass
'      ' 更新欄位輸入的內容
'      OnUpdateField
'      ' 存檔
'      'edit by nick 2004/11/03
'      'OnSaveData
'      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
'
'      ' 設定滑鼠游標為預設
'      Screen.MousePointer = vbDefault
'      ' 呼叫第一個畫面
'      frm030101_01.SetData 0, m_TM01, True
'      frm030101_01.SetData 1, m_TM02, False
'      frm030101_01.SetData 2, m_TM03, False
'      frm030101_01.SetData 3, m_TM04, False
'      frm030101_01.SetQueryFromTM
'      Unload Me
'      frm030101_01.Show
'      frm030101_01.radio(1).Value = True
'      frm030101_01.radio_Click 1
'      frm030101_01.QueryData
'   End If
'End Sub

'Private Sub Form_Activate()
    'Add By Cheng 2003/10/06
    '若有按下變更事項按鈕, 則重新讀取資料
    'edit by nickc 2005/08/23
    'If m_blnClkChgButton = True Then
'Modify By Sindy 2012/10/1 下列程式無意義Mark
'    If m_blnClkChgButton = True Or (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
'        pub_ModifyCaseNum = ""
'        QueryData
''        m_blnClkChgButton = False
'    End If
'End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08_2.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23_2.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F

   'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    textTM47.MaxLength = Pub_MaxCEL10
    textTM48.MaxLength = Pub_MaxCEL11
    textTM50.MaxLength = Pub_MaxCEL10
    textTM51.MaxLength = Pub_MaxCEL11
   'end 2016/09/10
   
   SSTab1.Tab = 0

   MoveFormToCenter Me
'    m_blnClkChgButton = False
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
      'Modify By Cheng 2002/09/18
'      ' 查名總收文號
'      Case 99: textCP09S = strData
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
      m_TM119 = CheckStr(rsTmp.Fields("TM119"))
      m_TM120 = CheckStr(rsTmp.Fields("TM120"))
      
      ' 審定號數
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         textTM20 = rsTmp.Fields("TM20")
      End If
'      ' 案件中文名稱
'      textTM05 = Empty
'      If IsNull(rsTmp.Fields("TM05")) = False Then
'         textTM05 = rsTmp.Fields("TM05")
'      End If
'      SetTMSPFieldOldData "TM05", textTM05, 0
      ' 案件名稱
      textTM05_1 = Empty
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textTM05_1 = rsTmp.Fields("TM05")
      End If
      SetTMSPFieldOldData "TM05", textTM05_1, 0
'      ' 案件英文名稱
'      textTM06 = Empty
'      If IsNull(rsTmp.Fields("TM06")) = False Then
'         textTM06 = rsTmp.Fields("TM06")
'      End If
'      SetTMSPFieldOldData "TM06", textTM06, 0
'      ' 案件日文名稱
'      textTM07 = Empty
'      If IsNull(rsTmp.Fields("TM07")) = False Then
'         textTM07 = rsTmp.Fields("TM07")
'      End If
'      SetTMSPFieldOldData "TM07", textTM07, 0
      ' 商標種類
      textTM08 = Empty
      If IsNull(rsTmp.Fields("TM08")) = False Then
         textTM08 = rsTmp.Fields("TM08")
         textTM08_2 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      End If
      SetTMSPFieldOldData "TM08", textTM08, 0
      ' 商品類別
      textTM09 = Empty
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      SetTMSPFieldOldData "TM09", textTM09, 0
      ' 申請國家
      'Add By Cheng 2002/07/18
      m_NA14 = Empty
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         m_NA14 = GetNationExtentYear(m_TM10)
      End If
      ' 專用期限起日
      'Add By Cheng 2002/07/18
      m_TM21 = Empty
      If IsNull(rsTmp.Fields("TM21")) = False Then
         m_TM21 = rsTmp.Fields("TM21")
      End If
      ' 專用期限止日
      'Add By Cheng 2002/07/18
      m_TM22 = Empty
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_TM22 = rsTmp.Fields("TM22")
      End If
      'add by nickc 2006/07/03
      mm_TM21 = m_TM21
      mm_TM22 = m_TM22
      
      ' 申請人1
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = rsTmp.Fields("TM23")
         textTM23_2 = GetCustomerName(textTM23, 0)
      End If
      'Add By Cheng 2002/08/23
      m_strCust1 = "" & Me.textTM23.Text
      
      'add by nickc 2005/11/18
      ' 中文地址
      m_TM24 = ""
      If IsNull(rsTmp.Fields("TM24")) = False Then
         m_TM24 = rsTmp.Fields("TM24")
      End If
      SetTMSPFieldOldData "TM24", m_TM24, 0
      ' 英文地址
      m_tm25 = ""
      If IsNull(rsTmp.Fields("TM25")) = False Then
         m_tm25 = rsTmp.Fields("TM25")
      End If
      SetTMSPFieldOldData "TM25", m_tm25, 0
      ' 日文地址
      m_tm26 = ""
      If IsNull(rsTmp.Fields("TM26")) = False Then
         m_tm26 = rsTmp.Fields("TM26")
      End If
      SetTMSPFieldOldData "TM26", m_tm26, 0
      
      'Add By Sindy 2011/1/26
      ' 申請人2
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = rsTmp.Fields("TM78")
      End If
      ' 申請人3
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = rsTmp.Fields("TM79")
      End If
      ' 申請人4
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = rsTmp.Fields("TM80")
      End If
      ' 申請人4
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = rsTmp.Fields("TM81")
      End If
      '2011/1/26 End
      
      ' 正商標號數
      textTM27 = Empty
      If IsNull(rsTmp.Fields("TM27")) = False Then
         textTM27 = rsTmp.Fields("TM27")
      End If
      SetTMSPFieldOldData "TM27", textTM27, 0
      ' 商品群組
      textTM32 = Empty
      If IsNull(rsTmp.Fields("TM32")) = False Then
         textTM32 = rsTmp.Fields("TM32")
      End If
      SetTMSPFieldOldData "TM32", textTM32, 0
      
      'Morgan 2003/11/20
      '代表人
      Dim i As Integer, j As Integer
      For i = 0 To 1
         Combo2(i).AddItem ""
      Next
      
      If rsTmp.Fields("TM23").Value <> "" Then
         strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM23").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 不用 dll 了   = objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(0).AddItem rsTmp.Fields("TM23").Value & "-" & j & strExc(0)
               Combo2(1).AddItem rsTmp.Fields("TM23").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      
      'Morgan 2003/11/20 -- end
      
      ' 代表人1(中)
      textTM47 = Empty
      If IsNull(rsTmp.Fields("TM47")) = False Then
         textTM47 = rsTmp.Fields("TM47")
      End If
      SetTMSPFieldOldData "TM47", textTM47, 0
      ' 代表人1(英)
      textTM48 = Empty
      If IsNull(rsTmp.Fields("TM48")) = False Then
         textTM48 = rsTmp.Fields("TM48")
      End If
      SetTMSPFieldOldData "TM48", textTM48, 0
      ' 代表人1(日)
      textTM49 = Empty
      If IsNull(rsTmp.Fields("TM49")) = False Then
         textTM49 = rsTmp.Fields("TM49")
      End If
      SetTMSPFieldOldData "TM49", textTM49, 0
      ' 代表人2(中)
      textTM50 = Empty
      If IsNull(rsTmp.Fields("TM50")) = False Then
         textTM50 = rsTmp.Fields("TM50")
      End If
      SetTMSPFieldOldData "TM50", textTM50, 0
      ' 代表人2(英)
      textTM51 = Empty
      If IsNull(rsTmp.Fields("TM51")) = False Then
         textTM51 = rsTmp.Fields("TM51")
      End If
      SetTMSPFieldOldData "TM51", textTM51, 0
      ' 代表人2(日)
      textTM52 = Empty
      If IsNull(rsTmp.Fields("TM52")) = False Then
         textTM52 = rsTmp.Fields("TM52")
      End If
      SetTMSPFieldOldData "TM52", textTM52, 0
      ' 案件備註
      textTM58 = Empty
      If IsNull(rsTmp.Fields("TM58")) = False Then
         textTM58 = rsTmp.Fields("TM58")
      End If
      SetTMSPFieldOldData "TM58", textTM58, 0
      ' 放棄專用權
      textTM67 = Empty
      If IsNull(rsTmp.Fields("TM67")) = False Then
         textTM67 = rsTmp.Fields("TM67")
      End If
      SetTMSPFieldOldData "TM67", textTM67, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub
'Morgan 2003/11/20
Private Sub Combo2_Click(Index As Integer)

   Dim i As Integer, strTmp As String
   
   If (Combo2(Index).Text = "") Then
      For i = 0 To 2
         Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = ""
      Next i
      Exit Sub
   End If
   
   strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
   strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
   strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 不用 dll 了   = objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
   
      For i = 0 To 2
         
         If Not IsNull(RsTemp.Fields(i)) Then
            Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = RsTemp.Fields(i)
         End If
         
      Next
   End If
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
        
      ' 案件中文名稱
      textTM05 = Empty
        Select Case m_TM01
        Case "S"
            If IsNull(rsTmp.Fields("SP05")) = False Then
               textTM05_1 = "" & rsTmp.Fields("SP05")
            End If
            SetTMSPFieldOldData "SP05", textTM05_1, 0
        Case Else
            If IsNull(rsTmp.Fields("SP05")) = False Then
               textTM05 = rsTmp.Fields("SP05")
            End If
            SetTMSPFieldOldData "SP05", textTM05, 0
        End Select
      ' 案件英文名稱
      textTM06 = Empty
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textTM06 = rsTmp.Fields("SP06")
      End If
      SetTMSPFieldOldData "SP06", textTM06, 0
      ' 案件日文名稱
      textTM07 = Empty
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textTM07 = rsTmp.Fields("SP07")
      End If
      SetTMSPFieldOldData "SP07", textTM07, 0
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
         textTM23 = rsTmp.Fields("TM23")
         textTM23_2 = GetCustomerName(textTM23, 0)
      End If
      'Add By Cheng 2002/08/23
      m_strCust1 = "" & Me.textTM23.Text
      
      'Add By Sindy 2011/1/26
      ' 申請人2
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("SP58")) = False Then
         m_TM78 = rsTmp.Fields("SP58")
      End If
      ' 申請人3
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("SP59")) = False Then
         m_TM79 = rsTmp.Fields("SP59")
      End If
      ' 申請人4
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("SP65")) = False Then
         m_TM80 = rsTmp.Fields("SP65")
      End If
      ' 申請人4
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("SP66")) = False Then
         m_TM81 = rsTmp.Fields("SP66")
      End If
      '2011/1/26 End
      
      ' 申請國家
      'Add By Cheng 2002/07/18
      m_NA14 = Empty
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
         m_NA14 = GetNationExtentYear(m_TM10)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      ' 專用期限起日
      'Add By Cheng 2002/07/18
      m_TM21 = Empty
      If IsNull(rsTmp.Fields("SP20")) = False Then
         m_TM21 = rsTmp.Fields("SP20")
      End If
      ' 專用期限止日
      'Add By Cheng 2002/07/18
      m_TM22 = Empty
      If IsNull(rsTmp.Fields("SP21")) = False Then
         m_TM22 = rsTmp.Fields("SP21")
      End If
      'add by nickc 2006/07/03
      mm_TM21 = m_TM21
      mm_TM22 = m_TM22
      ' 案件備註
      textTM58 = Empty
      If IsNull(rsTmp.Fields("SP18")) = False Then
         textTM58 = rsTmp.Fields("SP18")
      End If
      SetTMSPFieldOldData "SP18", textTM58, 0
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
   Dim strCP45 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      
      ' 案件性質
      'Add By Cheng 2002/07/18
      m_CP10 = Empty: m_CP14 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 業務區別
      If IsNull(rsTmp.Fields("CP12")) = False Then
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      'add by nickc 2006/09/07
      m_CP13 = ""
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
         m_CP13 = rsTmp.Fields("CP13")
      End If
      
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         '92.10.6 ADD BY SONIA
         m_CP14 = rsTmp.Fields("CP14")
         '92'10'6 END
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 發文日(預設為系統日)
      textCP27 = strSrvDate(1)
      strCP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         strCP27 = rsTmp.Fields("CP27")
      End If
      SetCPFieldOldData "CP27", strCP27, 1
      
      'Add By Sindy 2012/5/3
      '法定期限
      m_CP07 = Empty
      If IsNull(rsTmp.Fields("CP07")) = False Then: m_CP07 = rsTmp.Fields("CP07")
      '2012/5/3 End
      
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = rsTmp.Fields("CP44")
         SetCPFieldOldData "CP44", textCP44, 0 'Modify By Sindy 2013/5/23
      Else
         SetCPFieldOldData "CP44", "", 0 'Modify By Sindy 2013/5/23
      End If
      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      'Add By Sindy 2013/5/23
      If IsNull(rsTmp.Fields("CP116")) = False Then
         textCP44 = textCP44 & "-" & rsTmp.Fields("CP116")
         SetCPFieldOldData "CP116", rsTmp.Fields("CP116"), 0
      Else
         SetCPFieldOldData "CP116", "", 0
      End If
      '2013/5/23 End
      
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      SetCPFieldOldData "CP18", textCP18, 0
      ' 是否算案件數
      textCP26 = Empty
      If IsNull(rsTmp.Fields("CP26")) = False Then
         textCP26 = rsTmp.Fields("CP26")
      End If
      SetCPFieldOldData "CP26", textCP26, 0
      ' 彼所案號
      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         strCP45 = rsTmp.Fields("CP45")
      End If
      SetCPFieldOldData "CP45", strCP45, 0
      'add by nickc 2006/07/03 延展時，要更新專用期限
      'modify by sonia 2025/9/4 +109緩審延展CFT-016520
      If m_CP10 = "102" Or m_CP10 = "109" Then
        SetCPFieldOldData "CP53", CheckStr(rsTmp.Fields("CP53")), 1
        SetCPFieldOldData "CP54", CheckStr(rsTmp.Fields("CP54")), 1
      End If
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
      ' 代理人
      ClearAgentList
      'Add By Sindy 2013/5/23 若是原先有，也要加入
      If textCP44.Text <> "" Then
         If InStr(textCP44, "-") > 0 Then
            If ClsPDGetContact(textCP44, strCP44) Then
               AddAgent textCP44, strCP44
            End If
         Else
            strCP44 = GetFAgentName(textCP44)
            AddAgent textCP44, strCP44
         End If
      End If
      '2013/5/23 End
      '2009/2/3 modify by sonia B類收文之文件簽證711及申請英文證明304不要列入
      '2010/9/7 Modify by Sindy 文件簽證711及申請英文證明304不要列入
      'Modify By Sindy 2013/5/23 加聯絡人判斷
      strSubSQL = "SELECT CP44,CP116,MAX(CP27) AS CP27 FROM CASEPROGRESS " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null " & _
                        "AND CP10 NOT IN ('711','304') " & _
                  "GROUP BY CP44,CP116 " & _
                  "ORDER BY CP27 DESC "
      rsSubTmp.CursorLocation = adUseClient
      rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSubTmp.RecordCount > 0 Then
         rsSubTmp.MoveFirst
         ' 依序將代理人加入到系統串列中
         Do While rsSubTmp.EOF = False
            If IsNull(rsSubTmp.Fields("CP44")) = False Then
               'Modify By Sindy 2013/5/23
               If IsNull(rsSubTmp.Fields("CP116")) = False Then
                  If ClsPDGetContact(rsSubTmp.Fields("CP44") & "-" & rsSubTmp.Fields("CP116"), strCP44) Then
                     AddAgent rsSubTmp.Fields("CP44") & "-" & rsSubTmp.Fields("CP116"), strCP44
                  End If
               Else
               '2013/5/23 End
                  strCP44 = GetFAgentName(rsSubTmp.Fields("CP44"))
                  AddAgent rsSubTmp.Fields("CP44"), strCP44
               End If
            End If
            rsSubTmp.MoveNext
         Loop
      End If
      rsSubTmp.Close
      ' 從系統串列中取得所有代理人並放入Combo Box中
      For nIndex = 0 To m_AgentCount - 1
         textCP44.AddItem m_AgentList(nIndex).aiCode
      Next nIndex
      ' 設定顯示為第一筆
      If textCP44.ListCount > 0 Then
         textCP44.ListIndex = 0
         textCP44_Validate False
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strDay As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
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
    'Add By Cheng 2003/11/11
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF", "S"
        Me.Label9.Visible = False
        Me.textTM05.Visible = False
        Me.textTM05.Enabled = False
        Me.Label8.Visible = False
        Me.textTM06.Visible = False
        Me.textTM06.Enabled = False
        Me.Label7.Visible = False
        Me.textTM07.Visible = False
        Me.textTM07.Enabled = False
    Case Else
        Me.Label37.Visible = False
        Me.textTM05_1.Visible = False
        Me.textTM05_1.Enabled = False
    End Select
    'End
   ' 本所案號
'   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)

   ' 收文號
   textCP09 = m_CP09
   
   ' 取得國家代碼
   m_TM10 = GetNationNo(m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 取得案件進度檔的欄位
   QueryCaseProgress
   
   ' 取得基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "CFT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   
   ' 大約?可接獲回音(欄位)
   textCF09 = Empty
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
'CANCEL BY SONIA 2022/8/18 固定改１星期左右，寫在定稿裡
'      If IsNull(rsTmp.Fields("CF09")) = False Then
'         textCF09 = rsTmp.Fields("CF09")
'      End If
'END 2022/8/18
      'add by nickc 2005/08/31 提申期限
      If textPetition.Text = "" And textCP27.Text <> "" And IsNull(rsTmp.Fields("CF11").Value) = False Then
         textPetition.Text = DBDATE(ChangeWStringToTString(CompDate(2, rsTmp.Fields("CF11").Value, ChangeTStringToWString(textCP27))))
      End If
      'Add By Sindy 2019/6/11 檢查期限是否正確
      textPetition.Text = DBDATE(PUB_T997998LimitDate(textPetition.Text, m_CP07, 2))
      '2019/6/11 END
   End If
   rsTmp.Close
   
   ' 計算催審期限
   strDay = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
'   If IsEmptyText(strDay) = False Then
      textUargeDate = strDay
'   End If
   textCP27.Tag = textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   
   ' 案件性質為延展時才可輸入延展後專用期限
   'modify by sonia 2025/9/4 +109緩審延展CFT-016520
   If m_CP10 = "102" Or m_CP10 = "109" Then
      textTM21.BackColor = &H80000005
      textTM21.Locked = False
      textTM21.TabStop = True
      textTM22.BackColor = &H80000005
      textTM22.Locked = False
      textTM22.TabStop = True
   Else
      textTM21.BackColor = &H8000000F
      textTM21.Locked = True
      textTM21.TabStop = False
      textTM22.BackColor = &H8000000F
      textTM22.Locked = True
      textTM22.TabStop = False
   End If
   
   ' 讀取優先權資料
   m_Pa(1) = m_TM01
   m_Pa(2) = m_TM02
   m_Pa(3) = m_TM03
   m_Pa(4) = m_TM04
   
   'edit by nickc 2007/02/06 不用 dll 了 objPublicData.ReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify By Sindy 2017/10/12 + , m_Priority(6)
   ClsPDReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)
   Set rsTmp = Nothing
    'Add By Cheng 2002/11/05
    '延展發文預設延展後專用期限
    If m_CP10 = "102" Then
        If "" & m_TM22 <> "" And "" & m_NA14 <> "" Then
            'Modify By Cheng 2002/11/29
'             Me.textTM21.Text = DBDATE(DateSerial(Val(DBYEAR(m_TM22)), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22)) + 1))
'             Me.textTM22.Text = DBDATE(DateSerial(Val(DBYEAR(m_TM22)) + Val(m_NA14), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22)) + 1))
            'Modify By Cheng 2003/09/02
'             Me.textTM21.Text = DBDATE(DateSerial(Val(DBYEAR(m_TM22)), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22))))
'             Me.textTM22.Text = DBDATE(DateSerial(Val(DBYEAR(m_TM22)) + Val(m_NA14), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22))))
             Me.textTM21.Text = DBDATE(m_TM22)
             'Modified by Lydia 2019/11/13 2019/11/13 改用共用模組, 第1次專用期間=公告日+10年-1天，之後延展102沒有減１天；與專利不一樣
             'Me.textTM22.Text = DBDATE(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
             'Modify By Sindy 2022/3/7 + m_TM10 : 延展後之專用期限年度倘有2月29日時，專用期限止日應為2月29日，而非以加10年之方式計算為2月28日
             Me.textTM22.Text = PUB_GetEndDate(DBDATE(m_TM22), Val(m_NA14), "N", m_TM10)
             '91.12.26 ADD BY SONIA
             m_TM21 = textTM21: m_TM22 = textTM22
        End If
    'add by sonia 2025/9/5 +109緩審延展CFT-016520
    ElseIf m_CP10 = "109" Then
      Me.textTM21.Text = DBDATE(m_CP07)
      Me.textTM22.Text = PUB_GetEndDate(DBDATE(m_CP07), Val(m_NA14), "N", m_TM10)
      m_TM21 = textTM21: m_TM22 = textTM22
    'end 2025/9/5
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Add By Cheng 2002/07/19
   Set frm030101_03 = Nothing
End Sub

Private Sub textCP09_S_GotFocus()
   'Add By Cheng 2002/09/18
   InverseTextBox Me.textCP09_S
End Sub

Private Sub textCP09_S_KeyPress(KeyAscii As Integer)
   'Add By Cheng 2002/09/17
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP09_S_Validate(Cancel As Boolean)
   'Add By Cheng 2002/09/18
   Cancel = False
   If Me.textCP09_S.Text <> "" Then
      If Me.textCP09_S.Text <> "S" Then
         MsgBox "查名本所案號的系統類別類輸入錯誤!!!", vbExclamation + vbOKOnly
         Cancel = True
         Me.textCP09_S.SetFocus
         textCP09_S_GotFocus
      End If
   End If
End Sub

Private Sub textCP09_S1_GotFocus()
   'Add By Cheng 2002/09/18
   InverseTextBox Me.textCP09_S1
End Sub

Private Sub textCP09_S2_GotFocus()
   'Add By Cheng 2002/09/18
   InverseTextBox Me.textCP09_S2
End Sub

Private Sub textCP09_S3_GotFocus()
   'Add By Cheng 2002/09/18
   InverseTextBox Me.textCP09_S3
End Sub

Private Sub textCP09_S3_LostFocus()
   'Add By Cheng 2002/09/17
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   'Add By Cheng 2002/09/17
   If textCP09_S = "S" And IsEmptyText(textCP09_S1) = False Then
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text)
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         strTit = "檢核資料"
         strMsg = "查名本所案號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.textCP09_S.SetFocus
         textCP09_S_GotFocus
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Sub

' 點數
Private Sub textCP18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP18) = False Then
      If IsNumeric(textCP18) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "點數只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP18_GotFocus
      End If
   End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否算案件數
Private Sub textCP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP26) = False Then
      Select Case textCP26
         Case " ", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
   End If
   
End Sub

Private Sub textCP27_LostFocus()
Dim strDay As String
    
    'Add By Cheng 2003/10/14
    '若有輸發文日
    'Modified by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
    'If Me.textCP27.Text <> "" Then
    If Me.textCP27.Text <> "" And Me.textCP27.Tag <> Me.textCP27.Text Then
        ' 計算催審期限
        strDay = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
'        If IsEmptyText(strDay) = False Then
            textUargeDate = strDay
'        End If
    End If
    Me.textCP27.Tag = Me.textCP27.Text   'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
End Sub

' 發文日
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP27) = False Then
      ' 發文日日期不正確
      If CheckIsDate(textCP27, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' 發文日日期不可超過系統日
      If Val(DBDATE(textCP27)) > Val(strSrvDate(1)) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "發文日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
        'Modify By Cheng 2003/10/14
'      ' 計算催審期限
'      strDay = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
'      If IsEmptyText(strDay) = False Then
'         textUargeDate = strDay
'      End If
   End If
EXITSUB:
End Sub

Private Sub textCP44_Click()
   textCP44_2 = m_AgentList(textCP44.ListIndex).aiName
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCP44_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCP44.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 代理人
Private Sub textCP44_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTempName As String   '2010/11/24 add by sonia
   
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
   textCP44_2 = Empty
   If IsEmptyText(textCP44) = False Then
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      '2010/11/24 modify by sonia 取消basQuery的GetFAgentNameAndState
      'Dim oState As Boolean
      'oState = True
      ''textCP44_2 = GetFAgentName(textCP44)
      'textCP44_2 = GetFAgentNameAndState(textCP44, oState)
      'If oState = False Then
      '      Cancel = True
      '      Exit Sub
      'End If
      
      'Add By Sindy 2013/5/23 加判斷是否為聯絡人
      If InStr(textCP44, "-") > 0 Then
         textCP44_2.Text = ""
         If ClsPDGetContact(textCP44, strTempName) Then
            textCP44_2 = strTempName
         End If
      Else
      '2013/5/23 End
         If PUB_GetAgentNameAndState(m_TM01, textCP44.Text, strTempName) Then
            textCP44_2 = strTempName
         Else
            textCP44_2.Text = ""
            If strTempName <> "" Then
               Cancel = True
               Exit Sub
            End If
         End If
         '2010/11/24 end
      End If
      
      If IsEmptyText(textCP44_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44_GotFocus
      Else
         '依所選擇的代理人找出案件進度檔中其收文日最大的一筆其彼所案號更新到畫面上的彼所案號欄位
         'Modify By Sindy 2013/5/23 因代理人增加聯絡人,所以在取得代理人編碼時抓取前9碼
         strSql = "SELECT CP45 FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP44 = '" & Left(textCP44, 9) & "' AND " & _
                        "CP05 IN (SELECT MAX(CP05) FROM CASEPROGRESS " & _
                                 "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                       "CP02 = '" & m_TM02 & "' AND " & _
                                       "CP03 = '" & m_TM03 & "' AND " & _
                                       "CP04 = '" & m_TM04 & "' AND " & _
                                       "CP44 = '" & Left(textCP44, 9) & "')"
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

' 提申期限
Private Sub textPetition_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPetition) = False Then
      ' 日期不正確
      If CheckIsDate(textPetition, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的提申期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textPetition_GotFocus
      End If
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

Private Sub textPrintLetter_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否列印指示信
Private Sub textPrintLetter_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPrintLetter) = False Then
      Select Case textPrintLetter
         Case " ", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrintLetter_GotFocus
      End Select
   End If
End Sub

Private Sub textPrintTNT_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否列印TNT
Private Sub textPrintTNT_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPrintTNT) = False Then
      Select Case textPrintTNT
         Case " ", "Y":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrintTNT_GotFocus
      End Select
   End If
End Sub

Private Sub textPriorityDoc_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否附帶優先權證明文件
Private Sub textPriorityDoc_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textPriorityDoc) = False Then
      Select Case textPriorityDoc
         Case "Y", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入Y或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPriorityDoc_GotFocus
      End Select
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
Dim strCP64 As String
   
   ' 機關文號
   SetCPFieldNewData "CP18", textCP18
   ' 是否算案件數
   SetCPFieldNewData "CP26", textCP26
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   
   ' 代理人
   If IsEmptyText(textCP44) = False Then
      'Add By Sindy 2013/5/23 加判斷是否為聯絡人
      intI = InStr(textCP44, "-")
      If intI > 0 Then
         SetCPFieldNewData "CP44", Left(textCP44, intI - 1) & String(9 - Len(Left(textCP44, intI - 1)), "0")
         m_CP44New = Left(textCP44, intI - 1) & String(9 - Len(Left(textCP44, intI - 1)), "0")
         SetCPFieldNewData "CP116", Mid(textCP44, intI + 1)
      Else
      '2013/5/23 End
         SetCPFieldNewData "CP44", textCP44 & String(9 - Len(textCP44), "0")
         m_CP44New = textCP44 & String(9 - Len(textCP44), "0") 'add by nickc 2008/02/22
         SetCPFieldNewData "CP116", "" 'Add By Sindy 2013/5/23
      End If
   Else
      SetCPFieldNewData "CP44", textCP44
      m_CP44New = textCP44 'add by nickc 2008/02/22
      SetCPFieldNewData "CP116", "" 'Add By Sindy 2013/5/23
   End If
   
   ' 彼所案號
   SetCPFieldNewData "CP45", textTM45
   ' 進度備註
    strCP64 = Me.textCP64.Text
    'Modify By Cheng 2003/09/05
    '取消
    'Begin
'    'Add By Cheng 2003/06/16
'    '若有輸入查名本所案號
'    If Me.textCP09_S.Text <> "" And Me.textCP09_S1.Text <> "" Then
'        strCP64 = strCP64 & IIf(strCP64 <> "", ",", "") & "原查名本所案號：" & Me.textCP09_S.Text & "-" & Me.textCP09_S1.Text & "-" & Left(Me.textCP09_S2.Text & "0", 1) & "-" & Left(Me.textCP09_S3.Text & "00", 2)
'    End If
    'End
   SetCPFieldNewData "CP64", strCP64
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
   'add by nickc 2006/07/03 延展時，要存專用期限
   'modify by sonia 2025/9/4 +109緩審延展CFT-016520
   If m_CP10 = "102" Or m_CP10 = "109" Then
       SetCPFieldNewData "CP53", DBDATE(textTM21)
       SetCPFieldNewData "CP54", DBDATE(textTM22)
   End If
   Select Case m_TM01
      ' 系統類別為CFT的為更新商標基本檔
      Case "CFT":
'         ' 案件中文名稱
'         SetTMSPFieldNewData "TM05", textTM05
         ' 案件名稱
         SetTMSPFieldNewData "TM05", textTM05_1
'         ' 案件英文名稱
'         SetTMSPFieldNewData "TM06", textTM06
'         ' 案件日文名稱
'         SetTMSPFieldNewData "TM07", textTM07
         ' 商標種類
         SetTMSPFieldNewData "TM08", textTM08
         ' 商品類別
         SetTMSPFieldNewData "TM09", textTM09
         ' 申請人
         If IsEmptyText(textTM23) = False Then
            SetTMSPFieldNewData "TM23", textTM23 & String(9 - Len(textTM23), "0")
         Else
            SetTMSPFieldNewData "TM23", textTM23
         End If
         'add by nickc 2005/11/18 若有修改申請人時，要更新基本檔的申請地址
         If m_TM23 & String(9 - Len(m_TM23), "0") <> textTM23 & String(9 - Len(textTM23), "0") Then
            Dim rsTmp As New ADODB.Recordset
            Set rsTmp = New ADODB.Recordset
            If rsTmp.State = 1 Then rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open "select * from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM23.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM23.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <> 0 Then
                SetTMSPFieldNewData "TM24", CheckStr(rsTmp.Fields("cu23"))
                SetTMSPFieldNewData "TM25", CheckStr(rsTmp.Fields("CU24")) & IIf(CheckStr(rsTmp.Fields("cu25")) <> "", " " & CheckStr(rsTmp.Fields("cu25")), "") & IIf(CheckStr(rsTmp.Fields("cu26")) <> "", " " & CheckStr(rsTmp.Fields("cu26")), "") & IIf(CheckStr(rsTmp.Fields("cu27")) <> "", " " & CheckStr(rsTmp.Fields("cu27")), "") & IIf(CheckStr(rsTmp.Fields("cu28")) <> "", " " & CheckStr(rsTmp.Fields("cu28")), "")
                SetTMSPFieldNewData "TM26", CheckStr(rsTmp.Fields("cu29"))
            End If
         Else
            SetTMSPFieldNewData "TM24", m_TM24
            SetTMSPFieldNewData "TM25", m_tm25
            SetTMSPFieldNewData "TM26", m_tm26
         End If
         ' 正商標號數
         SetTMSPFieldNewData "TM27", textTM27
         '商品組群
         SetTMSPFieldNewData "TM32", textTM32
         ' 代表人
         SetTMSPFieldNewData "TM47", textTM47
         ' 代表人
         SetTMSPFieldNewData "TM48", textTM48
         ' 代表人
         SetTMSPFieldNewData "TM49", textTM49
         ' 代表人
         SetTMSPFieldNewData "TM50", textTM50
         ' 代表人
         SetTMSPFieldNewData "TM51", textTM51
         ' 代表人
         SetTMSPFieldNewData "TM52", textTM52
         ' 案件備註
         SetTMSPFieldNewData "TM58", textTM58
         ' 放棄專用權
         SetTMSPFieldNewData "TM67", textTM67
      Case Else:
         ' 案件中文名稱
        Select Case m_TM01
        Case "S"
            SetTMSPFieldNewData "SP05", textTM05_1
        Case Else
            SetTMSPFieldNewData "SP05", textTM05
        End Select
         ' 案件英文名稱
         SetTMSPFieldNewData "SP06", textTM06
         ' 案件日文名稱
         SetTMSPFieldNewData "SP07", textTM07
         ' 案件備註
         SetTMSPFieldNewData "SP18", textTM58
         ' 申請人
         If IsEmptyText(textTM23) = False Then
            SetTMSPFieldNewData "SP08", textTM23 & String(9 - Len(textTM23), "0")
         Else
            SetTMSPFieldNewData "SP08", textTM23
         End If
   End Select
      
End Sub

' 更新商標基本檔的相關欄位
Private Sub OnUpdateTradeMark()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
            strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
End Sub

' 更新服務業務基本檔的相關欄位
Private Sub OnUpdateServicePractice()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
   ' 更新案件進度檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
            strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
End Sub

'edit b nick 2004/11/03
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strTmp As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP22 As String
   'add by nickc 2006/09/07
   Dim strNP09 As String
   Dim strNP10 As String 'Add By Sindy 2014/9/11
   Dim strYear As String, strStartUpDay As String 'Add By Sindy 2019/10/16
   
'911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   'Modified by Lydia 2016/03/11 +案號
   'Call GetNP69("", m_TM10, m_CP13, strNP10) 'Add By Sindy 2014/9/11
   'Modified by Lydia 2017/05/12 GetNP69更名為GetNA69
   Call GetNA69("", m_TM10, m_CP13, strNP10, m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 更新案件進度檔
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
            strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & 0
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
   'Add By Cheng 2004/05/17
   '同時發文跨類資料
   If m_blnOutGoingMsg107 = True Then
      'Add By Sindy 2013/5/23 加判斷是否為聯絡人
      intI = InStr(textCP44, "-")
      If intI > 0 Then
         'modify by sonia 2017/2/17 +714超項費,711文件公／簽證
         strSql = "Update Caseprogress Set CP27=" & Val(DBDATE(Me.textCP27.Text)) & _
                  ", CP44='" & Left(textCP44, intI - 1) & String(9 - Len(Left(textCP44, intI - 1)), "0") & "'" & _
                  ", CP116='" & Mid(textCP44, intI + 1) & "'" & _
                  " Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10 in ('107','714','711') And CP27 Is Null And CP57 Is Null"
      Else
      '2013/5/23 End
         'modify by sonia 2017/2/17 +714超項費,711文件公／簽證
         strSql = "Update Caseprogress Set CP27=" & Val(DBDATE(Me.textCP27.Text)) & _
                  ", CP44='" & ChangeCustomerL(Me.textCP44.Text) & "'" & _
                  ", CP116=null" & _
                  " Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10 in ('107','714','711') And CP27 Is Null And CP57 Is Null"
      End If
      cnnConnection.Execute strSql
   End If
   'End
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為更新商標基本檔
      Case "CFT":
         OnUpdateTradeMark
      Case Else:
         OnUpdateServicePractice
   End Select
   
   'Add By Sindy 2019/10/16 新增下次緩審延展期限
   If m_CP10 = "109" Then
      strNP08 = "": strNP09 = ""
      strNP07 = "109" '緩審延展
      If ClsPDGetNationTax(11, "302", strStartUpDay, strYear) = True And Val(m_CP07) > 0 Then
         '法定期限=原法定期限+國家檔NA14延展年度
         If Val(strYear) > 0 Then
            strNP09 = DBDATE(DateAdd("yyyy", Val(strYear), ChangeWStringToWDateString(DBDATE(m_CP07))))
         End If
         '本所期限=法定期限-2個月
         strNP08 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
      End If
      If strNP09 = "" Then strNP09 = strSrvDate(1)
      If strNP08 = "" Then strNP08 = strSrvDate(1)
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      '下一程序序號
      strNP22 = GetNextProgressNo()
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & strNP07 & "'," & _
                           strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
   End If
   '2019/10/16 END
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有輸入催審期限時, 新增一筆催審的記錄到下一程序檔
   If IsEmptyText(textUargeDate) = False Then
      strNP07 = "305"
      strNP22 = GetNextProgressNo()
      '92.10.6 MODIFY BY SONIA
      'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
      '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
      '                    DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strUserNum & "'," & strNP22 & ")"
      'Modify By Sindy 2014/9/11 m_CP14=>strNP10
      'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                          DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strNP10 & "'," & strNP22 & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                          PUB_GetWorkDay1(textUargeDate, True) & "," & DBDATE(textUargeDate) & ",'" & strNP10 & "'," & strNP22 & ")"
      '92.10.6 END
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      'Modify By Sindy 2019/10/16 109緩審延展不印接洽結案單
      Select Case strNP07
         Case "102", "105", "702", "708", "305", "998", "997", "109":
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
         strNP08 = DBDATE(textCP27)
        'Modify By 2003/09/02
'         strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP08)), Val(DBMONTH(strNP08)), Val(DBDAY(strNP08)) + Val(rsTmp.Fields("CF23")))))
         strNP08 = DBDATE(DateAdd("d", Val(rsTmp.Fields("CF23")), ChangeWStringToWDateString(DBDATE(strNP08))))
         'Add By Sindy 2019/6/11 檢查期限是否正確
         strNP08 = PUB_T997998LimitDate(strNP08, m_CP07, 1)
         '2019/6/11 END
         strNP22 = GetNextProgressNo()
         '92.10.6 MODIFY BY SONIA
         'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
         '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
         '                   strNP08 & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
         'Modify By Sindy 2014/9/11 m_CP14=>strNP10
         'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            strNP08 & "," & strNP08 & ",'" & strNP10 & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & strNP10 & "'," & strNP22 & ")"
         '92.10.6 END
         cnnConnection.Execute strSql
         ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
         'Modify By Sindy 2019/10/16 109緩審延展不印接洽結案單
         Select Case strNP07
            Case "102", "105", "702", "708", "305", "998", "997", "109":
            Case Else:
               ' 列印國內案件接洽及結案記錄單
'               g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
                'Add By Cheng 2004/04/08
                '新增列印接洽結案單資料
                pub_AddressListSN = pub_AddressListSN + 1
                PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
         End Select
      End If
   End If
   rsTmp.Close

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 有輸入提申期限時, 新增一筆提申的記錄到下一程序檔
   If IsEmptyText(textPetition) = False Then
      strNP07 = "998"
      strNP22 = GetNextProgressNo()
      '92.10.6 MODIFY BY SONIA
      'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
      '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
      '                 DBDATE(textPetition) & "," & DBDATE(textPetition) & ",'" & strUserNum & "'," & strNP22 & ")"
      'Modify By Sindy 2014/9/11 m_CP14=>strNP10
      'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                       DBDATE(textPetition) & "," & DBDATE(textPetition) & ",'" & strNP10 & "'," & strNP22 & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                       PUB_GetWorkDay1(textPetition, True) & "," & DBDATE(textPetition) & ",'" & strNP10 & "'," & strNP22 & ")"
      '92.10.6 END
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      'Modify By Sindy 2019/10/16 109緩審延展不印接洽結案單
      Select Case strNP07
         Case "102", "105", "702", "708", "305", "998", "997", "109":
         Case Else:
            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   'Added by Lydia 2024/07/31
   Else
      '判斷案件國家收費表內有設定提申期限(天)CF11，要加掛提申(998)期限；
      '與秀玲討論：因為變更NA69會整批更新未發文和未續辦的下一程序，所以傳入模組統一使用CP14
       Call Pub_GetCF11to998(m_TM10, m_TM01, m_TM02, m_TM03, m_TM04, m_CP07, m_CP09, m_CP10, m_CP14, textCP27)
   'end 2024/07/31
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Modify By Cheng 2002/09/18
'   ' 有輸入查名總收文號時, 更新此收文號之本所案號為本案之本所案號
   ' 有輸入查名本所案號時, 更新此查名本所案號資料之本所案號為本案之本所案號
   If textCP09_S.Text = "S" And IsEmptyText(textCP09_S1.Text) = False Then
         'add by nickc 2005/10/28 清未結餘的可結餘日期
         strSql = "UPDATE CaseProgress SET cp109=null " & _
                  "WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text) & " and cp59 is null "
         cnnConnection.Execute strSql
         'edit by nickc 2006/07/18 加入 cp31=null
         strSql = "UPDATE CaseProgress SET cp31=null " & _
                  "WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text) & " "
         cnnConnection.Execute strSql
         
      strSql = "UPDATE CaseProgress SET CP01 = '" & m_TM01 & "', CP02 = '" & m_TM02 & "', " & _
                     "CP03 = '" & m_TM03 & "', CP04 = '" & m_TM04 & "', " & _
                     "CP64=CP64||Decode(CP64,Null,'','，')||'" & "原查名本所案號：" & Me.textCP09_S.Text & "-" & Me.textCP09_S1.Text & "-" & Left(Me.textCP09_S2.Text & "0", 1) & "-" & Left(Me.textCP09_S3.Text & "00", 2) & "' " & _
               " WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text)
      cnnConnection.Execute strSql
      'Add By Cheng 2003/06/16
      strSql = "Update ServicePractice Set SP18=SP18||Decode(SP18,Null,'','，')||'轉入商標：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' Where " & ChgService(Me.textCP09_S.Text & Me.textCP09_S1.Text & Left(Me.textCP09_S2.Text & "0", 1) & Left(Me.textCP09_S3.Text & "00", 2))
      cnnConnection.Execute strSql
      '2005/4/18 ADD BY SONIA 1~4欄原查名本所案號,5~8欄新商標本所案號
      If PUB_UpdOther(Me.textCP09_S.Text, Me.textCP09_S1.Text, Left(Me.textCP09_S2.Text & "0", 1), Left(Me.textCP09_S3.Text & "00", 2), m_TM01, m_TM02, m_TM03, m_TM04) = False Then
         GoTo CheckingErr
      End If
      '2005/4/18 END
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 儲存優先權資料
   'edit by nickc 2007/02/06 不用 dll 了 objPublicData.SavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify By Sindy 2017/10/12 + , m_Priority(6)
   ClsPDSavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'add by nick 2004/09/08
   'CFT發文frm030101_03：若案件性質為’申請’(101)且申請國家為’摩洛哥’(306)時，
   '於發文存檔時，於該案之商標基本檔的案件備註欄再加註’丹吉爾區。’；
   '並同時新增一筆商標基本檔資料，其案號為原案號子案，
   '但於案件備註欄再加註’卡薩布蘭加區。’，其他欄位資料都同於原案，
   '並新增該子案之案件進度資料：收文號為B類收文，收文日為系統日期，
   '費用、規費、點數皆為NULL，是否向客戶收款、是否算案件數、是否開電腦收據設為’N’
   '，其他欄位資料都同於原收文號。例原案號為 CFT-009488-0-00，新增的子案案號為 CFT-009488-1-00。
'''''edit by nickc 2008/03/13 外商阿蓮請作單  取消，因為已經不分區了
'''''   If m_CP10 = "101" And m_TM10 = "306" Then
'''''        Dim strCP09 As String
'''''        strSQL = "insert into trademark ("
'''''        For nIndex = 1 To TF_TM
'''''            strSQL = strSQL & "TM" & Format(nIndex, "00")
'''''            If nIndex <> TF_TM Then
'''''                strSQL = strSQL & ","
'''''            End If
'''''        Next nIndex
'''''        strSQL = strSQL & " ) select "
'''''        For nIndex = 1 To TF_TM
'''''            Select Case nIndex
'''''            Case 3
'''''                strSQL = strSQL & "to_char(to_number(TM" & Format(nIndex, "00") & ")+1)"
'''''            Case 58
'''''                strSQL = strSQL & "TM" & Format(nIndex, "00") & "||'卡薩布蘭加區。'"
'''''            Case Else
'''''                strSQL = strSQL & "TM" & Format(nIndex, "00")
'''''            End Select
'''''            If nIndex <> TF_TM Then
'''''                strSQL = strSQL & ","
'''''            End If
'''''        Next nIndex
'''''        strSQL = strSQL & " FROM trademark WHERE TM01 = '" & m_TM01 & "' AND TM02 = '" & m_TM02 & "' AND TM03 = '" & m_TM03 & "' AND TM04 = '" & m_TM04 & "' "
'''''        cnnConnection.Execute strSQL
'''''        strSQL = "update trademark set tm58=tm58||'丹吉爾區。' WHERE TM01 = '" & m_TM01 & "' AND " & _
'''''                        "TM02 = '" & m_TM02 & "' AND " & _
'''''                        "TM03 = '" & m_TM03 & "' AND " & _
'''''                        "TM04 = '" & m_TM04 & "' "
'''''        cnnConnection.Execute strSQL
'''''        strCP09 = AutoNo("B", 6)
'''''        strSQL = "insert into caseprogress ("
'''''        For nIndex = 1 To TF_CP
'''''            strSQL = strSQL & "CP" & Format(nIndex, "00")
'''''            If nIndex <> TF_CP Then
'''''                strSQL = strSQL & ","
'''''            End If
'''''         Next nIndex
'''''         strSQL = strSQL & " ) select "
'''''         For nIndex = 1 To TF_CP
'''''            Select Case nIndex
'''''            Case 3
'''''                strSQL = strSQL & "to_char(to_number(CP" & Format(nIndex, "00") & ")+1)"
'''''            Case 5
'''''                strSQL = strSQL & DBDATE(ServerDate)
'''''            Case 9
'''''                strSQL = strSQL & " '" & strCP09 & "' "
'''''            Case 16, 17, 18, 33, 34
'''''                strSQL = strSQL & "NULL"
'''''            Case 20, 26, 32
'''''                strSQL = strSQL & "'N'"
'''''            Case Else
'''''                strSQL = strSQL & "CP" & Format(nIndex, "00")
'''''            End Select
'''''            If nIndex <> TF_CP Then
'''''                strSQL = strSQL & ","
'''''            End If
'''''        Next nIndex
'''''        strSQL = strSQL & " from caseprogress where cp09='" & m_CP09 & "' "
'''''        cnnConnection.Execute strSQL
'''''
'''''   End If
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'add by sonia 2024/4/1 「坦尚尼亞聯合共和國」之申請案，分「TANZANIA區」(國家代碼為327)及「ZANZIBAR區」(國家代碼為328)
   '案件性質為’申請’(101)且申請國家為’坦尚尼亞聯合共和國’(327)時，
   '於發文存檔時，於該案之商標基本檔的案件備註欄再加註’丹吉爾區。’；
   '並同時新增一筆商標基本檔資料，其案號為原案號子案，
   '但於案件備註欄再加註’卡薩布蘭加區。’，其他欄位資料都同於原案，
   '並新增該子案之案件進度資料：收文號為B類收文，收文日為系統日期，
   '費用、規費、點數皆為NULL，是否向客戶收款、是否算案件數、是否開電腦收據設為’N’
   '，其他欄位資料都同於原收文號
   If m_CP10 = "101" And m_TM10 = "327" Then
        Dim strCP09 As String
        strSql = "insert into trademark ("
        For nIndex = 1 To TF_TM
            strSql = strSql & "TM" & Format(nIndex, "00")
            If nIndex <> TF_TM Then
                strSql = strSql & ","
            End If
        Next nIndex
        strSql = strSql & " ) select "
        For nIndex = 1 To TF_TM
            Select Case nIndex
            Case 3
                strSql = strSql & "to_char(to_number(TM" & Format(nIndex, "00") & ")+1)"
            Case 10
                strSql = strSql & "'328'"
            Case 58
                strSql = strSql & "TM" & Format(nIndex, "00") & "||'ZANZIBAR區。與" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "合併計算結餘。'"
            Case Else
                strSql = strSql & "TM" & Format(nIndex, "00")
            End Select
            If nIndex <> TF_TM Then
                strSql = strSql & ","
            End If
        Next nIndex
        strSql = strSql & " FROM trademark WHERE TM01 = '" & m_TM01 & "' AND TM02 = '" & m_TM02 & "' AND TM03 = '" & m_TM03 & "' AND TM04 = '" & m_TM04 & "' "
        cnnConnection.Execute strSql
        'strSql = "update trademark set tm58=tm58||'TANZANIA區。與'" & m_TM01 & "'-'" & m_TM02 & "'-'" & to_char(to_number(" & m_TM03 & " & ")+1) & "-" & m_TM04 & "合併計算結餘。' WHERE TM01 = '" & m_TM01 & "' AND "
        strSql = "update trademark set tm58=tm58||'TANZANIA區。與" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 + 1 & "-" & m_TM04 & "合併計算結餘。' WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND TM03 = '" & m_TM03 & "' AND TM04 = '" & m_TM04 & "' "
        cnnConnection.Execute strSql
        strCP09 = AutoNo("B", 6)
        strSql = "insert into caseprogress ("
        For nIndex = 1 To TF_CP
            strSql = strSql & "CP" & Format(nIndex, "00")
            If nIndex <> TF_CP Then
                strSql = strSql & ","
            End If
         Next nIndex
         strSql = strSql & " ) select "
         For nIndex = 1 To TF_CP
            Select Case nIndex
            Case 3
                strSql = strSql & "to_char(to_number(CP" & Format(nIndex, "00") & ")+1)"
            Case 5
                strSql = strSql & DBDATE(ServerDate)
            Case 9
                strSql = strSql & " '" & strCP09 & "' "
            Case 16, 17, 18, 33, 34, 60, 73, 74, 75, 76, 77, 78, 79, 140
                strSql = strSql & "NULL"
            Case 20, 26, 32
                strSql = strSql & "'N'"
            Case Else
                strSql = strSql & "CP" & Format(nIndex, "00")
            End Select
            If nIndex <> TF_CP Then
                strSql = strSql & ","
            End If
        Next nIndex
        strSql = strSql & " from caseprogress where cp09='" & m_CP09 & "' "
        cnnConnection.Execute strSql
   End If
   'end 2024/4/1
   
'2012/10/11 cancel by sonia 葡萄牙自20081001起取消使用宣誓期限,而且智權人員存成姓名
'   'add by nickc 2006/09/07 加入葡萄牙(213)使用宣誓(105)管制
'   If m_TM01 = "CFT" And m_CP10 = "102" And m_TM10 = "213" Then
'      strNP07 = "105"
'      strNP09 = ChangeWDateStringToWString(DateAdd("yyyy", -5, ChangeWStringToWDateString(frm030101_03.textTM22.Text)))
'      strNP08 = ChangeWDateStringToWString(DateAdd("m", -3, ChangeWStringToWDateString(strNP09)))
'      strNP22 = GetNextProgressNo()
'      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                       strNP08 & "," & strNP09 & ",'" & m_CP13 & "'," & strNP22 & ")"
'      cnnConnection.Execute strSql
'   End If
'2012/10/11 end
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
   
   Set rsTmp = Nothing
'911106 nick transation
   cnnConnection.CommitTrans
   
   ' 列印定稿
   If textPrint <> "N" Then
      PrintLetter
   End If
    
   'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
   PUB_CheckEMail m_CP44New, m_CP116
   PUB_CheckEMail m_TM44, m_TM119
   If m_TM120 <> "" Then
      PUB_CheckEMail m_TM44, m_TM120
   End If
   'end 2008/02/22
   
   Exit Function
CheckingErr:
   MsgBox (Err.Description)
   cnnConnection.RollbackTrans
   OnSaveData = False
End Function

Private Sub textTM05_1_GotFocus()
    TextInverse Me.textTM05_1
    'edit by nickc 2007/06/06 切換輸入法改用API
    OpenIme
End Sub

Private Sub textTM05_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    If CheckLengthIsOK(textTM05_1, 140) = False Then
        Cancel = True
        strTit = "檢核資料"
        strMsg = "案件名稱內容太長"
        textTM05_1_GotFocus
    End If
    'edit by nickc 2007/06/06 切換輸入法改用API
    If Cancel = False Then CloseIme
End Sub

' 案件中文名稱
Private Sub textTM05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM05, 40) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件中文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM05_GotFocus
   End If
End Sub

' 案件英文名稱
Private Sub textTM06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM06, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件英文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM06_GotFocus
   End If
End Sub

' 案件日文名稱
Private Sub textTM07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM07, 40) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件日文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM07_GotFocus
   End If
End Sub

' 商標種類
Private Sub textTM08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textTM08_2 = Empty
   If IsEmptyText(textTM08) = False Then
      textTM08_2 = GetTradeMarkName(textTM08, 0)
      If IsEmptyText(textTM08_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商標種類不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM08_GotFocus
      End If
   End If
End Sub

' 商品類別
Private Sub textTM09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textTM09) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textTM09)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM09, nIndex)
      If Len(strTemp) > 6 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商品類別<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM09_GotFocus
         GoTo EXITSUB
      End If
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM09, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTM09, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商品類別<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM09_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
'add by nickc 2005/06/03
textTM09 = Replace(textTM09, " ", "")
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 延展後專用期限起日
Private Sub textTM21_Validate(Cancel As Boolean)
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM21) = False Then
      If CheckIsDate(textTM21, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請輸入正確的延展後專用期限起日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM21_GotFocus
      End If
      
      If IsEmptyText(m_TM21) = False Then
         '91.12.26 MODIFY BY SONIA
         'strTemp = DBDATE(DateSerial(Val(DBYEAR(m_TM22)), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22)) + 1))
         'If DBDATE(textTM21) <> strTemp Then
         '   Cancel = True
         '   strTit = "檢核資料"
         '   strMsg = "延展後專用期限起日應為<" & strTemp & ">"
         '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '   textTM21_GotFocus
         'End If
         If textTM21 <> m_TM21 And textTM21 <> CompDate(2, -1, m_TM21) Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "專用期限起日應為<" & m_TM21 & "> 或 <" & CompDate(2, -1, m_TM21) & ">"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM21_GotFocus
         End If
         '91.12.26 END
      End If
   Else
      '910722 Sieg
      'modify by sonia 2025/9/4 +109緩審延展CFT-016520
      If m_CP10 = "102" Or m_CP10 = "109" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質為延展時，延展後專用期限不得空白 !"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

Private Sub textTM22_LostFocus()
   If textTM21 <> "" And textTM22 <> "" Then
      If Not ChkRange(textTM21, textTM22, "延展後專用期限") Then
         textTM21.SetFocus
      End If
   End If
End Sub

' 延展後專用期限止日
Private Sub textTM22_Validate(Cancel As Boolean)
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM22) = False Then
      If CheckIsDate(textTM22, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請輸入正確的延展後專用期限止日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      
      If IsEmptyText(m_TM22) = False And IsEmptyText(m_NA14) = False Then
         '91.12.26 MODIFY BY SONIA
         'strTemp = DBDATE(DateSerial(Val(DBYEAR(m_TM22)) + Val(m_NA14), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22)) + 1))
         'If DBDATE(textTM22) <> strTemp Then
         '   Cancel = True
         '   strTit = "檢核資料"
         '   strMsg = "延展後專用期限止日應為<" & strTemp & ">"
         '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         'End If
         If textTM22 <> m_TM22 And textTM22 <> CompDate(2, -1, m_TM22) Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "專用期限止日應為<" & m_TM22 & "> 或 <" & CompDate(2, -1, m_TM22) & ">"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM22_GotFocus
         End If
         '91.12.26 END
      End If
   Else
      '910722 Sieg
      'modify by sonia 2025/9/4 +109緩審延展CFT-016520
      If m_CP10 = "102" Or m_CP10 = "109" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質為延展時，期限不得空白 !"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   If Cancel Then TextInverse textTM22
End Sub

Private Sub textTM23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 申請人
Private Sub textTM23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM23_2 = Empty
   If IsEmptyText(textTM23) = False Then
        Me.textTM23.Text = ChangeCustomerL(Me.textTM23.Text)
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      oState = True
      'textTM23_2 = GetCustomerName(textTM23, 0)
      textTM23_2 = GetCustomerNameAndState(textTM23, 0, oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If textTM23_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM23 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM23_GotFocus
      End If
   End If
   'Add By Cheng 2002/08/22
   If Cancel = False Then
      If Me.textTM23.Text <> m_strCust1 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textTM23_GotFocus
   
End Sub

' 商品組群
Private Sub textTM32_Validate(Cancel As Boolean)
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textTM32) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textTM32)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM32, nIndex)
      If Len(strTemp) > 6 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "商品組群<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM32_GotFocus
         GoTo EXITSUB
      End If
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM32, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTM32, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商品組群<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM32_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 代表人1(中)
Private Sub textTM47_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM47, 40) = False Then
   If CheckLengthIsOK(textTM47, textTM47.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "代表人1(中)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM47_GotFocus
   End If
End Sub

' 代表人1(英)
Private Sub textTM48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM48, textTM48.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "代表人1(英)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM48_GotFocus
   End If
End Sub

' 代表人1(日)
Private Sub textTM49_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
    'Modified by Lydia 2016/09/10
   'If CheckLengthIsOK(textTM49, 40) = False Then
   If CheckLengthIsOK(textTM49, textTM49.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "代表人1(日)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM49_GotFocus
   End If
End Sub

' 代表人2(中)
Private Sub textTM50_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM50, 40) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "代表人2(中)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM50_GotFocus
   End If
End Sub

' 代表人2(英)
Private Sub textTM51_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM51, textTM51.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "代表人2(英)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM51_GotFocus
   End If
End Sub

' 代表人2(日)
Private Sub textTM52_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM52, 40) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "代表人2(日)內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM52_GotFocus
   End If
End Sub

' 案件備註
Private Sub textTM58_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textTM58, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM58_GotFocus
   End If
End Sub

' 催審期限
Private Sub textUargeDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textUargeDate) = False Then
      If CheckIsDate(textUargeDate, False) = False Then
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
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   CheckDataValid = False
   
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
   
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF", "S"
        ' 案件名稱
        If IsEmptyText(textTM05_1) = True Then
           strTit = "檢核資料"
           strMsg = "請輸入案件名稱"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textTM05_1.SetFocus
           GoTo EXITSUB
        End If
    Case Else
        ' 案件名稱(中, 英, 日)
        If IsEmptyText(textTM05) = True And IsEmptyText(textTM06) = True And IsEmptyText(textTM07) = True Then
           strTit = "檢核資料"
           strMsg = "請輸入案件名稱"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textTM05.SetFocus
           GoTo EXITSUB
        End If
    End Select
   ' 發文日
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' 申請人
   If IsEmptyText(textTM23) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入申請人"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM23.SetFocus
      GoTo EXITSUB
   End If
   ' 代理人
   If IsEmptyText(textCP44) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入代理人"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP44.SetFocus
      GoTo EXITSUB
   End If
   ' 商品類別
   If IsEmptyText(textTM09) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入商品類別"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM09.SetFocus
      GoTo EXITSUB
   End If
   ' 商品組群
   'If IsEmptyText(textTM32) = True Then
   '   strTit = "檢核資料"
   '   strMsg = "請輸入商品組群"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   GoTo EXITSUB
   'End If
   ' 商標種類為聯合商標, 防護商標, 聯合服務標章, 防護服務標章時正商標號數不可空白
   If IsEmptyText(textTM08) = True Then
      Select Case textTM08
         Case "2", "3", "5", "6":
            strTit = "檢核資料"
            strMsg = "請輸入正商標號數"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM08.SetFocus
            GoTo EXITSUB
      End Select
   End If
   
   'Add By Sindy 2012/5/3
   If m_CP10 = "102" Then
      '延展後專用期止日不可小於等於基本檔專用期止日
      If Val(DBDATE(textTM22)) <= Val(DBDATE(mm_TM22)) Then
         strTit = "檢核資料"
         strMsg = "延展後專用期止日不可小於等於基本檔專用期止日！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM22.SetFocus
         GoTo EXITSUB
      End If
      '延展後專用期止日不可小於等於法定期限
      If Val(DBDATE(textTM22)) <= Val(DBDATE(m_CP07)) Then
         strTit = "檢核資料"
         strMsg = "延展後專用期止日不可小於等於法定期限！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM22.SetFocus
         GoTo EXITSUB
      End If
   End If
   '2012/5/3 End
      
    'Add By Cheng 2004/05/17
    'CFT商申(101)發文時, 檢查本案是否有收文未發文未取消收文的跨類(107)資料
    m_blnOutGoingMsg107 = False
    If m_TM01 = "CFT" And m_CP10 = "101" Then
        'modify by sonia 2017/2/17 +714超項費,711文件公／簽證
        StrSQLa = "Select Count(*) From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10 in ('107','714','711') And CP27 Is Null And CP57 Is Null "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If Val("" & rsA.Fields(0).Value) > 0 Then
            If MsgBox("此案有 " & Val("" & rsA.Fields(0).Value) & "筆 跨類或超項費或文件公／簽證 收文資料, 確定是否同時發文???", vbExclamation + vbOKCancel) = vbOK Then
                m_blnOutGoingMsg107 = True
            End If
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
    'End
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textUargeDate_GotFocus()
   InverseTextBox textUargeDate
End Sub

Private Sub textPetition_GotFocus()
   InverseTextBox textPetition
End Sub

Private Sub textPriorityDoc_GotFocus()
   InverseTextBox textPriorityDoc
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textPrintLetter_GotFocus()
   InverseTextBox textPrintLetter
End Sub

Private Sub textPrintTNT_GotFocus()
   InverseTextBox textPrintTNT
End Sub

Private Sub textCP18_GotFocus()
   InverseTextBox textCP18
End Sub

Private Sub textTM05_GotFocus()
   InverseTextBox textTM05
End Sub

Private Sub textTM06_GotFocus()
   InverseTextBox textTM06
End Sub

Private Sub textTM07_GotFocus()
   InverseTextBox textTM07
End Sub

Private Sub textTM08_GotFocus()
   InverseTextBox textTM08
End Sub

Private Sub textTM09_GotFocus()
   InverseTextBox textTM09
End Sub

Private Sub textTM21_GotFocus()
   InverseTextBox textTM21
End Sub

Private Sub textTM22_GotFocus()
   InverseTextBox textTM22
End Sub

Private Sub textTM27_GotFocus()
   InverseTextBox textTM27
End Sub

Private Sub textTM23_GotFocus()
   InverseTextBox textTM23
End Sub

Private Sub textTM32_GotFocus()
   InverseTextBox textTM32
End Sub

Private Sub textTM47_GotFocus()
   InverseTextBox textTM47
End Sub

Private Sub textTM48_GotFocus()
   InverseTextBox textTM48
End Sub

Private Sub textTM49_GotFocus()
   InverseTextBox textTM49
End Sub

Private Sub textTM50_GotFocus()
   InverseTextBox textTM50
End Sub

Private Sub textTM51_GotFocus()
   InverseTextBox textTM51
End Sub

Private Sub textTM52_GotFocus()
   InverseTextBox textTM52
End Sub

Private Sub textTM58_GotFocus()
   InverseTextBox textTM58
End Sub

Private Sub textTM67_GotFocus()
   InverseTextBox textTM67
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP44_GotFocus()
   InverseTextBox textCP44
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strSql As String
   
   ' 系統類別為CFT
   If m_TM01 = "CFT" Then
      Select Case m_CP10
         ' 申請
         Case "101":
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "01", strUserNum
'CANCEL BY SONIA 2022/8/18 固定改１星期左右，寫在定稿裡
'            ' 回音
'            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
'                     "','回音','" & textCF09 & "')"
'            cnnConnection.Execute strSql
'END 2022/8/18
         ' 不續辦
         Case "703":
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "02", strUserNum
         ' 其它
         Case Else:
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "03", strUserNum
      End Select
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   ' 系統類別為CFT
   If m_TM01 = "CFT" Then
      Select Case m_CP10
         ' 申請
         Case "101":
            ' 列印定稿
            NowPrint m_CP09, "01", "01", False, strUserNum, 0
         ' 不續辦
         Case "703":
            ' 列印定稿
            NowPrint m_CP09, "01", "02", False, strUserNum, 0
         ' 其它
         Case Else:
            ' 列印定稿
            NowPrint m_CP09, "01", "03", False, strUserNum, 0
      End Select
   End If
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
   If Me.textCP09_S.Enabled = True Then
      Cancel = False
      textCP09_S_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP18.Enabled = True Then
      Cancel = False
      textCP18_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP26.Enabled = True Then
      Cancel = False
      textCP26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
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
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textPetition.Enabled = True Then
      Cancel = False
      textPetition_Validate Cancel
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
   
   If Me.textPrintLetter.Enabled = True Then
      Cancel = False
      textPrintLetter_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textPrintTNT.Enabled = True Then
      Cancel = False
      textPrintTNT_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textPriorityDoc.Enabled = True Then
      Cancel = False
      textPriorityDoc_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM05.Enabled = True Then
      Cancel = False
      textTM05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM05_1.Enabled = True Then
      Cancel = False
      textTM05_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM06.Enabled = True Then
      Cancel = False
      textTM06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM07.Enabled = True Then
      Cancel = False
      textTM07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM08.Enabled = True Then
      Cancel = False
      textTM08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM09.Enabled = True Then
      Cancel = False
      textTM09_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM21.Enabled = True Then
      Cancel = False
      textTM21_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM22.Enabled = True Then
      Cancel = False
      textTM22_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM23.Enabled = True Then
      Cancel = False
      textTM23_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM32.Enabled = True Then
      Cancel = False
      textTM32_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'add by nickc 2006/07/03 若是基本檔專用期間沒有值，不允許存檔
   If m_CP10 = "102" And mm_TM21 = Empty And mm_TM22 = Empty Then
       MsgBox "基本資料內的專用期間錯誤！", vbExclamation
       Cancel = True
       Exit Function
   End If
   
   'add by nickc 2007/04/17 加入檢查催審期限  阿蓮有請作單
   If textUargeDate.Enabled = True Then
       Cancel = False
       textUargeDate_Validate Cancel
       If Cancel = True Then
           Exit Function
       End If
   End If
   
   'Add By Sindy 2016/12/20
   '檢查有設定副本收受人需提醒並新增信函副本B類收文
   m_990CP09 = ""
   If textPrint = "N" Then '不印定稿
      If PUB_ChkCC(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_990CP09) = False Then
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   '2016/12/20 END
    'Added by Lydia 2021/06/04 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
    If Pub_ChkACS112isNull(m_TM01, m_TM02, m_TM03, m_TM04, txtCP113) = True Then
        SSTab1.Tab = 0
        txtCP113.SetFocus
        txtCP113_GotFocus
        Exit Function
    End If
    'end 2021/06/04
    
   'Added by Lydia 2021/08/20 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
   
   TxtValidate = True
End Function

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
