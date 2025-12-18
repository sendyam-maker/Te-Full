VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030203_02 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "°Ó¼Ð¥Ó½Ð®×¸¹¿é¤J"
   ClientHeight    =   5890
   ClientLeft      =   3650
   ClientTop       =   2990
   ClientWidth     =   8410
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5890
   ScaleWidth      =   8410
   Begin VB.CommandButton cmdOK 
      Caption         =   "°Ó«~¤ÎªA°È¸ê®Æ¬d¸ß(&I)"
      Height          =   400
      Index           =   6
      Left            =   3000
      TabIndex        =   28
      Top             =   0
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5445
      Left            =   60
      TabIndex        =   42
      Top             =   420
      Width           =   8295
      _ExtentX        =   14623
      _ExtentY        =   9596
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "°ò¥»¸ê®Æ"
      TabPicture(0)   =   "frm030203_02.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label26"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label23"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label22"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label17"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label38"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label39"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label3"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(8)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(7)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label36"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label37"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label9"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label11"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(18)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(17)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label15"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label16"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label32"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textPS"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textTM67"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdPriority"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textPriorityDoc"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textAddDate"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textCP05"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textPrint"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textAdd"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textTM27"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textTM11"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textTM12"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textTM32"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textTM08"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "cmbTM05"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textTMKey"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textTM10"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textTM09"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textDN"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtToEng"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "textNP08"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "textNP09"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "textAddDate2"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtLaw"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Frame1"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Frame2"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "textPrtTrans"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).ControlCount=   57
      TabCaption(1)   =   "¥Nªí¤H-1"
      TabPicture(1)   =   "frm030203_02.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5(8)"
      Tab(1).Control(1)=   "Label5(7)"
      Tab(1).Control(2)=   "Label5(6)"
      Tab(1).Control(3)=   "Label5(5)"
      Tab(1).Control(4)=   "Label5(4)"
      Tab(1).Control(5)=   "Label5(3)"
      Tab(1).Control(6)=   "Label14(1)"
      Tab(1).Control(7)=   "Label18(2)"
      Tab(1).Control(8)=   "Label18(0)"
      Tab(1).Control(9)=   "Label14(0)"
      Tab(1).Control(10)=   "Label5(1)"
      Tab(1).Control(11)=   "Label5(2)"
      Tab(1).Control(12)=   "Label5(9)"
      Tab(1).Control(13)=   "Label5(10)"
      Tab(1).Control(14)=   "Label5(11)"
      Tab(1).Control(15)=   "Label5(12)"
      Tab(1).Control(16)=   "Label18(1)"
      Tab(1).Control(17)=   "Label14(2)"
      Tab(1).Control(18)=   "Label5(13)"
      Tab(1).Control(19)=   "Label5(14)"
      Tab(1).Control(20)=   "Label5(15)"
      Tab(1).Control(21)=   "Label5(16)"
      Tab(1).Control(22)=   "Label5(17)"
      Tab(1).Control(23)=   "Label5(18)"
      Tab(1).Control(24)=   "Combo2(0)"
      Tab(1).Control(25)=   "textTM47"
      Tab(1).Control(26)=   "textTM48"
      Tab(1).Control(27)=   "textTM49"
      Tab(1).Control(28)=   "Combo2(2)"
      Tab(1).Control(29)=   "textTM94"
      Tab(1).Control(30)=   "textTM95"
      Tab(1).Control(31)=   "textTM96"
      Tab(1).Control(32)=   "Combo2(4)"
      Tab(1).Control(33)=   "textTM100"
      Tab(1).Control(34)=   "textTM101"
      Tab(1).Control(35)=   "textTM102"
      Tab(1).Control(36)=   "Combo2(1)"
      Tab(1).Control(37)=   "textTM50"
      Tab(1).Control(38)=   "textTM51"
      Tab(1).Control(39)=   "textTM52"
      Tab(1).Control(40)=   "Combo2(3)"
      Tab(1).Control(41)=   "textTM97"
      Tab(1).Control(42)=   "textTM98"
      Tab(1).Control(43)=   "textTM99"
      Tab(1).Control(44)=   "Combo2(5)"
      Tab(1).Control(45)=   "textTM103"
      Tab(1).Control(46)=   "textTM104"
      Tab(1).Control(47)=   "textTM105"
      Tab(1).ControlCount=   48
      TabCaption(2)   =   "¥Nªí¤H-2"
      TabPicture(2)   =   "frm030203_02.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TextTM114"
      Tab(2).Control(1)=   "Combo2(8)"
      Tab(2).Control(2)=   "TextTM113"
      Tab(2).Control(3)=   "TextTM112"
      Tab(2).Control(4)=   "TextTM106"
      Tab(2).Control(5)=   "TextTM107"
      Tab(2).Control(6)=   "Combo2(6)"
      Tab(2).Control(7)=   "TextTM108"
      Tab(2).Control(8)=   "TextTM117"
      Tab(2).Control(9)=   "Combo2(9)"
      Tab(2).Control(10)=   "TextTM116"
      Tab(2).Control(11)=   "TextTM115"
      Tab(2).Control(12)=   "TextTM111"
      Tab(2).Control(13)=   "Combo2(7)"
      Tab(2).Control(14)=   "TextTM110"
      Tab(2).Control(15)=   "TextTM109"
      Tab(2).Control(16)=   "Label18(3)"
      Tab(2).Control(17)=   "Label14(3)"
      Tab(2).Control(18)=   "Label5(19)"
      Tab(2).Control(19)=   "Label5(20)"
      Tab(2).Control(20)=   "Label5(21)"
      Tab(2).Control(21)=   "Label5(22)"
      Tab(2).Control(22)=   "Label5(23)"
      Tab(2).Control(23)=   "Label5(24)"
      Tab(2).Control(24)=   "Label18(4)"
      Tab(2).Control(25)=   "Label14(4)"
      Tab(2).Control(26)=   "Label5(25)"
      Tab(2).Control(27)=   "Label5(26)"
      Tab(2).Control(28)=   "Label5(27)"
      Tab(2).Control(29)=   "Label5(28)"
      Tab(2).Control(30)=   "Label5(29)"
      Tab(2).Control(31)=   "Label5(30)"
      Tab(2).ControlCount=   32
      Begin VB.TextBox textPrtTrans 
         Height          =   264
         Left            =   4590
         MaxLength       =   10
         TabIndex        =   15
         Top             =   3940
         Width           =   852
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   3960
         TabIndex        =   102
         Top             =   4830
         Width           =   4215
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   25
            Top             =   150
            Width           =   375
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   840
            MaxLength       =   2
            TabIndex        =   23
            Top             =   150
            Width           =   375
         End
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   2760
            MaxLength       =   7
            TabIndex        =   27
            Top             =   150
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Caption         =   "                      ¤é"
            Height          =   225
            Index           =   2
            Left            =   2520
            TabIndex        =   26
            Top             =   180
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "        ¤ë"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   24
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "¤å¨ì          ¤Ñ"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   180
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   1080
         TabIndex        =   101
         Top             =   4830
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "¤å¨ì¦¸¤é"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   21
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "¤å¨ì·í¤é"
            Height          =   180
            Index           =   0
            Left            =   144
            TabIndex        =   20
            Top             =   180
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtLaw 
         Height          =   285
         Left            =   6600
         TabIndex        =   19
         Top             =   4530
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox textAddDate2 
         Height          =   285
         Left            =   7080
         MaxLength       =   7
         TabIndex        =   13
         Top             =   3630
         Width           =   945
      End
      Begin VB.TextBox textNP09 
         Height          =   285
         Left            =   4170
         MaxLength       =   7
         TabIndex        =   18
         Top             =   4530
         Width           =   1095
      End
      Begin VB.TextBox textNP08 
         Height          =   285
         Left            =   1530
         MaxLength       =   7
         TabIndex        =   17
         Top             =   4530
         Width           =   1095
      End
      Begin VB.TextBox txtToEng 
         Height          =   285
         Left            =   4590
         MaxLength       =   1
         TabIndex        =   12
         Text            =   "N"
         Top             =   3630
         Width           =   372
      End
      Begin VB.TextBox textDN 
         Height          =   285
         Left            =   1470
         MaxLength       =   1
         TabIndex        =   11
         Top             =   3630
         Width           =   492
      End
      Begin VB.TextBox textTM09 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   395
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1260
         Width           =   6855
      End
      Begin VB.TextBox textTM10 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   360
         Width           =   2532
      End
      Begin VB.TextBox textTMKey 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   360
         Width           =   2532
      End
      Begin VB.ComboBox cmbTM05 
         Height          =   300
         Left            =   1200
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   660
         Width           =   6852
      End
      Begin VB.TextBox textTM08 
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   960
         Width           =   2532
      End
      Begin VB.TextBox textTM32 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   699
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1560
         Width           =   6852
      End
      Begin VB.TextBox textTM12 
         Height          =   285
         Left            =   5520
         MaxLength       =   9
         TabIndex        =   3
         Top             =   1860
         Width           =   2532
      End
      Begin VB.TextBox textTM11 
         Height          =   285
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1860
         Width           =   1332
      End
      Begin VB.TextBox textTM27 
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   5
         Top             =   2445
         Width           =   2532
      End
      Begin VB.TextBox textAdd 
         Height          =   285
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   10
         Top             =   3330
         Width           =   852
      End
      Begin VB.TextBox textPrint 
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   14
         Top             =   3930
         Width           =   372
      End
      Begin VB.TextBox textCP05 
         Height          =   285
         Left            =   5520
         MaxLength       =   8
         TabIndex        =   6
         Top             =   2445
         Width           =   2532
      End
      Begin VB.TextBox textAddDate 
         Height          =   285
         Left            =   5520
         MaxLength       =   7
         TabIndex        =   9
         Top             =   3045
         Width           =   2532
      End
      Begin VB.TextBox textPriorityDoc 
         Height          =   285
         Left            =   2190
         MaxLength       =   1
         TabIndex        =   8
         Top             =   3045
         Width           =   372
      End
      Begin VB.CommandButton cmdPriority 
         Caption         =   "¿é¤J(&V)"
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   2745
         Width           =   1332
      End
      Begin MSForms.TextBox textTM67 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   2160
         Width           =   6855
         VariousPropertyBits=   -1476378597
         MaxLength       =   200
         Size            =   "12086;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPS 
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   4230
         Width           =   6855
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12091;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM114 
         Height          =   285
         Left            =   -74130
         TabIndex        =   150
         Top             =   2580
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   8
         Left            =   -74130
         TabIndex        =   38
         Top             =   1675
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM113 
         Height          =   285
         Left            =   -74130
         TabIndex        =   149
         Top             =   2280
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM112 
         Height          =   285
         Left            =   -74130
         TabIndex        =   148
         Top             =   1985
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM106 
         Height          =   285
         Left            =   -74130
         TabIndex        =   147
         Top             =   790
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM107 
         Height          =   285
         Left            =   -74130
         TabIndex        =   146
         Top             =   1085
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   6
         Left            =   -74130
         TabIndex        =   36
         Top             =   480
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM108 
         Height          =   285
         Left            =   -74130
         TabIndex        =   145
         Top             =   1380
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM117 
         Height          =   285
         Left            =   -70050
         TabIndex        =   144
         Top             =   2580
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   9
         Left            =   -70050
         TabIndex        =   39
         Top             =   1675
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM116 
         Height          =   285
         Left            =   -70050
         TabIndex        =   143
         Top             =   2280
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM115 
         Height          =   285
         Left            =   -70050
         TabIndex        =   142
         Top             =   1985
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM105 
         Height          =   285
         Left            =   -70140
         TabIndex        =   141
         Top             =   3705
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM104 
         Height          =   285
         Left            =   -70140
         TabIndex        =   140
         Top             =   3415
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM103 
         Height          =   285
         Left            =   -70140
         TabIndex        =   139
         Top             =   3129
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   5
         Left            =   -70140
         TabIndex        =   35
         Top             =   2828
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM99 
         Height          =   285
         Left            =   -70140
         TabIndex        =   138
         Top             =   2542
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM98 
         Height          =   285
         Left            =   -70140
         TabIndex        =   137
         Top             =   2256
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM97 
         Height          =   285
         Left            =   -70140
         TabIndex        =   136
         Top             =   1970
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   3
         Left            =   -70140
         TabIndex        =   33
         Top             =   1669
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM52 
         Height          =   285
         Left            =   -70140
         TabIndex        =   135
         Top             =   1383
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM51 
         Height          =   285
         Left            =   -70140
         TabIndex        =   134
         Top             =   1097
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM50 
         Height          =   285
         Left            =   -70140
         TabIndex        =   133
         Top             =   811
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   1
         Left            =   -70140
         TabIndex        =   31
         Top             =   510
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM111 
         Height          =   285
         Left            =   -70050
         TabIndex        =   132
         Top             =   1380
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   7
         Left            =   -70050
         TabIndex        =   37
         Top             =   480
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM110 
         Height          =   285
         Left            =   -70050
         TabIndex        =   131
         Top             =   1085
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM109 
         Height          =   285
         Left            =   -70050
         TabIndex        =   130
         Top             =   790
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H8"
         Height          =   180
         Index           =   3
         Left            =   -70830
         TabIndex        =   129
         Top             =   540
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H7"
         Height          =   180
         Index           =   3
         Left            =   -74910
         TabIndex        =   128
         Top             =   540
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   19
         Left            =   -74580
         TabIndex        =   127
         Top             =   838
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   20
         Left            =   -74580
         TabIndex        =   126
         Top             =   1136
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   21
         Left            =   -74580
         TabIndex        =   125
         Top             =   1434
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   22
         Left            =   -70500
         TabIndex        =   124
         Top             =   838
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   23
         Left            =   -70500
         TabIndex        =   123
         Top             =   1136
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   24
         Left            =   -70500
         TabIndex        =   122
         Top             =   1434
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H10"
         Height          =   180
         Index           =   4
         Left            =   -70830
         TabIndex        =   121
         Top             =   1732
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H9"
         Height          =   180
         Index           =   4
         Left            =   -74910
         TabIndex        =   120
         Top             =   1732
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   25
         Left            =   -74580
         TabIndex        =   119
         Top             =   2030
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   26
         Left            =   -74580
         TabIndex        =   118
         Top             =   2328
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   27
         Left            =   -74580
         TabIndex        =   117
         Top             =   2632
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   28
         Left            =   -70500
         TabIndex        =   116
         Top             =   2030
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   29
         Left            =   -70500
         TabIndex        =   115
         Top             =   2328
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   30
         Left            =   -70500
         TabIndex        =   114
         Top             =   2632
         Width           =   345
      End
      Begin MSForms.TextBox textTM102 
         Height          =   285
         Left            =   -74190
         TabIndex        =   113
         Top             =   3705
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM101 
         Height          =   285
         Left            =   -74190
         TabIndex        =   112
         Top             =   3415
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM100 
         Height          =   285
         Left            =   -74190
         TabIndex        =   111
         Top             =   3129
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   4
         Left            =   -74190
         TabIndex        =   34
         Top             =   2828
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM96 
         Height          =   285
         Left            =   -74190
         TabIndex        =   110
         Top             =   2542
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM95 
         Height          =   285
         Left            =   -74190
         TabIndex        =   109
         Top             =   2256
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM94 
         Height          =   285
         Left            =   -74190
         TabIndex        =   108
         Top             =   1970
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   2
         Left            =   -74190
         TabIndex        =   32
         Top             =   1669
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM49 
         Height          =   285
         Left            =   -74190
         TabIndex        =   107
         Top             =   1383
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM48 
         Height          =   285
         Left            =   -74190
         TabIndex        =   106
         Top             =   1097
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM47 
         Height          =   285
         Left            =   -74190
         TabIndex        =   105
         Top             =   811
         Width           =   3255
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "5741;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   0
         Left            =   -74190
         TabIndex        =   30
         Top             =   510
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label32 
         Caption         =   "¨Ó¨ç´Á­­:"
         Height          =   255
         Left            =   150
         TabIndex        =   103
         Top             =   4950
         Width           =   855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "©w½Z¤Wªk±ø :"
         Height          =   180
         Left            =   5520
         TabIndex        =   100
         Top             =   4575
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   18
         Left            =   -70530
         TabIndex        =   99
         Top             =   3757
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   17
         Left            =   -70530
         TabIndex        =   98
         Top             =   3460
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   16
         Left            =   -70530
         TabIndex        =   97
         Top             =   3171
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   15
         Left            =   -74610
         TabIndex        =   96
         Top             =   3757
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   14
         Left            =   -74610
         TabIndex        =   95
         Top             =   3460
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   13
         Left            =   -74610
         TabIndex        =   94
         Top             =   3171
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H5"
         Height          =   180
         Index           =   2
         Left            =   -74940
         TabIndex        =   93
         Top             =   2882
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H6"
         Height          =   180
         Index           =   1
         Left            =   -70860
         TabIndex        =   92
         Top             =   2882
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   12
         Left            =   -70530
         TabIndex        =   91
         Top             =   2593
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   11
         Left            =   -70530
         TabIndex        =   90
         Top             =   2304
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   10
         Left            =   -70530
         TabIndex        =   89
         Top             =   2015
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   9
         Left            =   -74610
         TabIndex        =   88
         Top             =   2593
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   2
         Left            =   -74610
         TabIndex        =   87
         Top             =   2304
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   1
         Left            =   -74610
         TabIndex        =   86
         Top             =   2015
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H3"
         Height          =   180
         Index           =   0
         Left            =   -74940
         TabIndex        =   85
         Top             =   1726
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H4"
         Height          =   180
         Index           =   0
         Left            =   -70860
         TabIndex        =   84
         Top             =   1726
         Width           =   630
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "¸É¤å¥ó´Á­­ :"
         Height          =   180
         Left            =   6000
         TabIndex        =   83
         Top             =   3690
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¤l®×·sªk©w´Á­­ :"
         Height          =   180
         Index           =   17
         Left            =   2760
         TabIndex        =   82
         Top             =   4575
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¤l®×·s¥»©Ò´Á­­ :"
         Height          =   180
         Index           =   18
         Left            =   150
         TabIndex        =   81
         Top             =   4575
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "(N:¤£¦L)"
         Height          =   180
         Left            =   5010
         TabIndex        =   80
         Top             =   3690
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "¦C¦LÄ¶¤å :"
         Height          =   180
         Left            =   3540
         TabIndex        =   79
         Top             =   3690
         Width           =   810
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H2"
         Height          =   180
         Index           =   2
         Left            =   -70860
         TabIndex        =   78
         Top             =   570
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H1"
         Height          =   180
         Index           =   1
         Left            =   -74940
         TabIndex        =   77
         Top             =   570
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   3
         Left            =   -74610
         TabIndex        =   76
         Top             =   859
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   4
         Left            =   -74610
         TabIndex        =   75
         Top             =   1148
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   5
         Left            =   -74610
         TabIndex        =   74
         Top             =   1437
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   6
         Left            =   -70530
         TabIndex        =   73
         Top             =   859
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   7
         Left            =   -70530
         TabIndex        =   72
         Top             =   1148
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   8
         Left            =   -70530
         TabIndex        =   71
         Top             =   1437
         Width           =   345
      End
      Begin VB.Label Label37 
         Caption         =   "(Y:¿é¤J)"
         Height          =   255
         Left            =   2010
         TabIndex        =   70
         Top             =   3660
         Width           =   855
      End
      Begin VB.Label Label36 
         Caption         =   "¬O§_¿é¤JD/N :"
         Height          =   255
         Left            =   150
         TabIndex        =   69
         Top             =   3645
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "°Ó«~Ãþ§O :"
         Height          =   180
         Index           =   7
         Left            =   150
         TabIndex        =   68
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð°ê®a :"
         Height          =   180
         Index           =   8
         Left            =   4470
         TabIndex        =   67
         Top             =   405
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥»©Ò®×¸¹ :"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   66
         Top             =   405
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "®×¥ó¦WºÙ :"
         Height          =   180
         Left            =   150
         TabIndex        =   65
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "°Ó¼ÐºØÃþ :"
         Height          =   180
         Index           =   4
         Left            =   150
         TabIndex        =   64
         Top             =   1020
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "°Ó«~²Õ¸s :"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   63
         Top             =   1605
         Width           =   810
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð®×¸¹ :"
         Height          =   180
         Left            =   4470
         TabIndex        =   62
         Top             =   1905
         Width           =   810
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤é :"
         Height          =   180
         Left            =   150
         TabIndex        =   61
         Top             =   1905
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¥¿°Ó¼Ð¸¹¼Æ:"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   60
         Top             =   2505
         Width           =   945
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_¸É¥ó(¥i½Æ¿ï) :"
         Height          =   180
         Left            =   150
         TabIndex        =   59
         Top             =   3390
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "(1:©e¥ôª¬ 2:Àu¥ýÅvÃÒ©ú 3:¤½¥qµù¥UÃÒ©ú 4:­Ó¤H¨­¥÷ÃÒ©ú)"
         Height          =   180
         Left            =   2610
         TabIndex        =   58
         Top             =   3390
         Width           =   5595
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_¦C¦LÂ½Ä¶¨ç :"
         Height          =   180
         Left            =   3030
         TabIndex        =   57
         Top             =   3982
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(N:¤£¦L)"
         Height          =   180
         Index           =   0
         Left            =   5490
         TabIndex        =   56
         Top             =   3982
         Width           =   645
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "¦C¦L©w½Z :"
         Height          =   180
         Left            =   150
         TabIndex        =   55
         Top             =   3975
         Width           =   810
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:¤£¦L)"
         Height          =   180
         Left            =   1650
         TabIndex        =   54
         Top             =   3982
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "¦C¦L³Æµù :"
         Height          =   180
         Left            =   150
         TabIndex        =   53
         Top             =   4275
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "¨Ó¨ç¦¬¤å¤é :"
         Height          =   180
         Left            =   4470
         TabIndex        =   52
         Top             =   2505
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "©ñ±ó±M¥ÎÅv :"
         Height          =   180
         Left            =   150
         TabIndex        =   51
         Top             =   2190
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "¸ÉÀu¥ýÅvÃÒ©ú´Á­­ :"
         Height          =   180
         Left            =   3960
         TabIndex        =   50
         Top             =   3090
         Width           =   1530
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó :"
         Height          =   180
         Left            =   150
         TabIndex        =   49
         Top             =   3090
         Width           =   1890
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "(Y / N)"
         Height          =   180
         Left            =   2640
         TabIndex        =   48
         Top             =   3090
         Width           =   495
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Àu¥ýÅv¸ê®Æ :"
         Height          =   180
         Left            =   150
         TabIndex        =   47
         Top             =   2790
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7140
      TabIndex        =   41
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4980
      TabIndex        =   29
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "¦^«eµe­±(&U)"
      Height          =   400
      Left            =   5940
      TabIndex        =   40
      Top             =   0
      Width           =   1152
   End
   Begin VB.Label LabNP07 
      Height          =   255
      Left            =   0
      TabIndex        =   104
      Top             =   0
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "frm030203_02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/11 §ï¦¨Form2.0 ;cmbTM05¡BtextPS¡BCombo2(index)¡BtextTM47~52¡BtextTM94~117¡BtextTM67(111/8/8)
'Memo By Sindy 2012/12/4 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/16 SQLDate¤wÀË¬d
'Memo By Sindy 2010/11/29 ­û¤u½s¸¹Äæ¤w­×§ï
'Memo By Sindy 2010/8/11 ¤é´ÁÄæ¤w­×§ï
Option Explicit

' ¥»©Ò®×¸¹
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' °Ó¼ÐºØÃþ
Dim m_TM08 As String
' °ê®a¥N½X
Dim m_TM10 As String
' ·~°È°Ï
Dim m_CP12 As String
' ´¼Åv¤H­û
Dim m_CP13 As String
Dim m_CP14 As String 'Add By Sindy 2010/10/25
' Á`¦¬¤å¸¹
Dim m_CP09 As String
' ©ñ±ó±M¥ÎÅv
Dim m_TM67 As String
'®×¥ó©Ê½è
Dim m_CP10 As String
'Add By Cheng 2003/01/23
Dim m_blnPriDate As Boolean '¬O§_¦³Àu¥ýÅv
'Add By Cheng 2004/05/10
Dim m_strLanguage As String '©w½Z»y¤å
'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
Dim m_Priority(1 To 6) As String 'Modify by Amy 2014/03/27 +pd08,pd09
Dim m_Pa(1 To 4) As String '¥»©Ò®×¸¹
'add by nick 2004/08/13
Dim m_CP27 As String
Dim strCP05 As String
Dim ii As Integer
Dim rsTmp As New ADODB.Recordset
'add by nickc 2006/07/28
Public UpForm As Form
Dim m_MonTM01 As String     '¬ö¿ý¤À³Î¥À®×®×¸¹
Dim m_MonTM02 As String
Dim m_MonTM03 As String
Dim m_MonTM04 As String
Public m_MonCP09 As String  '¶Ç¤J¤À³Î¥À®×¦¬¤å¸¹
Dim m_MonNP08 As String
Dim m_MonNP09 As String
'92.04.16 nick ¬ö¿ý§@¥Î«öÁä
Public cmdState As Integer
'add by nick 2004/10/05 ÀË¬d¬O§_¤w¸g¦³°Ó«~¤ÎªA°È
Public ChkTG As Boolean
Dim strRvType As String 'Add By Sindy 2012/5/18
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
Dim m_TM44 As String, m_fa76 As String 'Add By Sindy 2017/3/9
'Added by Morgan 2017/6/14 ¹q¤l¤½¤å
Public m_DocWord As String
Public m_DocNo As String
Public m_DocPdf As String
Public m_DocPdfDate As String
Public m_DocPdfTime As String
'end 2017/6/14
Dim m_ET03 As String 'Add By Sindy 2022/3/10


' ­ì¸ê®Æ¬O§_¦³¹ê»Úµ²ªG
Private Sub cmdCancel_Click()
'add by nickc 2008/01/23 ¥[¤J¥i¥H¨ú®ø
If UpForm Is Nothing Or Me.Visible = False Then
   Unload Me
   frm030203_01.Show
Else
    'add by nickc 2008/01/23 ¥[¤J¥i¥H¨ú®ø
    If UpForm Is frm02010401_6 Then
        frm02010401_6.m_IsCancal = True
        Unload Me
    End If
End If
End Sub

Private Sub cmdExit_Click()
   Unload frm030203_01
   Unload Me
End Sub

Public Sub cmdok_Click(Index As Integer)
'92.04.16 nick ¬ö¿ý§@¥Î«öÁä
cmdState = Index
PubShowNextData
Exit Sub
End Sub

'Add By Sindy 2009/05/14
Public Sub PubShowNextData()
Select Case cmdState
Case 0
      If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
      If TxtValidate = False Then Exit Sub
      'add by nickc 2006/08/02
      If UpForm Is Nothing Or Me.Visible = False Then
            ' ³]©w·Æ¹«´å¼Ð¬°µ¥«Ýª¬ºA
            Screen.MousePointer = vbHourglass
            ' Àx¦s¸ê®Æ
            'edit by  nick 2004/11/03
            'OnSaveData
            If OnSaveData = False Then MsgBox "¦sÀÉ¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
            ' ³]©w·Æ¹«´å¼Ð¬°¹w³]
            Screen.MousePointer = vbDefault
            
            'add by nickc 2006/09/29
            If Me.Visible = True And UpForm Is Nothing Then
                  If textDN = "Y" Then
                     'Add By Cheng 2003/03/19
                     '·s¼W¦a§}±ø¦Cªí¸ê®Æ
                     'Modify By Sindy 2025/10/2 ¨ú®ø¦a§}±ø
'                     pub_AddressListSN = pub_AddressListSN + 1
'                     PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
                      Screen.MousePointer = vbHourglass
                      Frmacc21h0.Show
                      mdiMain.ToolShow
                      mdiMain.tool1_enabled
                      Screen.MousePointer = vbDefault
                      Set Frmacc21h0.frmlink = frm030203_01
                      'add by nick 2004/11/24
                      Frmacc21h0.IsPrintAddress = False
                      Unload Me
                   Else
                     Unload Me
                     frm030203_01.Show
                   End If
           End If
      End If
       'add by nickc 2006/08/02
       If UpForm Is frm02010401_6 Then
          '­Y¬Oµe­±¦³¥X²{¥i¥H¿é¸ê®Æ¡A­n±N¸ê®Æ¥á¦^«e­±¦s
          If Me.Visible = True Then
            frm02010401_6.PutSeekData01 = textTM09
            frm02010401_6.PutSeekData02 = textTM32
            frm02010401_6.PutSeekData03 = textTM11
            frm02010401_6.PutSeekData04 = textTM12
            frm02010401_6.PutSeekData05 = textTM27
            frm02010401_6.PutSeekData06 = textCP05
            frm02010401_6.PutSeekData07 = textPrint
            frm02010401_6.PutSeekData08 = textPriorityDoc
            frm02010401_6.PutSeekData09 = textAddDate
            frm02010401_6.PutSeekData10 = textAdd
            frm02010401_6.PutSeekData11 = textDN
            frm02010401_6.PutSeekData12 = txtToEng
            frm02010401_6.PutSeekData13 = textPrtTrans
            frm02010401_6.PutSeekData14 = textPS
            frm02010401_6.PutSeekData15 = textTM67
            frm02010401_6.PutSeekData16 = textNP08
            frm02010401_6.PutSeekData17 = textNP09
            'add by nickc 2007/05/01 ¥[¤J¥Nªí¤H
            frm02010401_6.PutSeekData18 = Combo2(0).Text
            frm02010401_6.PutSeekData19 = Combo2(1).Text
            frm02010401_6.PutSeekData20 = Combo2(2).Text
            frm02010401_6.PutSeekData21 = Combo2(3).Text
            frm02010401_6.PutSeekData22 = Combo2(4).Text
            frm02010401_6.PutSeekData23 = Combo2(5).Text
            frm02010401_6.PutSeekData24 = Combo2(6).Text
            frm02010401_6.PutSeekData25 = Combo2(7).Text
            frm02010401_6.PutSeekData26 = Combo2(8).Text
            frm02010401_6.PutSeekData27 = Combo2(9).Text
            frm02010401_6.PutSeekData28 = textTM47
            frm02010401_6.PutSeekData29 = textTM48
            frm02010401_6.PutSeekData30 = textTM49
            frm02010401_6.PutSeekData31 = textTM50
            frm02010401_6.PutSeekData32 = textTM51
            frm02010401_6.PutSeekData33 = textTM52
            frm02010401_6.PutSeekData34 = textTM94
            frm02010401_6.PutSeekData35 = textTM95
            frm02010401_6.PutSeekData36 = textTM96
            frm02010401_6.PutSeekData37 = textTM97
            frm02010401_6.PutSeekData38 = textTM98
            frm02010401_6.PutSeekData39 = textTM99
            frm02010401_6.PutSeekData40 = textTM100
            frm02010401_6.PutSeekData41 = textTM101
            frm02010401_6.PutSeekData42 = textTM102
            frm02010401_6.PutSeekData43 = textTM103
            frm02010401_6.PutSeekData44 = textTM104
            frm02010401_6.PutSeekData45 = textTM105
            frm02010401_6.PutSeekData46 = TextTM106
            frm02010401_6.PutSeekData47 = TextTM107
            frm02010401_6.PutSeekData48 = TextTM108
            frm02010401_6.PutSeekData49 = TextTM109
            frm02010401_6.PutSeekData50 = TextTM110
            frm02010401_6.PutSeekData51 = TextTM111
            frm02010401_6.PutSeekData52 = TextTM112
            frm02010401_6.PutSeekData53 = TextTM113
            frm02010401_6.PutSeekData54 = TextTM114
            frm02010401_6.PutSeekData55 = TextTM115
            frm02010401_6.PutSeekData56 = TextTM116
            frm02010401_6.PutSeekData57 = TextTM117
            'add by nickc 2007/08/10
            frm02010401_6.PutSeekData58 = txtLaw
          End If
          Unload Me
       End If
   End If
   
'add by nick 2004/10/05
Case 6
    'frm03010303_04.Hide 'Modify By Sindy 2009/09/17
    Set frm03010303_04.UpForm = Me
    frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 'textTMKey 'lbl1(0).Caption
    frm03010303_04.AllClass = textTM09 'txt1(0).Text
    frm03010303_04.cmdOK(0).Visible = False
    frm03010303_04.cmd.Visible = False
    frm03010303_04.cmd2.Visible = False
    frm03010303_04.txt2(0).Visible = False
    frm03010303_04.Line1.Visible = False
    frm03010303_04.txt2(1).Visible = False
    frm03010303_04.txt2(2).Visible = False
    frm03010303_04.txt2(3).Visible = False
    frm03010303_04.Caption = "°Ó«~¤ÎªA°È¸ê®Æ"
    'edit by nickc 2008/02/12 §ï¦¨¥i¥H½Æ»s
    'frm03010303_04.TXT1(0).Enabled = False
    'frm03010303_04.TXT1(1).Enabled = False
    'frm03010303_04.TXT1(2).Enabled = False
    frm03010303_04.TXT1(0).Locked = True
    frm03010303_04.TXT1(1).Locked = True
    frm03010303_04.TXT1(2).Locked = True
    frm03010303_04.Label2.Visible = False
    'Me.Hide 'Modify By Sindy 2009/09/17
    frm03010303_04.QueryData
    frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 §ï¬°±j¨î¦^À³ªí³æ
End Select
End Sub

Private Sub Form_Load()
   ' ³]©w±±¨î¶µªº­I´ºÃC¦â
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   'edit by nick 2004/08/18
   'textTM09.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   'edit by nick 2004/08/18
   'textTM32.BackColor = &H8000000F
   
   'Added by Lydia 2016/09/10 ³]©w¥Nªí¤H¤¤¤å¦WºÙ©M­^¤å¦WºÙªø«×
    textTM47.MaxLength = Pub_MaxCEL10
    textTM48.MaxLength = Pub_MaxCEL11
    textTM50.MaxLength = Pub_MaxCEL10
    textTM51.MaxLength = Pub_MaxCEL11
    textTM94.MaxLength = Pub_MaxCEL10
    textTM95.MaxLength = Pub_MaxCEL11
    textTM97.MaxLength = Pub_MaxCEL10
    textTM98.MaxLength = Pub_MaxCEL11
    textTM100.MaxLength = Pub_MaxCEL10
    textTM101.MaxLength = Pub_MaxCEL11
    textTM103.MaxLength = Pub_MaxCEL10
    textTM104.MaxLength = Pub_MaxCEL11
    TextTM106.MaxLength = Pub_MaxCEL10
    TextTM107.MaxLength = Pub_MaxCEL11
    TextTM109.MaxLength = Pub_MaxCEL10
    TextTM110.MaxLength = Pub_MaxCEL11
    TextTM112.MaxLength = Pub_MaxCEL10
    TextTM113.MaxLength = Pub_MaxCEL11
    TextTM115.MaxLength = Pub_MaxCEL10
    TextTM116.MaxLength = Pub_MaxCEL11
   'end 2016/09/10
   
   MoveFormToCenter Me
   'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
   frm880002.m_blnAddNew = False
   'add by nick 2004/08/18
   SSTab1.Tab = 0
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' ²M°£·j´MªºKey
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
   End If
   
   Select Case nType
      ' ¥»©Ò®×¸¹ Äæ¦ì1
      Case 0: m_TM01 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì2
      Case 1: m_TM02 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì3
      Case 2: m_TM03 = strData
      ' ¥»©Ò®×¸¹ Äæ¦ì4
      Case 3: m_TM04 = strData
      ' Á`¦¬¤å¸¹
      Case 4: m_CP09 = strData
   End Select
End Sub

' Åª¨ú°Ó¼Ð°ò¥»ÀÉ
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim strSub As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSub As ADODB.Recordset
   
   ' ¨ú±o°Ó¼Ð°ò¥»ÀÉªº¬ÛÃö¶µ¥Ø
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
      ' ¥Ó½Ð°ê®a
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"))
      End If
      ' °Ó¼Ð¦WºÙ(¤¤)
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' °Ó¼Ð¦WºÙ(­^)
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' °Ó¼Ð¦WºÙ(¤é)
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' Åã¥Ü°Ó¼Ð¦WºÙ
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' °Ó¼ÐºØÃþ
      If IsNull(rsTmp.Fields("TM08")) = False Then
         m_TM08 = rsTmp.Fields("TM08")
         textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      End If
      ' °Ó«~Ãþ§O
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      ' °Ó«~²Õ¸s
      'edit by nick 2004/08/18 ³¯ª÷½¬»¡¹w³]¬°°Ó«~Ãþ§O
      'If IsNull(rsTmp.Fields("TM32")) = False Then
      '   textTM32 = rsTmp.Fields("TM32")
      'End If
      textTM32 = textTM09
      ' ©ñ±ó±M¥ÎÅv
      If IsNull(rsTmp.Fields("TM67")) = False Then
         m_TM67 = rsTmp.Fields("TM67")
      End If
      '93.9.30 ADD BY SONIA
      ' ¥Ó½Ð¤é´Á
      If IsNull(rsTmp.Fields("TM11")) = False Then
         textTM11 = ChangeWStringToTString((rsTmp.Fields("TM11")))
      End If
      ' ¥Ó½Ð®×¸¹
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      'Add By Sindy 2017/3/9
      ' FC¥N²z¤H
      m_TM44 = ""
      If IsNull(rsTmp.Fields("TM44")) = False Then
         m_TM44 = rsTmp.Fields("TM44")
      End If
      '2017/3/9 END
      '93.9.30 END
      'Add By Cheng 2003/03/11
      'Åã¥Ü©ñ±ó±M¥ÎÅv
      Me.textTM67.Text = "" & rsTmp("TM67").Value
      '¥Nªí¤H
      Dim i As Integer, j As Integer
      For i = 0 To 9 ' edit by nickc 2007/05/01  1
         Combo2(i).AddItem ""
      Next i
      
      'Modified by Lydia 2019/03/18 ­×§ï¥Nªí¤H1~10ªº¤U©Ô¿ï³æ,³£±a¥X©Ò¦³¥Ó½Ð¤H¦³¿é¤Jªº¥Nªí¤H
'      If rsTmp.Fields("TM23").Value <> "" Then
'         'edit by nickc 2008/04/08 §ï¦¨  ­^->¤¤->¤é
'         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM23").Value)
'         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM23").Value)
'
'         intI = 1
'         'edit by nickc 2007/02/06 ¤£¥Î dll ¤F
'         'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            For j = 1 To 6
'               If IsNull(RsTemp.Fields(j - 1)) Then
'                  strExc(0) = ""
'               Else
'                  strExc(0) = "-" & RsTemp.Fields(j - 1)
'               End If
'               Combo2(0).AddItem rsTmp.Fields("TM23").Value & "-" & j & strExc(0)
'               Combo2(1).AddItem rsTmp.Fields("TM23").Value & "-" & j & strExc(0)
'            Next
'         End If
'      End If
'      'add by nickc 2007/05/01 ¥[¦h¥Nªí¤H
'      If rsTmp.Fields("TM78").Value <> "" Then
'         'edit by nickc 2008/04/08 §ï¦¨  ­^->¤¤->¤é
'         strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM78").Value)
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            For j = 1 To 6
'               If IsNull(RsTemp.Fields(j - 1)) Then
'                  strExc(0) = ""
'               Else
'                  strExc(0) = "-" & RsTemp.Fields(j - 1)
'               End If
'               Combo2(2).AddItem rsTmp.Fields("TM78").Value & "-" & j & strExc(0)
'               Combo2(3).AddItem rsTmp.Fields("TM78").Value & "-" & j & strExc(0)
'            Next
'         End If
'      End If
'      If rsTmp.Fields("TM79").Value <> "" Then
'         'edit by nickc 2008/04/08 §ï¦¨  ­^->¤¤->¤é
'         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM79").Value)
'         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM79").Value)
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            For j = 1 To 6
'               If IsNull(RsTemp.Fields(j - 1)) Then
'                  strExc(0) = ""
'               Else
'                  strExc(0) = "-" & RsTemp.Fields(j - 1)
'               End If
'               Combo2(4).AddItem rsTmp.Fields("TM79").Value & "-" & j & strExc(0)
'               Combo2(5).AddItem rsTmp.Fields("TM79").Value & "-" & j & strExc(0)
'            Next
'         End If
'      End If
'      If rsTmp.Fields("TM80").Value <> "" Then
'         'edit by nickc 2008/04/08 §ï¦¨  ­^->¤¤->¤é
'         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM80").Value)
'         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM80").Value)
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            For j = 1 To 6
'               If IsNull(RsTemp.Fields(j - 1)) Then
'                  strExc(0) = ""
'               Else
'                  strExc(0) = "-" & RsTemp.Fields(j - 1)
'               End If
'               Combo2(6).AddItem rsTmp.Fields("TM80").Value & "-" & j & strExc(0)
'               Combo2(7).AddItem rsTmp.Fields("TM80").Value & "-" & j & strExc(0)
'            Next
'         End If
'      End If
'      If rsTmp.Fields("TM81").Value <> "" Then
'         'edit by nickc 2008/04/08 §ï¦¨  ­^->¤¤->¤é
'         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM81").Value)
'         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM81").Value)
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            For j = 1 To 6
'               If IsNull(RsTemp.Fields(j - 1)) Then
'                  strExc(0) = ""
'               Else
'                  strExc(0) = "-" & RsTemp.Fields(j - 1)
'               End If
'               Combo2(8).AddItem rsTmp.Fields("TM81").Value & "-" & j & strExc(0)
'               Combo2(9).AddItem rsTmp.Fields("TM81").Value & "-" & j & strExc(0)
'            Next
'         End If
'      End If
      If "" & rsTmp.Fields("TM23") <> "" Then
            '­×§ï¥Nªí¤H1~10ªº¤U©Ô¿ï³æ,³£±a¥X©Ò¦³¥Ó½Ð¤H¦³¿é¤Jªº¥Nªí¤H
            strExc(0) = "SELECT '1' ord1, nvl(CU40,nvl(cu39,cu41)) B1,nvl(CU43,nvl(cu42,cu44)) B2,nvl(CU46,nvl(cu45,cu47)) B3,nvl(CU49,nvl(cu48,cu50)) B4,nvl(CU52,nvl(cu51,cu53)) B5,nvl(CU55,nvl(cu54,cu56)) B6, CU01||CU02 as CustNo FROM CUSTOMER WHERE " & ChgCustomer("" & rsTmp.Fields("TM23"))
            If "" & rsTmp.Fields("TM78") <> "" Then strExc(0) = strExc(0) & " Union SELECT '2' ord1, nvl(CU40,nvl(cu39,cu41)) B1,nvl(CU43,nvl(cu42,cu44)) B2,nvl(CU46,nvl(cu45,cu47)) B3,nvl(CU49,nvl(cu48,cu50)) B4,nvl(CU52,nvl(cu51,cu53)) B5,nvl(CU55,nvl(cu54,cu56)) B6, CU01||CU02 as CustNo FROM CUSTOMER WHERE " & ChgCustomer("" & rsTmp.Fields("TM78"))
            If "" & rsTmp.Fields("TM79") <> "" Then strExc(0) = strExc(0) & " Union SELECT '3' ord1, nvl(CU40,nvl(cu39,cu41)) B1,nvl(CU43,nvl(cu42,cu44)) B2,nvl(CU46,nvl(cu45,cu47)) B3,nvl(CU49,nvl(cu48,cu50)) B4,nvl(CU52,nvl(cu51,cu53)) B5,nvl(CU55,nvl(cu54,cu56)) B6, CU01||CU02 as CustNo FROM CUSTOMER WHERE " & ChgCustomer("" & rsTmp.Fields("TM79"))
            If "" & rsTmp.Fields("TM80") <> "" Then strExc(0) = strExc(0) & " Union SELECT '4' ord1, nvl(CU40,nvl(cu39,cu41)) B1,nvl(CU43,nvl(cu42,cu44)) B2,nvl(CU46,nvl(cu45,cu47)) B3,nvl(CU49,nvl(cu48,cu50)) B4,nvl(CU52,nvl(cu51,cu53)) B5,nvl(CU55,nvl(cu54,cu56)) B6, CU01||CU02 as CustNo FROM CUSTOMER WHERE " & ChgCustomer("" & rsTmp.Fields("TM80"))
            If "" & rsTmp.Fields("TM81") <> "" Then strExc(0) = strExc(0) & " Union SELECT '5' ord1, nvl(CU40,nvl(cu39,cu41)) B1,nvl(CU43,nvl(cu42,cu44)) B2,nvl(CU46,nvl(cu45,cu47)) B3,nvl(CU49,nvl(cu48,cu50)) B4,nvl(CU52,nvl(cu51,cu53)) B5,nvl(CU55,nvl(cu54,cu56)) B6, CU01||CU02 as CustNo FROM CUSTOMER WHERE " & ChgCustomer("" & rsTmp.Fields("TM81"))
            
            strExc(0) = strExc(0) & " order by 1"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                RsTemp.MoveFirst
                Do While Not RsTemp.EOF
                    For j = 1 To 6
                       If "" & RsTemp.Fields(j) <> "" Then
                           strExc(0) = "-" & RsTemp.Fields(j)
                           For i = 0 To 9
                               Combo2(i).AddItem "" & RsTemp.Fields("CustNo").Value & "-" & j & strExc(0)
                           Next i
                       End If
                    Next j
                    RsTemp.MoveNext
                Loop
            End If
      End If
      'end 2019/03/18
      
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' ¨ú±o®×¥ó¶i«×ÀÉªºÄæ¦ì¤º®e
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' ¨t²Î¤é
   strDate = DBDATE(SystemDate())
   ' ¦¬¤å¸¹
   'textCP09 = m_CP09
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' ·~°È°Ï
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' ´¼Åv¤H­û
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
      End If
      'Add By Sindy 2010/10/25 ©Ó¿ì¤H
      m_CP14 = Empty
      If IsNull(rsTmp.Fields("CP14")) = False Then
         m_CP14 = rsTmp.Fields("CP14")
      End If
      ' ®×¥ó©Ê½è
      If IsNull(rsTmp.Fields("CP10")) = False Then
        m_CP10 = rsTmp.Fields("CP10")
      End If
      'add by nick 2004/08/13 ºâ¸É¤å¥ó¤é¥Î
      'µo¤å¤é
      If IsNull(rsTmp.Fields("CP27")) = False Then
        m_CP27 = ChangeWStringToTString(CheckStr(rsTmp.Fields("CP27")))
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   If m_CP10 = "308" Then textPrint = "N" '2009/4/30 ADD BY SONIA¤À³Î¤l®×¹w³]¤£¦L©w½Z
End Sub

Public Sub QueryData()
   
   ' Åª¨ú°ò¥»ÀÉ
   QueryTradeMark
   ' Åª¨ú®×¥ó¶i«×ÀÉ
   QueryCaseProgress
   
   'add by nickc 2006/09/29 Åª¨ú¥À®×¸ê®Æ
   If UpForm Is frm02010401_6 Then
       QueryMonTradeMark
   End If
   
   'add by nickc 2006/08/03
   If UpForm Is frm02010401_6 Then
        textCP05 = TAIWANDATE(UpForm.oStrCDate)
   Else
        'add by nick 2004/08/18 ¨Ó¨ç¦¬¤å¤é¹w³]¬°¨t²Î¤é
        textCP05 = ChangeWStringToTString(ServerDate)
   End If
   
   'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
   ' ¸É¤å¥ó´Á­­
   If textPriorityDoc = "N" Then
      EnableTextBox textAddDate, True
   Else
      textAddDate = Empty
      EnableTextBox textAddDate, False
   End If
   
   'add by nickc 2006/12/21
   textAddDate2 = Empty
   EnableTextBox textAddDate2, False
   
   'add by nickc 2006/08/03 ­Y¬O¤À³Î¤l®×¶i¤J¤£°µÀu¥ýÅv
   If UpForm Is Nothing Then
        ' Åª¨úÀu¥ýÅv¸ê®Æ
        m_Pa(1) = m_TM01
        m_Pa(2) = m_TM02
        m_Pa(3) = m_TM03
        m_Pa(4) = m_TM04
        'edit by nickc 2007/02/06 ¤£¥Î dll ¤F
        'objPublicData.ReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
        'Modify by Amy 2014/03/27 +pd08,pd09
        'Modify By Sindy 2017/10/12 + , m_Priority(6)
        ClsPDReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)
        '92.10.19 ADD BY SONIA
        If m_Priority(1) <> "" Then
           frm880002.m_blnAddNew = True
        End If
   End If
   
   Call ChgType 'Add By Sindy 2012/5/18 Åª¨ú¨Ó¨ç´Á­­
End Sub


Private Sub textAdd_KeyUp(KeyCode As Integer, Shift As Integer)
   If InStr(1, textAdd, "1") = 0 And InStr(1, textAdd, "3") = 0 Then
      textAddDate2 = Empty    '2011/1/4 add by sonia ®³±¼textadd©Î§ï¬°2®É,textAddDate2¤]­n²M
      EnableTextBox textAddDate2, False
   Else
      textAddDate2 = Empty
      EnableTextBox textAddDate2, True
      textAddDate2 = TAIWANDATE(DateAdd("d", -5, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_CP27)))))
   End If
End Sub

'add by nickc 2006/12/21
Private Sub textAddDate2_GotFocus()
InverseTextBox textAddDate2
End Sub
Private Sub textAddDate2_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textAddDate2) = False Then
      If CheckIsTaiwanDate(textAddDate2, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½Tªº¸É¤å¥ó´Á­­"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textAddDate2_GotFocus
      End If
   End If
End Sub

'add by nick 2004/08/18 ±q frm030202_03  copy ¨Ó
Private Sub textDN_GotFocus()
   InverseTextBox textDN
End Sub
Private Sub textDN_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
' ¬O§_¿é¤JD/N
Private Sub textDN_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textDN) = False Then
      Select Case textDN
         Case " ", "Y":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎY"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDN_GotFocus
      End Select
   End If
End Sub

Private Sub textNP08_GotFocus()
InverseTextBox textNP08
End Sub

Private Sub textNP08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDate As String
   
   Cancel = False
   If IsEmptyText(textNP08) = False Then
      If CheckIsTaiwanDate(textNP08, False) = False Then
         Cancel = True
         strMsg = "¤é´Á¤£¥¿½T"
         strTit = "¤l®×·s¥»©Ò´Á­­"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP08.SetFocus
         textNP08_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
      Else
          textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08, True), 1)
      'end 2020/07/07
      End If
   End If
EXITSUB:
End Sub

Private Sub textNP09_GotFocus()
    InverseTextBox textNP09
End Sub

Private Sub textNP09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDate As String
   
   Cancel = False
   If IsEmptyText(textNP09) = False Then
      If CheckIsTaiwanDate(textNP09, False) = False Then
         Cancel = True
         strMsg = "¤é´Á¤£¥¿½T"
         strTit = "¤l®×·sªk©w´Á­­"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP09_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
Private Sub textPriorityDoc_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Modified by Lydia 2019/03/18 ­ì¥»¥u¦³¥Nªí¤H1~2, ²{¦b¦³¥Nªí¤H1~10
'Private Sub Combo2_Click(Index As Integer)
'
'   Dim i As Integer, strTmp As String
'
'   If (Combo2(Index).Text = "") Then
'      For i = 0 To 2
'         Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = ""
'      Next i
'      Exit Sub
'   End If
'
'   strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
'   strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
'   strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
'
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 ¤£¥Î dll ¤F   = objLawDll.ReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      For i = 0 To 2
'         If Not IsNull(RsTemp.Fields(i)) Then
'            Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = RsTemp.Fields(i)
'         'add by nickc 2008/04/07
'         Else
'            Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = ""
'         End If
'      Next
'   End If
'End Sub
Private Sub Combo2_Click(Index As Integer)
Dim intA As Integer, intP As Integer
Dim strTmp As String

   If Index < 2 Then
       intA = 47 + 3 * Index
   Else
       intA = 94 + 3 * (Index - 2)
   End If
   
   '²MªÅ¥Nªí¤HXÄæ¦ì
   If Trim(Combo2(Index).Text) = "" Then
      For intP = 0 To 2
         Me.Controls("textTM" & Format(intA + intP, "#")).Text = ""
      Next intP
   '±a¥X¥Nªí¤Hªº¦WºÙ
   Else
       strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
       strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
       strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
       
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
          For intP = 0 To 2
             If Not IsNull(RsTemp.Fields(intP)) Then
                Me.Controls("textTM" & Format(intA + intP, "#")).Text = "" & RsTemp.Fields(intP)
             Else
                Me.Controls("textTM" & Format(intA + intP, "#")).Text = ""
             End If
          Next intP
       End If
   End If
End Sub
'end 2019/03/18

'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
' ¬O§_ªþ±aÀu¥ýÅvÃÒ©ú¤å¥ó
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
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JY©ÎN"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPriorityDoc_GotFocus
      End Select
   End If
   
   If textPriorityDoc = "N" Then
      EnableTextBox textAddDate, True
'     '¸ÉÀu¥ýÅvÃÒ©ú´Á­­¹w³]¬°µo¤å¤é¥[¤T­Ó¤ë´î¤@¤Ñ
      If IsEmptyText(textAddDate) = True And IsEmptyText(m_CP27) = False Then
        'Modify By Cheng 2003/09/02
'         textAddDate = TAIWANDATE(DateSerial(Val(DBYEAR(textCP27)), Val(DBMONTH(textCP27)) + 3, Val(DBDAY(textCP27)) - 1))
        'Modify By Cheng 2004/03/18
         '¸ÉÀu¥ýÅvÃÒ©ú´Á­­¹w³]¬°µo¤å¤é¥[¤T­Ó¤ë´î¤­¤Ñ
'         textAddDate = TAIWANDATE(DateAdd("d", -1, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(textCP27)))))
         'Modify By Sindy 2015/7/29
         '¸ÉÀu¥ýÅvÃÒ©ú´Á­­¹w³]¬°µe­±¤W¥Ó½Ð¤é¥[¤T­Ó¤ë
'         textAddDate = TAIWANDATE(DateAdd("d", -5, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_CP27)))))
         textAddDate = TAIWANDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(textTM11))))
         '2015/7/29 END
        'End
      End If
   Else
      textAddDate = Empty
      EnableTextBox textAddDate, False
   End If
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
Dim nIndex As Integer
Dim strSql As String
Dim strCP09 As String
Dim strCP10 As String
'Dim strCP12 As String
Dim strCP27 As String
'add by nick 2004/08/13
Dim strNP07 As String
Dim strNP08 As String
Dim strNP22 As String
Dim strCP06 As String
Dim strCP07 As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'add by nickc 2006/12/21
Dim tmpNP15 As String
Dim m_CP110 As String   'add by sonia 2018/12/13
    
   OnSaveData = True
   'add by nickc 2006/08/03
   If Me.Visible = True Then
        '911107 nick transation
       On Error GoTo CheckingErr
       cnnConnection.BeginTrans
   End If
   ' §ó·s°Ó¼Ð°ò¥»ÀÉ (¥Ó½Ð¤é, ¥Ó½Ð®×¸¹, ¥¿°Ó¼Ð¸¹¼Æ)
    'Modify By Cheng 2003/03/11
    '¥[§ó·s©ñ±ó±M¥ÎÅv
'   strSQL = "UPDATE TradeMark SET TM11 = " & DBDATE(textTM11) & ", " & _
'                                 "TM12 = '" & textTM12 & "', " & _
'                                 "TM27 = '" & textTM27 & "' " & _
'            "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                  "TM02 = '" & m_TM02 & "' AND " & _
'                  "TM03 = '" & m_TM03 & "' AND " & _
'                  "TM04 = '" & m_TM04 & "' "
'edit by nick 2004/08/18 ¥[¤J§ó·s°Ó«~Ãþ§O¤Î²Õ¸s¤Î¥Nªí¤H1 & 2
'   strSQL = "UPDATE TradeMark SET TM11 = " & DBDATE(textTM11) & ", " & _
'                                 "TM12 = '" & textTM12 & "', " & _
'                                 "TM27 = '" & textTM27 & "', " & _
'                                 "TM67 = '" & ChgSQL(textTM67) & "' " & _
'            "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                  "TM02 = '" & m_TM02 & "' AND " & _
'                  "TM03 = '" & m_TM03 & "' AND " & _
'                  "TM04 = '" & m_TM04 & "' "
'edit by nickc 2007/05/01 ¥[¤J¥Nªí¤H 3-10
'   strSQL = "UPDATE TradeMark SET TM11 = " & DBDATE(textTM11) & ", " & _
'                                 "TM12 = '" & textTM12 & "', " & _
'                                 "TM27 = '" & textTM27 & "', " & _
'                                 "TM47 = '" & ChgSQL(textTM47) & "', " & _
'                                 "TM48 = '" & ChgSQL(textTM48) & "', " & _
'                                 "TM49 = '" & ChgSQL(textTM49) & "', " & _
'                                 "TM50 = '" & ChgSQL(textTM50) & "', " & _
'                                 "TM51 = '" & ChgSQL(textTM51) & "', " & _
'                                 "TM52 = '" & ChgSQL(textTM52) & "', " & _
'                                 "TM09 = '" & textTM09 & "', " & _
'                                 "TM32 = '" & textTM32 & "', " & _
'                                 "TM67 = '" & ChgSQL(textTM67) & "' " & _
'            "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                  "TM02 = '" & m_TM02 & "' AND " & _
'                  "TM03 = '" & m_TM03 & "' AND " & _
'                  "TM04 = '" & m_TM04 & "' "
   strSql = "UPDATE TradeMark SET TM11 = " & DBDATE(textTM11) & ", " & _
                                 "TM12 = '" & textTM12 & "', " & _
                                 "TM27 = '" & textTM27 & "', " & _
                                 "TM47 = '" & ChgSQL(textTM47) & "', TM48 = '" & ChgSQL(textTM48) & "', " & _
                                 "TM49 = '" & ChgSQL(textTM49) & "', TM50 = '" & ChgSQL(textTM50) & "', " & _
                                 "TM51 = '" & ChgSQL(textTM51) & "', TM52 = '" & ChgSQL(textTM52) & "', " & _
                                 "TM94 = '" & ChgSQL(textTM94) & "', TM95 = '" & ChgSQL(textTM95) & "', " & _
                                 "TM96 = '" & ChgSQL(textTM96) & "', TM97 = '" & ChgSQL(textTM97) & "', " & _
                                 "TM98 = '" & ChgSQL(textTM98) & "', TM99 = '" & ChgSQL(textTM99) & "', " & _
                                 "TM100= '" & ChgSQL(textTM100) & "', TM101= '" & ChgSQL(textTM101) & "', " & _
                                 "TM102= '" & ChgSQL(textTM102) & "', TM103= '" & ChgSQL(textTM103) & "', " & _
                                 "TM104= '" & ChgSQL(textTM104) & "', TM105= '" & ChgSQL(textTM105) & "', " & _
                                 "TM106= '" & ChgSQL(TextTM106) & "', TM107= '" & ChgSQL(TextTM107) & "', " & _
                                 "TM108= '" & ChgSQL(TextTM108) & "', TM109= '" & ChgSQL(TextTM109) & "', " & _
                                 "TM110= '" & ChgSQL(TextTM110) & "', TM111= '" & ChgSQL(TextTM111) & "', " & _
                                 "TM112= '" & ChgSQL(TextTM112) & "', TM113= '" & ChgSQL(TextTM113) & "', " & _
                                 "TM114= '" & ChgSQL(TextTM114) & "', TM115= '" & ChgSQL(TextTM115) & "', " & _
                                 "TM116= '" & ChgSQL(TextTM116) & "', TM117= '" & ChgSQL(TextTM117) & "', " & _
                                 "TM09 = '" & textTM09 & "', " & _
                                 "TM32 = '" & textTM32 & "', " & _
                                 "TM67 = '" & ChgSQL(textTM67) & "' " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
    cnnConnection.Execute strSql
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'add by nickc 2006/08/03
   If UpForm Is Nothing Then
       ' ·s¼W¸ê®Æ¨ì®×¥ó¶i«×ÀÉ
       ' ¦¬¤å¸¹
       strCP09 = Empty
       strCP09 = AutoNo("C", 6)
       ' ®×¥ó©Ê½è¬°³qª¾¥Ó½Ð®×¸¹
       strCP10 = "1101"
       ' ·~°È°Ï§O 91.8.26 MODIFY BY SONIA
       'strCP12 = GetStaffDepartment(m_CP13)
       ' µo¤å¤é
       strCP27 = DBDATE(SystemDate())
       ' ·s¼W®×¥ó¶i«×¸ê®Æ
        'Modify By Cheng 2003/09/05
    '   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43) " & _
    '            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(textCP05) & "," & _
    '                    "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
    '                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "') "
       '2012/10/2 MODIFY BY SONIA ·~°È°Ï§ï§ì·s´¼Åv¤H­ûªº·~°È°Ï
       'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43) " & _
                "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(textCP05) & "," & _
                        "'" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
                        "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "') "
       'Modified by Lydia 2018/02/06 PUB_GetFCTSalesNo+®×¥ó©Ê½è1101
       strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43) " & _
                "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(textCP05) & "," & _
                        "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101")) & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101") & "','" & strUserNum & "'," & _
                        "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "') "
       cnnConnection.Execute strSql
   End If
   
   
   'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
   ' ¦³¿é¤J¸É¤å¥ó´Á­­®É, ·s¼W¤@µ§¸É¤å¥óªº°O¿ý¨ì¤U¤@µ{§ÇÀÉ
   If IsEmptyText(textAddDate) = False Then
      strNP07 = "208"
      strNP22 = GetNextProgressNo()
      'Modify By Sindy 2015/7/29 ¥»©Ò´Á­­§ï¬°Àu¥ýÅvÃÒ©ú´Á­­-2¤u§@¤Ñ
      'strNP08 = DBDATE(DateAdd("d", -25, ChangeWStringToWDateString(DBDATE(textAddDate))))
      strNP08 = DBDATE(PUB_GetOurDeadline(textAddDate)) '¥»©Ò´Á­­=Àu¥ýÅvÃÒ©ú´Á­­-2¤u§@¤Ñ
      '2015/7/29 END
      'Modify By Sindy 2015/8/20 ­Y¤U¤@µ{§Ç¦³¥¼Äò¿ì¤§208´Á­­«h§ó·s,¤£¥i·s¼W
      strExc(0) = "SELECT np01 FROM nextprogress" & _
                  " WHERE np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "'" & _
                  " and np06 is null and np07='208'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
          'Modified by Lydia 2018/02/06 PUB_GetFCTSalesNo+®×¥ó©Ê½è1101
         strSql = "update nextprogress" & _
                  " set NP01='" & strCP09 & "',NP08=" & strNP08 & ",NP09=" & DBDATE(textAddDate) & ",NP10='" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101") & "'" & _
                  " WHERE np02='" & m_TM01 & "' and np03='" & m_TM02 & "' and np04='" & m_TM03 & "' and np05='" & m_TM04 & "'" & _
                  " and np06 is null and np07='208'"
      Else
      '2015/8/20 END
         'edit by nickc 2007/02/06 ¤£¥Î dll ¤F
         'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
         'edit by nickc 2005/05/24 À³¸Ó¬O¥Ó½Ð®×ªº¦¬¤å¸¹
   '      StrSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                           strNP08 & "," & DBDATE(textAddDate) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
          'Modified by Lydia 2018/02/06 PUB_GetFCTSalesNo+®×¥ó©Ê½è1101
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                           strNP08 & "," & DBDATE(textAddDate) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101") & "'," & strNP22 & ")"
      End If
      cnnConnection.Execute strSql
   End If
   
   'add by nickc 2006/12/21
   If IsEmptyText(textAddDate2) = False Then
      tmpNP15 = ""
      If InStr(1, textAdd, "1") > 0 Then
        tmpNP15 = tmpNP15 & "1:©e¥ôª¬¡F"
      End If
      If InStr(1, textAdd, "3") > 0 Then
        tmpNP15 = tmpNP15 & "3:¤½¥qµù¥UÃÒ©ú¡F"
      End If
      strNP07 = "201"
      strNP22 = GetNextProgressNo()
      strNP08 = DBDATE(DateAdd("d", -25, ChangeWStringToWDateString(DBDATE(textAddDate2))))
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
      'Modified by Lydia 2018/02/06 PUB_GetFCTSalesNo+®×¥ó©Ê½è1101
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,np15,NP22) " & _
               "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                        strNP08 & "," & DBDATE(textAddDate2) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101") & "','" & tmpNP15 & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
   End If
   
   'add by nickc 2006/08/03
   If UpForm Is Nothing Then
        ' Àx¦sÀu¥ýÅv¸ê®Æ
        'edit by nickc 2007/02/06 ¤£¥Î dll ¤F
        'If objPublicData.SavePriority(m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)) = False Then GoTo CheckingErr
        'Modify by Amy 2014/03/27 +pd08,pd09
        'Modify By Sindy 2017/10/12 + , m_Priority(6)
        If ClsPDSavePriority(m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)) = False Then GoTo CheckingErr
         '­Y¬°°Ó¥Ó®×¥B¦³Àu¥ýÅv¸ê®Æ, «hºÞ¨î"¥D±iÀu¥ýÅv"(108)ªº´Á­­
         'edit by nick 2004/12/23 ¥[¤J¤À³Îµ¥©ó¥Ó½Ð
         'If m_CP10 = "101" And m_Priority(1) <> "" Then
         If (m_CP10 = "101" Or m_CP10 = "308") And m_Priority(1) <> "" Then
             'ªk©w´Á­­
             strCP07 = DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(m_CP27))))
             '¥»©Ò´Á­­
             'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
             If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                strCP06 = PUB_GetOurDeadline(DBDATE(strCP07))
             Else
             '2014/10/6 END
                strCP06 = DBDATE(DateAdd("d", -4, ChangeWStringToWDateString(DBDATE(strCP07))))
             End If
             strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
             StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='108' "
             rsA.CursorLocation = adUseClient
             rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
             '­Y¦³¦¬¤å¥D±iÀu¥ýÅv, §ó·s¶i«×ÀÉ
             If rsA.RecordCount > 0 Then
                 StrSQLa = "Update CaseProgress Set CP06=" & strCP06 & ", CP07=" & strCP07 & " Where CP09='" & rsA("CP09").Value & "' "
                 cnnConnection.Execute StrSQLa
             '­Y¥¼¦¬¤å¥D±iÀu¥ýÅv, ·s¼W¤U¤@µ{§ÇÀÉ
             Else
                 'edit by nickc 2005/12/02 ªü½¬¸ò¨q¬Â»¡ §ï¦¨ ¨S¦¬¤£¥i¦sÀÉ¡A­n¦¬¤F¤§«á¤~¥i¥H¦s
                 'strNP07 = "108"
                 'strNP22 = GetNextProgressNo()
                 'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                 '                "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                 '                DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
                 'cnnConnection.Execute strSQL
                 'add by nickc 2006/08/03
                 If Me.Visible = True Then
                     cnnConnection.RollbackTrans
                 End If
                 OnSaveData = False
                 MsgBox "¦³Àu¥ýÅv¸ê®Æ¡A¦ý¥¼¦¬¤å¥D±iÀu¥ýÅv", , "¤£¥i¦sÀÉ¡I"
                 Exit Function
             End If
             If rsA.State <> adStateClosed Then rsA.Close
             Set rsA = Nothing
         End If
    End If
    'add by nickc 2006/07/24
    If m_CP10 = "308" Then
      '·s¼W¤l®×®Ö­ã¨Ó¤å
      strCP09 = AutoNo("C", 6)
      strCP05 = DBDATE(UpForm.oStrCDate)
      strCP27 = DBDATE(SystemDate())
      ' ²Õ¦¨SQL»yªk
      '2012/10/2 MODIFY BY SONIA ´¼Åv¤H­û§ì·s´¼Åv¤H­û, ·~°È°Ï§ï§ì·s´¼Åv¤H­ûªº·~°È°Ï
      'strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14,  CP26, cp27,  CP43) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "1001" & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "','" & "N" & "'," & strCP27 & ",'" & m_CP09 & "')"
      'Modify By Sindy 2015/8/4 ·s¼W¤l®×®Ö­ã¨Ó¤å ¤Î ·s¼W¤l®×¥Ó½Ð,·s¼W¦¹¤G¶i«×®É, ½Ð¥[¤J CP20='N'
      'Modified by Lydia 2018/02/06 PUB_GetFCTSalesNo+®×¥ó©Ê½è1101
      strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14,  CP26, cp27,  CP43,CP20) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "1001" & "','" & GetSalesArea(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101")) & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101") & "','" & strUserNum & "','" & "N" & "'," & strCP27 & ",'" & m_CP09 & "','N')"
      ' ·s¼W¸ê®Æ¨ì¸ê®Æ®w
      cnnConnection.Execute strSql
      
      'Added by Morgan 2017/6/14 ¹q¤l¤½¤å
      If m_DocNo <> "" Then
         '§ó·s¾÷Ãö¤å¸¹
         strSql = "update caseprogress set cp08='" & m_DocWord & "¦r²Ä" & PUB_GetEDocNo(m_DocNo) & "¸¹' where cp09='" & strCP09 & "'"
         cnnConnection.Execute strSql, intI
         '½Æ»s¥À®×¤½¤å¹q¤lÀÉ
         strExc(0) = PUB_GetEDocFileName(m_TM01, m_TM02, m_TM03, m_TM04, "1001")
         SaveAttFile_PDF strCP09, m_DocPdf, strExc(0), Format(m_DocPdfDate), Format(m_DocPdfTime), False, , , True
      End If
      'end 2017/6/14
      
      '·s¼W¤l®×¥Ó½Ð¡A¦Û°Êµo¤å
      strCP09 = AutoNo("B", 6)
      strCP05 = DBDATE("111111")
      strCP27 = DBDATE("111111")  '2010/11/3 MODIFY BY SONIA ­ì©ñ¨t²Î¤é
      ' ²Õ¦¨SQL»yªk
      '2012/10/2 MODIFY BY SONIA ´¼Åv¤H­û§ì·s´¼Åv¤H­û, ·~°È°Ï§ï§ì·s´¼Åv¤H­ûªº·~°È°Ï
      'strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14,  CP26,cp27,   CP43) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "101" & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "','" & "N" & "'," & strCP27 & ",'" & m_CP09 & "')"
      'Modify By Sindy 2015/8/4 ·s¼W¤l®×®Ö­ã¨Ó¤å ¤Î ·s¼W¤l®×¥Ó½Ð,·s¼W¦¹¤G¶i«×®É, ½Ð¥[¤J CP20='N'
      'Modified by Lydia 2018/02/06 PUB_GetFCTSalesNo+®×¥ó©Ê½è1101
      'modify by sonia 2018/12/13 +CP110, FCT-043065¤l®×®Ö­ã¤é¤å©w½ZÄ¶¤å§ì¤£¨ì¥X¦W¥N²z¤H
      'strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14, CP26, cp27, CP43, CP20) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "101" & "','" & GetSalesArea(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101")) & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101") & "','" & strUserNum & "','" & "N" & "'," & strCP27 & ",'" & m_CP09 & "','N')"
      m_CP110 = ""
      strExc(0) = "SELECT CP110 FROM CASEPROGRESS" & _
                  " WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "'" & _
                  " AND CP10='308'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_CP110 = "" & RsTemp.Fields("CP110")
      End If
      strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14, CP26, cp27, CP43, CP20, CP110) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & ",'" & strCP09 & "','" & "101" & "','" & GetSalesArea(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101")) & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101") & "','" & strUserNum & "','" & "N" & "'," & strCP27 & ",'" & m_CP09 & "','N','" & m_CP110 & "')"
      ' ·s¼W¸ê®Æ¨ì¸ê®Æ®w
      cnnConnection.Execute strSql
      'add by nickc 2006/11/09 ±N¤l«ö¤U¤@µ{§Ç¤À³Î¶Ê¼f¡A§ï¦¨¥Ó½Ð¶Ê¼f
'      strSql = "update nextprogress set np01='" & strCP09 & "' where np01='" & m_CP09 & "' and np07=305 "
'      cnnConnection.Execute strSql
      'Add By Sindy 2010/10/25 §ï¬°§ó·s¤l®×¤À³Î¶Ê¼f¤§NP06=Y
      strSql = "update nextprogress set np06='Y' where np01='" & m_CP09 & "' and np07=305 "
      cnnConnection.Execute strSql
      'Add By Sindy 2010/10/25 ¥t·s¼WBÃþ¥Ó½Ð¤§¶Ê¼f´Á­­
      strNP08 = GetUrgeDate(m_TM01, m_TM10, "101", DBDATE(UpForm.oStrCDate))
      'modify by sonia 2017/9/6¤l®×¥Ó½Ð¶Ê¼f§ï±¾´¼Åv¤H­û m_CP14->PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      'Modified by Lydia 2018/02/06 PUB_GetFCTSalesNo+®×¥ó©Ê½è1101
      'Modified by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
      'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',305," & _
                          strNP08 & "," & strNP08 & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101") & "'," & GetNextProgressNo & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',305," & _
                          PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04, "1101") & "'," & GetNextProgressNo & ")"
      cnnConnection.Execute strSql
      '§ó·s¤l®×®Ö­ã¤Îµ²ªG¤é
      strCP05 = DBDATE(UpForm.oStrCDate)
      strSql = "update caseprogress set cp24='1',cp25=" & strCP05 & " where cp09='" & m_CP09 & "' "
      cnnConnection.Execute strSql
         '¦³´Á­­®É
         If textNP08.Enabled = True And textNP09.Enabled = True Then
                '­Yµe­±¦³¿é¤J·s´Á­­¥H·s´Á­­¬°¥D¡A¨S¦³ªº¸Ü±NÄ~©Ó¥À®×´Á­­
                If Trim(textNP08) <> "" And Trim(textNP09) <> "" Then
                   If UpForm.IsHaveNp202 Then
                         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                             "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',202," & _
                             DBDATE(textNP08) & "," & DBDATE(textNP09) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo & ")"
                         cnnConnection.Execute strSql
                   ElseIf UpForm.IsHaveCp202 Then
'                         strCP09 = AutoNo("B", 6)
'                         strCP05 = DBDATE("111111")
'                         strCP27 = "null"
'                         strSQL = "insert into caseprogress select "
'                         For ii = 1 To TF_CP
'                             Select Case ii
'                             Case 1
'                                 strSQL = strSQL & "'" & m_TM01 & "',"
'                             Case 2
'                                strSQL = strSQL & "'" & m_TM02 & "',"
'                             Case 3
'                                strSQL = strSQL & "'" & m_TM03 & "',"
'                             Case 4
'                                strSQL = strSQL & "'" & m_TM04 & "',"
'                             Case 9
'                                 strSQL = strSQL & "'" & strCP09 & "',"
'                             Case 27
'                                 strSQL = strSQL & strCP27 & ","
'                             Case 5
'                                 strSQL = strSQL & strCP05 & ","
'                             Case 6
'                                 strSQL = strSQL & DBDATE(textNP08) & ","
'                             Case 7
'                                 strSQL = strSQL & DBDATE(textNP09) & ","
'                             Case Else
'                                 If ii < 100 Then
'                                     strSQL = strSQL & "CP" & Format(ii, "00") & ","
'                                 Else
'                                     strSQL = strSQL & "CP" & Format(ii, "000") & ","
'                                 End If
'                             End Select
'                         Next ii
'                         strSQL = Left(strSQL, Len(strSQL) - 1) & " from caseprogress where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202' and cp27 is null  "
'                         cnnConnection.Execute strSQL
                        If Trim(textNP08) <> "" Then
                            strSql = "update caseprogress set cp06=" & DBDATE(textNP08) & ",cp07=" & DBDATE(textNP09) & ",cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        Else
                            strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        End If
                        cnnConnection.Execute strSql
                        '2010/11/17 modify by sonia cp43§ï±¾¤À³Î®×¤§BÃþ¥Ó½Ð
                        'strSql = "update caseprogress set cp43='" & m_CP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        strSql = "update caseprogress set cp43='" & strCP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        cnnConnection.Execute strSql
                        strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        cnnConnection.Execute strSql
                   End If
                Else
                   If UpForm.IsHaveNp202 Then
                         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                             "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',202," & _
                             m_MonNP08 & "," & m_MonNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo & ")"
                         cnnConnection.Execute strSql
                   ElseIf UpForm.IsHaveCp202 Then
'                         strCP09 = AutoNo("B", 6)
'                         strCP05 = DBDATE("111111")
'                         strCP27 = "null"
'                         strSQL = "insert into caseprogress select "
'                         For ii = 1 To TF_CP
'                             Select Case ii
'                             Case 1
'                                 strSQL = strSQL & "'" & m_TM01 & "',"
'                             Case 2
'                                strSQL = strSQL & "'" & m_TM02 & "',"
'                             Case 3
'                                strSQL = strSQL & "'" & m_TM03 & "',"
'                             Case 4
'                                strSQL = strSQL & "'" & m_TM04 & "',"
'                             Case 9
'                                 strSQL = strSQL & "'" & strCP09 & "',"
'                             Case 27
'                                 strSQL = strSQL & strCP27 & ","
'                             Case 5
'                                 strSQL = strSQL & strCP05 & ","
'                             Case Else
'                                 If ii < 100 Then
'                                     strSQL = strSQL & "CP" & Format(ii, "00") & ","
'                                 Else
'                                     strSQL = strSQL & "CP" & Format(ii, "000") & ","
'                                 End If
'                             End Select
'                         Next ii
'                         strSQL = Left(strSQL, Len(strSQL) - 1) & " from caseprogress where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202' and cp27 is null  "
'                         cnnConnection.Execute strSQL
                        If Trim(textNP08) <> "" Then
                            strSql = "update caseprogress set cp06=" & DBDATE(textNP08) & ",cp07=" & DBDATE(textNP09) & ",cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        Else
                            strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & ";­ì¬ÛÃö¦¬¤å¸¹¡G'||cp43||';' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        End If
                        cnnConnection.Execute strSql
                        '2010/11/17  modify by sonia cp43§ï±¾¤À³Î®×¤§BÃþ¥Ó½Ð
                        'strSql = "update caseprogress set cp43='" & m_CP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        strSql = "update caseprogress set cp43='" & strCP09 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        cnnConnection.Execute strSql
                        strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27 is null and cp57 is null and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202'  "
                        cnnConnection.Execute strSql

                   End If
                End If
                If UpForm.IsHaveNp202 Then
                     strSql = "update nextprogress set np06='N',np15=np15||'Âà¤J¤l®×¡A¤l®×®×¸¹¡G" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' where np02='" & m_MonTM01 & "' and np03='" & m_MonTM02 & "' and np04='" & m_MonTM03 & "' and np05='" & m_MonTM04 & "' and np06 is null and np07=202 "
                     cnnConnection.Execute strSql
                ElseIf UpForm.IsHaveCp202 Then
                     strSql = "update caseprogress set cp57=to_number(to_char(sysdate,'YYYYMMDD')),cp64=cp64||'Âà¤J¤l®×¡A¤l®×®×¸¹¡G" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' where cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "' and cp10='202' and cp27 is null "
                     cnnConnection.Execute strSql
                End If
                '¥À®×¤À³Îµo¤å«áªº¦¬¤å¤Îµo¤å®×¥ó¬ÒÂà¤J¦³´Á­­ªº¤l®×
                Dim m_MonCP27 As String
                strSql = "select cp27 from caseprogress where cp09='" & m_MonCP09 & "' "
                m_MonCP27 = ""
                Set rsTmp = New ADODB.Recordset
                If rsTmp.State = 1 Then rsTmp.Close
                rsTmp.CursorLocation = adUseClient
                rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 ­ì¥ý¬O  °ÊºA¶}±Ò
                If rsTmp.RecordCount > 0 Then
                    m_MonCP27 = CheckStr(rsTmp.Fields("cp27"))
                End If
                If m_MonCP27 <> "" Then
                    strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & "' where cp05>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                    cnnConnection.Execute strSql
                    strSql = "update caseprogress set cp64=cp64||'±q¥À®×Âà¤J¡A®×¸¹¡G" & m_MonTM01 & "-" & m_MonTM02 & "-" & m_MonTM03 & "-" & m_MonTM04 & "' where cp27>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                    cnnConnection.Execute strSql
                    strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp05>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                    cnnConnection.Execute strSql
                    strSql = "update caseprogress set cp01='" & m_TM01 & "',cp02='" & m_TM02 & "',cp03='" & m_TM03 & "',cp04='" & m_TM04 & "' where cp27>" & m_MonCP27 & " and cp01='" & m_MonTM01 & "' and cp02='" & m_MonTM02 & "' and cp03='" & m_MonTM03 & "' and cp04='" & m_MonTM04 & "'  and cp10<>'1001' "
                    cnnConnection.Execute strSql
                End If
         End If
    End If
   'add by nickc 2006/08/03
   If Me.Visible = True Then
      '911107 nick transation
      cnnConnection.CommitTrans
   End If
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'add by nickc 2006/08/03
   If UpForm Is Nothing Or Me.Visible = False Then
       ' ¦C¦L©w½Z
       If textPrint <> "N" Then
           PrintLetter
           'Add By Cheng 2003/02/17
           '·s¼W¦a§}±ø¦Cªí¸ê®Æ
           'edit by nick 2004/09/14 ©w½Z¤£¦L¦a§}±ø¤F
           'pub_AddressListSN = pub_AddressListSN + 1
           'PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
       End If
   End If

   '911107 nick transation
   Exit Function
     
CheckingErr:
   'add by nickc 2006/08/03
   If Me.Visible = True Then
      MsgBox (Err.Description)
      cnnConnection.RollbackTrans
   End If
   'edit by nick 2004/11/03
   OnSaveData = False
End Function

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nCount As Integer
Dim nIndex As Integer
Dim strTemp As String
   
   CheckDataValid = False
   
   ' ¥Ó½Ð¤é¤£¥iªÅ¥Õ
   'edit by nickc 2006/09/29
   'If IsEmptyText(textTM11) = True Then
   If IsEmptyText(textTM11) = True And textTM11.Enabled = True Then
      strTit = "¸ê®ÆÀË®Ö"
      strMsg = "½Ð¿é¤J¥Ó½Ð¤é"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM11.SetFocus
      GoTo EXITSUB
   End If
   
   ' ¥Ó½Ð®×¸¹¤£¥iªÅ¥Õ
   If IsEmptyText(textTM12) = True Then
      strTit = "¸ê®ÆÀË®Ö"
      strMsg = "½Ð¿é¤J¥Ó½Ð®×¸¹"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM12.SetFocus
      GoTo EXITSUB
   End If
   
   'Add By Sindy 2011/6/3
   ' ¥Ó½Ð°ê®a¬°¥xÆW®É¡AÀË¬d¥Ó½Ð®×¸¹ªº«e¤G(¤T)½X¥²¶·¬°¥Ó½Ð¦~«×
   '2011/9/22 modify by sonia ¤À³Î®×¤£ÀË¬d
   'If m_TM10 = "000" Then
   If m_TM10 = "000" And m_CP10 <> "308" Then
      If IsEmptyText(textTM11) = False And IsEmptyText(textTM12) = False Then
         If Val(Left(textTM12, 1)) > "1" Then
            strExc(1) = Val(Left(textTM12, 2))
         Else
            strExc(1) = Val(Left(textTM12, 3))
         End If
         strExc(2) = Trim(Val(textTM11) \ 10000)
         If strExc(1) <> strExc(2) Then
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥Ó½Ð®×¸¹ªº«e¤G(¤T)½X¥²¶·¬°¥Ó½Ð¦~«×!"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM12.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   ' ¨Ó¨ç¦¬¤å¤é¤£¥iªÅ¥Õ
   If IsEmptyText(textCP05) = True Then
      strTit = "¸ê®ÆÀË®Ö"
      strMsg = "½Ð¿é¤J¨Ó¨ç¦¬¤å¤é"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP05.SetFocus
      GoTo EXITSUB
   End If
   
   'add by nick 2004/09/09 ±q frm030202_03 ²¾¨Ó
   'edit by nick 2004/12/23 ¥[¤J¤À³Îµ¥©ó¥Ó½Ð
   'If m_CP10 = "101" Then
   '2007/10/18 modify by sonia ÃÒ©ú¼Ð³¹,¹ÎÅé¼Ð³¹¤£¥²ÀË¬d
   'If m_CP10 = "101" Or m_CP10 = "308" Then
   If (m_CP10 = "101" Or m_CP10 = "308") And m_TM08 < "7" Then
      If IsEmptyText(textTM32) = True Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "½Ð¿é¤J°Ó«~²Õ¸s"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM32.SetFocus
         GoTo EXITSUB
      End If
   Else
      'Modify By Sindy 2024/4/18 °Ó«~²Õ¸sÄæ¤H­û¶K¤W¸ê®Æ«á±N¥þ§Î©Î¥b§Îªº¡u¡F¡v¤À¸¹¡AÂà¬°¥b§Îªº³r¸¹¦s¤JTM32¡C
      textTM32 = Replace(Replace(textTM32, ";", ","), "¡F", ",")
      nCount = GetSubStringCount(textTM32)
      For nIndex = 1 To nCount
         strTemp = GetSubString(textTM32, nIndex)
         If Len(strTemp) > 6 Then
            strTit = "ÀË®Ö¸ê®Æ"
            strMsg = "°Ó«~²Õ¸s<" & strTemp & ">¤£¥¿½T"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM32.SetFocus
            GoTo EXITSUB
         End If
      Next nIndex
      For nIndex = 1 To nCount
         strTemp = GetSubString(textTM32, nIndex)
         For nCount = 1 To nCount
            If nIndex <> nCount Then
               If strTemp = GetSubString(textTM32, nCount) Then
                  strTit = "ÀË®Ö¸ê®Æ"
                  strMsg = "°Ó«~²Õ¸s<" & strTemp & ">¤£¥i­«ÂÐ"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM32.SetFocus
                  GoTo EXITSUB
               End If
            End If
         Next nCount
      Next nIndex
      '2024/4/18 END
   End If
   
   ' °Ó¼ÐºØÃþ¬°Áp¦X°Ó¼Ð, ¨¾Å@°Ó¼Ð, Áp¦XªA°È¼Ð³¹, ¨¾Å@ªA°È¼Ð³¹®É, ¥¿°Ó¼Ð¸¹¼Æ¤£¥i¬°ªÅ¥Õ
   If IsEmptyText(textTM27) = True Then
      Select Case m_TM08
         Case "2", "3", "5", "6":
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "°Ó¼ÐºØÃþ¬°Áp¦X°Ó¼Ð, ¨¾Å@°Ó¼Ð, Áp¦XªA°È¼Ð³¹, ¨¾Å@ªA°È¼Ð³¹®É, ¥¿°Ó¼Ð¸¹¼Æ¤£¥i¬°ªÅ¥Õ"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM27.SetFocus
            GoTo EXITSUB
         Case Else
      End Select
   End If
    'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
    If m_Priority(1) = "" Then
         If Me.textPriorityDoc.Text <> "" Then
            MsgBox "µLÀu¥ýÅv¸ê®Æ, ¤£¥i¿é¤J¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó !!!", vbExclamation + vbOKOnly
            Me.textPriorityDoc.SetFocus
            GoTo EXITSUB
         End If
      Else
         If Me.textPriorityDoc.Text = "" Then
            MsgBox "¦³Àu¥ýÅv¸ê®Æ, ½Ð¿é¤J¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó !!!", vbExclamation + vbOKOnly
            Me.textPriorityDoc.SetFocus
            GoTo EXITSUB
         End If
      End If
   'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
   'ÀË¬d­Y®×¥ó©Ê½è¬°"¥Ó½Ð"(101)®É, ·í¬O§_¥IÀu¥ýÅvÃÒ©ú¤å¥ó¿é¤J"Y"®É, ¤@©w­n¿é¤JÀu¥ýÅv¸ê®Æ
   'edit by nick 2004/12/23 ¥[¤J¤À³Îµ¥©ó¥Ó½Ð
   'If m_CP10 = "101" And Me.textPriorityDoc.Text = "Y" Then
   If (m_CP10 = "101" Or m_CP10 = "308") And Me.textPriorityDoc.Text = "Y" Then
      If frm880002.m_blnAddNew = False Then
         MsgBox "½Ð¿é¤JÀu¥ýÅv¸ê®Æ!!!", vbExclamation + vbOKOnly
         Me.cmdPriority.SetFocus
         GoTo EXITSUB
      End If
   End If
    'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
    '­Y¬O§_¸É¥ó¦³¿ï¾Ü2¥B¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó¥¼¿é"N"
    If InStr(Me.textAdd.Text, "2") > 0 And Me.textPriorityDoc.Text <> "N" Then
        MsgBox "­Y¬O§_¸É¥óÄæ¦ì¦³¿ï¾Ü2, «h¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó¥²¶·¬°N!!!", vbExclamation + vbOKOnly
        Me.textPriorityDoc.Text = "N"
        Me.textAdd.SetFocus
        textAdd_GotFocus
        GoTo EXITSUB
    End If
    'add by nickc 2007/08/10 ªü½¬»¡À³¸Ó­n¦b¤l®×¿é
    If txtLaw.Visible = True Then
      If txtLaw.Text = "" Then
        strMsg = "©w½Z¤Wªºªk±ø¤£¥iªÅ¥Õ"
        strTit = "¸ê®ÆÀË®Ö"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        txtLaw.SetFocus
        GoTo EXITSUB
      End If
    End If
    
   'Add By Sindy 2012/5/18
   If LabNP07.Caption <> "" Then
      'ÀË¬d¨Ó¨ç´Á­­--¤é´Á
      If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
         '2015/3/5 ADD BY SONIA FCT®×¤£¹w³]¨Ó¨ç¤Ñ¼Æ©Î¤ë¼Æ,¦]¬°¥i¯à30¤Ñ¥i¯à1­Ó¤ë,¦ý¥[ÀË¬d¤@©w­n¿é,FCT-036284
         If Me.Option4(0).Value = False And Me.Option4(1).Value = False And Me.Option4(2).Value = False Then
            MsgBox "½Ð¿ï¾Ü¨Ó¨ç´Á­­¤Ñ¼Æ,¤ë¼Æ©Î¤é´Á!!!", vbExclamation + vbOKOnly
            Me.Text10.SetFocus
            GoTo EXITSUB
         End If
         If Me.Option4(0).Value = True Then
            If Me.Text10.Text = "" Then
               MsgBox "½Ð¿é¤J¨Ó¨ç´Á­­¤Ñ¼Æ!!!", vbExclamation + vbOKOnly
               Me.Text10.SetFocus
               GoTo EXITSUB
            End If
         End If
         If Me.Option4(1).Value = True Then
            If Me.Text11.Text = "" Then
               MsgBox "½Ð¿é¤J¨Ó¨ç´Á­­¤ë¼Æ!!!", vbExclamation + vbOKOnly
               Me.Text11.SetFocus
               GoTo EXITSUB
            End If
         End If
         '2015/3/5 END
         If Me.Option4(2).Value = True Then
            If Me.Text12.Text = "" Then
               MsgBox "½Ð¿é¤J¨Ó¨ç´Á­­!!!", vbExclamation + vbOKOnly
               Me.Text12.SetFocus
               GoTo EXITSUB
            End If
         End If
      End If
   End If
    
    '2010/6/25 add by sonia ªü½¬»¡­n±±¨î
    If txtLaw.Visible = True Then
      If IsEmptyText(textNP08) = True Then
        strMsg = "¤l®×·s¥»©Ò´Á­­¤£¥iªÅ¥Õ"
        strTit = "¸ê®ÆÀË®Ö"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        textNP08.SetFocus
        GoTo EXITSUB
      End If
      If IsEmptyText(textNP09) = True Then
        strMsg = "¤l®×·sªk©w´Á­­¤£¥iªÅ¥Õ"
        strTit = "¸ê®ÆÀË®Ö"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        textNP09.SetFocus
        GoTo EXITSUB
      End If
    End If
    '2010/6/25 end
   
    'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
    'edit by nickc 2005/05/24
    'If Me.textPriorityDoc.Text <> "N" And InStr(Me.textAdd.Text, "2") <> 0 Then
    If Me.textPriorityDoc.Text = "N" And InStr(Me.textAdd.Text, "2") = 0 Then
        MsgBox "­Y¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó¬° N ®É, «h¬O§_¸É¥óÄæ¦ì¥²¶·¦³¿ï¾Ü2!!!", vbExclamation + vbOKOnly
        Me.textAdd.SetFocus
        textAdd_GotFocus
        GoTo EXITSUB
    End If
    'add by nickc 2006/03/17 ¥[¤JÅçÃÒ
    Dim Cancel As Boolean
    Cancel = False
    textTM11_Validate Cancel
    If Cancel = True Then GoTo EXITSUB
    textCP05_Validate Cancel
    If Cancel = True Then GoTo EXITSUB
   
   'Add By Sindy 2012/7/9 ¥H¨¾­×§ï´Á­­¤Ñ¼Æ©Î¤ë¼Æ,­«·s­pºâ´Á­­
   If Me.Text10.Enabled = True Then
      Cancel = False
      Text10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.Text11.Enabled = True Then
      Cancel = False
      Text11_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2012/7/9 End
   
   CheckDataValid = True
EXITSUB:
End Function

'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
Private Sub cmdPriority_Click()
   ' ­×§ïÀu¥ýÅv¸ê®Æ
   'Modify by Amy 2014/03/27 +pd08,pd09
   'Modify By Sindy 2017/10/12 + , m_Priority(6)
   'Modify by Sindy 2019/1/23 + m_TM01 & m_TM02 & m_TM03 & m_TM04
   ModifyPriority m_Priority(1), m_Priority(2), m_Priority(3), , , m_TM01 & m_TM02 & m_TM03 & m_TM04, , , m_Priority(4), m_Priority(5), m_Priority(6)
End Sub

'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
Private Sub textPriorityDoc_GotFocus()
   InverseTextBox textPriorityDoc
End Sub

'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
Private Sub textAddDate_GotFocus()
   InverseTextBox textAddDate
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030203_02 = Nothing
End Sub

Private Sub textAdd_KeyPress(KeyAscii As Integer)
    'Modify By Cheng 2003/12/09
'    '¥Ó½Ð¤é¤p©ó921128
'    If Val(Me.textTM11.Text) < 921128 Then
'        If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 Then
'            KeyAscii = 0
'        End If
'    '¥Ó½Ð¤é¤j©óµ¥©ó921128
'    Else
        'Modify By Sindy 2011/2/17
        'If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 Then
        If KeyAscii <> 8 And KeyAscii <> 44 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 Then
            KeyAscii = 0
        End If
'    End If
End Sub

' ¬O§_¸É¥ó
Private Sub textAdd_Validate(Cancel As Boolean)
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Cancel = False
   
   ' µL¸ê®Æ®É¤£°µ¥ô¦óÀË¬d
   If IsEmptyText(textAdd) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textAdd)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textAdd, nIndex)
      Select Case strTemp
         Case "1", "2", "3", "4":
         Case Else
            Cancel = True
            strTit = "ÀË®Ö¸ê®Æ"
            strMsg = "¬O§_¸É¥ó¶µ¥Ø<" & strTemp & ">¤£¥¿½T"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textAdd_GotFocus
            GoTo EXITSUB
      End Select
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textAdd, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textAdd, nCount) Then
               Cancel = True
               strTit = "ÀË®Ö¸ê®Æ"
               strMsg = "¬O§_¸É¥ó¶µ¥Ø<" & strTemp & ">¤£¥i­«ÂÐ"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textAdd_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   'add by nickc 2006/12/21
   If InStr(1, textAdd, "1") > 0 Or InStr(1, textAdd, "3") > 0 Then
      EnableTextBox textAddDate2, True
   End If
EXITSUB:
End Sub

' ¨Ó¨ç¦¬¤å¤é
Private Sub textCP05_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If CheckIsTaiwanDate(textCP05, False) = False Then
         Cancel = True
         strMsg = "½Ð¿é¤J¥¿½Tªº¨Ó¨ç¦¬¤å¤é"
         strTit = "¨Ó¨ç¦¬¤å¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
         GoTo EXITSUB
      End If
      If Val(DBDATE(textCP05)) > Val(DBDATE(SystemDate())) Then
         Cancel = True
         strMsg = "¨Ó¨ç¦¬¤å¤é¤£¥i¶W¹L¨t²Î¤é"
         strTit = "¨Ó¨ç¦¬¤å¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP05_GotFocus
         GoTo EXITSUB
      End If
'edit by nick 2004/08/31 ²¾°£
'      ' ÀË¬d¨Ó¨ç°O¿ýÀÉ¬O§_¦³¸Óµ§µL´Á­­ªº°O¿ý
'      If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, textCP05) = False Then
'         strTit = "ÀË®Ö¸ê®Æ"
'         strMsg = "»PÂd¥x¤§¨Ó¨ç°O¿ý¤£²Å, ½Ð½T»{?"
'         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
'         If nResponse <> vbYes Then
'            Cancel = True
'            textCP05_GotFocus
'            GoTo EXITSUB
'         End If
'      End If
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

' ¬O§_¦C¦L©w½Z
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case "", " ", "N":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎN"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            'textPrint_GotFocus
      End Select
   End If
End Sub

' ¬O§_¦C¦LÂ½Ä¶¨ç
Private Sub textPrtTrans_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrtTrans) = False Then
      Select Case textPrtTrans
         Case "", " ", "N":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎN"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrtTrans_GotFocus
      End Select
   End If
End Sub

Private Sub textTM09_Validate(Cancel As Boolean)
'add by nickc 2005/06/03
textTM09 = Replace(textTM09, " ", "")
End Sub

Private Sub textTM11_Change()
    'Modify By Cheng 2003/12/09
    '¨Ï¥Î·sªº³W©w
'    If Val(Me.textTM11.Text) < 921128 Then
'        Me.Label4.Caption = "(1:©e¥ôª¬ 2:¨Ï¥Î«Å»} 3:Àu¥ýÅvÃÒ©ú 4:­»´ä¤½¥qµù¥UÃÒ©ú)"
'    Else
'        Me.Label4.Caption = "(1:©e¥ôª¬ 2:Àu¥ýÅvÃÒ©ú 3:¤½¥qµù¥UÃÒ©ú)"
'    End If
End Sub

' ¥Ó½Ð¤é
Private Sub textTM11_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM11) = False Then
      If CheckIsTaiwanDate(textTM11, False) = False Then
         Cancel = True
         strMsg = "½Ð¿é¤J¥¿½Tªº¥Ó½Ð¤é"
         strTit = "ÀË®Ö¸ê®Æ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM11_GotFocus
         GoTo EXITSUB
      End If
      'If Val(textTM11) > Val(TAIWANDATE(Date)) Then
      If Val(textTM11) > Val(strSrvDate(2)) Then
         Cancel = True
         strMsg = "¥Ó½Ð¤é¤£¥i¶W¹L¨t²Î¤é"
         strTit = "ÀË®Ö¸ê®Æ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM11_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub

' ¦C¦L©w½Z«e±N¨Ò¥~Äæ¦ì¥[¤J¨ì¦C¦L©w½Z¨Ò¥~Äæ¦ìÀÉ®×¤¤
Private Sub InsExpField()
Dim strSql As String
Dim strTemp As String
Dim nCount As Integer
Dim nIndex As Integer
'Add By Cheng 2003/01/23
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Add By Cheng 2003/01/24
Dim strTemp1 As String
Dim strDate As String   '2011/5/16 ADD BY SONIA

   'Add  By Cheng 2003/01/23
   '§PÂ_¬O§_¦³Àu¥ýÅv¸ê®Æ
   StrSQLa = "Select Count(*) From PriDate Where PD01='" & m_TM01 & "' And PD02='" & m_TM02 & "' And PD03='" & m_TM03 & "' And PD04='" & m_TM04 & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.Fields(0).Value > 0 Then
       m_blnPriDate = True
   Else
       m_blnPriDate = False
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   
   ' ¬O§_¸É¥ó
   strTemp = Empty
   ' ¨Ì®×¥ó©Ê½è¤£¦P
   Select Case m_CP10
     'add by nick 2004/12/23
     Case "308":
        Select Case m_strLanguage
        Case "1"
        Case "2"
        Case "3"
        Case Else
        End Select
      ' ¥Ó½Ð
      Case "101":
         nCount = GetSubStringCount(textAdd)
         For nIndex = 1 To nCount
            'Modify By Cheng 2003/01/24
'            strTemp = GetSubString(textAdd, nIndex)
            strTemp1 = GetSubString(textAdd, nIndex)
            'Modify By Cheng 2003/12/09
            '¨Ï¥Î·sªº³W©w
'            '­Y¥Ó½Ð¤é¤p©ó20031128
'            If DBDATE(textTM11) < 20031128 Then
'                Select Case strTemp1
'                   Case "1":
'                      If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
'                        'Modify By Cheng 2003/02/20
'    '                  strTemp = strTemp & "* Power of Attorney."
'                      strTemp = strTemp & "    * Power of Attorney."
'                   Case "2":
'                      If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
'                        'Modify By Cheng 2003/02/20
'    '                  strTemp = strTemp & "* Intend to Use Declaration."
'                      strTemp = strTemp & "    * Intent-to-Use Declaration."
'                   Case "3":
'                      If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
'                        'Modify By Cheng 2003/02/20
'    '                  strTemp = strTemp & "* A Certified copy of the home application (priority document)."
'                      strTemp = strTemp & "    * A certified copy of the home application (priority document) must " & vbCrLf & _
'                                                        "      be submitted to the IPO within three months from the day of filing " & vbCrLf & _
'                                                        "      the subject application."
'                   Case "4":
'                      If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
'                        'Modify By Cheng 2003/02/20
'    '                  strTemp = strTemp & "* Disclaim the excusive right of use of " & m_TM67 & "."
'                        'Modify By Cheng 2003/03/10
'    '                  strTemp = strTemp & "    * Disclaim the excusive right of use of " & m_TM67 & "."
'                      strTemp = strTemp & "    * A certified copy of Certificate of Incorporation from Register of Companies " & vbCrLf & _
'                                                        "      in Hong Kong."
'                End Select
'            '­Y¥Ó½Ð¤é¤j©óµ¥©ó20031128
'            Else
                Select Case strTemp1
                Case "1":
                    Select Case m_strLanguage
                    Case "2"
                        If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                        strTemp = strTemp & "    * Power of Attorney."
                    Case "3"
                        'edit by nick 2004/09/09 ¦]¬°¤µ¤Ñªü½¬¸É¤é¤å©e¥ôª¬ªº©w½Z
                        'strTemp = strTemp & "©e¥ôûì"
                        'edit by nick 2004/10/06 May ´£¥X¥¿¦¡ªº¼Ë¥»
                        'strTemp = strTemp & "¡@¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¤FþèþîÇrþòÇhÇR¡B²K¥IÇU©e¥ôûìÇR¥NªíªÌ¦LÇyþç±Ì¦L³»þàþçªð°eþâþóþèÆê¡C¤S¡B©e¥ôûìÇU°l¸É¶O¥ÎÇRþ÷þàÇeþêþùÇV¡BµL®ÆÇOþèþîþùÆêþòþóþàÇeþì¡C"
                        If InStr(1, textAdd, "2") = 0 Then
                            '2011/5/12 MODIFY BY SONIA ¸­©ö¶³»¡§ï¤º®e
                            'strTemp = "¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¤FþèþîÇrþòÇhÇR¡B©e¥ôûìÇyþç°e¥IþâþóþèÇrÇoÆìþÝÄ@Æê­PþêÇeþì¡C¤S¡B©e¥ôûìÇU°l¸É¶O¥ÎÇRþ÷þàÇeþêþùÇV¡BµL®ÆÇOþèþîþùÆêþòþóþàÇeþì¡C"
                            strDate = CompDate(1, "3", ChangeTStringToWString(Trim(textTM11)))
                            'MODIFY BY Sindy 2012/6/8 ¸­©ö¶³»¡§ï¤º®e
                            'strTemp = "¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¤FþèþîÇrþòÇhÇR¡B©e¥ôûìÇy" & Mid(strDate, 1, 4) & "¦~" & Mid(strDate, 5, 2) & "¤ë" & Mid(strDate, 7, 2) & "¤éÇeþúÇRþç°e¥IþâþóþèÇrÇoÆìþÝÄ@Æê­PþêÇeþì¡C¤S¡B©e¥ôûìÇU°l¸É¶O¥ÎÇRþ÷þàÇeþêþùÇV¡BµL®ÆÇOþèþîþùÆêþòþóþàÇeþì¡C"
                            'Modify By Sindy 2016/12/6 ²{¦b­n¦¬¶O¤F
                            'strTemp = "¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¤FþèþîÇrþòÇhÇR¡B©e¥ôûìÇy" & Mid(strDate, 1, 4) & "¦~" & Mid(strDate, 5, 2) & "¤ë" & Mid(strDate, 7, 2) & "¤éÇeþúÇRþç°e¥IþâþóþèÇrÇoÆìþÝÄ@Æê­PþêÇeþì¡C¤S¡B©e¥ôûìÇU°l¸É¶O¥ÎÇVúú¥ÍþêÇeþîÇz¡C"
                            'Modified by Morgan 2023/2/8
                            'strTemp = "¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¤FþèþîÇrþòÇhÇR¡B©e¥ôûìÇy" & Mid(strDate, 1, 4) & "¦~" & Mid(strDate, 5, 2) & "¤ë" & Mid(strDate, 7, 2) & "¤éÇeþúÇRþç°e¥IþâþóþèÇrÇoÆìþÝÄ@Æê­PþêÇeþì¡C"
                            strTemp = PUB_GetUniText(Me.Name, "¸É©e¥ôª¬")
                            strTemp = strTemp & Mid(strDate, 1, 4) & "¦~" & Mid(strDate, 5, 2) & "¤ë" & Mid(strDate, 7, 2) & "¤é"
                            strTemp = strTemp & PUB_GetUniText(Me.Name, "¸É¥ó±Ô­z")
                            'end 2023/2/8
                            '2016/12/6 END
                        End If
                    End Select
                Case "2":
                    Select Case m_strLanguage
                    Case "2"
                        If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                        'Modify By Sindy 2013/12/3
                        strDate = DBDATE(DateAdd("m", 3, Format(DBDATE(textTM11), "####/##/##")))
                        strDate = TranslateKeyWord(incCNV_ENGLISH_DATE, strDate, Empty)
                        'modify by sonia 2024/8/6 Ú{«º°Æ²z³qª¾¡A­ì¬°namely before (" & strDate & ")¡A¨ú®ø¤é´Á«e«áªº¬A¸¹()
                        strTemp = strTemp & "    * A certified copy of the home application (priority document) must " & vbCrLf & _
                                                            "      be submitted to the IPO within three months from the day of filing " & vbCrLf & _
                                                            "      the subject application, namely before " & strDate & "."
                        '2013/12/3 END
                    Case "3"
                        'edit by nick 2004/10/06 May ´£¥X¥¿¦¡ªº¼Ë¥»
                        'strTemp = strTemp & "¡NÀu¥ý“¸¥D±iÇR¥²­nÇQ¤é¥»¥XÄ@µý©ú®Ñ"
                        If InStr(1, textAdd, "1") = 0 Then
                           'MODIFY BY Sindy 2012/6/8 ¸­©ö¶³»¡§ï¤º®e
                           'strTemp = "¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¤FþèþîÇrþòÇhÇR¡BÀu¥ý“¸¥D±iÇR¥²­nÇQ¤é¥»¥XÄ@µý©ú®ÑÇyþç°e¥IþâþóþèÇrÇoÆìþÝÄ@Æê­PþêÇeþì¡C¤S¡B¤é¥»¥XÄ@µý©ú®ÑÇU°l¸É¶O¥ÎÇRþ÷þàÇeþêþùÇV¡BµL®ÆÇOþèþîþùÆêþòþóþàÇeþì¡C"
                           'Modify By Sindy 2013/4/19 ¼W¥[¤º®e:" & Mid(strDate, 1, 4) & "¦~" & Mid(strDate, 5, 2) & "¤ë" & Mid(strDate, 7, 2) & "¤é
                           strDate = CompDate(1, "3", ChangeTStringToWString(Trim(textTM11)))
                           'Modify By Sindy 2016/8/10
                           'strTemp = "¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¤FþèþîÇrþòÇhÇR¡BÀu¥ý“¸¥D±iÇR¥²­nÇQ¤é¥»¥XÄ@µý©ú®ÑÇy" & Mid(strDate, 1, 4) & "¦~" & Mid(strDate, 5, 2) & "¤ë" & Mid(strDate, 7, 2) & "¤éÇeþúÇRþç°e¥IþâþóþèÇrÇoÆìþÝÄ@Æê­PþêÇeþì¡C¤S¡B¤é¥»¥XÄ@µý©ú®ÑÇU°l¸É¶O¥ÎÇVúú¥ÍþêÇeþîÇz¡C"
                           'Modified by Morgan 2023/2/8
                           'strTemp = "¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¤FþèþîÇrþòÇhÇR¡BÀu¥ý“¸¥D±iÇR¥²­nÇQ¤é¥»¥XÄ@µý©ú®ÑÇy" & Mid(strDate, 1, 4) & "¦~" & Mid(strDate, 5, 2) & "¤ë" & Mid(strDate, 7, 2) & "¤éÇeþúÇRþç°e¥IþâþóþèÇrÇoÆìþÝÄ@Æê­PþêÇeþì¡C"
                            strTemp = PUB_GetUniText(Me.Name, "¸ÉÀu¥ýÅvÃÒ©ú")
                            strTemp = strTemp & Mid(strDate, 1, 4) & "¦~" & Mid(strDate, 5, 2) & "¤ë" & Mid(strDate, 7, 2) & "¤é"
                            strTemp = strTemp & PUB_GetUniText(Me.Name, "¸É¥ó±Ô­z")
                            'end 2023/2/8
                           '2016/8/10 END
                        End If
                    End Select
               Case "3":
                    If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                    Select Case m_strLanguage
                    Case "2"
                        strTemp = strTemp & "    * A certified copy of Certificate of Incorporation from Register of Companies."
                    End Select
                'Add By Sindy 2011/2/17 ­Ó¤H¨­¥÷ÃÒ©ú
                Case "4":
                    If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                    Select Case m_strLanguage
                    Case "2"
                        strTemp = strTemp & "    * A certified or true copy of the Applicant's passport or ID card."
                    End Select
                '2011/2/17 End
                End Select
'            End If
         Next nIndex
            'Modify By Cheng 2003/02/20
'         If strTemp <> Empty Then: strTemp = "  The remaining documents we need for the renewal application are : " & Chr(13) & Chr(10) & strTemp
        'Modify By Cheng 2003/02/26
'         If strTemp <> Empty Then: strTemp = "    The remaining documents we need for the referenced application are : " & Chr(13) & Chr(10) & strTemp
            'Modify By Cheng 2003/12/09
            '¨Ï¥Î·sªº³W©w
'            '­Y¥Ó½Ð¤é¤p©ó20031128
'            If DBDATE(textTM11) < 20031128 Then
'                If strTemp <> Empty Then: strTemp = vbCrLf & "    The remaining documents we need for the referenced application are : " & Chr(13) & Chr(10) & strTemp
'            '­Y¥Ó½Ð¤é¤j©óµ¥©ó20031128
'            Else
                Select Case m_strLanguage
                Case "2"
                    If strTemp <> Empty Then: strTemp = vbCrLf & "    The remaining document(s) we need for the referenced application is/are : " & Chr(13) & Chr(10) & strTemp
                Case "3"
                    'edit by nick 2004/09/09 ¦]¬°¤µ¤Ñªü½¬¸É¤é¤å©e¥ôª¬ªº©w½Z
                    'If strTemp <> Empty Then: strTemp = "¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¦¨þìÇrþòÇhÇR¡B¥XÄ@¤HÇU" & strTemp & "Çyþç°e¥IþâþóþèÇrÇoÆìþÝÄ@Æê­PþêÇeþì¡C"
                    'edit by nick 2004/10/06 May ´£¥X¥¿¦¡ªº¼Ë¥»
                    If InStr(1, textAdd, "1") <> 0 And InStr(1, textAdd, "2") <> 0 Then
                        'MODIFY BY Sindy 2012/6/8 ¸­©ö¶³»¡§ï¤º®e
                        'strTemp = "¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¤FþèþîÇrþòÇhÇR¡B©e¥ôûì¤ÎÇZÀu¥ý“¸¥D±iÇR¥²­nÇQ¤é¥»¥XÄ@µý©ú®ÑÇyþç°e¥IþâþóþèÇrÇoÆìþÝÄ@Æê­PþêÇeþì¡C¤S¡B©e¥ôûì¤ÎÇZ¤é¥»¥XÄ@µý©ú®ÑÇU°l¸É¶O¥ÎÇRþ÷þàÇeþêþùÇV¡BµL®ÆÇOþèþîþùÆêþòþóþàÇeþì¡C"
                        strDate = CompDate(1, "3", ChangeTStringToWString(Trim(textTM11))) 'Add By Sindy 2012/9/10 ªü½¬»¡¦³¸É©e¥ôª¬¤ÎÀu¥ýÃÒ©ú®É¶·±a´Á­­
                        'Modify By Sindy 2016/12/6 ²{¦b­n¦¬¶O¤F
                        'strTemp = "¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¤FþèþîÇrþòÇhÇR¡B©e¥ôûì¤ÎÇZÀu¥ý“¸¥D±iÇR¥²­nÇQ¤é¥»¥XÄ@µý©ú®ÑÇy" & Mid(strDate, 1, 4) & "¦~" & Mid(strDate, 5, 2) & "¤ë" & Mid(strDate, 7, 2) & "¤éÇeþúÇRþç°e¥IþâþóþèÇrÇoÆìþÝÄ@Æê­PþêÇeþì¡C¤S¡B©e¥ôûì¤ÎÇZ¤é¥»¥XÄ@µý©ú®ÑÇU°l¸É¶O¥ÎÇVúú¥ÍþêÇeþîÇz¡C"
                        'Modified by Morgan 2023/2/8
                        'strTemp = "¡@©|¡B¥»¥ó°Ó¼Ðµn“÷¥XÄ@ÇR“¡þìÇr¤è¦¡¤â“dÇy§¹¤FþèþîÇrþòÇhÇR¡B©e¥ôûì¤ÎÇZÀu¥ý“¸¥D±iÇR¥²­nÇQ¤é¥»¥XÄ@µý©ú®ÑÇy" & Mid(strDate, 1, 4) & "¦~" & Mid(strDate, 5, 2) & "¤ë" & Mid(strDate, 7, 2) & "¤éÇeþúÇRþç°e¥IþâþóþèÇrÇoÆìþÝÄ@Æê­PþêÇeþì¡C"
                        strTemp = PUB_GetUniText(Me.Name, "¸É©e¥ôª¬¤ÎÀu¥ýÅvÃÒ©ú")
                        strTemp = strTemp & Mid(strDate, 1, 4) & "¦~" & Mid(strDate, 5, 2) & "¤ë" & Mid(strDate, 7, 2) & "¤é"
                        strTemp = strTemp & PUB_GetUniText(Me.Name, "¸É¥ó±Ô­z")
                        '2016/12/6 END
                    End If
                End Select
'            End If
      Case Else:
   End Select
   
   'Add By Sindy 2022/3/10 ³]©w¯S©w¤á¤§¯S§O³qª¾¨ç©w½Z
   If m_strLanguage = "2" Then '­^¤å
      If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, m_ET03, , "02") = True Then
         '©w½Z
         EndLetter "02", m_CP09, m_ET03, strUserNum
         Exit Sub
      End If
   End If
   '2022/3/10 END
            
   ' ®×¥ó©Ê½è
   Select Case m_CP10
      'add by nick 2004/12/23
      Case "308":
        Select Case m_strLanguage
        Case "1"
            ' ¤¤¤å
            Case "1":
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "02", m_CP09, "01", strUserNum
        Case "2"
            '©w½Z
            EndLetter "02", m_CP09, "01", strUserNum
            'Add By Sindy 2012/11/23 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
            If bolEmail = True And bolPlusPaper = False Then
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "02" & "','" & m_CP09 & "','01','" & strUserNum & _
                        "','¨Ò¥~¤º¤å','Enclosed please find a scanned copy of the Official Notice for your reference.')"
               cnnConnection.Execute strSql
            Else '¶l¥ó
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "02" & "','" & m_CP09 & "','01','" & strUserNum & _
                        "','¨Ò¥~¤º¤å','A copy of the Official Notice will be mailed to you with the confirmation copy of this letter for your records.')"
               cnnConnection.Execute strSql
            End If
            '2012/11/23 End
            
            'Ä¶¤å©w½Z
            If txtToEng = "" Then
                '¨Ò¥~Äæ¦ì  ¨ä¾l¤l®×¸ê®Æ
                strTemp = ""
                '¬d¥À®×¦A¬d¤l®×
                strSql = "select tm01||'-'||tm02||decode(tm03,'0','','-'||tm03)||decode(tm04,'00','','-'||tm04) as RefNum,tm12,tm09 from trademark,divisioncase a,divisioncase b where b.dc01='" & m_TM01 & "' and b.dc02='" & m_TM02 & "' and b.dc03='" & m_TM03 & "' and b.dc04='" & m_TM04 & "' and b.dc05=a.dc05(+) and b.dc06=a.dc06(+) and b.dc07=a.dc07(+) and b.dc08=a.dc08(+) and a.dc01=tm01(+) and a.dc02=tm02(+) and a.dc03=tm03(+) and a.dc04=tm04(+) "
                CheckOC3
                With AdoRecordSet3
                    .CursorLocation = adUseClient
                    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                    If .RecordCount <> 0 Then
                        .MoveFirst
                        Do While Not .EOF
                            strTemp = strTemp & "Application No. " & CheckStr(.Fields("tm12").Value) & "(Our Ref: " & .Fields("RefNum").Value & ")" & vbCrLf
                            strTemp = strTemp & "Class: " & CheckStr(.Fields("tm09").Value) & vbCrLf
                            strTemp = strTemp & "Goods/services designated:" & vbCrLf & vbCrLf
                            .MoveNext
                        Loop
                    End If
                    CheckOC3
                End With
                EndLetter "02", m_CP09, "02", strUserNum
            End If
            If txtToEng = "" Then
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "02" & "','" & m_CP09 & "','02','" & strUserNum & _
                         "','¨ä¾l¤l®×¸ê®Æ','" & strTemp & "')"
                cnnConnection.Execute strSql
            End If
        Case "3"
        Case Else
        End Select
        
      ' ¥Ó½Ð
      Case "101":
         ' ©w½Z»y¤å
'         Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
         Select Case m_strLanguage
            ' ¤¤¤å
            Case "1":
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "02", m_CP09, "01", strUserNum
            ' ­^¤å
            Case "2":
                'Modify By Cheng 2003/01/23
                '¦h§PÂ_¬O§_¦³Àu¥ýÅv¥X¿ï¾Ü¤£¦P³B²z¤è¦¡
                '¦ýÁp¦X¤ÎÁp¦XªA°È°Ó¼Ðªº­^Ä¶¥»»P¨ä¥L¿ï¾Ü¤£¦P³B²z¤è¦¡
'               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'               EndLetter "02", m_CP09, "02", strUserNum
'               ' ¬O§_¸É¥ó
'               If IsEmptyText(strTemp) = False Then
'                  ' ¬O§_¸É¥ó
'                  strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                           "VALUES ('" & "02" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
'                           "','¬O§_¸É¥ó','" & strTemp & "')"
'                  cnnConnection.Execute strSQL
'               End If
'               ' ¬O§_¦C¦LÂ½Ä¶¨ç
'               If textPrtTrans <> "N" Then
'                  ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                  EndLetter "02", m_CP09, "03", strUserNum
'               End If
                'Modify By Cheng 2003/12/09
                '¨Ï¥Î·sªº³W©w
'                '­Y¥Ó½Ð¤é¤p©ó20031128
'                If DBDATE(Me.textTM11.Text) < 20031128 Then
'                       ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                       EndLetter "02", m_CP09, IIf(m_blnPriDate, "02", "04"), strUserNum
'                       ' ¬O§_¸É¥ó
'                       If IsEmptyText(strTemp) = False Then
'                          ' ¬O§_¸É¥ó
'                          strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "02", "04") & "','" & strUserNum & _
'                                   "','¬O§_¸É¥ó','" & strTemp & "')"
'                          cnnConnection.Execute strSQL
'                       End If
'                       ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                       If textPrtTrans <> "N" Then
'                          ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                          '­Y°Ó¼ÐºØÃþ¬°Áp¦X©ÎÁp¦XªA°È¼Ð³¹
'                          If m_TM08 = "2" Or m_TM08 = "5" Then
'                            EndLetter "02", m_CP09, IIf(m_blnPriDate, "06", "07"), strUserNum
'                            'Add By Cheng 2003/02/26
'                            '­Y¦³©ñ±ó±M¥ÎÅv
'                            'Modify By Cheng 2003/03/11
'        '                    If m_TM67 <> "" Then
'                            If Me.textTM67.Text <> "" Then
'                                ' ©ñ±ó±M¥ÎÅv
'                                strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "06", "07") & "','" & strUserNum & _
'                                         "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & Me.textTM67.Text & "')"
'                                cnnConnection.Execute strSQL
'                            End If
'                          '¨ä¥L
'                          Else
'                            EndLetter "02", m_CP09, IIf(m_blnPriDate, "03", "05"), strUserNum
'                            'Add By Cheng 2003/02/26
'                            '­Y¦³©ñ±ó±M¥ÎÅv
'                            If Me.textTM67.Text <> "" Then
'                                ' ©ñ±ó±M¥ÎÅv
'                                strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                         "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "03", "05") & "','" & strUserNum & _
'                                         "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & Me.textTM67.Text & "')"
'                                cnnConnection.Execute strSQL
'                            End If
'                          End If
'                       End If
'                '­Y¥Ó½Ð¤é¤j©óµ¥©ó20031128
'                Else
                        ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                        EndLetter "02", m_CP09, IIf(m_blnPriDate, "10", "12"), strUserNum
                        'Add By Sindy 2012/11/23 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
                        If bolEmail = True And bolPlusPaper = False Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "10", "12") & "','" & strUserNum & _
                                    "','¨Ò¥~¤º¤å','scanned ')"
                           cnnConnection.Execute strSql
                        End If
                        '2012/11/23 End
                        ' ¬O§_¸É¥ó
                        If IsEmptyText(strTemp) = False Then
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "10", "12") & "','" & strUserNum & _
                                    "','¬O§_¸É¥ó','" & ChgSQL(strTemp) & "')"
                           cnnConnection.Execute strSql
                        End If
                        ' ¬O§_¦C¦LÂ½Ä¶¨ç
                        If textPrtTrans <> "N" Then
                           ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                           EndLetter "02", m_CP09, IIf(m_blnPriDate, "11", "13"), strUserNum
                           'Add By Cheng 2003/02/26
                           '­Y¦³©ñ±ó±M¥ÎÅv
                           If Me.textTM67.Text <> "" Then
                               ' ©ñ±ó±M¥ÎÅv
                               If m_blnPriDate Then
                                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                            "VALUES ('" & "02" & "','" & m_CP09 & "','11','" & strUserNum & _
                                            "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & Me.textTM67.Text & "')"
                               Else
                                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                            "VALUES ('" & "02" & "','" & m_CP09 & "','13','" & strUserNum & _
                                            "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & Me.textTM67.Text & "')"
                               End If
                               cnnConnection.Execute strSql
                           End If
                        End If
'                End If

            ' ¤é¤å
            Case "3":
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "02", m_CP09, "08", strUserNum
                ' ¬O§_¸É¥ó
                If IsEmptyText(strTemp) = False Then
                   ' ¬O§_¸É¥ó
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "02" & "','" & m_CP09 & "','08','" & strUserNum & _
                            "','¬O§_¸É¥ó','" & strTemp & "')"
                   cnnConnection.Execute strSql
                End If
                'Add By Sindy 2017/3/9
                m_fa76 = PUB_GetFAgentFA76(m_TM44 & String(9 - Len(m_TM44), "0"))
                If bolEmail = True Or m_fa76 = "B" Then
                   '¥÷¼Æ:
                   '©w½Z³Ì«á¦P«Êª«³¡¥÷¤§¥÷¼Æ(§Y¬õ¦â³¡¥÷)­ì³]©w¬°¡u2¡v¡A
                   '½Ð©ó¤U¦C±¡§Î®É§ï¬°¡u1¡v
                   '(1)¥N²z¤H©Î¥Ó½Ð¤H©Î®×¥ó¦³³]©w¥HE-MAIL³qª¾(§tE+±H);
                   '(2)¥N²z¤H©Ê½è¬°¡uB¡v¡C
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "02" & "','" & m_CP09 & "','08','" & strUserNum & _
                            "','¥÷¼Æ','1')"
                   cnnConnection.Execute strSql
                End If
                '2017/3/9 END
                ' ¬O§_¦C¦LÂ½Ä¶¨ç
                If textPrtTrans <> "N" Then
                  If m_blnPriDate Then
                     '²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                     EndLetter "02", m_CP09, "14", strUserNum
                     'add by nick 2004/12/16 may ¥[ªº
                     If Me.textTM67.Text <> "" Then
                        '©ñ±ó±M¥ÎÅv
                        'Modify By Sindy 2017/7/31 ­ì:°Ó¼Ð¨£¥»ÇRÆèÇr¡u" & Me.textTM67.Text & "¡vÇUˆü¥e“¸Çy“mˆüþú¥D±iþìÇrþæÇOÇVþúþàÇQÆê¡C
                        'Modified by Morgan 2023/2/8
                        'strExc(0) = "ÇÃÇ}ÇµÇ«Çè¡ÐÇÜ¡G¥»¥ó°Ó¼Ð¨£¥»ÇRþÝþäÇr¡u" & Me.textTM67.Text & "¡vÇU°Ó¼ÐÅvÇy¥D±iþêÇQÆê¡C"
                        strExc(0) = PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv1")
                        strExc(0) = strExc(0) & "¡u" & Me.textTM67.Text & "¡v"
                        strExc(0) = strExc(0) & PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv2")
                        'end 2023/2/8
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "02" & "','" & m_CP09 & "','14','" & strUserNum & "','©ñ±ó±M¥ÎÅv','" & strExc(0) & "')"
                        cnnConnection.Execute strSql
                     End If
                     '2004/12/16 END
                     'Add By Sindy 2017/10/12
                     'Àu¥ýÅv¸ê®Æ­ì¥u§ì³æµ§§ï¬°¦hµ§
                     strExc(0) = "select pd05,pd07,na03,pd06,pd10 from pridate,nation " & _
                                 "where pd01='" & m_TM01 & "' and pd02='" & m_TM02 & "' and pd03='" & m_TM03 & "' and pd04='" & m_TM04 & "' " & _
                                 "and pd07=na01 "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     strTemp = ""
                     If intI = 1 Then
                         RsTemp.MoveFirst
                         Do While Not RsTemp.EOF
                            If strTemp <> "" Then strTemp = strTemp & vbCrLf & vbCrLf
                            strExc(1) = "" & RsTemp.Fields("pd05")
                            If strExc(1) <> "" Then strExc(1) = Left(strExc(1), 4) & "¦~" & Mid(strExc(1), 5, 2) & "¤ë" & Right(strExc(1), 2) & "¤é"
                            strExc(2) = "" & RsTemp.Fields("na03")
                            'Modified by Morgan 2023/2/8
                            'If strExc(2) = "¤é¥»" Then strExc(2) = strExc(2) & "üÂ"
                            'strTemp = strTemp & "Àu¥ý“¸ûÐ¥Í¤é¡G" & strExc(1) & vbCrLf & _
                                                "Àu¥ý“¸¥D±iüÂ¡G" & strExc(2) & vbCrLf & _
                                                "Àu¥ý“¸¥D±iÇU°òÂ¦¥XÄ@¡G" & RsTemp.Fields("pd06")
                            If strExc(2) = "¤é¥»" Then strExc(2) = strExc(2) & PUB_GetUniText(Me.Name, "°ê")
                            strTemp = strTemp & PUB_GetUniText(Me.Name, "Àu¥ýÅv¤é") & "¡G" & strExc(1) & vbCrLf & _
                                                PUB_GetUniText(Me.Name, "Àu¥ýÅv°ê") & "¡G" & strExc(2) & vbCrLf & _
                                                PUB_GetUniText(Me.Name, "Àu¥ýÅv¸¹") & "¡G" & RsTemp.Fields("pd06")
                            'end 2023/2/8
                            RsTemp.MoveNext
                         Loop
                     End If
                     If strTemp <> "" Then
                         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "02" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & _
                                  "','¥D±iÀu¥ýÅv','" & strTemp & "')"
                         cnnConnection.Execute strSql
                     End If
                     '2017/10/12 END
                  Else
                     '²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                     EndLetter "02", m_CP09, "09", strUserNum
                     'add by nick 2004/12/16 may ¥[ªº
                     If Me.textTM67.Text <> "" Then
                        ' ©ñ±ó±M¥ÎÅv
                        'Modify By Sindy 2017/7/31 ­ì:°Ó¼Ð¨£¥»ÇRÆèÇr¡u" & Me.textTM67.Text & "¡vÇUˆü¥e“¸Çy“mˆüþú¥D±iþìÇrþæÇOÇVþúþàÇQÆê¡C
                        'Modified by Morgan 2023/2/8
                        'strExc(0) ="ÇÃÇ}ÇµÇ«Çè¡ÐÇÜ¡G¥»¥ó°Ó¼Ð¨£¥»ÇRþÝþäÇr¡u" & Me.textTM67.Text & "¡vÇU°Ó¼ÐÅvÇy¥D±iþêÇQÆê¡C"
                        strExc(0) = PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv1")
                        strExc(0) = strExc(0) & "¡u" & Me.textTM67.Text & "¡v"
                        strExc(0) = strExc(0) & PUB_GetUniText(Me.Name, "©ñ±ó±M¥ÎÅv2")
                        'end 2023/2/8
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "02" & "','" & m_CP09 & "','09','" & strUserNum & "','©ñ±ó±M¥ÎÅv','" & strExc(0) & "')"
                        cnnConnection.Execute strSql
                     End If
                  End If
                End If
         End Select
      Case Else:
   End Select
End Sub

Private Sub PrintLetter()
   'Add by Morgan 2008/6/12
   Dim ET03 As String, ET03_1 As String, stContent As String
   Dim strFilePath As String, strFN01 As String, strFN02 As String  'Added by Lydia 2023/05/03
   
    'Add By Cheng 2004/05/10
    '¨ú±o©w½Z»y¤å
    m_strLanguage = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
    'End
    
   'Add By Sindy 2012/11/23 ±q¤U­±µ{¦¡©¹¤WMove¦Ü¦¹
   bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper) 'ÀË¬d¬O§_¥HE-Mail³qª¾
   '2012/11/23 End
   
   ' ¥ý©I¥s©w½Zµ{¦¡ªº²M°£­ì©w½Z¸ê®Æªº¨ç¦¡¥h²M°£¤§«e´Ý¯d¦b¨Ò¥~Äæ¦ìÀÉ¤¤ªº¸ê®Æ
   InsExpField
      
   ' ®×¥ó©Ê½è
   Select Case m_CP10
      'add by nick 2004/12/23
      Case "308":
        Select Case m_strLanguage
        Case "1"
            '¤¤¤å 2005/8/26 ADD BY SONIA
            ET03 = "03"
            '2005/8/26 END
            
        Case "2"
            'Modify By Sindy 2022/3/10
            If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, ET03, , "02") = False Then
            '2022/3/10 END
               ET03 = "01"
            End If
            'Ä¶¤å©w½Z
            If txtToEng = "" Then
               ET03 = "02"
            End If
        Case "3"
        Case Else
        End Select
      ' ¥Ó½Ð
      Case "101":
         ' ©w½Z»y¤å
'         Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
         Select Case m_strLanguage
            ' ¤¤¤å
            Case "1":
               ET03 = "01"
            ' ­^¤å
            Case "2":
                'Modify By Cheng 2003/01/23
                '¬O§_¦³Àu¥ýÅv¥X¤£¦P©w½Z
'               ' ¦C¦L©w½Z
'               NowPrint m_CP09, "02", "02", False, strUserNum, 0
'               ' ¬O§_¦C¦LÂ½Ä¶¨ç
'               If textPrtTrans <> "N" Then
'                  ' ¦C¦L©w½Z
'                  NowPrint m_CP09, "02", "03", False, strUserNum, 0
'               End If
                'Modify By Cheng 2003/12/09
                '¨Ï¥Î·sªº³W©w
'                '­Y¥Ó½Ð¤é¤p©ó20031128
'                If DBDATE(Me.textTM11.Text) < 20031128 Then
'                    ' ¦C¦L©w½Z
'                    NowPrint m_CP09, "02", IIf(m_blnPriDate, "02", "04"), False, strUserNum, 0
'                    ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                    If textPrtTrans <> "N" Then
'                       ' ¦C¦L©w½Z
'                         'Áp¦X°Ó¼Ð
'                         If m_TM08 = "2" Or m_TM08 = "5" Then
'                           NowPrint m_CP09, "02", IIf(m_blnPriDate, "06", "07"), False, strUserNum, 0
'                         '«DÁp¦X°Ó¼Ð
'                         Else
'                           NowPrint m_CP09, "02", IIf(m_blnPriDate, "03", "05"), False, strUserNum, 0
'                         End If
'                    End If
'                '­Y¥Ó½Ð¤é¤j©óµ¥©ó20031128
'                Else
                  
                  'Modify By Sindy 2022/3/10
                  If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, ET03, , "02") = False Then
                  '2022/3/10 END
                     If m_blnPriDate Then
                        ET03 = "10"
                     Else
                        ET03 = "12"
                     End If
                  End If
                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
                  If textPrtTrans <> "N" Then
                     If m_blnPriDate Then
                        ET03_1 = "11"
                     Else
                        ET03_1 = "13"
                     End If
                  End If
'                End If
            ' ¤é¤å
            Case "3":
                ET03 = "08"
                ' ¬O§_¦C¦LÂ½Ä¶¨ç
                If textPrtTrans <> "N" Then
                    'edit by nick 2005/02/05 ¥[¤JÀu¥ýÅvªº©w½Z
                    'NowPrint m_CP09, "02", "09", False, strUserNum, 0
                    If m_blnPriDate Then
                        ET03_1 = "14"
                    Else
                        ET03_1 = "09"
                    End If
                End If
         End Select
      Case Else:
   End Select
   
   If ET03 <> "" Then
      'Add by Morgan 2008/6/12
'      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      'If bolEmail Then 'Mark by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW
         'Add by Morgan 2009/10/20 +§PÂ_¬O§_EMail¦P®É±H¯È¥»
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'end 2009/10/20
         'Added by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW: ­^¤å²Õ¤À¦¨«H¨ç©MÂ½Ä¶¨â­ÓÀÉ®×
         If m_strLanguage <> "3" Then
            strFilePath = Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04)
            If Pub_GetFCTeFileName(strFilePath, m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, , strFN01, strFN02) = False Then
              Exit Sub
            End If
            NowPrint m_CP09, "02", ET03, True, strUserNum, , , , , iCopy, , True
            If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                Sleep 100
            End If
            If ET03_1 <> "" Then
               NowPrint m_CP09, "02", ET03_1, True, strUserNum, , , , , iCopy, , True
               If PUB_PrintWord2File(g_WordAp, strFilePath, strFN02) = True Then
                   Sleep 100
               End If
            End If
         Else '¤é¤å²Õ:¤£§ïÅÜ¦sÀÉ¼Ò¦¡
         'end 2023/05/03
            'Added by Lydia 2024/11/14 ¦]¤é¥»¥N²z¤H¯S§O­n¨D¡A»Ý±N³qª¾«H¨ç»PÄ¶¤åµ¥¤À¶}¡A¨Ã¥B²Î¤@¦WºÙ¦p¤U(¼Ò²Õ¨ú±o)¡F­ì¥»ªºÀÉ®×(®×¸¹_¤é´Á=³qª¾¨ç+Ä¶¤å)¤´­n²£¥Í¡A¥H§K¤é«á¤S¦³¥N²z¤H­n¨D¦X¨Ö
            strFilePath = Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04)
            If Pub_GetFCTeFileName(strFilePath, m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, , strFN01, strFN02) = False Then
              Exit Sub
            End If
            NowPrint m_CP09, "02", ET03, True, strUserNum, , , , , iCopy, , True
            If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                Sleep 100
            End If
            If ET03_1 <> "" Then
               NowPrint m_CP09, "02", ET03_1, True, strUserNum, , , , , iCopy, , True
               If PUB_PrintWord2File(g_WordAp, strFilePath, strFN02) = True Then
                   Sleep 100
               End If
            End If
            'end 2024/11/14
            If ET03_1 <> "" Then
               NowPrint m_CP09, "02", ET03, False, strUserNum, , , , , iCopy
               NowPrint m_CP09, "02", ET03_1, False, strUserNum, , , , , iCopy
               NowPrint m_CP09, "02", ET03, False, strUserNum, , , True, stContent, , , , True
               NowPrint m_CP09, "02", ET03_1, False, strUserNum, , stContent, , , , , True, True
            Else
               NowPrint m_CP09, "02", ET03, False, strUserNum, , , , , iCopy, , True, True
            End If
         End If 'Added by Lydia 2023/05/03
         MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(m_TM01) & " ]¡I"
      'Mark by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW
      'Else
      ''end 2008/6/12
       '  NowPrint m_CP09, "02", ET03, False, strUserNum, 0
      '   If ET03_1 <> "" Then
      '      NowPrint m_CP09, "02", ET03_1, False, strUserNum, 0
      '   End If
      'End If
      'end 2023/05/03
      
   End If
End Sub

Private Sub textTM11_GotFocus()
   InverseTextBox textTM11
End Sub

Private Sub textTM12_GotFocus()
   InverseTextBox textTM12
End Sub

'Add By Sindy 2010/9/1
Private Sub textTM12_Validate(Cancel As Boolean)
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   If IsEmptyText(textTM12) = False Then
      'ÀË¬d¥Ó½Ð®×¸¹©Ò¿é¤Jªºªø«×¬O§_¥¿½T
      'Add By Sindy 2017/5/17 + strRetrunText
      If PUB_ChkTm12Tm15Length("1", textTM12, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10, , , strRetrunText) = False Then
         Cancel = True
         textTM12_GotFocus
         Exit Sub
      'Add By Sindy 2017/5/17
      Else
         textTM12 = strRetrunText
      '2017/5/17 END
      End If
   End If
End Sub

Private Sub textTM27_GotFocus()
   InverseTextBox textTM27
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub

Private Sub textAdd_GotFocus()
   InverseTextBox textAdd
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textPrtTrans_GotFocus()
   InverseTextBox textPrtTrans
End Sub

Private Sub textPS_GotFocus()
   InverseTextBox textPS
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

'Add By Sindy 2010/12/24
If Me.textTM12.Enabled = True Then
   Cancel = False
   textTM12_Validate Cancel
   If Cancel = True Then
      textTM12.SetFocus
      Exit Function
   End If
End If

'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
If Me.textPriorityDoc.Enabled = True Then
   Cancel = False
   textPriorityDoc_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
If Me.textAddDate.Enabled = True Then
   Cancel = False
   textAddDate_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'add by nickc 2006/12/21
If Me.textAddDate2.Enabled = True Then
   Cancel = False
   textAddDate2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textAdd.Enabled = True Then
   Cancel = False
   textAdd_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP05.Enabled = True Then
   Cancel = False
   textCP05_Validate Cancel
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

If Me.textPrtTrans.Enabled = True Then
   Cancel = False
   textPrtTrans_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM11.Enabled = True Then
   Cancel = False
   textTM11_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

'add by nick 2004/08/13 ±q frm030202_03 ²¾¨Ó
' ¸É¤å¥ó´Á­­
Private Sub textAddDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textAddDate) = False Then
      If CheckIsTaiwanDate(textAddDate, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         'edit by nickc 2006/12/21
         'strMsg = "½Ð¿é¤J¥¿½Tªº¸É¤å¥ó´Á­­"
         strMsg = "½Ð¿é¤J¥¿½Tªº¸ÉÀu¥ýÅvÃÒ©ú´Á­­"
         
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textAddDate_GotFocus
      End If
   End If
End Sub

Private Sub textTM67_GotFocus()
InverseTextBox textTM67
End Sub

'add by nick 2004/12/30
Private Sub txtToEng_GotFocus()
   InverseTextBox txtToEng
End Sub

Private Sub txtToEng_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(txtToEng) = False Then
      Select Case txtToEng
         Case "", " ", "N":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎN"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            'textPrint_GotFocus
      End Select
   End If
End Sub

' Åª¨ú°Ó¼Ð°ò¥»ÀÉ
Private Sub QueryMonTradeMark()
   Dim strSql As String
   Dim strSub As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSub As ADODB.Recordset
   
   ' ¨ú±o°Ó¼Ð°ò¥»ÀÉªº¬ÛÃö¶µ¥Ø
   strSql = "SELECT * FROM TradeMark,divisioncase " & _
            "WHERE dc01 = '" & m_TM01 & "' AND " & _
                  "dc02 = '" & m_TM02 & "' AND " & _
                  "dc03 = '" & m_TM03 & "' AND " & _
                  "dc04 = '" & m_TM04 & "' and dc05=tm01(+) and dc06=tm02(+) and dc07=tm03(+) and dc08=tm04(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' ¥Ó½Ð¤é
      textTM11 = ChangeWStringToTString(CheckStr(rsTmp.Fields("TM11")))
        m_MonTM01 = CheckStr(rsTmp.Fields("tm01"))
        m_MonTM02 = CheckStr(rsTmp.Fields("tm02"))
        m_MonTM03 = CheckStr(rsTmp.Fields("tm03"))
        m_MonTM04 = CheckStr(rsTmp.Fields("tm04"))
        If textNP08.Enabled = True And textNP09.Enabled = True Then
             strSql = "SELECT * FROM nextprogress " & _
                      "WHERE np02 = '" & m_MonTM01 & "' AND " & _
                           " np03 = '" & m_MonTM02 & "' AND " & _
                           " np04 = '" & m_MonTM03 & "' AND " & _
                           " np05 = '" & m_MonTM04 & "' and np06 is null and np07=202 "
            rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
                m_MonNP08 = CheckStr(rsTmp.Fields("np08"))
                m_MonNP09 = CheckStr(rsTmp.Fields("np09"))
            End If
        End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub txtLaw_GotFocus()
    InverseTextBox txtLaw
End Sub

'Add By Sindy 2012/5/18
Private Sub Option1_Click(Index As Integer)
   If Me.Option4(0).Value Then
      Text10_Validate False
   ElseIf Me.Option4(1).Value Then
      Text11_Validate False
   ElseIf Me.Option4(2).Value Then
      Text12_Validate False
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme
End Sub

Private Sub Text10_LostFocus()
   '«D¥xÆW"¤Ñ"¸õÂ÷®É¨ì"¥»©Ò´Á­­"Äæ¦ì
   If m_TM10 <> ¥xÆW°ê®a¥N¸¹ Then
      If textNP08.Enabled = True Then textNP08.SetFocus
   End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
   CloseIme
End Sub

Private Sub Text11_LostFocus()
   '«D¥xÆW"¤ë"¸õÂ÷®É¨ì"¥»©Ò´Á­­"Äæ¦ì
   'If m_TM10 <> ¥xÆW°ê®a¥N¸¹ Then
   '   If textNP08.Enabled = True Then textNP08.SetFocus
   'End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_LostFocus()
   '«D¥xÆW"¤é"¸õÂ÷®É¨ì"¥»©Ò´Á­­"Äæ¦ì
   If m_TM10 <> ¥xÆW°ê®a¥N¸¹ Then
      If textNP08.Enabled = True Then textNP08.SetFocus
   End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
   Else
      If ChkDate(Text12) Then
         If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
            If Val(Text12) < Val(strSrvDate(2)) Then
               MsgBox "¨Ó¨ç´Á­­¤£¥i¤p©ó¨t²Î¤é !", vbCritical
               Cancel = True
            Else
               textNP09 = Text12
               'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
               If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                  textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
               Else
               '2014/10/6 END
                  textNP08 = TransDate(CompDate(2, -2, TransDate(textNP09, 2)), 1)
               End If
               textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08.Text, True), 1) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
            End If
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub GetTime()
   Dim i As Integer
   Dim strFromDate As String '´Á­­°_ºâ¤é
   
   'Add By Sindy 2012/8/30
   If Option4(0).Value = False And Option4(1).Value = False Then Exit Sub
   '2012/8/30 End
   
   strFromDate = DBDATE(textCP05)
   
   If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
      '¤å¨ì¤Ñ¼Æ
      If Option4(0).Value = True Then
         textNP09 = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
         If Option1(0).Value = True Then textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
         If Val(Text10) >= 60 Then
            i = -4
         Else
            i = -2
         End If
      '¤å¨ì¤ë¼Æ
      ElseIf Option4(1).Value = True Then
         textNP09 = TAIWANDATE(AddMonth(strFromDate, Val(Text11)))
         If Option1(0).Value = True Then textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
         If Val(Text11) >= 2 Then
            i = -4
         Else
            i = -2
         End If
      End If
      If textNP09 <> "" Then
         'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
         If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
            textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
         Else
         '2014/10/6 END
            textNP08 = TransDate(CompDate(2, i, TransDate(textNP09, 2)), 1)
         End If
         textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08.Text, True), 1) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
      End If
   End If
End Sub

'Åª¨ú¨Ó¨ç´Á­­
Private Function ChgType() As Boolean
Dim strTempName As String, bolTmp As Boolean
Dim i As Integer
Dim strFromDate As String '´Á­­°_ºâ¤é
   
   strFromDate = DBDATE(textCP05)
   
   ChgType = False
   If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
      bolTmp = False
   Else
      bolTmp = True
   End If
   
   ' ®×¥ó©Ê½è
   strRvType = LabNP07.Caption '202.¥Ó½Ð·N¨£®Ñ
   If strRvType = "" Then Exit Function
   
   If ClsPDGetCaseProperty(m_TM01, strRvType, strTempName, bolTmp) Then
      textNP08 = ""
      textNP09 = ""
      
      If m_TM10 = ¥xÆW°ê®a¥N¸¹ Then
         strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & m_TM01 & "' AND CPM02='" & strRvType & "'"
         If strExc(0) <> "" Then
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            With RsTemp
               If intI = 1 Then
                  '2015/3/5 add by sonia FCT®×¤£¹w³]¤å¨ì¤Ñ¼Æ©Î¤ë¼Æ,¦]¬°¥i¯à30¤Ñ¥i¯à1­Ó¤ë FCT-036284
                  If m_TM01 = "FCT" Then
                     Option4(0).Value = False
                     If Not IsNull(.Fields(0)) Then
                        '¤å¨ì·í¤é
                        If .Fields(0) = "1" Then
                           Option1(0).Value = True
                        '¤å¨ì¦¸¤é
                        Else
                           Option1(1).Value = True
                        End If
                     End If
                  Else
                  '2015/3/5 END
                     If Not IsNull(.Fields(1)) Then
                        '¤å¨ì¤Ñ¼Æ
                        Option4(0).Value = True
                        Text10 = .Fields(1)
                        textNP09 = TransDate(CompDate(2, Text10, TransDate(strFromDate, 2)), 1)
                     ElseIf Not IsNull(.Fields(2)) Then
                        '¤å¨ì¤ë¼Æ
                        Option4(1).Value = True
                        Text11 = .Fields(2)
                        textNP09 = TransDate(CompDate(1, .Fields(2), TransDate(strFromDate, 2)), 1)
                     Else
                        '¤å¨ì¤Ñ¼Æ
                        Option4(0).Value = True
                        Text10 = ""
                        Text11 = ""
                     End If
                     If textNP09 <> "" And Not IsNull(.Fields(0)) Then
                        '¤å¨ì·í¤é
                        If .Fields(0) = "1" Then
                           Option1(0).Value = True
                           textNP09 = TransDate(CompDate(2, -1, TransDate(textNP09, 2)), 1)
                        '¤å¨ì¦¸¤é
                        Else
                           Option1(1).Value = True
                        End If
                     End If
                     '¤å¨ì¤Ñ¼Æ
                     If Text10 <> "" Then
                        If Val(Text10) >= 60 Then
                           i = -4
                        Else
                           i = -2
                        End If
                     '¤å¨ì¤ë¼Æ
                     ElseIf Not IsNull(.Fields(2)) Then
                        If Val(.Fields(2)) >= 2 Then
                           i = -4
                        Else
                           i = -2
                        End If
                     End If
                     If textNP09 <> "" Then
                        'Modify By Sindy 2014/10/6 ¥xÆW®×¤§¥»©Ò´Á­­³]©w
                        If m_TM10 = "000" And Val(strSrvDate(1)) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
                           textNP08 = TransDate(PUB_GetOurDeadline(DBDATE(textNP09)), 1)
                        Else
                        '2014/10/6 END
                           textNP08 = TransDate(CompDate(2, i, TransDate(textNP09, 2)), 1)
                        End If
                        textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08.Text, True), 1) 'Added by Lydia 2020/07/07 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
                     End If
                  End If   '2015/3/5 ADD BY SONIA
               End If
            End With
         End If
      End If
      ChgType = True
   End If
End Function
