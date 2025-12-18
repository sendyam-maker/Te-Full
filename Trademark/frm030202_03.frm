VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030202_03 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "µo¤å(¥Ó½Ð, ©µ®i, ¸É´«µoÃÒ®Ñ, ­^¤åÃÒ©ú)"
   ClientHeight    =   6252
   ClientLeft      =   84
   ClientTop       =   1404
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6252
   ScaleWidth      =   9156
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   280
      Left            =   1344
      Locked          =   -1  'True
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   1014
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   280
      Left            =   5664
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   1980
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   280
      Left            =   1344
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   1335
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   280
      Left            =   1344
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   372
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   280
      Left            =   1344
      Locked          =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   693
      Width           =   2532
   End
   Begin VB.TextBox textTM12S 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   280
      Left            =   5664
      Locked          =   -1  'True
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   1335
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   280
      Left            =   5664
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   1014
      Width           =   2532
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   280
      Left            =   5664
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   693
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8160
      TabIndex        =   48
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6060
      TabIndex        =   40
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "¦^«eµe­±(&U)"
      Height          =   350
      Left            =   7020
      TabIndex        =   41
      Top             =   0
      Width           =   1092
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H008080FF&
      Caption         =   "ÅÜ§ó¨Æ¶µ(&R)"
      Height          =   350
      Left            =   4920
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   39
      Top             =   0
      Width           =   1092
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "¬ÛÃö¨÷¸¹(&F)"
      Height          =   350
      Left            =   3780
      TabIndex        =   38
      Top             =   0
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "¦P®Éµo¤å(&N)"
      Height          =   350
      Index           =   1
      Left            =   2664
      TabIndex        =   37
      Top             =   0
      Width           =   1092
   End
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "®×¥ó¶i«×(&C)"
      Height          =   350
      Left            =   1410
      TabIndex        =   36
      Top             =   0
      Width           =   1212
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3930
      Left            =   180
      TabIndex        =   49
      Top             =   2280
      Width           =   8895
      _ExtentX        =   15706
      _ExtentY        =   6922
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   420
      TabCaption(0)   =   "°ò¥»¸ê®Æ1"
      TabPicture(0)   =   "frm030202_03.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label16"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(10)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label22"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label5(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label14(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label25"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label17"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label11"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label19"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label15"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label42"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label37"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label36"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(5)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label39"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label10"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label12"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label13"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label21"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(12)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label43"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textTM81_2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textTM80_2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textTM79_2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textTM78_2"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textTM23_2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textTM05"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textTM06"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textTM07"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textTM05_1"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cboTM08"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cboTM72"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textTM08_2"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textTM23"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCP18"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textTM27"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textTM08"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textPrint"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textTM21"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textTM22"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textUargeDate"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "textCP27"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "textAdd"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "textAdd_2"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "textPrtTrans"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "textCP26"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "textDN"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "textTM72_2"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "textTM72"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "textCP84"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "textTM78"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "textTM79"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "textTM80"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "textTM81"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "textCP113"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "textCP118"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).ControlCount=   64
      TabCaption(1)   =   "°ò¥»¸ê®Æ2"
      TabPicture(1)   =   "frm030202_03.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "textCP09_S"
      Tab(1).Control(2)=   "textCP09_S1"
      Tab(1).Control(3)=   "textCP09_S2"
      Tab(1).Control(4)=   "textCP09_S3"
      Tab(1).Control(5)=   "textTM09"
      Tab(1).Control(6)=   "textTM32"
      Tab(1).Control(7)=   "textMail"
      Tab(1).Control(8)=   "lstNameAgent"
      Tab(1).Control(9)=   "textCP64"
      Tab(1).Control(10)=   "textTM58"
      Tab(1).Control(11)=   "lblNameAgent"
      Tab(1).Control(12)=   "Line2"
      Tab(1).Control(13)=   "Label1(14)"
      Tab(1).Control(14)=   "Label1(13)"
      Tab(1).Control(15)=   "Label41"
      Tab(1).Control(16)=   "Label40"
      Tab(1).Control(17)=   "Label27"
      Tab(1).Control(18)=   "Label28"
      Tab(1).Control(19)=   "Label29"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "¥Nªí¤H-1"
      TabPicture(2)   =   "frm030202_03.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label18(2)"
      Tab(2).Control(1)=   "Label14(1)"
      Tab(2).Control(2)=   "Label5(3)"
      Tab(2).Control(3)=   "Label5(4)"
      Tab(2).Control(4)=   "Label5(5)"
      Tab(2).Control(5)=   "Label5(6)"
      Tab(2).Control(6)=   "Label5(7)"
      Tab(2).Control(7)=   "Label5(8)"
      Tab(2).Control(8)=   "Label5(1)"
      Tab(2).Control(9)=   "Label5(2)"
      Tab(2).Control(10)=   "Label5(9)"
      Tab(2).Control(11)=   "Label5(10)"
      Tab(2).Control(12)=   "Label5(11)"
      Tab(2).Control(13)=   "Label5(12)"
      Tab(2).Control(14)=   "Label14(2)"
      Tab(2).Control(15)=   "Label18(1)"
      Tab(2).Control(16)=   "Label5(13)"
      Tab(2).Control(17)=   "Label5(14)"
      Tab(2).Control(18)=   "Label5(15)"
      Tab(2).Control(19)=   "Label5(16)"
      Tab(2).Control(20)=   "Label5(17)"
      Tab(2).Control(21)=   "Label5(18)"
      Tab(2).Control(22)=   "Label14(3)"
      Tab(2).Control(23)=   "Label18(3)"
      Tab(2).Control(24)=   "textTM105"
      Tab(2).Control(25)=   "textTM104"
      Tab(2).Control(26)=   "textTM103"
      Tab(2).Control(27)=   "Combo2(5)"
      Tab(2).Control(28)=   "textTM99"
      Tab(2).Control(29)=   "textTM98"
      Tab(2).Control(30)=   "textTM97"
      Tab(2).Control(31)=   "Combo2(3)"
      Tab(2).Control(32)=   "textTM52"
      Tab(2).Control(33)=   "textTM51"
      Tab(2).Control(34)=   "textTM50"
      Tab(2).Control(35)=   "Combo2(1)"
      Tab(2).Control(36)=   "textTM102"
      Tab(2).Control(37)=   "textTM101"
      Tab(2).Control(38)=   "textTM100"
      Tab(2).Control(39)=   "Combo2(4)"
      Tab(2).Control(40)=   "textTM96"
      Tab(2).Control(41)=   "textTM95"
      Tab(2).Control(42)=   "textTM94"
      Tab(2).Control(43)=   "Combo2(2)"
      Tab(2).Control(44)=   "textTM49"
      Tab(2).Control(45)=   "textTM48"
      Tab(2).Control(46)=   "textTM47"
      Tab(2).Control(47)=   "Combo2(0)"
      Tab(2).ControlCount=   48
      TabCaption(3)   =   "¥Nªí¤H-2"
      TabPicture(3)   =   "frm030202_03.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TextTM106"
      Tab(3).Control(1)=   "TextTM107"
      Tab(3).Control(2)=   "TextTM109"
      Tab(3).Control(3)=   "TextTM110"
      Tab(3).Control(4)=   "Combo2(7)"
      Tab(3).Control(5)=   "Combo2(6)"
      Tab(3).Control(6)=   "TextTM108"
      Tab(3).Control(7)=   "TextTM111"
      Tab(3).Control(8)=   "TextTM112"
      Tab(3).Control(9)=   "TextTM113"
      Tab(3).Control(10)=   "TextTM115"
      Tab(3).Control(11)=   "TextTM116"
      Tab(3).Control(12)=   "Combo2(9)"
      Tab(3).Control(13)=   "Combo2(8)"
      Tab(3).Control(14)=   "TextTM114"
      Tab(3).Control(15)=   "TextTM117"
      Tab(3).Control(16)=   "Label18(5)"
      Tab(3).Control(17)=   "Label14(5)"
      Tab(3).Control(18)=   "Label5(30)"
      Tab(3).Control(19)=   "Label5(29)"
      Tab(3).Control(20)=   "Label5(28)"
      Tab(3).Control(21)=   "Label5(27)"
      Tab(3).Control(22)=   "Label5(26)"
      Tab(3).Control(23)=   "Label5(25)"
      Tab(3).Control(24)=   "Label18(4)"
      Tab(3).Control(25)=   "Label14(4)"
      Tab(3).Control(26)=   "Label5(24)"
      Tab(3).Control(27)=   "Label5(23)"
      Tab(3).Control(28)=   "Label5(22)"
      Tab(3).Control(29)=   "Label5(21)"
      Tab(3).Control(30)=   "Label5(20)"
      Tab(3).Control(31)=   "Label5(19)"
      Tab(3).ControlCount=   32
      Begin VB.TextBox textCP118 
         Height          =   270
         Left            =   5730
         MaxLength       =   1
         TabIndex        =   19
         Top             =   2437
         Width           =   375
      End
      Begin VB.TextBox textCP113 
         Height          =   285
         Left            =   5910
         MaxLength       =   4
         TabIndex        =   2
         Top             =   300
         Width           =   600
      End
      Begin VB.TextBox textTM81 
         Height          =   285
         Left            =   840
         MaxLength       =   9
         TabIndex        =   18
         Top             =   2430
         Width           =   1092
      End
      Begin VB.TextBox textTM80 
         Height          =   264
         Left            =   5250
         MaxLength       =   9
         TabIndex        =   17
         Top             =   2134
         Width           =   1092
      End
      Begin VB.TextBox textTM79 
         Height          =   285
         Left            =   840
         MaxLength       =   9
         TabIndex        =   16
         Top             =   2124
         Width           =   1092
      End
      Begin VB.TextBox textTM78 
         Height          =   264
         Left            =   5250
         MaxLength       =   9
         TabIndex        =   15
         Top             =   1830
         Width           =   1092
      End
      Begin VB.TextBox Text7 
         Height          =   288
         Left            =   -74790
         MaxLength       =   1
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   1530
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox textCP84 
         Alignment       =   1  '¾a¥k¹ï»ô
         Height          =   285
         Left            =   3570
         TabIndex        =   1
         Top             =   300
         Width           =   1425
      End
      Begin VB.TextBox textTM72 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3924
         MaxLength       =   1
         TabIndex        =   7
         Top             =   1224
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textTM72_2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   4308
         Locked          =   -1  'True
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   1224
         Visible         =   0   'False
         Width           =   828
      End
      Begin VB.TextBox textDN 
         Height          =   285
         Left            =   6630
         MaxLength       =   1
         TabIndex        =   10
         Top             =   908
         Width           =   492
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   8280
         MaxLength       =   1
         TabIndex        =   24
         Top             =   1290
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textPrtTrans 
         Height          =   285
         Left            =   6630
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1212
         Width           =   372
      End
      Begin VB.TextBox textCP09_S 
         Height          =   285
         Left            =   -70590
         MaxLength       =   1
         TabIndex        =   29
         Top             =   960
         Width           =   465
      End
      Begin VB.TextBox textCP09_S1 
         Height          =   285
         Left            =   -70050
         MaxLength       =   6
         TabIndex        =   30
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox textCP09_S2 
         Height          =   285
         Left            =   -68970
         MaxLength       =   1
         TabIndex        =   31
         Top             =   960
         Width           =   345
      End
      Begin VB.TextBox textCP09_S3 
         Height          =   285
         Left            =   -68520
         MaxLength       =   2
         TabIndex        =   32
         Top             =   960
         Width           =   465
      End
      Begin VB.TextBox textAdd_2 
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   1526
         Width           =   6072
      End
      Begin VB.TextBox textAdd 
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   13
         Top             =   1516
         Width           =   852
      End
      Begin VB.TextBox textTM09 
         Height          =   285
         Left            =   -73560
         MaxLength       =   395
         TabIndex        =   26
         Top             =   360
         Width           =   7272
      End
      Begin VB.TextBox textTM32 
         Height          =   285
         Left            =   -73560
         MaxLength       =   699
         TabIndex        =   27
         Top             =   660
         Width           =   7272
      End
      Begin VB.TextBox textMail 
         Height          =   285
         Left            =   -73560
         MaxLength       =   1
         TabIndex        =   28
         Top             =   960
         Width           =   492
      End
      Begin VB.TextBox textCP27 
         Height          =   285
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   0
         Top             =   300
         Width           =   1092
      End
      Begin VB.TextBox textUargeDate 
         Height          =   285
         Left            =   1440
         MaxLength       =   7
         TabIndex        =   3
         Top             =   604
         Width           =   1092
      End
      Begin VB.TextBox textTM22 
         Height          =   285
         Left            =   2880
         MaxLength       =   7
         TabIndex        =   9
         Top             =   908
         Width           =   852
      End
      Begin VB.TextBox textTM21 
         Height          =   285
         Left            =   1704
         MaxLength       =   7
         TabIndex        =   8
         Top             =   908
         Width           =   852
      End
      Begin VB.TextBox textPrint 
         Height          =   285
         Left            =   1704
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1212
         Width           =   372
      End
      Begin VB.TextBox textTM08 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Left            =   3936
         MaxLength       =   1
         TabIndex        =   6
         Top             =   888
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textTM27 
         Height          =   288
         Left            =   8715
         MaxLength       =   20
         TabIndex        =   25
         Top             =   3075
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.TextBox textCP18 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   285
         Left            =   7170
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   300
         Width           =   1635
      End
      Begin VB.TextBox textTM23 
         Height          =   285
         Left            =   840
         MaxLength       =   9
         TabIndex        =   14
         Top             =   1820
         Width           =   1092
      End
      Begin VB.TextBox textTM08_2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   888
         Visible         =   0   'False
         Width           =   912
      End
      Begin MSForms.ComboBox cboTM72 
         Height          =   300
         Left            =   6630
         TabIndex        =   5
         Top             =   600
         Width           =   2124
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3746;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboTM08 
         Height          =   288
         Left            =   3576
         TabIndex        =   4
         Top             =   600
         Width           =   1932
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3408;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM106 
         Height          =   285
         Left            =   -74100
         TabIndex        =   164
         Top             =   630
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM107 
         Height          =   285
         Left            =   -74100
         TabIndex        =   193
         Top             =   930
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM109 
         Height          =   285
         Left            =   -69690
         TabIndex        =   192
         Top             =   630
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM110 
         Height          =   285
         Left            =   -69690
         TabIndex        =   191
         Top             =   930
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   7
         Left            =   -69690
         TabIndex        =   53
         Top             =   330
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   6
         Left            =   -74100
         TabIndex        =   52
         Top             =   330
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM108 
         Height          =   285
         Left            =   -74100
         TabIndex        =   190
         Top             =   1230
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM111 
         Height          =   285
         Left            =   -69690
         TabIndex        =   189
         Top             =   1230
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM112 
         Height          =   285
         Left            =   -74100
         TabIndex        =   188
         Top             =   1830
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM113 
         Height          =   285
         Left            =   -74100
         TabIndex        =   187
         Top             =   2115
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM115 
         Height          =   285
         Left            =   -69690
         TabIndex        =   186
         Top             =   1830
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM116 
         Height          =   285
         Left            =   -69690
         TabIndex        =   185
         Top             =   2115
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   9
         Left            =   -69690
         TabIndex        =   55
         Top             =   1530
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   8
         Left            =   -74100
         TabIndex        =   54
         Top             =   1515
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM114 
         Height          =   285
         Left            =   -74100
         TabIndex        =   184
         Top             =   2415
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM117 
         Height          =   285
         Left            =   -69690
         TabIndex        =   183
         Top             =   2415
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   0
         Left            =   -74100
         TabIndex        =   42
         Top             =   300
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM47 
         Height          =   285
         Left            =   -74100
         TabIndex        =   182
         Top             =   603
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM48 
         Height          =   285
         Left            =   -74100
         TabIndex        =   181
         Top             =   891
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM49 
         Height          =   285
         Left            =   -74100
         TabIndex        =   180
         Top             =   1179
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   2
         Left            =   -74100
         TabIndex        =   44
         Top             =   1467
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM94 
         Height          =   285
         Left            =   -74100
         TabIndex        =   179
         Top             =   1770
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM95 
         Height          =   285
         Left            =   -74100
         TabIndex        =   178
         Top             =   2058
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM96 
         Height          =   285
         Left            =   -74100
         TabIndex        =   177
         Top             =   2346
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   4
         Left            =   -74100
         TabIndex        =   46
         Top             =   2634
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM100 
         Height          =   285
         Left            =   -74100
         TabIndex        =   176
         Top             =   2937
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM101 
         Height          =   285
         Left            =   -74100
         TabIndex        =   175
         Top             =   3225
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM102 
         Height          =   285
         Left            =   -74100
         TabIndex        =   174
         Top             =   3520
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   1
         Left            =   -69690
         TabIndex        =   43
         Top             =   300
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM50 
         Height          =   285
         Left            =   -69690
         TabIndex        =   173
         Top             =   603
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM51 
         Height          =   285
         Left            =   -69690
         TabIndex        =   172
         Top             =   891
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM52 
         Height          =   285
         Left            =   -69690
         TabIndex        =   171
         Top             =   1179
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   3
         Left            =   -69690
         TabIndex        =   45
         Top             =   1467
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM97 
         Height          =   285
         Left            =   -69690
         TabIndex        =   170
         Top             =   1770
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM98 
         Height          =   285
         Left            =   -69690
         TabIndex        =   169
         Top             =   2058
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM99 
         Height          =   285
         Left            =   -69690
         TabIndex        =   168
         Top             =   2346
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   5
         Left            =   -69690
         TabIndex        =   47
         Top             =   2634
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM103 
         Height          =   285
         Left            =   -69690
         TabIndex        =   167
         Top             =   2937
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM104 
         Height          =   285
         Left            =   -69690
         TabIndex        =   166
         Top             =   3225
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM105 
         Height          =   285
         Left            =   -69690
         TabIndex        =   165
         Top             =   3520
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   315
         Left            =   -73470
         TabIndex        =   33
         Top             =   1320
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;556"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   615
         Left            =   -73560
         TabIndex        =   34
         Top             =   2340
         Width           =   7272
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12827;1085"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM58 
         Height          =   615
         Left            =   -73560
         TabIndex        =   35
         Top             =   3000
         Width           =   7272
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12827;1085"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05_1 
         Height          =   825
         Left            =   1290
         TabIndex        =   20
         Top             =   2760
         Width           =   7395
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "13044;1455"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM07 
         Height          =   285
         Left            =   1290
         TabIndex        =   23
         Top             =   3330
         Width           =   7395
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "13039;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM06 
         Height          =   285
         Left            =   1290
         TabIndex        =   22
         Top             =   3045
         Width           =   7395
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "13039;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05 
         Height          =   285
         Left            =   1290
         TabIndex        =   21
         Top             =   2775
         Width           =   7395
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "13039;503"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_2 
         Height          =   285
         Left            =   1950
         TabIndex        =   163
         Top             =   1820
         Width           =   2400
         VariousPropertyBits=   671105055
         Size            =   "4233;503"
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM78_2 
         Height          =   285
         Left            =   6390
         TabIndex        =   162
         Top             =   1820
         Width           =   2400
         VariousPropertyBits=   671105055
         Size            =   "4233;503"
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM79_2 
         Height          =   285
         Left            =   1950
         TabIndex        =   161
         Top             =   2124
         Width           =   2400
         VariousPropertyBits=   671105055
         Size            =   "4233;503"
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM80_2 
         Height          =   285
         Left            =   6390
         TabIndex        =   160
         Top             =   2124
         Width           =   2400
         VariousPropertyBits=   671105055
         Size            =   "4233;503"
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM81_2 
         Height          =   285
         Left            =   1950
         TabIndex        =   159
         Top             =   2430
         Width           =   2400
         VariousPropertyBits=   671105055
         Size            =   "4233;503"
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_¹q¤l°e¥ó:          (Y: ¬O)"
         Height          =   180
         Left            =   4560
         TabIndex        =   155
         Top             =   2482
         Width           =   2085
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¤u§@®É¼Æ:"
         Height          =   180
         Index           =   12
         Left            =   5115
         TabIndex        =   154
         Top             =   345
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H10"
         Height          =   180
         Index           =   5
         Left            =   -70395
         TabIndex        =   153
         Top             =   1564
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H9"
         Height          =   180
         Index           =   5
         Left            =   -74775
         TabIndex        =   152
         Top             =   1564
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   30
         Left            =   -74490
         TabIndex        =   151
         Top             =   1865
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   29
         Left            =   -74490
         TabIndex        =   150
         Top             =   2166
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   28
         Left            =   -74490
         TabIndex        =   149
         Top             =   2467
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   27
         Left            =   -70110
         TabIndex        =   148
         Top             =   1865
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   26
         Left            =   -70110
         TabIndex        =   147
         Top             =   2166
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   25
         Left            =   -70110
         TabIndex        =   146
         Top             =   2467
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H8"
         Height          =   180
         Index           =   4
         Left            =   -70395
         TabIndex        =   145
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H7"
         Height          =   180
         Index           =   4
         Left            =   -74775
         TabIndex        =   144
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   24
         Left            =   -74490
         TabIndex        =   143
         Top             =   661
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   23
         Left            =   -74490
         TabIndex        =   142
         Top             =   962
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   22
         Left            =   -74490
         TabIndex        =   141
         Top             =   1263
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   21
         Left            =   -70110
         TabIndex        =   140
         Top             =   661
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   20
         Left            =   -70110
         TabIndex        =   139
         Top             =   962
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   19
         Left            =   -70110
         TabIndex        =   138
         Top             =   1263
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H6"
         Height          =   180
         Index           =   3
         Left            =   -70395
         TabIndex        =   137
         Top             =   2664
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H5"
         Height          =   180
         Index           =   3
         Left            =   -74775
         TabIndex        =   136
         Top             =   2655
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   18
         Left            =   -74490
         TabIndex        =   135
         Top             =   2962
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   17
         Left            =   -74490
         TabIndex        =   134
         Top             =   3240
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   16
         Left            =   -74490
         TabIndex        =   133
         Top             =   3532
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   15
         Left            =   -70110
         TabIndex        =   132
         Top             =   2952
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   14
         Left            =   -70110
         TabIndex        =   131
         Top             =   3240
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   13
         Left            =   -70110
         TabIndex        =   130
         Top             =   3532
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H4"
         Height          =   180
         Index           =   1
         Left            =   -70395
         TabIndex        =   129
         Top             =   1512
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H3"
         Height          =   180
         Index           =   2
         Left            =   -74775
         TabIndex        =   128
         Top             =   1512
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   12
         Left            =   -74490
         TabIndex        =   127
         Top             =   1800
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   11
         Left            =   -74490
         TabIndex        =   126
         Top             =   2088
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   10
         Left            =   -74490
         TabIndex        =   125
         Top             =   2376
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   9
         Left            =   -70110
         TabIndex        =   124
         Top             =   1800
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   2
         Left            =   -70110
         TabIndex        =   123
         Top             =   2088
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   1
         Left            =   -70110
         TabIndex        =   122
         Top             =   2376
         Width           =   345
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H5 :"
         Height          =   180
         Left            =   90
         TabIndex        =   121
         Top             =   2482
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H4 :"
         Height          =   180
         Left            =   4530
         TabIndex        =   120
         Top             =   2176
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H3 :"
         Height          =   180
         Left            =   90
         TabIndex        =   119
         Top             =   2176
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H2 :"
         Height          =   180
         Left            =   4530
         TabIndex        =   118
         Top             =   1872
         Width           =   720
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "¥X¦W¥N²z¤H"
         Height          =   180
         Left            =   -74475
         TabIndex        =   117
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "µo¤å³W¶O¡G"
         Height          =   180
         Left            =   2640
         TabIndex        =   115
         Top             =   352
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¯S®í°Ó¼Ð :"
         Height          =   180
         Index           =   5
         Left            =   5730
         TabIndex        =   114
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_¿é¤JD/N :"
         Height          =   180
         Left            =   5280
         TabIndex        =   112
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "(Y:¿é¤J)"
         Height          =   180
         Left            =   7230
         TabIndex        =   111
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   8
         Left            =   -70110
         TabIndex        =   110
         Top             =   1224
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   7
         Left            =   -70110
         TabIndex        =   109
         Top             =   936
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   6
         Left            =   -70110
         TabIndex        =   108
         Top             =   648
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   5
         Left            =   -74490
         TabIndex        =   107
         Top             =   1224
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   4
         Left            =   -74490
         TabIndex        =   106
         Top             =   936
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   3
         Left            =   -74490
         TabIndex        =   105
         Top             =   648
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H1"
         Height          =   180
         Index           =   1
         Left            =   -74775
         TabIndex        =   104
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H2"
         Height          =   180
         Index           =   2
         Left            =   -70395
         TabIndex        =   103
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label42 
         Caption         =   "®×¥ó¦WºÙ :"
         Height          =   255
         Left            =   90
         TabIndex        =   102
         Top             =   2775
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "(N:¤£ºâ)"
         Height          =   255
         Left            =   7800
         TabIndex        =   60
         Top             =   1380
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_¦C¦LÂ½Ä¶¨ç :"
         Height          =   180
         Left            =   5250
         TabIndex        =   100
         Top             =   1264
         Width           =   1350
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "(N:¤£¦L)"
         Height          =   180
         Left            =   7050
         TabIndex        =   99
         Top             =   1264
         Width           =   645
      End
      Begin VB.Line Line2 
         X1              =   -70320
         X2              =   -68250
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label17 
         Caption         =   "¬O§_¸É¥ó(¥i½Æ¿ï) :"
         Height          =   255
         Left            =   90
         TabIndex        =   96
         Top             =   1531
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "°Ó«~Ãþ§O :"
         Height          =   252
         Index           =   14
         Left            =   -74880
         TabIndex        =   95
         Top             =   376
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "°Ó«~²Õ¸s :"
         Height          =   252
         Index           =   13
         Left            =   -74880
         TabIndex        =   94
         Top             =   676
         Width           =   852
      End
      Begin VB.Label Label41 
         Caption         =   "(Y:¶l±H)"
         Height          =   252
         Left            =   -72960
         TabIndex        =   93
         Top             =   976
         Width           =   852
      End
      Begin VB.Label Label40 
         Caption         =   "¬O§_¶l±H¥Ó½Ð :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   92
         Top             =   976
         Width           =   1212
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "µo¤å¤é :"
         Height          =   180
         Left            =   90
         TabIndex        =   73
         Top             =   352
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¶Ê¼f´Á­­ :"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   72
         Top             =   656
         Width           =   810
      End
      Begin VB.Label Label4 
         Caption         =   "¥N²z¤H :"
         Height          =   252
         Left            =   120
         TabIndex        =   71
         Top             =   -360
         Width           =   972
      End
      Begin VB.Line Line1 
         X1              =   2640
         X2              =   2760
         Y1              =   1008
         Y2              =   1008
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "©µ®i«á±M¥Î´Á­­ :"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   70
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "¦C¦L©w½Z :"
         Height          =   180
         Left            =   90
         TabIndex        =   69
         Top             =   1264
         Width           =   810
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:¤£¦L)"
         Height          =   180
         Left            =   2160
         TabIndex        =   68
         Top             =   1264
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H1 :"
         Height          =   180
         Left            =   90
         TabIndex        =   67
         Top             =   1872
         Width           =   720
      End
      Begin VB.Label Label7 
         Caption         =   "®×¥ó¤é¤å¦WºÙ :"
         Height          =   255
         Left            =   90
         TabIndex        =   66
         Top             =   3330
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "®×¥ó­^¤å¦WºÙ :"
         Height          =   255
         Left            =   90
         TabIndex        =   65
         Top             =   3045
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "®×¥ó¤¤¤å¦WºÙ :"
         Height          =   255
         Left            =   90
         TabIndex        =   64
         Top             =   2775
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "°Ó¼ÐºØÃþ :"
         Height          =   180
         Index           =   4
         Left            =   2580
         TabIndex        =   63
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "¥¿°Ó¼Ð¸¹¼Æ:"
         Height          =   255
         Index           =   8
         Left            =   8520
         TabIndex        =   62
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ÂI¼Æ :"
         Height          =   180
         Index           =   10
         Left            =   6660
         TabIndex        =   61
         Top             =   352
         Width           =   450
      End
      Begin VB.Label Label16 
         Caption         =   "¬O§_ºâ®×¥ó¼Æ :"
         Height          =   255
         Left            =   7590
         TabIndex        =   59
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label27 
         Caption         =   "¬d¦W¥»©Ò®×¸¹ :"
         Height          =   255
         Left            =   -71910
         TabIndex        =   58
         Top             =   975
         Width           =   1275
      End
      Begin VB.Label Label28 
         Caption         =   "¶i«×³Æµù :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   57
         Top             =   2340
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "®×¥ó³Æµù :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   56
         Top             =   3000
         Width           =   975
      End
   End
   Begin MSForms.TextBox textTM44 
      Height          =   285
      Left            =   1344
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   1980
      Width           =   2535
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   280
      Left            =   5664
      TabIndex        =   157
      Top             =   1656
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;494"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   280
      Left            =   1344
      TabIndex        =   156
      TabStop         =   0   'False
      Top             =   1656
      Width           =   2535
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4471;494"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "S°Ó«~Ãþ§O¿é¦b""®×¥ó³Æµù""Äæ!!!"
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   10.2
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4680
      TabIndex        =   101
      Top             =   405
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Label Label18 
      Caption         =   "¥N²z¤H :"
      Height          =   280
      Index           =   0
      Left            =   264
      TabIndex        =   97
      Top             =   1980
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "¼f©w¸¹¼Æ :"
      Height          =   280
      Left            =   270
      TabIndex        =   91
      Top             =   1014
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "´¼Åv¤H­û :"
      Height          =   280
      Index           =   11
      Left            =   4704
      TabIndex        =   90
      Top             =   1656
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "©¼©Ò®×¸¹ :"
      Height          =   280
      Index           =   9
      Left            =   4704
      TabIndex        =   89
      Top             =   1980
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "®×¥ó©Ê½è :"
      Height          =   280
      Index           =   6
      Left            =   264
      TabIndex        =   88
      Top             =   1335
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "¦¬¤å¸¹ :"
      Height          =   280
      Index           =   1
      Left            =   264
      TabIndex        =   87
      Top             =   372
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "¥»©Ò®×¸¹ :"
      Height          =   280
      Index           =   0
      Left            =   264
      TabIndex        =   86
      Top             =   693
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "¥Ó½Ð®×¸¹ :"
      Height          =   280
      Left            =   4710
      TabIndex        =   85
      Top             =   1335
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "µoÃÒ¤é :"
      Height          =   280
      Index           =   3
      Left            =   4710
      TabIndex        =   84
      Top             =   1014
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "·~°È°Ï§O :"
      Height          =   280
      Index           =   2
      Left            =   4710
      TabIndex        =   83
      Top             =   693
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "©Ó¿ì¤H :"
      Height          =   280
      Left            =   264
      TabIndex        =   82
      Top             =   1656
      Width           =   852
   End
End
Attribute VB_Name = "frm030202_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/02 §ï¦¨Form2.0 ; textCP14¡BtextCP13¡BtextTM44¡BtextTM23_2¡BtextTM78_2¡BtextTM79_2¡BtextTM80_2¡BtextTM81_2¡BtextCP64¡BtextTM58¡BlstNameAgent¡BCombo2(index)¡BtextTM47~52¡BtextTM94~117
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
' ¦¬¤å¸¹
Dim m_CP09 As String
' ¥Ó½Ð°ê®a
Dim m_TM10 As String
' ®×¥ó©Ê½è¥N¸¹
Dim m_CP10 As String
' ´¼Åv¤H­û
Dim m_CP13 As String
' ©Ó¿ì¤H Add By Sindy 98/03/11
Dim m_CP14 As String
' ­ì±M¥Î´Á­­°_¤é
Dim m_TM21 As String
' ­ì±M¥Î´Á­­¤î¤é
Dim m_TM22 As String
' «Å§iÄæ¦ì¤º®eµ²ºc
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' Àx¦s°Ó¼Ð°ò¥»ÀÉ©ÎªA°È·~°È°ò¥»ÀÉÀÉ®×Äæ¦ìªº¦ê¦C
Dim m_TMSPList() As FIELDITEM
Dim m_TMSPCount As Integer
' Àx¦s®×¥ó¶i«×ÀÉÀÉ®×Äæ¦ìªº¦ê¦C
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer
' Àu¥ýÅvµe­±©Ò¨Ï¥ÎªºÅÜ¼Æ
'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'Dim m_Pa(1 To 4) As String '¥»©Ò®×¸¹
'Dim m_Priority(1 To 3) As String
'910801 Sieg 602
Dim m_CP64 As String
'Add By Cheng 2002/08/23
Dim m_strCust1 As String '¥Ó½Ð¤H1
'add by nickc 2007/01/18
Dim m_strCust2 As String
Dim m_strCust3 As String
Dim m_strCust4 As String
Dim m_strCust5 As String

'Modify By Cheng 2003/03/11
'ª½±µ¥Îµe­±¤WªºÄæ¦ì§PÂ_¬O§_©ñ±ó±M¥ÎÅv
''Add By Cheng 2003/01/13
'Dim m_TM67 As String '©ñ±ó±M¥ÎÅv
'Add By Cheng 2003/01/23
'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'Dim m_blnPriDate As Boolean '¬O§_¦³Àu¥ýÅv
'Add By Cheng 2003/10/06
Public m_blnClkChgButton As Boolean '¬O§_«ö¤UÅÜ§ó¨Æ¶µ«ö¶s
'Add By Cheng 2004/05/17
Dim m_blnOutGoingMsg108 As Boolean
'End
'add by nick 2004/08/13
Dim m_CP84 As String       'µo¤å³W¶O
'add by nickc 2005/11/18
Dim m_TM23 As String
Dim m_TM24 As String
Dim m_tm25 As String
Dim m_tm26 As String
'add by nickc 2006/01/26
Dim m_CP110 As String
'add by nickc 2007/01/16
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
Dim m_TM82 As String
Dim m_TM83 As String
Dim m_TM84 As String
Dim m_TM85 As String
Dim m_TM86 As String
Dim m_TM87 As String
Dim m_TM88 As String
Dim m_TM89 As String
Dim m_TM90 As String
Dim m_TM91 As String
Dim m_TM92 As String
Dim m_TM93 As String
'add by nickc 2008/02/22
Dim m_CP44 As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP09s As String, m_CP123s As String 'Add by Sindy 98/3/24 ¦¬¤å¸¹,¬O§_ºâµo¤å«Ç®×¥ó
Dim m_CP130s As String 'Add by Sindy 2009/4/24 µo¤å-¥DºÞ¾÷Ãö
Dim m_CP07 As String 'Add By Sindy 2012/5/3
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
Dim m_TM09 As String, tmpArr As Variant 'Add By Sindy 2014/2/20
'Add By Sindy 2016/12/6
Dim m_strCE04 As String
Dim m_strCE23CE24CE25 As String
'2016/12/6 END
Dim m_ET03 As String 'Add By Sindy 2022/3/10
Dim strPTM As String, strSPT As String 'Added by Lydia 2023/11/16 ¼È¦s°Ó¼ÐºØÃþ¤Î¯S®í°Ó¼ÐªºCombo.ItemData

Private Sub cmdCancel_Click()
   frm030202_01.Show
   Unload Me
End Sub

' ®×¥ó¶i«×¬d¸ß
Private Sub cmdCaseProgress_Click()
   frm030202_04.SetData 0, m_TM01, True
   frm030202_04.SetData 1, m_TM02, False
   frm030202_04.SetData 2, m_TM03, False
   frm030202_04.SetData 3, m_TM04, False
   frm030202_04.SetData 4, m_CP09, False
   frm030202_04.SetParent Me
   Me.Hide
   frm030202_04.Show
   frm030202_04.QueryData
End Sub

Private Sub cmdExit_Click()
   Unload frm030202_01
   Unload Me
   'frm030202_01.Show
End Sub

' ÅÜ§ó¨Æ¶µ
Private Sub cmdMod_Click()
   frm030202_05.SetData 0, m_TM01, True
   frm030202_05.SetData 1, m_TM02, False
   frm030202_05.SetData 2, m_TM03, False
   frm030202_05.SetData 3, m_TM04, False
   frm030202_05.SetData 4, m_CP09, False
   'Add By Sindy 2009/06/03
   frm030202_05.SetData 5, m_TM23, False
   frm030202_05.SetData 6, m_TM78, False
   frm030202_05.SetData 7, m_TM79, False
   frm030202_05.SetData 8, m_TM80, False
   frm030202_05.SetData 9, m_TM81, False
   If textCP27.Text = "" Then
      frm030202_05.SetData 10, strSrvDate(1), False
   Else
      frm030202_05.SetData 10, DBDATE(Trim(textCP27.Text)), False
   End If
   '2009/06/03 End
   
   'frm030202_05.SetParent Me
   frm030202_05.SetParent "frm030202_03"
   Me.Hide
   frm030202_05.Show
   frm030202_05.QueryData
'    m_blnClkChgButton = True
End Sub

Private Sub cmdOK_Click(Index As Integer)
   'Modify By Sindy 2010/11/19 §â¡u½T©w¡v¤Î¡u¦P®Éµo¤å¡v«ö¶sµ{¦¡½X¦X¨Ö
   Select Case Index
      Case 0, 1
         If CheckDataValid = True Then
            'Add By Cheng 2002/05/23
            '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
            If TxtValidate = False Then Exit Sub
            
            'Add by Sindy 98/3/24 ³]©w¬O§_ºâµo¤å«Ç®×¥ó
            If m_TM10 = "000" Then
               'Modify By Sindy 2012/12/20 ­Y¬°¹q¤l°e¥ó«h¤£¸gµo¤å«Ç
               'Modify By Sindy 2023/8/1 ¹q¤l°e¥óÄæ¦ì­È¤£¬OªÅ¥ÕªÌ,§Y¬°¹q¤l°e¥ó
               If (textCP118.Visible = True And textCP118 <> "") Then
                  'Added by Morgan 2016/5/16 ¹q¤l°e¥ó¤]­n°O¿ý¥DºÞ¾÷Ãö
                  If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27, , True) = False Then
                     Exit Sub
                  End If
                  'end 2016/5/16
               
                  'Add By Sindy 2013/8/26
                  strExc(0) = Trim(InputBox("½Ð¿é¤J´¼¼z§½¦¬¤å¤å¸¹!!"))
                  If strExc(0) = "" Then
                     Exit Sub
                  Else
                     textCP64 = "´¼¼z§½¦¬¤å¤å¸¹:" & strExc(0) & ";" & Trim(textCP64)
                  End If
                  '2013/8/26 END
               Else
                  'Add by Sindy 2009/4/24
                  If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
                     Exit Sub
                  Else
                     If m_CP123s = "Y" Then
                        'modify by sonia 2014/6/23 ¥[¶Çµo¤å³W¶O, P-108903
                        If ModifyDispatch(textCP09, m_CP09s, m_CP123s, textCP84, textCP27) = False Then
                            Exit Sub
                        End If
                     End If
                  End If
               End If '2012/12/20 End
            End If
            
            ' ³]©w·Æ¹«´å¼Ð¬°µ¥«Ýª¬ºA
            Screen.MousePointer = vbHourglass
            ' §ó·sÄæ¦ì¿é¤Jªº¤º®e
            OnUpdateField
            ' ¦sÀÉ
            'edit by  nick 2004/11/03
            'OnSaveData
            If OnSaveData = False Then MsgBox "¦sÀÉ¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
            ' ³]©w·Æ¹«´å¼Ð¬°¹w³]
            Screen.MousePointer = vbDefault
            
            'Add By Sindy 2012/4/5 CFT,FCT©Ò¦³®×¥ó©Ê½èµo¤å®É,ÀË¬d¥Nªí¹Ï¬O§_¦s¦b
            'Mark by Amy 2018/07/31 ¦]ChkIsExistImg¤£¨Ï¥Î,»PSindy½T»{FCT¤£¼uMsg¬G®³±¼
            'Call ChkIsExistImg(m_TM01, m_TM02, m_TM03, m_TM04)
            
            'Added by Lydia 2018/07/19 FCTµo¤å¦Û°Ê±N¤U¸üªºPDFÀÉ,¤W¶Ç¨ì¨÷©v°Ï
            If Pub_AutoSavePdf_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_CP10) = False Then
            End If
            'end 2018/07/19
      
            'Added by Lydia 2018/07/19
            
            '************   90.11.23 nick   ²Mµe­±
            'frm030202_01.radio(0).Value = True
            'frm030202_01.textCP09.Enabled = True
            'frm030202_01.textCP09.Text = ""
            'frm030202_01.textTM01.Enabled = False
            'frm030202_01.textTM01.Text = ""
            'frm030202_01.textTM02.Enabled = False
            'frm030202_01.textTM02.Text = ""
            'frm030202_01.textTM02_2.Enabled = False
            'frm030202_01.textTM02_2.Text = ""
            'frm030202_01.textTM03.Enabled = False
            'frm030202_01.textTM03.Text = ""
            'frm030202_01.textTM04.Enabled = False
            'frm030202_01.textTM04.Text = ""
            'frm030202_01.grdList.Clear
            'frm030202_01.grdList.Rows = 2
            'frm030202_01.QueryData
            'frm030202_01.Show
            '*************************************
            'Ken 91.04.09 -- Start
            If textDN = "Y" Then
              'Add By Cheng 2003/03/19
              '·s¼W¦a§}±ø¦Cªí¸ê®Æ
              'edit by nick 2004/09/03 ¥Ó½Ð¤]¤£¦L¦W±ø
      'edit by nick 2004/11/17  ¦]¬°½Ð´Ú¤w¸g¦³²£¥Í¤F
      '         If m_CP10 <> "101" Then
      '            pub_AddressListSN = pub_AddressListSN + 1
      '            PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
      '         End If
               Screen.MousePointer = vbHourglass
               Frmacc21h0.Show
               mdiMain.ToolShow
               mdiMain.tool1_enabled
               Screen.MousePointer = vbDefault
               Set Frmacc21h0.frmlink = frm030202_01
               'add by nick 2004/11/24
               Frmacc21h0.IsPrintAddress = False
'            Else
'               'Add By Cheng 2002/04/30
'               '­Y¦³¥¼µo¤å¸ê®ÆÅã¥ÜÄµ§i
'               PUB_GetCPunIssueDatas "" & Me.textTMKey.Text
'
'               frm030202_01.Show
'               ' 90.12.07 modify by louis
'               frm030202_01.Clear1
            End If
            'Ken 91.04.09 -- End
            
            'Add By Sindy 2013/11/19 FCT·s¥Ó½Ð®×¥B¬°¹q¤l°e¥ó®É,±a¥X¥Ó½Ð®×¸¹¿é¤J§@·~
            If m_CP10 = "101" And (textCP118.Visible = True And textCP118 = "Y") Then
               ' ¥»©Ò®×¸¹
               frm030203_02.SetData 0, m_TM01, True
               frm030203_02.SetData 1, m_TM02, False
               frm030203_02.SetData 2, m_TM03, False
               frm030203_02.SetData 3, m_TM04, False
               ' Á`¦¬¤å¸¹
               frm030203_02.SetData 4, m_CP09, False
               Me.Hide
               frm030203_02.QueryData
               frm030203_02.Show vbModal
               Unload frm030203_01
            End If
            '2013/11/19 END
            
            Call PUB_FCTSendRecvMail(m_CP09) 'Add By Sindy 2024/10/30 ¥~°Óµo¤å®É,¼W¥[µoMail³qª¾©Ó¿ì¤H¤Î°Æ¥»µ¹§Pµo¥DºÞ
            'Add By Sindy 2024/8/19
            If frm030202_01.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2024/8/19 End
            If Index = 0 Then '½T©wÁä
               'Ken 91.04.09 -- Start
               If textDN <> "Y" Then
                  'Add By Cheng 2002/04/30
                  '­Y¦³¥¼µo¤å¸ê®ÆÅã¥ÜÄµ§i
                  If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = True Then
                     frm030202_01.Show
                     ' 90.12.07 modify by louis
                     frm030202_01.Clear1
                  Else
                     'Add By Sindy 2024/8/19
                     If frm030202_01.bolIsEMPFlow = True Then
                        Unload frm030202_01
                        frm090202_4.Show
                     Else
                     '2024/8/19 End
                        frm030202_01.Show
                        frm030202_01.Clear1
                     End If
                  End If
               End If
               'Ken 91.04.09 -- End
               Unload Me
            ElseIf Index = 1 Then '¦P®Éµo¤åÁä
               If textDN <> "Y" Then
                  ' ©I¥s²Ä¤@­Óµe­±
                  frm030202_01.SetData 0, m_TM01, True
                  frm030202_01.SetData 1, m_TM02, False
                  frm030202_01.SetData 2, m_TM03, False
                  frm030202_01.SetData 3, m_TM04, False
                  frm030202_01.SetQueryFromTM
                  Unload Me
                  frm030202_01.Show
                  frm030202_01.radio(1).Value = True
                  frm030202_01.radio_Click 1
                  frm030202_01.QueryData
               Else
                  Unload Me
               End If
            End If
         End If
      Case Else
   End Select
End Sub

'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'Private Sub cmdPriority_Click()
'   ' ­×§ïÀu¥ýÅv¸ê®Æ
'   ModifyPriority m_Priority(1), m_Priority(2), m_Priority(3)
'End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

'Private Sub cmdTogether_Click()
'   If CheckDataValid = True Then
'      'Add By Cheng 2002/07/15
'      '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
'      If TxtValidate = False Then Exit Sub
'
'      'Add by Sindy 98/3/24 ³]©w¬O§_ºâµo¤å«Ç®×¥ó
'      If m_TM10 = "000" Then
'         'Add by Sindy 2009/4/24
'         If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
'            Exit Sub
'         Else
'            If m_CP123s = "Y" Then
'               If ModifyDispatch(textCP09, m_CP09s, m_CP123s, textCP27) = False Then
'                   Exit Sub
'               End If
'            End If
'         End If
'      End If
'
'      ' ³]©w·Æ¹«´å¼Ð¬°µ¥«Ýª¬ºA
'      Screen.MousePointer = vbHourglass
'      ' §ó·sÄæ¦ì¿é¤Jªº¤º®e
'      OnUpdateField
'      ' ¦sÀÉ
'      'edit by nick 2004/11/03
'      'OnSaveData
'      If OnSaveData = False Then MsgBox "¦sÀÉ¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
'
'      ' ³]©w·Æ¹«´å¼Ð¬°¹w³]
'      Screen.MousePointer = vbDefault
'
'      ' ©I¥s²Ä¤@­Óµe­±
'      frm030202_01.SetData 0, m_TM01, True
'      frm030202_01.SetData 1, m_TM02, False
'      frm030202_01.SetData 2, m_TM03, False
'      frm030202_01.SetData 3, m_TM04, False
'      frm030202_01.SetQueryFromTM
'      Unload Me
'      frm030202_01.Show
'      frm030202_01.radio(1).Value = True
'      frm030202_01.radio_Click 1
'      frm030202_01.QueryData
'   End If
'End Sub

'Morgan 2003/11/20
Private Sub Combo2_Click(Index As Integer)

   Dim i As Integer, strTmp As String
   
   If (Combo2(Index).Text = "") Then
      'edit by nickc 2007/01/16
      If Index <= 1 Then
        For i = 0 To 2
           Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = ""
        Next i
      'add by nickc 2007/01/16
      Else
        For i = 0 To 2
           Me.Controls("textTM" & Format(94 + i + 3 * Index, "#")).Text = ""
        Next i
      End If
      Exit Sub
   End If
   
   strTmp = Mid(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") + 1, 1)
   strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
   strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & ChgCustomer(Left(Combo2(Index).Text, InStr(Combo2(Index).Text, "-") - 1))
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 ¤£¥Î dll ¤F   = objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'edit by nickc 2007/01/16
      If Index <= 1 Then
        For i = 0 To 2
           If Not IsNull(RsTemp.Fields(i)) Then
              Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = RsTemp.Fields(i)
           'add by nickc 2008/04/08 ­×¥¿¿ù»~
           Else
              Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = ""
           End If
        Next
      'add by nickc 2007/01/16
      Else
        For i = 0 To 2
           If Not IsNull(RsTemp.Fields(i)) Then
              Me.Controls("textTM" & Format(94 + i + 3 * Index, "#")).Text = RsTemp.Fields(i)
           'add by nickc 2008/04/08 ­×¥¿¿ù»~
           Else
              Me.Controls("textTM" & Format(94 + i + 3 * Index, "#")).Text = ""
           End If
        Next
      End If
   End If
End Sub

'Private Sub Form_Activate()
'    'Add By Cheng 2003/10/06
'    '­Y¦³«ö¤UÅÜ§ó¨Æ¶µ«ö¶s, «h­«·sÅª¨ú¸ê®Æ
'    'edit by nickc 2005/08/23
'    'If m_blnClkChgButton = True Then
'    If m_blnClkChgButton = True Or (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
'        pub_ModifyCaseNum = ""
'        QueryData
''        m_blnClkChgButton = False
'    End If
'End Sub

Private Sub Form_Load()
   ' ³]©w±±¨î¶µªº­I´ºÃC¦â
   textTMKey.BackColor = &H8000000F
   textTM08_2.BackColor = &H8000000F
'edit by nick 2004/08/13  ¤£­n¤F
'   textTM12S.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23_2.BackColor = &H8000000F
   textTM44.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   textTM72_2.BackColor = &H8000000F
   
   'add by nickc 2007/01/16
   textTM78_2.BackColor = &H8000000F
   textTM79_2.BackColor = &H8000000F
   textTM80_2.BackColor = &H8000000F
   textTM81_2.BackColor = &H8000000F
   textTM12S.BackColor = &H8000000F
   
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   
   textAdd_2.BackColor = &H8000000F
   
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
   'Add By Cheng 2002/06/05
'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'   frm880002.m_blnAddNew = False
'    m_blnClkChgButton = False
   'Add by nickc 2006/01/26
   '¥xÆW¥[¥X¦W¥N²z¤H²M³æ¨Ñ¤Ä¿ï,­ì¬O§_¥X¦WÄæ¦ì¤£Åã¥Ü
   Text7.Visible = False
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Added by Lydia 2021/09/02 ¦pªG¤@¶}©l±NListBox©Ô¨ì»Ý­nªº¤j¤p¡A¦r«¬·|¦Û°Ê©ñ¤j¡F©Ò¥Hµe­±¹w³]¬°¤@¦C°ª«×¡AForm_Load¤~©ñ¤j¨ì»Ý­nªº¤j¤p
   lstNameAgent.Height = 900
   lstNameAgent.Width = 1300
   
   'Add By Sindy 2021/3/16 FCT·s¥Ó½Ð®×­Y­n­×§ï¥Ó½Ð¤H³£¨ì®×¥ó°ò¥»ÀÉ­×§ï, ¬Gµo¤åµ{¦¡°µ¥H¤U­×§ï:
   '1.5­Ó¥Ó½Ð¤HÄæ¦ìÂê¦í¡F
   '2.µ{¦¡¦sÀÉ®É·|¨Ì¥Ó½Ð¤H­«·s±a¦a§}§ó·s®×¥ó¤§¥Ó½Ð¦a§}¡A½Ð¨ú®ø¡F¦]¬°¦³¨Ç¤é¥»¨Óªº®×¥ó¤£­n­«·s±a«È¤áÀÉ¦a§}¡C
   textTM23.Enabled = False
   textTM78.Enabled = False
   textTM79.Enabled = False
   textTM80.Enabled = False
   textTM81.Enabled = False
   
   Me.SSTab1.Tab = 0 'Added by Lydia 2021/09/03
   
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2012/4/17
   
   ' ²M°£·j´MªºKey
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' ¦¬¤å¸¹
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
      ' ¬d¦WÁ`¦¬¤å¸¹
'      Case 99: textCP09S = strData
   End Select
End Sub

' ²M°£°Ó¼Ð°ò¥»ÀÉÀÉ®×Äæ¦ì¦ê¦C
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' ³]©w°Ó¼Ð°ò¥»ÀÉ©ÎªA°È·~°È°ò¥»ÀÉÄæ¦ì¦ê¦C¤¤ªºÄæ¦ì¤º®e
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

' ³]©w°Ó¼Ð°ò¥»ÀÉ©ÎªA°È·~°È°ò¥»ÀÉÄæ¦ì¦ê¦C¤¤ªºÄæ¦ì¤º®e
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

' ²M°£®×¥ó¶i«×ÀÉÀÉ®×Äæ¦ì¦ê¦C
Private Sub ClearCPFieldList()
   If m_CPCount > 0 Then
      Erase m_CPList
   End If
   m_CPCount = 0
End Sub

' ³]©w®×¥ó¶i«×ÀÉÄæ¦ì¦ê¦C¤¤ªºÄæ¦ì¤º®e
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

' ³]©w®×¥ó¶i«×ÀÉÄæ¦ì¦ê¦C¤¤ªºÄæ¦ì¤º®e
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

' ¨ú±o°Ó¼Ð°ò¥»ÀÉªºÄæ¦ì¤º®e
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   
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
      ' ¼f©w¸¹¼Æ
      If IsNull(rsTmp.Fields("TM15")) = False Then: textTM15 = rsTmp.Fields("TM15")
      ' ¥Ó½Ð¤é
      strTemp = Empty
'edit by nick 2004/08/13  ¤£­n¤F
'      textTM11 = Empty
'      If IsNull(rsTmp.Fields("TM11")) = False Then
'         textTM11 = TAIWANDATE(rsTmp.Fields("TM11"))
'         strTemp = rsTmp.Fields("TM11")
'      End If
'      SetTMSPFieldOldData "TM11", strTemp, 0
      '2013/4/26 add by sonia FCT·s¥Ó½Ð®×µo¤å¤é¦P®É§ó·s¦Ü°ò¥»ÀÉªº¥Ó½Ð¤é,¬G­n¥ý©ñ¤JSetTMSPFieldOldData
      SetTMSPFieldOldData "TM11", "" & rsTmp.Fields("TM11"), 0
      ' ¥Ó½Ð®×¸¹
'edit by nick 2004/08/13  ¤£­n¤F
'      If IsNull(rsTmp.Fields("TM12")) = False Then
'         textTM12S = rsTmp.Fields("TM12")
'         textTM12 = rsTmp.Fields("TM12")
'      End If
'      SetTMSPFieldOldData "TM12", textTM12, 0
      ' µoÃÒ¤é
      If IsNull(rsTmp.Fields("TM20")) = False Then: textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            ' ®×¥ó¤¤¤å¦WºÙ
            textTM05_1 = Empty
            If IsNull(rsTmp.Fields("TM05")) = False Then: textTM05_1 = rsTmp.Fields("TM05")
            SetTMSPFieldOldData "TM05", textTM05_1, 0
        Case Else
            ' ®×¥ó¤¤¤å¦WºÙ
            textTM05 = Empty
            If IsNull(rsTmp.Fields("TM05")) = False Then: textTM05 = rsTmp.Fields("TM05")
            SetTMSPFieldOldData "TM05", textTM05, 0
            ' ®×¥ó­^¤å¦WºÙ
            textTM06 = Empty
            If IsNull(rsTmp.Fields("TM06")) = False Then: textTM06 = rsTmp.Fields("TM06")
            SetTMSPFieldOldData "TM06", textTM06, 0
            ' ®×¥ó¤é¤å¦WºÙ
            textTM07 = Empty
            If IsNull(rsTmp.Fields("TM07")) = False Then: textTM07 = rsTmp.Fields("TM07")
            SetTMSPFieldOldData "TM07", textTM07, 0
        End Select
      ' °Ó¼ÐºØÃþ
      textTM08 = Empty
      If IsNull(rsTmp.Fields("TM08")) = False Then: textTM08 = rsTmp.Fields("TM08")
      SetTMSPFieldOldData "TM08", textTM08, 0
      textTM08_Validate False
      ' °Ó«~Ãþ§O
      textTM09 = Empty
      m_TM09 = "" 'Add By Sindy 2014/2/20
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
         m_TM09 = rsTmp.Fields("TM09") 'Add By Sindy 2014/2/20
      End If
      SetTMSPFieldOldData "TM09", textTM09, 0
      ' ¥Ó½Ð°ê®a
      If IsNull(rsTmp.Fields("TM10")) = False Then: m_TM10 = rsTmp.Fields("TM10")
      ' ­ì±M¥Î´Á­­°_¤é
      If IsNull(rsTmp.Fields("TM21")) = False Then
         m_TM21 = rsTmp.Fields("TM21")
      End If
      ' ­ì±M¥Î´Á­­¤î¤é
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_TM22 = rsTmp.Fields("TM22")
      End If
      ' ¥Ó½Ð¤H
      'add by nickc 2005/11/18
      m_TM23 = ""
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = rsTmp.Fields("TM23")
         textTM23_2 = GetCustomerName(rsTmp.Fields("TM23"), 0)
         'add by nickc 2005/11/18
         m_TM23 = rsTmp.Fields("tm23")
      End If
      SetTMSPFieldOldData "TM23", textTM23, 0
      'Add By Cheng 2002/08/23
      m_strCust1 = "" & Me.textTM23.Text
      
      'add by nickc 2005/11/18
      ' ¤¤¤å¦a§}
      m_TM24 = ""
      If IsNull(rsTmp.Fields("TM24")) = False Then
         m_TM24 = rsTmp.Fields("TM24")
      End If
      SetTMSPFieldOldData "TM24", m_TM24, 0
      ' ­^¤å¦a§}
      m_tm25 = ""
      If IsNull(rsTmp.Fields("TM25")) = False Then
         m_tm25 = rsTmp.Fields("TM25")
      End If
      SetTMSPFieldOldData "TM25", m_tm25, 0
      ' ¤é¤å¦a§}
      m_tm26 = ""
      If IsNull(rsTmp.Fields("TM26")) = False Then
         m_tm26 = rsTmp.Fields("TM26")
      End If
      SetTMSPFieldOldData "TM26", m_tm26, 0
      
      'add by nickc 2007/01/16
      m_TM78 = ""
      If IsNull(rsTmp.Fields("TM78")) = False Then
         textTM78 = rsTmp.Fields("TM78")
         textTM78_2 = GetCustomerName(rsTmp.Fields("TM78"), 0)
         m_TM78 = rsTmp.Fields("tm78")
      End If
      SetTMSPFieldOldData "TM78", textTM78, 0
      m_strCust2 = "" & Me.textTM78.Text
      m_TM79 = ""
      If IsNull(rsTmp.Fields("TM79")) = False Then
         textTM79 = rsTmp.Fields("TM79")
         textTM79_2 = GetCustomerName(rsTmp.Fields("TM79"), 0)
         m_TM79 = rsTmp.Fields("tm79")
      End If
      SetTMSPFieldOldData "TM79", textTM79, 0
      m_strCust3 = "" & Me.textTM79.Text
      m_TM80 = ""
      If IsNull(rsTmp.Fields("TM80")) = False Then
         textTM80 = rsTmp.Fields("TM80")
         textTM80_2 = GetCustomerName(rsTmp.Fields("TM80"), 0)
         m_TM80 = rsTmp.Fields("tm80")
      End If
      SetTMSPFieldOldData "TM80", textTM80, 0
      m_strCust4 = "" & Me.textTM80.Text
      m_TM81 = ""
      If IsNull(rsTmp.Fields("TM81")) = False Then
         textTM81 = rsTmp.Fields("TM81")
         textTM81_2 = GetCustomerName(rsTmp.Fields("TM81"), 0)
         m_TM81 = rsTmp.Fields("tm81")
      End If
      SetTMSPFieldOldData "TM81", textTM81, 0
      m_strCust5 = "" & Me.textTM81.Text
      '¦a§}
      m_TM82 = ""
      If IsNull(rsTmp.Fields("TM82")) = False Then
         m_TM82 = rsTmp.Fields("TM82")
      End If
      SetTMSPFieldOldData "TM82", m_TM82, 0
      m_TM83 = ""
      If IsNull(rsTmp.Fields("TM83")) = False Then
         m_TM83 = rsTmp.Fields("TM83")
      End If
      SetTMSPFieldOldData "TM83", m_TM83, 0
      m_TM84 = ""
      If IsNull(rsTmp.Fields("TM84")) = False Then
         m_TM84 = rsTmp.Fields("TM84")
      End If
      SetTMSPFieldOldData "TM84", m_TM84, 0
      m_TM85 = ""
      If IsNull(rsTmp.Fields("TM85")) = False Then
         m_TM85 = rsTmp.Fields("TM85")
      End If
      SetTMSPFieldOldData "TM85", m_TM85, 0
      m_TM86 = ""
      If IsNull(rsTmp.Fields("TM86")) = False Then
         m_TM86 = rsTmp.Fields("TM86")
      End If
      SetTMSPFieldOldData "TM86", m_TM86, 0
      m_TM87 = ""
      If IsNull(rsTmp.Fields("TM87")) = False Then
         m_TM87 = rsTmp.Fields("TM87")
      End If
      SetTMSPFieldOldData "TM87", m_TM87, 0
      m_TM88 = ""
      If IsNull(rsTmp.Fields("TM88")) = False Then
         m_TM88 = rsTmp.Fields("TM88")
      End If
      SetTMSPFieldOldData "TM88", m_TM88, 0
      m_TM89 = ""
      If IsNull(rsTmp.Fields("TM89")) = False Then
         m_TM89 = rsTmp.Fields("TM89")
      End If
      SetTMSPFieldOldData "TM89", m_TM89, 0
      m_TM90 = ""
      If IsNull(rsTmp.Fields("TM90")) = False Then
         m_TM90 = rsTmp.Fields("TM90")
      End If
      SetTMSPFieldOldData "TM90", m_TM90, 0
      m_TM91 = ""
      If IsNull(rsTmp.Fields("TM91")) = False Then
         m_TM91 = rsTmp.Fields("TM91")
      End If
      SetTMSPFieldOldData "TM91", m_TM91, 0
      m_TM92 = ""
      If IsNull(rsTmp.Fields("TM92")) = False Then
         m_TM92 = rsTmp.Fields("TM92")
      End If
      SetTMSPFieldOldData "TM92", m_TM92, 0
      m_TM93 = ""
      If IsNull(rsTmp.Fields("TM93")) = False Then
         m_TM93 = rsTmp.Fields("TM93")
      End If
      SetTMSPFieldOldData "TM93", m_TM93, 0
      
      ' ¥¿°Ó¼Ð¸¹¼Æ
      textTM27 = Empty
      If IsNull(rsTmp.Fields("TM27")) = False Then: textTM27 = rsTmp.Fields("TM27")
      SetTMSPFieldOldData "TM27", textTM27, 0
      ' °Ó«~²Õ¸s
      textTM32 = Empty
      If IsNull(rsTmp.Fields("TM32")) = False Then: textTM32 = rsTmp.Fields("TM32")
      SetTMSPFieldOldData "TM32", textTM32, 0
      ' FC¥N²z¤H
      textTM44 = Empty
      If IsNull(rsTmp.Fields("TM44")) = False Then: textTM44 = GetFAgentName(rsTmp.Fields("TM44"))
      ' ©¼©Ò®×¸¹
      If IsNull(rsTmp.Fields("TM45")) = False Then: textTM45 = rsTmp.Fields("TM45")
      
      
      'Morgan 2003/11/20
      '¥Nªí¤H
      Dim i As Integer, j As Integer
      For i = 0 To 9 'edit by nickc 2007/01/16  1
         Combo2(i).AddItem ""
      Next
      
      If rsTmp.Fields("TM23").Value <> "" Then
         'edit by nickc 2008/04/08 §ï¦¨  ­^->¤¤->¤é
         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM23").Value)
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM23").Value)
         
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 ¤£¥Î dll ¤F   = objLawDll.ReadRstMsg(intI, strExc(0))
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
      
      'add by nickc 2007/01/18
      If rsTmp.Fields("TM78").Value <> "" Then
         'edit by nickc 2008/04/08 §ï¦¨  ­^->¤¤->¤é
         strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM78").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 ¤£¥Î dll ¤F   = objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(2).AddItem rsTmp.Fields("TM78").Value & "-" & j & strExc(0)
               Combo2(3).AddItem rsTmp.Fields("TM78").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      If rsTmp.Fields("TM79").Value <> "" Then
         'edit by nickc 2008/04/08 §ï¦¨  ­^->¤¤->¤é
         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM79").Value)
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM79").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 ¤£¥Î dll ¤F   = objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(4).AddItem rsTmp.Fields("TM79").Value & "-" & j & strExc(0)
               Combo2(5).AddItem rsTmp.Fields("TM79").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      If rsTmp.Fields("TM80").Value <> "" Then
         'edit by nickc 2008/04/08 §ï¦¨  ­^->¤¤->¤é
         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM80").Value)
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM80").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 ¤£¥Î dll ¤F   = objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(6).AddItem rsTmp.Fields("TM80").Value & "-" & j & strExc(0)
               Combo2(7).AddItem rsTmp.Fields("TM80").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      If rsTmp.Fields("TM81").Value <> "" Then
         'edit by nickc 2008/04/08 §ï¦¨  ­^->¤¤->¤é
         'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM81").Value)
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("TM81").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 ¤£¥Î dll ¤F   = objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(8).AddItem rsTmp.Fields("TM81").Value & "-" & j & strExc(0)
               Combo2(9).AddItem rsTmp.Fields("TM81").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      'Morgan 2003/11/20 -- end
      
      
      ' ¥Nªí¤H1(¤¤)
      textTM47 = Empty
      If IsNull(rsTmp.Fields("TM47")) = False Then: textTM47 = rsTmp.Fields("TM47")
      SetTMSPFieldOldData "TM47", textTM47, 0
      ' ¥Nªí¤H1(­^)
      textTM48 = Empty
      If IsNull(rsTmp.Fields("TM48")) = False Then: textTM48 = rsTmp.Fields("TM48")
      SetTMSPFieldOldData "TM48", textTM48, 0
      ' ¥Nªí¤H1(¤é)
      textTM49 = Empty
      If IsNull(rsTmp.Fields("TM49")) = False Then: textTM49 = rsTmp.Fields("TM49")
      SetTMSPFieldOldData "TM49", textTM49, 0
      ' ¥Nªí¤H2(¤¤)
      textTM50 = Empty
      If IsNull(rsTmp.Fields("TM50")) = False Then: textTM50 = rsTmp.Fields("TM50")
      SetTMSPFieldOldData "TM50", textTM50, 0
      ' ¥Nªí¤H2(­^)
      textTM51 = Empty
      If IsNull(rsTmp.Fields("TM51")) = False Then: textTM51 = rsTmp.Fields("TM51")
      SetTMSPFieldOldData "TM51", textTM51, 0
      ' ¥Nªí¤H2(¤é)
      textTM52 = Empty
      If IsNull(rsTmp.Fields("TM52")) = False Then: textTM52 = rsTmp.Fields("TM52")
      SetTMSPFieldOldData "TM52", textTM52, 0
      'add by nickc 2007/01/18
      textTM94 = Empty
      If IsNull(rsTmp.Fields("TM94")) = False Then: textTM94 = rsTmp.Fields("TM94")
      SetTMSPFieldOldData "TM94", textTM94, 0
      textTM95 = Empty
      If IsNull(rsTmp.Fields("TM95")) = False Then: textTM95 = rsTmp.Fields("TM95")
      SetTMSPFieldOldData "TM95", textTM95, 0
      textTM96 = Empty
      If IsNull(rsTmp.Fields("TM96")) = False Then: textTM96 = rsTmp.Fields("TM96")
      SetTMSPFieldOldData "TM96", textTM96, 0
      textTM97 = Empty
      If IsNull(rsTmp.Fields("TM97")) = False Then: textTM97 = rsTmp.Fields("TM97")
      SetTMSPFieldOldData "TM97", textTM97, 0
      textTM98 = Empty
      If IsNull(rsTmp.Fields("TM98")) = False Then: textTM98 = rsTmp.Fields("TM98")
      SetTMSPFieldOldData "TM98", textTM98, 0
      textTM99 = Empty
      If IsNull(rsTmp.Fields("TM99")) = False Then: textTM99 = rsTmp.Fields("TM99")
      SetTMSPFieldOldData "TM99", textTM99, 0
      textTM100 = Empty
      If IsNull(rsTmp.Fields("TM100")) = False Then: textTM100 = rsTmp.Fields("TM100")
      SetTMSPFieldOldData "TM100", textTM100, 0
      textTM101 = Empty
      If IsNull(rsTmp.Fields("TM101")) = False Then: textTM101 = rsTmp.Fields("TM101")
      SetTMSPFieldOldData "TM101", textTM101, 0
      textTM102 = Empty
      If IsNull(rsTmp.Fields("TM102")) = False Then: textTM102 = rsTmp.Fields("TM102")
      SetTMSPFieldOldData "TM102", textTM102, 0
      textTM103 = Empty
      If IsNull(rsTmp.Fields("TM103")) = False Then: textTM103 = rsTmp.Fields("TM103")
      SetTMSPFieldOldData "TM103", textTM103, 0
      textTM104 = Empty
      If IsNull(rsTmp.Fields("TM104")) = False Then: textTM104 = rsTmp.Fields("TM104")
      SetTMSPFieldOldData "TM104", textTM104, 0
      textTM105 = Empty
      If IsNull(rsTmp.Fields("TM105")) = False Then: textTM105 = rsTmp.Fields("TM105")
      SetTMSPFieldOldData "TM105", textTM105, 0
      TextTM106 = Empty
      If IsNull(rsTmp.Fields("TM106")) = False Then: TextTM106 = rsTmp.Fields("TM106")
      SetTMSPFieldOldData "TM106", TextTM106, 0
      TextTM107 = Empty
      If IsNull(rsTmp.Fields("TM107")) = False Then: TextTM107 = rsTmp.Fields("TM107")
      SetTMSPFieldOldData "TM107", TextTM107, 0
      TextTM108 = Empty
      If IsNull(rsTmp.Fields("TM108")) = False Then: TextTM108 = rsTmp.Fields("TM108")
      SetTMSPFieldOldData "TM108", TextTM108, 0
      TextTM109 = Empty
      If IsNull(rsTmp.Fields("TM109")) = False Then: TextTM109 = rsTmp.Fields("TM109")
      SetTMSPFieldOldData "TM109", TextTM109, 0
      TextTM110 = Empty
      If IsNull(rsTmp.Fields("TM110")) = False Then: TextTM110 = rsTmp.Fields("TM110")
      SetTMSPFieldOldData "TM110", TextTM110, 0
      TextTM111 = Empty
      If IsNull(rsTmp.Fields("TM111")) = False Then: TextTM111 = rsTmp.Fields("TM111")
      SetTMSPFieldOldData "TM111", TextTM111, 0
      TextTM112 = Empty
      If IsNull(rsTmp.Fields("TM112")) = False Then: TextTM112 = rsTmp.Fields("TM112")
      SetTMSPFieldOldData "TM112", TextTM112, 0
      TextTM113 = Empty
      If IsNull(rsTmp.Fields("TM113")) = False Then: TextTM113 = rsTmp.Fields("TM113")
      SetTMSPFieldOldData "TM113", TextTM113, 0
      TextTM114 = Empty
      If IsNull(rsTmp.Fields("TM114")) = False Then: TextTM114 = rsTmp.Fields("TM114")
      SetTMSPFieldOldData "TM114", TextTM114, 0
      TextTM115 = Empty
      If IsNull(rsTmp.Fields("TM115")) = False Then: TextTM115 = rsTmp.Fields("TM115")
      SetTMSPFieldOldData "TM115", TextTM115, 0
      TextTM116 = Empty
      If IsNull(rsTmp.Fields("TM116")) = False Then: TextTM116 = rsTmp.Fields("TM116")
      SetTMSPFieldOldData "TM116", TextTM116, 0
      TextTM117 = Empty
      If IsNull(rsTmp.Fields("TM117")) = False Then: TextTM117 = rsTmp.Fields("TM117")
      SetTMSPFieldOldData "TM117", TextTM117, 0
      
      ' ®×¥ó³Æµù
      textTM58 = Empty
      If IsNull(rsTmp.Fields("TM58")) = False Then: textTM58 = rsTmp.Fields("TM58")
      SetTMSPFieldOldData "TM58", textTM58, 0
      ' ©ñ±ó±M¥ÎÅv
'edit by nick 2004/08/13 ¤£­n¤F
'      textTM67 = Empty
'      If IsNull(rsTmp.Fields("TM67")) = False Then: textTM67 = rsTmp.Fields("TM67")
'      SetTMSPFieldOldData "TM67", textTM67, 0
        'Modify By Cheng 2003/03/11
'      'Add By Cheng 2003/01/13
'      m_TM67 = "" & rsTmp.Fields("TM67")
      ' ¯S®í°Ó¼Ð
      textTM72 = Empty
      If IsNull(rsTmp.Fields("TM72")) = False Then: textTM72 = rsTmp.Fields("TM72")
      SetTMSPFieldOldData "TM72", textTM72, 0
      textTM72_Validate False
      'Added by Lydia 2023/11/16 ¤º¥~°Ó¤§¤À®×¤Î°Ó¼Ð°ò¥»¸ê®ÆºûÅ@¤§°Ó¼ÐºØÃþ¡B¯S®í°Ó¼ÐÄæ¦ì¼W¥[¤U©Ô¥\¯à
      Pub_SetTMcombo "1", cboTM08, textTM08, IIf(m_TM10 <> "000", True, False), strPTM '°Ó¼ÐºØÃþ
      Pub_SetTMcombo "2", cboTM72, textTM72, IIf(m_TM10 <> "000", True, False), strSPT '¯S®í°Ó¼ÐºØÃþ
      'end 2023/11/16
      
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' ¨ú±o®×¥ó¶i«×ÀÉªºÄæ¦ì¤º®e
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strDate As String
   Dim strCP27 As String
   Dim strCP44 As String
   Dim strCP45 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   Dim strTemp As String
   
   ' ¨t²Î¤é
   strDate = DBDATE(SystemDate())
   ' ¦¬¤å¸¹
   textCP09 = m_CP09
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      m_CP44 = CheckStr(rsTmp.Fields("CP44"))
      ' ®×¥ó©Ê½è
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' ·~°È°Ï§O
      
      '910718 Sieg
      If IsNull(rsTmp.Fields("CP12")) = False Then
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' ´¼Åv¤H­û
      m_CP13 = ""
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
         m_CP13 = rsTmp.Fields("CP13")
      End If
      
      'Add By Sindy 98/03/11
      '¤u§@®É¼Æ
      textCP113 = "" & rsTmp.Fields("CP113")
      SetCPFieldOldData "CP113", textCP113, 1
      ' ©Ó¿ì¤H
      m_CP14 = "" & rsTmp.Fields("CP14")
      '98/03/11 End
      
      ' ©Ó¿ì¤H­û
      If IsNull(rsTmp.Fields("CP14")) = False Then: textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      ' µo¤å¤é(¹w³]¬°¨t²Î¤é)
      strCP27 = Empty
      'Modify By Sindy 2010/01/22 §PÂ_µo¤å¤é¬°ªÅ¥Õ¤~¹w³]¬°¨t²Î¤é
      If textCP27 = "" Then
         textCP27 = TAIWANDATE(strDate)
      End If
      If IsNull(rsTmp.Fields("CP27")) = False Then: strCP27 = rsTmp.Fields("CP27")
      SetCPFieldOldData "CP27", strCP27, 1
      
      'Add By Sindy 2012/5/3
      'ªk©w´Á­­
      m_CP07 = Empty
      If IsNull(rsTmp.Fields("CP07")) = False Then: m_CP07 = rsTmp.Fields("CP07")
      '2012/5/3 End
      
      ' ÂI¼Æ
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then: textCP18 = rsTmp.Fields("CP18")
      ' ¬O§_ºâ®×¥ó¼Æ
      textCP26 = Empty
      If IsNull(rsTmp.Fields("CP26")) = False Then: textCP26 = rsTmp.Fields("CP26")
      SetCPFieldOldData "CP26", textCP26, 0
      If m_CP10 = "102" Then
         strTemp = Empty
         If IsNull(rsTmp.Fields("CP53")) = False Then
            strTemp = rsTmp.Fields("CP53")
         End If
         SetCPFieldOldData "CP53", strTemp, 1
         strTemp = Empty
         If IsNull(rsTmp.Fields("CP54")) = False Then
            strTemp = rsTmp.Fields("CP54")
         End If
         SetCPFieldOldData "CP54", strTemp, 1
      End If
      ' ¶i«×³Æµù
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then: textCP64 = rsTmp.Fields("CP64")
      SetCPFieldOldData "CP64", textCP64, 0
      
      'Add By Sindy 2012/12/20
      ' ¬O§_¹q¤l°e¥ó
      textCP118 = Empty
      If IsNull(rsTmp.Fields("CP118")) = False Then
         textCP118 = rsTmp.Fields("CP118")
      End If
      SetCPFieldOldData "CP118", textCP118, 0
      
      'add by nick 2004/08/13 µo¤å³W¶O
      If IsNull(rsTmp.Fields("CP17")) = False And textCP84.Enabled = True Then
          m_CP84 = CheckStr(rsTmp.Fields("CP17"))
      End If
      'Add By Sindy 2012/12/20 ¹q¤l°e¥óµo¤å³W¶O¹w³]¬°©Ó¿ì¤H¤w¿é¤Jªºª÷ÃB
      If rsTmp.Fields("cp118") = "Y" Then
         textCP84 = Val("" & rsTmp.Fields("cp84"))
      End If
      'end 2012/12/20
      
      'add by nickc 2006/02/10
      Text7 = CheckStr(rsTmp.Fields("CP22"))
      SetCPFieldOldData "CP22", Text7, 0
   End If
   'add by nickc 2006/01/26
   'm_CP110 = CheckStr(rsTmp.Fields("cp110"))   '2010/5/6 ADD BY SONIA
   'SetCPFieldOldData "CP110", m_CP110, 0
   'Modify By Sindy 2010/9/20
   If m_CP110 = "" Then m_CP110 = CheckStr(rsTmp.Fields("cp110"))
   SetCPFieldOldData "CP110", CheckStr(rsTmp.Fields("cp110")), 0
   '2010/9/20 End
   
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' Åª¨ú¸ê®Æ®w
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   'Modified by Lydia 2019/11/13 §ó§ïÅÜ¼Æ¦WºÙ
   'Dim strTemp As String
   Dim strNA14 As String
   'add by nickc 2006/01/26
   Dim tm(1 To 4) As String
   ' ¥ý²M°£°Ó¼Ð°ò¥»ÀÉ©ÎªA°È·~°È°ò¥»ÀÉÄæ¦ì¦ê¦C
   ClearTMSPFieldList
   ' ¥ý²M°£®×¥ó¶i«×ÀÉÄæ¦ì¦ê¦C
   ClearCPFieldList
   
   m_TM21 = Empty
   m_TM22 = Empty
      
   ' ¥ý¨ú±o¥»©Ò®×¸¹
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' ¥»©Ò®×¸¹
      If IsNull(rsTmp.Fields("CP01")) = False Then: m_TM01 = rsTmp.Fields("CP01")
      If IsNull(rsTmp.Fields("CP02")) = False Then: m_TM02 = rsTmp.Fields("CP02")
      If IsNull(rsTmp.Fields("CP03")) = False Then: m_TM03 = rsTmp.Fields("CP03")
      If IsNull(rsTmp.Fields("CP04")) = False Then: m_TM04 = rsTmp.Fields("CP04")
   End If
   rsTmp.Close
    'Add By Cheng 2003/11/11
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF", "S"
        Me.Label42.Visible = True
        Me.textTM05_1.Visible = True
        Me.textTM05_1.Enabled = True
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
        Me.Label42.Visible = False
        Me.textTM05_1.Visible = False
        Me.textTM05_1.Enabled = False
        Me.Label9.Visible = True
        Me.textTM05.Visible = True
        Me.textTM05.Enabled = True
        Me.Label8.Visible = True
        Me.textTM06.Visible = True
        Me.textTM06.Enabled = True
        Me.Label7.Visible = True
        Me.textTM07.Visible = True
        Me.textTM07.Enabled = True
    End Select
    
    'End
   ' ¥»©Ò®×¸¹
'   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)
   
   'add by nickc 2006/01/26
   tm(1) = m_TM01
   tm(2) = m_TM02
   tm(3) = m_TM03
   tm(4) = m_TM04
   '2010/5/6 CANCEL BY SONIA ·s¥Ó½Ð®×¹w³]¥X¦W¥N²z¤H,²¾¨ì¤U­±Åª§¹CP¦A°µ
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110
   '2010/5/6 END
   
   ' Åª¨ú°Ó¼Ð°ò¥»ÀÉ
   QueryTradeMark

   ' ¨ú±o®×¥ó¶i«×ÀÉªºÄæ¦ì
   QueryCaseProgress
   
   '2010/5/6 ADD BY SONIA ·s¥Ó½Ð®×¹w³]¥X¦W¥N²z¤H
   'Modified by Lydia 2021/09/02 + Form 2.0 = True
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True
   '2010/5/6 END

   ' ¨ú±o¶Ê¼f´Á­­ªº¤é´Á
   textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
   textCP27.Tag = textCP27.Text 'Added by Lydia 2019/11/08 °O¿ýµo¤å¤é¡A¦³­×§ïµo¤å¤é¦A­«·s­pºâ¶Ê¼f´Á­­
   
'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'   ' ¸É¤å¥ó´Á­­
'   If textPriorityDoc = "N" Then
'      EnableTextBox textAddDate, True
'   Else
'      textAddDate = Empty
'      EnableTextBox textAddDate, False
'   End If
   
   ' ®×¥ó©Ê½è¬°©µ®i®É¤~¥i¿é¤J©µ®i«á±M¥Î´Á­­
   If m_CP10 = "102" Then
      textTM21.BackColor = &H80000005
      textTM21.Locked = False
      textTM21.TabStop = True
      textTM22.BackColor = &H80000005
      textTM22.Locked = False
      textTM22.TabStop = True
      ' ±M¥Î´Á­­°_¤é
      If IsEmptyText(m_TM21) = False Then
         'Memo by Lydia 2019/11/13 FCT®×±M¥Î´Á¶¡°_¬O±q²Ä1¦¸µù¥UÃÒªº´Á¶¡°_ºâ; T®×«h¬O±q©µ®i´Á¶¡°_ºâ(³Q©µ®i´Á¶¡-1¤Ñ)
         textTM21 = TAIWANDATE(m_TM21)
      End If
      ' ±M¥Î´Á­­¤î¤é
      If IsEmptyText(m_TM22) = False Then
         strNA14 = GetNationExtentYear(m_TM10)
        'Modify By Cheng 2003/09/02
'         textTM22 = TAIWANDATE(DateSerial(Val(DBYEAR(m_TM22)) + Val(strNA14), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22))))
         'edit by nickc 2006/03/08 ·í¤é´Á¬°2/28 ®É¡A§PÂ_¦³µL2/29 ­Y¬O¦³¡A¤w 2/29 ¬°·Ç
         'textTM22 = TAIWANDATE(DateAdd("yyyy", Val(strNA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
         'Modified by Lydia 2019/11/13 §ï¥Î¦@¥Î¼Ò²Õ, ²Ä1¦¸±M¥Î´Á¶¡=¤½§i¤é+10¦~-1¤Ñ¡A¤§«á©µ®i102¨S¦³´î¢°¤Ñ¡F»P±M§Q¤£¤@¼Ë
         'If Mid(ChangeWDateStringToWString(DateAdd("yyyy", Val(strNA14), ChangeWStringToWDateString(DBDATE(m_TM22)))), 5) = "0228" Then
         '   If Mid(ChangeWDateStringToWString(DateAdd("d", 1, (DateAdd("yyyy", Val(strNA14), ChangeWStringToWDateString(DBDATE(m_TM22)))))), 5) = "0229" Then
         '       textTM22 = TAIWANDATE(DateAdd("d", 1, (DateAdd("yyyy", Val(strNA14), ChangeWStringToWDateString(DBDATE(m_TM22))))))
         '   Else
         '       textTM22 = TAIWANDATE(DateAdd("yyyy", Val(strNA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
         '   End If
         'Else
         '   textTM22 = TAIWANDATE(DateAdd("yyyy", Val(strNA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
         'End If
         'Modify By Sindy 2022/3/7 + m_TM10 : ©µ®i«á¤§±M¥Î´Á­­¦~«×­Õ¦³2¤ë29¤é®É¡A±M¥Î´Á­­¤î¤éÀ³¬°2¤ë29¤é¡A¦Ó«D¥H¥[10¦~¤§¤è¦¡­pºâ¬°2¤ë28¤é
         textTM22 = TAIWANDATE(PUB_GetEndDate(DBDATE(m_TM22), Val(strNA14), "N", m_TM10))
         'end 2019/11/13
      End If
   Else
      textTM21.BackColor = &H8000000F
      textTM21.Locked = True
      textTM21.TabStop = False
      textTM22.BackColor = &H8000000F
      textTM22.Locked = True
      textTM22.TabStop = False
   End If
   
   ' ¨Ì®×¥ó©Ê½è¤£¦PÅã¥Ü¤£¦P
   Select Case m_CP10
      ' ¥Ó½Ð
        'Modify By Cheng 2003/03/10
'      Case "101": textAdd_2 = "1:©e¥ôª¬ 2:¨Ï¥Î«Å»} 3:Àu¥ýÅvÃÒ©ú 4:©ñ±ó±M¥ÎÅvÁn©ú"
        'Modify By Cheng 2003/12/09
'      Case "101": textAdd_2 = "1:©e¥ôª¬ 2:¨Ï¥Î«Å»} 3:Àu¥ýÅvÃÒ©ú 4:­»´ä¤½¥qµù¥UÃÒ©ú"
      Case "101": textAdd_2 = "1:©e¥ôª¬ 2:Àu¥ýÅvÃÒ©ú 3:¤½¥qµù¥UÃÒ©ú"
      ' ©µ®i
      Case "102": textAdd_2 = "1:©e¥ôª¬ 2:µù¥UÃÒ 3:§ó¦WÃÒ©ú" 'Modify By Sindy 2017/5/9 3:¨Ï¥ÎÃÒ©ú==>3:§ó¦WÃÒ©ú
      ' ¸É´«µoÃÒ®Ñ
      Case "103": textAdd_2 = "1:©e¥ôª¬ 2:µù¥UÃÒ "
      ' ¥Ó½Ð­^¤åÃÒ©ú
      Case "304": textAdd_2 = "1:©e¥ôª¬"
   End Select
   
   'Add By Sindy 2012/12/20 ¥~°Ó000¥xÆW®×©Ò¦³®×¥ó©Ê½è¥[¹q¤l°e¥ó¥\¯à
   If m_TM01 = "FCT" And m_TM10 = "000" Then
      Label43.Visible = True
      textCP118.Visible = True
   Else
      Label43.Visible = False
      textCP118.Visible = False
   End If
   '2012/12/20 End
   
'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'   ' Åª¨úÀu¥ýÅv¸ê®Æ
'   m_Pa(1) = m_TM01
'   m_Pa(2) = m_TM02
'   m_Pa(3) = m_TM03
'   m_Pa(4) = m_TM04
'   objPublicData.ReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
'   '92.10.19 ADD BY SONIA
'   If m_Priority(1) <> "" Then
'      frm880002.m_blnAddNew = True
'   End If
   '92.10.19 END
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030202_03 = Nothing
End Sub

'add by nickc 2006/01/26
'ÀË¬d¨Ã³]©wcp110¸ê®Æ
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/5 ­û¤u½s¸¹¤w¥i«D¼Æ¦r»Ý°µÂà´«
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Lydia 2021/09/02 §ï¼Ò²Õ
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      Text7 = ""
   Else
      Text7 = "N"
      MsgBox "¥¼¤Ä¿ï¥N²z¤H!", vbInformation, "¥²­nÄæ¦ì¡I"
      Cancel = True
   End If

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
      Select Case m_CP10
         ' ¥Ó½Ð
         Case "101":
            Select Case strTemp
               Case "1", "2", "3", "4"
               Case Else
                  Cancel = True
                  strTit = "ÀË®Ö¸ê®Æ"
                  strMsg = "¬O§_¸É¥ó¶µ¥Ø<" & strTemp & ">¤£¥¿½T"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textAdd_GotFocus
                  GoTo EXITSUB
            End Select
         ' ©µ®i
         Case "102":
            Select Case strTemp
               Case "1", "2", "3"
               Case Else
                  Cancel = True
                  strTit = "ÀË®Ö¸ê®Æ"
                  strMsg = "¬O§_¸É¥ó¶µ¥Ø<" & strTemp & ">¤£¥¿½T"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textAdd_GotFocus
                  GoTo EXITSUB
            End Select
         ' ¸É´«µoÃÒ®Ñ
         Case "103":
            Select Case strTemp
               Case "1", "2"
               Case Else
                  Cancel = True
                  strTit = "ÀË®Ö¸ê®Æ"
                  strMsg = "¬O§_¸É¥ó¶µ¥Ø<" & strTemp & ">¤£¥¿½T"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textAdd_GotFocus
                  GoTo EXITSUB
            End Select
         ' ¥Ó½Ð­^¤åÃÒ©ú
         Case "304":
            Select Case strTemp
               Case "1"
               Case Else
                  Cancel = True
                  strTit = "ÀË®Ö¸ê®Æ"
                  strMsg = "¬O§_¸É¥ó¶µ¥Ø<" & strTemp & ">¤£¥¿½T"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textAdd_GotFocus
                  GoTo EXITSUB
            End Select
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
   
EXITSUB:
End Sub

'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'' ¸É¤å¥ó´Á­­
'Private Sub textAddDate_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'
'   If IsEmptyText(textAddDate) = False Then
'      If CheckIsTaiwanDate(textAddDate, False) = False Then
'         Cancel = True
'         strTit = "¸ê®ÆÀË®Ö"
'         strMsg = "½Ð¿é¤J¥¿½Tªº¸É¤å¥ó´Á­­"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textAddDate_GotFocus
'      End If
'   End If
'End Sub

'Modify By Cheng 2002/09/18
'' ¬d¦WÁ`¦¬¤å¸¹
'Private Sub textCP09S_Validate(Cancel As Boolean)
'   Dim rsTmp As New ADODB.Recordset
'   Dim strSQL As String
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP09S) = False Then
'      If textCP09S = m_CP09 Then
'         Cancel = True
'         strTit = "¸ê®ÆÀË®Ö"
'         strMsg = "¬d¦WÁ`¦¬¤å¸¹¤£¥i¬°¥»®×¤§¦¬¤å¸¹"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP09S_GotFocus
'         GoTo EXITSUB
'      End If
'
'      strSQL = "SELECT * FROM CaseProgress " & _
'               "WHERE CP01 = 'S' AND " & _
'                     "CP09 = '" & textCP09S & "' "
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount <= 0 Then
'         rsTmp.Close
'         Cancel = True
'         strTit = "¸ê®ÆÀË®Ö"
'         strMsg = "¬d¦WÁ`¦¬¤å¸¹¸ê®Æ¤£¦s¦b"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP09S_GotFocus
'         GoTo EXITSUB
'      End If
'      rsTmp.Close
'   End If
'EXITSUB:
'   Set rsTmp = Nothing
'End Sub

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
         MsgBox "¬d¦W¥»©Ò®×¸¹ªº¨t²ÎÃþ§OÃþ¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
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
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¬d¦W¥»©Ò®×¸¹¤£¦s¦b"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.textCP09_S.SetFocus
         textCP09_S_GotFocus
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' ¬O§_ºâ®×¥ó¼Æ
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
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎN"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
   End If
End Sub

Private Sub textCP27_LostFocus()
    'Add By Cheng 2003/10/14
    '­Y¦³¿éµo¤å¤é
    'Modified by Lydia 2019/11/08 °O¿ýµo¤å¤é¡A¦³­×§ïµo¤å¤é¦A­«·s­pºâ¶Ê¼f´Á­­
    'If Me.textCP27.Text <> "" Then
    If Me.textCP27.Text <> "" And Me.textCP27.Tag <> Me.textCP27.Text Then
        ' ¨ú±o¶Ê¼f´Á­­ªº¤é´Á
        textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
    End If
    Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 °O¿ýµo¤å¤é¡A¦³­×§ïµo¤å¤é¦A­«·s­pºâ¶Ê¼f´Á­­
End Sub

' µo¤å¤é
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP27) = False Then
      ' µo¤å¤é¤é´Á¤£¥¿½T
      If CheckIsTaiwanDate(textCP27, False) = False Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         strMsg = "½Ð¿é¤J¥¿½Tªºµo¤å¤é"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' µo¤å¤é¤é´Á¤£¥i¶W¹L¨t²Î¤é
      'edit by nick 2004/08/31 ¨t²Î¤é¥[¤@¤Ñ
      'If Val(DBDATE(textCP27)) > Val(DBDATE(SystemDate())) Then
      If Val(DBDATE(textCP27)) > Val(DBDATE(PUB_GetWorkDay(2))) Then
         Cancel = True
         strTit = "¸ê®ÆÀË®Ö"
         'edit by nick 2004/08/31
         'strMsg = "µo¤å¤é¤£¥i¶W¹L¨t²Î¤é"
         strMsg = "µo¤å¤é¤£¥i¶W¹L¨t²Î¤é¥[¤@¤Ñ"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
        'Modify By Cheng 2003/10/14
'      ' ¨ú±o¶Ê¼f´Á­­ªº¤é´Á
'      textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))

      'Added by Lydia 2015/11/24 ºÞ±±©µ®i®×102,µo¤å¤é¤£¥i¤p©ó"©µ®i´Áº¡«e6­Ó¤ë"
      'modify by sonia 2016/7/5 §ï¬°µo¤å¤é¤£¥i¤p©ó"©µ®i´Áº¡«e6­Ó¤ë+1¤Ñ"  T-093656(ªk©w1051224¤£¥i©ó1050624µo¤å)
      'If m_CP10 = "102" And TransDate(textCP27, 2) < CompDate(1, -6, m_CP07) Then
      'Modified by Lydia 2017/06/01 ©µ®i´Áº¡¤é´Á§ï¥Î¼Ò²Õ±±¨î
      'If m_CP10 = "102" And TransDate(textCP27, 2) < CompDate(2, 1, CompDate(1, -6, m_CP07)) Then
      If m_CP10 = "102" And TransDate(textCP27, 2) < PUB_Get102DeadLine("3", m_CP07) Then
          Cancel = True
          strTit = "¸ê®ÆÀË®Ö"
          strMsg = "©µ®i®×µo¤å¤é¤£±o¦­©ó©µ®i´Áº¡«e6­Ó¤ë+1¤Ñ!"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
          textCP27_GotFocus
          GoTo EXITSUB
      End If
      'end 2015/11/24
   End If
EXITSUB:
End Sub

'edit by nickc 2006/01/27
'Private Sub textCP64_2_GotFocus()
'   TextInverse textCP64_2
'End Sub

'add by nick 2004/08/13
Private Sub textCP84_GotFocus()
   InverseTextBox textCP84
End Sub

Private Sub textCP84_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
If IsEmptyText(textCP84) = False Then
    If IsNumeric(textCP84) = False Then
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "½Ð¿é¤J¼Æ¦r"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP84_GotFocus
    Else
        textCP84.Text = Trim(Val(textCP84.Text))
    End If
End If
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

' ¬O§_¶l±H¥Ó½Ð
Private Sub textMail_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textMail) = False Then
      Select Case textMail
         Case " ", "Y":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎY"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textMail_GotFocus
      End Select
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' ¦C¦L©w½Z
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case "N", "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ,N,1,2©Î3D"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'Private Sub textPriorityDoc_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub

'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'' ¬O§_ªþ±aÀu¥ýÅvÃÒ©ú¤å¥ó
'Private Sub textPriorityDoc_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textPriorityDoc) = False Then
'      Select Case textPriorityDoc
'         Case "Y", "N":
'         Case Else
'            Cancel = True
'            strTit = "¸ê®ÆÀË®Ö"
'            strMsg = "¥u¥i¿é¤JY©ÎN"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textPriorityDoc_GotFocus
'      End Select
'   End If
'
'   If textPriorityDoc = "N" Then
'      EnableTextBox textAddDate, True
''      ' ¸É¤å¥ó´Á­­¹w³]¬°µo¤å¤é¥[¤T­Ó¤ë´î¤@¤Ñ
'      If IsEmptyText(textAddDate) = True And IsEmptyText(textCP27) = False Then
'        'Modify By Cheng 2003/09/02
''         textAddDate = TAIWANDATE(DateSerial(Val(DBYEAR(textCP27)), Val(DBMONTH(textCP27)) + 3, Val(DBDAY(textCP27)) - 1))
'        'Modify By Cheng 2004/03/18
'      ' ¸É¤å¥ó´Á­­¹w³]¬°µo¤å¤é¥[¤T­Ó¤ë´î¤­¤Ñ
''         textAddDate = TAIWANDATE(DateAdd("d", -1, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(textCP27)))))
'         textAddDate = TAIWANDATE(DateAdd("d", -5, DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(textCP27)))))
'        'End
'      End If
'   Else
'      textAddDate = Empty
'      EnableTextBox textAddDate, False
'   End If
'End Sub

' §ó·sÄæ¦ìªº¤º®e
Private Sub OnUpdateField()
Dim strCP64 As String
'add by nickc 2007/01/29 move from down
Dim rsTmp As New ADODB.Recordset
   ' Add By Sindy 98/03/11
   SetCPFieldNewData "CP113", textCP113
   ' ¬O§_ºâ®×¥ó¼Æ
   SetCPFieldNewData "CP26", textCP26
   
   'Add By Sindy 2012/12/20
   ' ¬O§_¹q¤l°e¥ó
   SetCPFieldNewData "CP118", textCP118
   
   ' µo¤å¤é
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' ±ÂÅv´Á¶¡(©µ®i´Á¶¡)
   If m_CP10 = "102" Then
      If IsEmptyText(textTM21) = False Then
         SetCPFieldNewData "CP53", DBDATE(textTM21)
      Else
         SetCPFieldNewData "CP53", textTM21
      End If
      If IsEmptyText(textTM22) = False Then
         SetCPFieldNewData "CP54", DBDATE(textTM22)
      Else
         SetCPFieldNewData "CP54", textTM22
      End If
   End If
   ' ¶i«×³Æµù
   '910801 Sieg 602
    strCP64 = Me.textCP64.Text
    'edit by nickc 2006/01/27
'   If textCP64_2 <> "" Then
'      If strCP64 = "" Then
'         strCP64 = textCP64_2
'      Else
'         strCP64 = strCP64 & "," & textCP64_2
'      End If
'   End If
    'Modify By Cheng 2003/09/05
    '¨ú®ø
    'Begin
'    'Add By Cheng 2003/06/16
'    '­Y¦³¿é¤J¬d¦W¥»©Ò®×¸¹
'    If Me.textCP09_S.Text <> "" And Me.textCP09_S1.Text <> "" Then
'        strCP64 = strCP64 & IIf(strCP64 <> "", ",", "") & "­ì¬d¦W¥»©Ò®×¸¹¡G" & Me.textCP09_S.Text & "-" & Me.textCP09_S1.Text & "-" & Left(Me.textCP09_S2.Text & "0", 1) & "-" & Left(Me.textCP09_S3.Text & "00", 2)
'    End If
    'End
   SetCPFieldNewData "CP64", strCP64
   
   'add by nickc 2006/01/27
   SetCPFieldNewData "CP110", m_CP110
   'add by nickc 2006/02/10
   SetCPFieldNewData "CP22", Text7
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        ' °Ó¼Ð¦WºÙ
        SetTMSPFieldNewData "TM05", textTM05_1
    Case Else
        ' °Ó¼Ð¦WºÙ(¤¤)
        SetTMSPFieldNewData "TM05", textTM05
        ' °Ó¼Ð¦WºÙ(­^)
        SetTMSPFieldNewData "TM06", textTM06
        ' °Ó¼Ð¦WºÙ(¤é)
        SetTMSPFieldNewData "TM07", textTM07
    End Select
   ' °Ó¼ÐºØÃþ¥N¸¹
   SetTMSPFieldNewData "TM08", textTM08
   ' ¥Ó½Ð¤é
'edit by nick 2004/08/13 ¤£­n¤F
'   SetTMSPFieldNewData "TM11", DBDATE(textTM11)
   ' ¥Ó½Ð®×¸¹
'edit by nick 2004/08/13 ¤£­n¤F
'   SetTMSPFieldNewData "TM12", textTM12
   ' °Ó«~Ãþ§O
   SetTMSPFieldNewData "TM09", textTM09
   ' ¥Ó½Ð¤H
   If IsEmptyText(textTM23) = False Then
      SetTMSPFieldNewData "TM23", textTM23 & String(9 - Len(textTM23), "0")
   Else
      SetTMSPFieldNewData "TM23", textTM23
   End If
    'add by nickc 2005/11/18 ­Y¦³­×§ï¥Ó½Ð¤H®É¡A­n§ó·s°ò¥»ÀÉªº¥Ó½Ð¦a§}
    'Modify By Sindy 2009/06/22 ¤Z¥Ó½Ð®×®É¡A§¡§ó·s°ò¥»ÀÉªº¥Ó½Ð¦a§}
   'Add By Sindy 2021/3/16 FCT·s¥Ó½Ð®×­Y­n­×§ï¥Ó½Ð¤H³£¨ì®×¥ó°ò¥»ÀÉ­×§ï
   '¬Gµ{¦¡¦sÀÉ®É·|¨Ì¥Ó½Ð¤H­«·s±a¦a§}§ó·s®×¥ó¤§¥Ó½Ð¦a§}¡A½Ð¨ú®ø¡F¦]¬°¦³¨Ç¤é¥»¨Óªº®×¥ó¤£­n­«·s±a«È¤áÀÉ¦a§}¡C
'    If m_CP10 = "101" Or (m_TM23 & String(9 - Len(m_TM23), "0") <> textTM23 & String(9 - Len(textTM23), "0")) Then
'       'edit by nickc 2007/01/29 move to up
'       'Dim rsTmp As New ADODB.Recordset
'       Set rsTmp = New ADODB.Recordset
'       If rsTmp.State = 1 Then rsTmp.Close
'       rsTmp.CursorLocation = adUseClient
'       rsTmp.Open "select * from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM23.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM23.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
'       If rsTmp.RecordCount <> 0 Then
'           SetTMSPFieldNewData "TM24", CheckStr(rsTmp.Fields("cu23"))
'           SetTMSPFieldNewData "TM25", CheckStr(rsTmp.Fields("CU24")) & IIf(CheckStr(rsTmp.Fields("cu25")) <> "", " " & CheckStr(rsTmp.Fields("cu25")), "") & IIf(CheckStr(rsTmp.Fields("cu26")) <> "", " " & CheckStr(rsTmp.Fields("cu26")), "") & IIf(CheckStr(rsTmp.Fields("cu27")) <> "", " " & CheckStr(rsTmp.Fields("cu27")), "") & IIf(CheckStr(rsTmp.Fields("cu28")) <> "", " " & CheckStr(rsTmp.Fields("cu28")), "")
'           SetTMSPFieldNewData "TM26", CheckStr(rsTmp.Fields("cu29"))
'       End If
'    Else
'       SetTMSPFieldNewData "TM24", m_TM24
'       SetTMSPFieldNewData "TM25", m_tm25
'       SetTMSPFieldNewData "TM26", m_tm26
'    End If
   ' ¥¿°Ó¼Ð¸¹¼Æ
   SetTMSPFieldNewData "TM27", textTM27
   ' °Ó«~²Õ¸s
   SetTMSPFieldNewData "TM32", textTM32
   ' ¥Nªí¤H1(¤¤)
   SetTMSPFieldNewData "TM47", textTM47
   ' ¥Nªí¤H1(­^)
   SetTMSPFieldNewData "TM48", textTM48
   ' ¥Nªí¤H1(¤é)
   SetTMSPFieldNewData "TM49", textTM49
   ' ¥Nªí¤H2(¤¤)
   SetTMSPFieldNewData "TM50", textTM50
   ' ¥Nªí¤H2(­^)
   SetTMSPFieldNewData "TM51", textTM51
   ' ¥Nªí¤H2(¤é)
   SetTMSPFieldNewData "TM52", textTM52
   
   'add by nickc 2007/01/24
   SetTMSPFieldNewData "TM94", textTM94
   SetTMSPFieldNewData "TM95", textTM95
   SetTMSPFieldNewData "TM96", textTM96
   SetTMSPFieldNewData "TM97", textTM97
   SetTMSPFieldNewData "TM98", textTM98
   SetTMSPFieldNewData "TM99", textTM99
   SetTMSPFieldNewData "TM100", textTM100
   SetTMSPFieldNewData "TM101", textTM101
   SetTMSPFieldNewData "TM102", textTM102
   SetTMSPFieldNewData "TM103", textTM103
   SetTMSPFieldNewData "TM104", textTM104
   SetTMSPFieldNewData "TM105", textTM105
   SetTMSPFieldNewData "TM106", TextTM106
   SetTMSPFieldNewData "TM107", TextTM107
   SetTMSPFieldNewData "TM108", TextTM108
   SetTMSPFieldNewData "TM109", TextTM109
   SetTMSPFieldNewData "TM110", TextTM110
   SetTMSPFieldNewData "TM111", TextTM111
   SetTMSPFieldNewData "TM112", TextTM112
   SetTMSPFieldNewData "TM113", TextTM113
   SetTMSPFieldNewData "TM114", TextTM114
   SetTMSPFieldNewData "TM115", TextTM115
   SetTMSPFieldNewData "TM116", TextTM116
   SetTMSPFieldNewData "TM117", TextTM117
   
   ' ®×¥ó³Æµù
   SetTMSPFieldNewData "TM58", textTM58
   ' ©ñ±ó±M¥ÎÅv
'edit by nick 2004/08/13 ¤£­n¤F
'   SetTMSPFieldNewData "TM67", textTM67
   ' ¯S®í°Ó¼Ð
   SetTMSPFieldNewData "TM72", textTM72
   'add by nickc 2007/01/24 ¥Ó½Ð¤H2
   If IsEmptyText(textTM78) = False Then
      SetTMSPFieldNewData "TM78", textTM78 & String(9 - Len(textTM78), "0")
   Else
      SetTMSPFieldNewData "TM78", textTM78
   End If
    'Modify By Sindy 2009/06/22 ¤Z¥Ó½Ð®×®É¡A§¡§ó·s°ò¥»ÀÉªº¥Ó½Ð¦a§}
   'Add By Sindy 2021/3/16 FCT·s¥Ó½Ð®×­Y­n­×§ï¥Ó½Ð¤H³£¨ì®×¥ó°ò¥»ÀÉ­×§ï
   '¬Gµ{¦¡¦sÀÉ®É·|¨Ì¥Ó½Ð¤H­«·s±a¦a§}§ó·s®×¥ó¤§¥Ó½Ð¦a§}¡A½Ð¨ú®ø¡F¦]¬°¦³¨Ç¤é¥»¨Óªº®×¥ó¤£­n­«·s±a«È¤áÀÉ¦a§}¡C
'    If m_CP10 = "101" Or (m_TM78 & String(9 - Len(m_TM78), "0") <> textTM78 & String(9 - Len(textTM78), "0")) Then
'       Set rsTmp = New ADODB.Recordset
'       If rsTmp.State = 1 Then rsTmp.Close
'       rsTmp.CursorLocation = adUseClient
'       rsTmp.Open "select * from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM78.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM78.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
'       If rsTmp.RecordCount <> 0 Then
'           SetTMSPFieldNewData "TM82", CheckStr(rsTmp.Fields("cu23"))
'           SetTMSPFieldNewData "TM86", CheckStr(rsTmp.Fields("CU24")) & IIf(CheckStr(rsTmp.Fields("cu25")) <> "", " " & CheckStr(rsTmp.Fields("cu25")), "") & IIf(CheckStr(rsTmp.Fields("cu26")) <> "", " " & CheckStr(rsTmp.Fields("cu26")), "") & IIf(CheckStr(rsTmp.Fields("cu27")) <> "", " " & CheckStr(rsTmp.Fields("cu27")), "") & IIf(CheckStr(rsTmp.Fields("cu28")) <> "", " " & CheckStr(rsTmp.Fields("cu28")), "")
'           SetTMSPFieldNewData "TM90", CheckStr(rsTmp.Fields("cu29"))
'       End If
'    Else
'       SetTMSPFieldNewData "TM82", m_TM82
'       SetTMSPFieldNewData "TM86", m_TM86
'       SetTMSPFieldNewData "TM90", m_TM90
'    End If
   'add by nickc 2007/01/24 ¥Ó½Ð¤H3
   If IsEmptyText(textTM79) = False Then
      SetTMSPFieldNewData "TM79", textTM79 & String(9 - Len(textTM79), "0")
   Else
      SetTMSPFieldNewData "TM79", textTM79
   End If
    'Modify By Sindy 2009/06/22 ¤Z¥Ó½Ð®×®É¡A§¡§ó·s°ò¥»ÀÉªº¥Ó½Ð¦a§}
   'Add By Sindy 2021/3/16 FCT·s¥Ó½Ð®×­Y­n­×§ï¥Ó½Ð¤H³£¨ì®×¥ó°ò¥»ÀÉ­×§ï
   '¬Gµ{¦¡¦sÀÉ®É·|¨Ì¥Ó½Ð¤H­«·s±a¦a§}§ó·s®×¥ó¤§¥Ó½Ð¦a§}¡A½Ð¨ú®ø¡F¦]¬°¦³¨Ç¤é¥»¨Óªº®×¥ó¤£­n­«·s±a«È¤áÀÉ¦a§}¡C
'    If m_CP10 = "101" Or (m_TM79 & String(9 - Len(m_TM79), "0") <> textTM79 & String(9 - Len(textTM79), "0")) Then
'       Set rsTmp = New ADODB.Recordset
'       If rsTmp.State = 1 Then rsTmp.Close
'       rsTmp.CursorLocation = adUseClient
'       rsTmp.Open "select * from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM79.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM79.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
'       If rsTmp.RecordCount <> 0 Then
'           SetTMSPFieldNewData "TM83", CheckStr(rsTmp.Fields("cu23"))
'           SetTMSPFieldNewData "TM87", CheckStr(rsTmp.Fields("CU24")) & IIf(CheckStr(rsTmp.Fields("cu25")) <> "", " " & CheckStr(rsTmp.Fields("cu25")), "") & IIf(CheckStr(rsTmp.Fields("cu26")) <> "", " " & CheckStr(rsTmp.Fields("cu26")), "") & IIf(CheckStr(rsTmp.Fields("cu27")) <> "", " " & CheckStr(rsTmp.Fields("cu27")), "") & IIf(CheckStr(rsTmp.Fields("cu28")) <> "", " " & CheckStr(rsTmp.Fields("cu28")), "")
'           SetTMSPFieldNewData "TM91", CheckStr(rsTmp.Fields("cu29"))
'       End If
'    Else
'       SetTMSPFieldNewData "TM83", m_TM83
'       SetTMSPFieldNewData "TM87", m_TM87
'       SetTMSPFieldNewData "TM91", m_TM91
'    End If
   'add by nickc 2007/01/24 ¥Ó½Ð¤H4
   If IsEmptyText(textTM80) = False Then
      SetTMSPFieldNewData "TM80", textTM80 & String(9 - Len(textTM80), "0")
   Else
      SetTMSPFieldNewData "TM80", textTM80
   End If
    'Modify By Sindy 2009/06/22 ¤Z¥Ó½Ð®×®É¡A§¡§ó·s°ò¥»ÀÉªº¥Ó½Ð¦a§}
   'Add By Sindy 2021/3/16 FCT·s¥Ó½Ð®×­Y­n­×§ï¥Ó½Ð¤H³£¨ì®×¥ó°ò¥»ÀÉ­×§ï
   '¬Gµ{¦¡¦sÀÉ®É·|¨Ì¥Ó½Ð¤H­«·s±a¦a§}§ó·s®×¥ó¤§¥Ó½Ð¦a§}¡A½Ð¨ú®ø¡F¦]¬°¦³¨Ç¤é¥»¨Óªº®×¥ó¤£­n­«·s±a«È¤áÀÉ¦a§}¡C
'    If m_CP10 = "101" Or (m_TM80 & String(9 - Len(m_TM80), "0") <> textTM80 & String(9 - Len(textTM80), "0")) Then
'       Set rsTmp = New ADODB.Recordset
'       If rsTmp.State = 1 Then rsTmp.Close
'       rsTmp.CursorLocation = adUseClient
'       rsTmp.Open "select * from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM80.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM80.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
'       If rsTmp.RecordCount <> 0 Then
'           SetTMSPFieldNewData "TM84", CheckStr(rsTmp.Fields("cu23"))
'           SetTMSPFieldNewData "TM88", CheckStr(rsTmp.Fields("CU24")) & IIf(CheckStr(rsTmp.Fields("cu25")) <> "", " " & CheckStr(rsTmp.Fields("cu25")), "") & IIf(CheckStr(rsTmp.Fields("cu26")) <> "", " " & CheckStr(rsTmp.Fields("cu26")), "") & IIf(CheckStr(rsTmp.Fields("cu27")) <> "", " " & CheckStr(rsTmp.Fields("cu27")), "") & IIf(CheckStr(rsTmp.Fields("cu28")) <> "", " " & CheckStr(rsTmp.Fields("cu28")), "")
'           SetTMSPFieldNewData "TM92", CheckStr(rsTmp.Fields("cu29"))
'       End If
'    Else
'       SetTMSPFieldNewData "TM84", m_TM84
'       SetTMSPFieldNewData "TM88", m_TM88
'       SetTMSPFieldNewData "TM92", m_TM92
'    End If
   'add by nickc 2007/01/24 ¥Ó½Ð¤H2
   If IsEmptyText(textTM81) = False Then
      SetTMSPFieldNewData "TM81", textTM81 & String(9 - Len(textTM81), "0")
   Else
      SetTMSPFieldNewData "TM81", textTM81
   End If
   'Modify By Sindy 2009/06/22 ¤Z¥Ó½Ð®×®É¡A§¡§ó·s°ò¥»ÀÉªº¥Ó½Ð¦a§}
   'Add By Sindy 2021/3/16 FCT·s¥Ó½Ð®×­Y­n­×§ï¥Ó½Ð¤H³£¨ì®×¥ó°ò¥»ÀÉ­×§ï
   '¬Gµ{¦¡¦sÀÉ®É·|¨Ì¥Ó½Ð¤H­«·s±a¦a§}§ó·s®×¥ó¤§¥Ó½Ð¦a§}¡A½Ð¨ú®ø¡F¦]¬°¦³¨Ç¤é¥»¨Óªº®×¥ó¤£­n­«·s±a«È¤áÀÉ¦a§}¡C
'   If m_CP10 = "101" Or (m_TM81 & String(9 - Len(m_TM81), "0") <> textTM81 & String(9 - Len(textTM81), "0")) Then
'      Set rsTmp = New ADODB.Recordset
'      If rsTmp.State = 1 Then rsTmp.Close
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open "select * from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM81.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM81.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount <> 0 Then
'          SetTMSPFieldNewData "TM85", CheckStr(rsTmp.Fields("cu23"))
'          SetTMSPFieldNewData "TM89", CheckStr(rsTmp.Fields("CU24")) & IIf(CheckStr(rsTmp.Fields("cu25")) <> "", " " & CheckStr(rsTmp.Fields("cu25")), "") & IIf(CheckStr(rsTmp.Fields("cu26")) <> "", " " & CheckStr(rsTmp.Fields("cu26")), "") & IIf(CheckStr(rsTmp.Fields("cu27")) <> "", " " & CheckStr(rsTmp.Fields("cu27")), "") & IIf(CheckStr(rsTmp.Fields("cu28")) <> "", " " & CheckStr(rsTmp.Fields("cu28")), "")
'          SetTMSPFieldNewData "TM93", CheckStr(rsTmp.Fields("cu29"))
'      End If
'   Else
'      SetTMSPFieldNewData "TM85", m_TM85
'      SetTMSPFieldNewData "TM89", m_TM89
'      SetTMSPFieldNewData "TM93", m_TM93
'   End If
   
   '2013/4/26 add by sonia FCT·s¥Ó½Ð®×µo¤å¤é¦P®É§ó·s¦Ü°ò¥»ÀÉªº¥Ó½Ð¤é
   If m_CP10 = "101" Then
      SetTMSPFieldNewData "TM11", DBDATE(textCP27)
   End If
   '2013/4/26
End Sub

' §ó·s°Ó¼Ð°ò¥»ÀÉªº¬ÛÃöÄæ¦ì
Private Sub OnUpdateTradeMark()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
   ' §ó·s®×¥ó¶i«×ÀÉ
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            'Modify By Cheng 2003/01/16
            'Á×§K³æ¤Þ¸¹²£¥Í¿ù»~
'            strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
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
   ' ³]©wSQL»yªk§ó·sªº±ø¥ó
   strSql = strSql & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' °õ¦æSQL«ü¥O
   If bDifference = True Then: cnnConnection.Execute strSql
End Sub

' §ó·s°Ó¼Ð°ò¥»ÀÉªº¬ÛÃöÄæ¦ì
Private Sub OnUpdateCaseProperty()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   
   ' §ó·s®×¥ó¶i«×ÀÉ
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            'Modified by Lydia 2021/09/01 +ChgSQL
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
   ' ³]©wSQL»yªk§ó·sªº±ø¥ó
   strSql = strSql & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
   ' °õ¦æSQL«ü¥O
   If bDifference = True Then: cnnConnection.Execute strSql
   
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
Dim strSql As String
Dim strNP07 As String
Dim strNP08 As String
Dim strNP22 As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCP06 As String
Dim strCP07 As String
Dim i As Integer

 '911107 nick transation
On Error GoTo CheckingErr

cnnConnection.BeginTrans

   ' §ó·s®×¥ó¶i«×ÀÉ
   OnUpdateCaseProperty
    'Add By Cheng 2004/05/17
    '¦P®Éµo¤å¥D±iÀu¥ýÅv¸ê®Æ
    If m_blnOutGoingMsg108 = True Then
        strSql = "Update Caseprogress Set CP27=" & Val(DBDATE(Me.textCP27.Text)) & " Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='108' And CP27 Is Null And CP57 Is Null "
        cnnConnection.Execute strSql
    End If
    'End
   '§ó·s°Ó¼Ð°ò¥»ÀÉ
   OnUpdateTradeMark
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' ­Y¦³¿é¤J¶Ê¼f´Á­­®É, ·s¼W¤@µ§¶Ê¼fªº°O¿ý¨ì¤U¤@µ{§ÇÀÉ
   If IsEmptyText(textUargeDate) = False Then
      'Add By Sindy 2023/5/5 FCT­«·sµo¤å¡A­Y¤U¤@µ{§Ç¤w¦³¸Ó¦¬¤å¸¹¥¼Äò¿ì¤§¶Ê¼f´Á­­¡A«h§ó·s´Á­­§Y¥i¡A¤£­n¥t·s¼W´Á­­
      strExc(0) = "SELECT NP01,NP22 from NextProgress" & _
                  " Where NP01='" & m_CP09 & "' and NP07='305' and NP06 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "UPDATE NextProgress SET NP08=" & PUB_GetWorkDay1(textUargeDate, True) & ",NP09=" & DBDATE(textUargeDate) & _
                  " Where NP01='" & m_CP09 & "' and NP07='305' and NP06 is null"
         cnnConnection.Execute strSql
      Else
      '2023/5/5 END
         strNP07 = "305"
         strNP22 = GetNextProgressNo()
       'Modify By Cheng 2003/09/05
   '      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
   '               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
   '                          DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strUserNum & "'," & strNP22 & ")"
         'Modified by Lydia 2020/07/09 ¥»©Ò´Á­­ÀË¬d¡G­Y¥»©Ò´Á­­«D¤u§@¤Ñ«hª½±µ½Õ¾ã¦Ü³Ìªñªº¤u§@¤Ñ
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                             DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                             PUB_GetWorkDay1(textUargeDate, True) & "," & DBDATE(textUargeDate) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
         cnnConnection.Execute strSql
      End If
      
      ' ©µ®i, ¨Ï¥Î«Å»}, ¥Zµn¼s§i, Ãº¦~¶O, ¶Ê¼f, ´£¥Ó, ¦¬¹F¤£¦L±µ¬¢µ²®×³æ
      Select Case strNP07
         Case "102", "105", "702", "708", "305", "998", "997":
         Case Else:
            'Modify By Cheng 2002/01/15
            '¨ú®ø¥~°ÓFCT¦C¦L±µ¬¢µ²®×³æ
'            ' ¦C¦L°ê¤º®×¥ó±µ¬¢¤Îµ²®×°O¿ý³æ
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
      End Select
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Modify By Cheng 2002/09/18
'   ' ¦³¿é¤J¬d¦WÁ`¦¬¤å¸¹®É, §ó·s¦¹¦¬¤å¸¹¤§¥»©Ò®×¸¹¬°¥»®×¤§¥»©Ò®×¸¹
   ' ¦³¿é¤J¬d¦W¥»©Ò®×¸¹®É, §ó·s¦¹¥»©Ò®×¸¹¸ê®Æ¤§¥»©Ò®×¸¹¬°¥»®×¤§¥»©Ò®×¸¹
'   If textCP09_S.Text = False And IsEmptyText(textCP09_S1.Text) = False Then
   If IsEmptyText(textCP09_S.Text) = False And IsEmptyText(textCP09_S1.Text) = False Then

      'add by nickc 2006/07/18 ²M¥¼µ²¾lªº¥iµ²¾l¤é´Á
      strSql = "UPDATE CaseProgress SET cp109=null " & _
               "WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text) & " and cp59 is null "
      cnnConnection.Execute strSql
      'add by nickc 2006/07/18 ¥[¤J cp31=null
      strSql = "UPDATE CaseProgress SET cp31=null " & _
               "WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text) & " "
      cnnConnection.Execute strSql
      strSql = "UPDATE CaseProgress SET CP01 = '" & m_TM01 & "', CP02 = '" & m_TM02 & "', " & _
                     "CP03 = '" & m_TM03 & "', CP04 = '" & m_TM04 & "', " & _
                     "CP64=CP64||Decode(CP64,Null,'','¡A')||'" & "­ì¬d¦W¥»©Ò®×¸¹¡G" & Me.textCP09_S.Text & "-" & Me.textCP09_S1.Text & "-" & Left(Me.textCP09_S2.Text & "0", 1) & "-" & Left(Me.textCP09_S3.Text & "00", 2) & "' " & _
               " WHERE " & ChgCaseprogress(Me.textCP09_S.Text & Me.textCP09_S1.Text & Me.textCP09_S2.Text & Me.textCP09_S3.Text)
      cnnConnection.Execute strSql
      'Add By Cheng 2003/06/16
      strSql = "Update ServicePractice Set SP18=SP18||Decode(SP18,Null,'','¡A')||'Âà¤J°Ó¼Ð¡G" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' Where " & ChgService(Me.textCP09_S.Text & Me.textCP09_S1.Text & Left(Me.textCP09_S2.Text & "0", 1) & Left(Me.textCP09_S3.Text & "00", 2))
      cnnConnection.Execute strSql

      '2005/4/18 ADD BY SONIA 1~4Äæ­ì¬d¦W¥»©Ò®×¸¹,5~8Äæ·s°Ó¼Ð¥»©Ò®×¸¹
      If PUB_UpdOther(Me.textCP09_S.Text, Me.textCP09_S1.Text, Left(Me.textCP09_S2.Text & "0", 1), Left(Me.textCP09_S3.Text & "00", 2), m_TM01, m_TM02, m_TM03, m_TM04) = False Then
         GoTo CheckingErr
      End If
      '2005/4/18 END
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'   ' ¦³¿é¤J¸É¤å¥ó´Á­­®É, ·s¼W¤@µ§¸É¤å¥óªº°O¿ý¨ì¤U¤@µ{§ÇÀÉ
'   If IsEmptyText(textAddDate) = False Then
'      '92.5.6 MODIFY BY SONIA
'      'strNP07 = "201"
'      strNP07 = "208"
'      '92.5.6 END
'      strNP22 = GetNextProgressNo()
'    'Modify By Cheng 2003/09/02
''      strNP08 = DBDATE(DateSerial(Val(DBYEAR(textAddDate)), Val(DBMONTH(textAddDate)) - 1, Val(DBDAY(textAddDate))))
'        'Modify By Cheng 2004/03/18
'        '¥»©Ò=ªk©w-25¤Ñ
''      strNP08 = DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(DBDATE(textAddDate))))
'      strNP08 = DBDATE(DateAdd("d", -25, ChangeWStringToWDateString(DBDATE(textAddDate))))
'        'End
'      '92.5.6 MODIFY BY SONIA
'      'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22) " & _
'      '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'      '                  strNP08 & "," & DBDATE(textAddDate) & ",'" & m_CP13 & "','" & "¸ÉÀu¥ýÅvÃÒ©ú¤å¥ó" & "'," & strNP22 & ")"
'        'Modify By Cheng 2003/09/05
''      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
''               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
''                        strNP08 & "," & DBDATE(textAddDate) & ",'" & m_CP13 & "'," & strNP22 & ")"
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                        strNP08 & "," & DBDATE(textAddDate) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
'      '92.5.6 END
'      cnnConnection.Execute strSQL
'      ' ©µ®i, ¨Ï¥Î«Å»}, ¥Zµn¼s§i, Ãº¦~¶O, ¶Ê¼f, ´£¥Ó, ¦¬¹F¤£¦L±µ¬¢µ²®×³æ
'      Select Case strNP07
'         Case "102", "105", "702", "708", "305", "998", "997":
'         Case Else:
'            'Modify By Cheng 2002/01/15
'            '¨ú®ø¥~°ÓFCT¦C¦L±µ¬¢µ²®×³æ
''            ' ¦C¦L°ê¤º®×¥ó±µ¬¢¤Îµ²®×°O¿ý³æ
''            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
'      End Select
'   End If
   
'edit by nick 2004/08/13  ²¾¨ì frm030203_02
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ' Àx¦sÀu¥ýÅv¸ê®Æ
''   objPublicData.SavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
'   If objPublicData.SavePriority(m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)) = False Then GoTo CheckingErr
'    'Add By Cheng 2003/11/12
'    '­Y¬°°Ó¥Ó®×¥B¦³Àu¥ýÅv¸ê®Æ, «hºÞ¨î"¥D±iÀu¥ýÅv"(108)ªº´Á­­
'    If m_CP10 = "101" And m_Priority(1) <> "" Then
'        'ªk©w´Á­­
'        strCP07 = DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(Me.textCP27.Text))))
'        '¥»©Ò´Á­­
'        strCP06 = DBDATE(DateAdd("d", -4, ChangeWStringToWDateString(DBDATE(strCP07))))
'        strSQLA = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='108' "
'        rsA.CursorLocation = adUseClient
'        rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'        '­Y¦³¦¬¤å¥D±iÀu¥ýÅv, §ó·s¶i«×ÀÉ
'        If rsA.RecordCount > 0 Then
'            strSQLA = "Update CaseProgress Set CP06=" & strCP06 & ", CP07=" & strCP07 & " Where CP09='" & rsA("CP09").Value & "' "
'            cnnConnection.Execute strSQLA
'        '­Y¥¼¦¬¤å¥D±iÀu¥ýÅv, ·s¼W¤U¤@µ{§ÇÀÉ
'        Else
'            strNP07 = "108"
'            strNP22 = GetNextProgressNo()
'            strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                            "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                            DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
'            cnnConnection.Execute strSQL
'        End If
'        If rsA.State <> adStateClosed Then rsA.Close
'        Set rsA = Nothing
'    End If
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'add by nick 2004/08/13 §ó·s¹ê»Úµo¤å³W¶O
   If textCP84.Enabled = True Then
            strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2012/12/20 ­Y¬°¹q¤l°e¥ó«h¦Û°Ê³]©w¬°¤£¸gµo¤å«Ç
   '¥H¨¾°Ê§@¬°­«·sµo¤å, ©Ò¥H¤@¨Ö§âµo¤å«Ç¬ÛÃöÄæ¦ì²MªÅ
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
      
      'Add By Sindy 2014/2/20 §ó·s°Ó«~ÀÉ¥¼©µ®iµù°O
      If m_CP10 = "102" Then
         '¨Ì°ò¥»ÀÉªº°Ó«~Ãþ§O,³vµ§ÀË¬dÃþ§O¬O§_¤w¦s¦b°Ó«~ÀÉ¸Ì,­Y¨S¦³«h·s¼W¸ê®Æ
         m_TM09 = IIf(Right(m_TM09, 1) = ",", Mid(m_TM09, 1, Len(m_TM09) - 1), m_TM09)
         m_TM09 = IIf(Left(m_TM09, 1) = ",", Mid(m_TM09, 2, Len(m_TM09)), m_TM09)
         tmpArr = Split(m_TM09, ",")
         For i = 0 To UBound(tmpArr)
            strExc(0) = "SELECT tg01 from TMGoods" & _
                        " Where TG01='" & m_TM01 & "' and TG02='" & m_TM02 & "' and TG03='" & m_TM03 & "' and TG04='" & m_TM04 & "'" & _
                        " and TG05='" & tmpArr(i) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               strSql = "insert into TMGoods(tg01,tg02,tg03,tg04,tg05)" & _
                        " values('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & tmpArr(i) & "')"
               cnnConnection.Execute strSql
            End If
         Next i
         '¤£¦s¦bµe­±¤Wªº°Ó«~Ãþ§O,¦b°Ó«~ÀÉ¥²¶·§ó·s¬°¥¼©µ®i
         strSql = "Update TMGoods Set TG18='N'" & _
                  " Where TG01='" & m_TM01 & "' and TG02='" & m_TM02 & "' and TG03='" & m_TM03 & "' and TG04='" & m_TM04 & "'" & _
                  " and TG05 not in('" & Replace(textTM09, ",", "','") & "')"
         cnnConnection.Execute strSql
      End If
      '2014/2/20 END
   End If
   
   'Add By Sindy 2010/7/8 ÀË¬d°Ó«~¸ê®Æ»P°ò¥»ÀÉ°Ó«~Ãþ§O¬O§_¤@­P
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
 '911107 nick transation
  cnnConnection.CommitTrans
   
     'Add by nickc 2008/02/22 ÀË¬d¥N²z¤HEmail(»Ý¦Ò¼{¥i¯à¬°FF®×¥ó)
    PUB_CheckEMail m_CP44, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22
   
   ' ¦C¦L©w½Z
   'edit by nick 2004/09/03 ¥Ó½Ðµo¤å®É¡A¤£¦L¦W±ø¡A¤]¤£¦L©w½Z
   If m_CP10 <> "101" Then
        If textPrint <> "N" Then
           PrintLetter
          
'edit by nick 2004/09/22 §ï¥Ñ¿é d/n ®É¦b¦L
'             'Add By Cheng 2003/02/17
'             '·s¼W¦a§}±ø¦Cªí¸ê®Æ
'             pub_AddressListSN = pub_AddressListSN + 1
'             PUB_AddNewAddressList strUserNum, m_tm01, m_tm02, m_tm03, m_tm04, "" & pub_AddressListSN, "0"
        End If
   End If
 '911107 nick transation
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
    Resume
     cnnConnection.RollbackTrans
     OnSaveData = False
End Function

Private Sub textPrtTrans_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/01/13
    'Âà´«¤j¼g
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 78 Then
        KeyAscii = 0
    End If
End Sub

Private Sub textTM05_1_GotFocus()
    TextInverse Me.textTM05_1
End Sub

' °Ó¼ÐºØÃþ
Private Sub textTM08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   textTM08_2 = Empty
   Cancel = False
   If IsEmptyText(textTM08) = False Then
      textTM08_2 = GetTradeMarkName(textTM08, 0)
      If IsEmptyText(textTM08_2) = True Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "°Ó¼ÐºØÃþ¤£¦s¦b"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   Set rsTmp = Nothing
End Sub

' °Ó«~Ãþ§O
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
   ' µL¸ê®Æ®É¤£°µ¥ô¦óÀË¬d
   If IsEmptyText(textTM09) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textTM09)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM09, nIndex)
      If Len(strTemp) > 6 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "°Ó«~Ãþ§O<" & strTemp & ">¤£¥¿½T"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM09, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTM09, nCount) Then
               Cancel = True
               strTit = "ÀË®Ö¸ê®Æ"
               strMsg = "°Ó«~Ãþ§O<" & strTemp & ">¤£¥i­«ÂÐ"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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

Private Sub textTM23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' ¥Ó½Ð¤H
Private Sub textTM23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM23_2 = Empty
   If IsEmptyText(textTM23) = False Then
      Me.textTM23.Text = ChangeCustomerL(Me.textTM23.Text)
      'edit by 2004/07/22 nick  ÀË¬d¸Ó¥Ó½Ð¤H©Î¥N²z¤Hª¬ºA¡A­Y¬°¤£¦A¨Ï¥Î«h°±¦b­ì¦a
      Dim oState As Boolean
      oState = True
      'textTM23_2 = GetCustomerName(textTM23)
      textTM23_2 = GetCustomerNameAndState(textTM23, "0", oState)
      If oState = False Then
         Cancel = True
         Exit Sub
      End If
      If textTM23_2 = Empty Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¥Ó½Ð¤H¥N½X<" & textTM23 & ">¤£¥¿½T"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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

' °Ó«~²Õ¸s
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
   ' µL¸ê®Æ®É¤£°µ¥ô¦óÀË¬d
   If IsEmptyText(textTM32) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textTM32)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM32, nIndex)
      If Len(strTemp) > 6 Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "°Ó«~²Õ¸s<" & strTemp & ">¤£¥¿½T"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM32, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTM32, nCount) Then
               Cancel = True
               strTit = "ÀË®Ö¸ê®Æ"
               strMsg = "°Ó«~²Õ¸s<" & strTemp & ">¤£¥i­«ÂÐ"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   CheckDataValid = False
   
   'Add By Sindy 2012/4/17
   If m_blnClkChgButton = False Then
      MsgBox "½Ð¿é¤JÅÜ§ó¨Æ¶µ!!!", vbExclamation + vbOKOnly
      Me.cmdMod.SetFocus
      GoTo EXITSUB
   End If
   
   ' ®×¥ó©Ê½è¬°©µ®i®É¥²¶·¿é¤J©µ®i«á±M¥Î´Á¶¡
   If m_CP10 = "102" Then
      If IsEmptyText(textTM21) = True Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "½Ð¿é¤J©µ®i´Á¶¡"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM21.SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(textTM22) = True Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "½Ð¿é¤J©µ®i´Á¶¡"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM22.SetFocus
         GoTo EXITSUB
      End If
      'add by nickc 2008/03/13 ¥~°Óªü½¬½Ð§@³æ¡A©µ®i®É¡Aµo¤å¤é¶·¦b±M¥Î´Áº¡¤é«e¥b¦~¤º
      '2008/3/31 MODIFY BY SONIA ¹L±M¥Î´ÁªÌ¤£¦b¦¹­­,FCT-027811
      'If DBDATE(textCP27) < DBDATE(DateAdd("d", 1, DateAdd("m", -6, ChangeWStringToWDateString(DBDATE(m_TM22))))) Or DBDATE(textCP27) > DBDATE(m_TM22) Then
      '2013/10/31 MODIFY BY SONIA §ï¤î¤é¥ý¥[¤@¤Ñ¦A´î¥b¦~,§_«h4/30´Á­­10/31§Y¥iµo¤å FCT-035360
      'If DBDATE(textCP27) < DBDATE(DateAdd("d", 1, DateAdd("m", -6, ChangeWStringToWDateString(DBDATE(m_TM22))))) Then
      If DBDATE(textCP27) < DBDATE(DateAdd("m", -6, DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22))))) Then
      '2008/3/31 END
         strTit = "ÀË®Ö¸ê®Æ"
         '2013/10/31 MODIFY BY SONIA §ï¤î¤é¥ý¥[¤@¤Ñ¦A´î¥b¦~,§_«h4/30´Á­­10/31§Y¥iµo¤å FCT-035360
         'strMsg = "½Ð¿é¤J¥¿½Tµo¤å®É¶¡¡A©µ®i®×ªº¥i¿ì´Á¶¡¬°±M¥Î´Áº¡¤é«e¥b¦~¤º(" & ChangeWStringToWDateString(DBDATE(DateAdd("d", 1, DateAdd("m", -6, ChangeWStringToWDateString(DBDATE(m_TM22)))))) & "-" & ChangeWStringToWDateString(m_TM22) & ") "
         strMsg = "½Ð¿é¤J¥¿½Tµo¤å®É¶¡¡A©µ®i®×ªº¥i¿ì´Á¶¡¬°±M¥Î´Áº¡¤é«e¥b¦~¤º(" & ChangeWStringToWDateString(DBDATE(DateAdd("m", -6, DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22)))))) & "-" & ChangeWStringToWDateString(m_TM22) & ") "
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27.SetFocus
         GoTo EXITSUB
      End If
      'Add By Sindy 2012/5/3
      '©µ®i«á±M¥Î´Á¤î¤é¤£¥i¤p©óµ¥©ó°ò¥»ÀÉ±M¥Î´Á¤î¤é
      If Val(DBDATE(textTM22)) <= Val(DBDATE(m_TM22)) Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "©µ®i«á±M¥Î´Á¤î¤é¤£¥i¤p©óµ¥©ó°ò¥»ÀÉ±M¥Î´Á¤î¤é¡I"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM22.SetFocus
         GoTo EXITSUB
      End If
      '©µ®i«á±M¥Î´Á¤î¤é¤£¥i¤p©óµ¥©óªk©w´Á­­
      If Val(DBDATE(textTM22)) <= Val(DBDATE(m_CP07)) Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "©µ®i«á±M¥Î´Á¤î¤é¤£¥i¤p©óµ¥©óªk©w´Á­­¡I"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM22.SetFocus
         GoTo EXITSUB
      End If
      '2012/5/3 End
   End If
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF", "S"
        ' ®×¥ó¦WºÙ
        If IsEmptyText(textTM05_1) = True Then
           strTit = "ÀË®Ö¸ê®Æ"
           strMsg = "½Ð¿é¤J®×¥ó¦WºÙ"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textTM05_1.SetFocus
           GoTo EXITSUB
        End If
    Case Else
        ' ®×¥ó¦WºÙ(¤¤, ­^, ¤é)
        If IsEmptyText(textTM05) = True And IsEmptyText(textTM06) = True And IsEmptyText(textTM07) = True Then
           strTit = "ÀË®Ö¸ê®Æ"
           strMsg = "½Ð¿é¤J®×¥ó¦WºÙ"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textTM05.SetFocus
           GoTo EXITSUB
        End If
    End Select
   ' µo¤å¤é
   If IsEmptyText(textCP27) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "½Ð¿é¤Jµo¤å¤é"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   
   'Add By Sindy 2011/01/06
   '¥~°Ó(S)¥Ó½Ð¤H1©ÎFC¥N²z¤H¦Ü¤Ö­n¿é¤J¤@­Ó
   '¨ä¥Lªº¤@©w­n¿é¤J¥Ó½Ð¤H1
   If m_TM01 = "S" Then
        If textTM23 = "" And m_TM44 = "" Then
            MsgBox "¥Ó½Ð¤H1©ÎFC¥N²z¤H¦Ü¤Ö­n¿é¤J¤@­Ó!!!", vbExclamation + vbOKOnly
            Me.textTM23.SetFocus
            textTM23_GotFocus
            GoTo EXITSUB
        End If
   '2011/01/06 End
   Else
      ' ¥Ó½Ð¤H
      If IsEmptyText(textTM23) = True Then
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "½Ð¿é¤J¥Ó½Ð¤H"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM23.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   ' °Ó¼ÐºØÃþ
   If IsEmptyText(textTM08) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "½Ð¿é¤J°Ó¼ÐºØÃþ"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      'Modified by Lydia 2023/11/16
      'textTM08.SetFocus
      cboTM08.SetFocus
      GoTo EXITSUB
   End If
   'edit  by nickc 2007/10/17 ¹ÎÅé¼Ð³¹ÃÒ©ú¼Ð³¹¤£¥Î¿é¤J
    If Me.textTM08.Text <> "8" And Me.textTM08.Text <> "7" Then
        ' °Ó«~Ãþ§O
        If IsEmptyText(textTM09) = True Then
           strTit = "ÀË®Ö¸ê®Æ"
           strMsg = "½Ð¿é¤J°Ó«~Ãþ§O"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textTM09.SetFocus
           GoTo EXITSUB
        End If
   End If
   ' ®×¥ó©Ê½è¬°101®É, °Ó«~²Õ¸s¤@©w­n¿é¤J
   If m_CP10 = "101" Then
'edit by nick 2004/09/09 ²¾¨ì     frm030203_02
'      If IsEmptyText(textTM32) = True Then
'         strTit = "ÀË®Ö¸ê®Æ"
'         strMsg = "½Ð¿é¤J°Ó«~²Õ¸s"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textTM32.SetFocus
'         GoTo EXITSUB
'      End If
      '92.10.19 ADD BY SONIA
'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'      If m_Priority(1) = "" Then
'         If Me.textPriorityDoc.Text <> "" Then
'            MsgBox "µLÀu¥ýÅv¸ê®Æ, ¤£¥i¿é¤J¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó !!!", vbExclamation + vbOKOnly
'            Me.SSTab1.Tab = 0
'            Me.textPriorityDoc.SetFocus
'            GoTo EXITSUB
'         End If
'      Else
'         If Me.textPriorityDoc.Text = "" Then
'            MsgBox "¦³Àu¥ýÅv¸ê®Æ, ½Ð¿é¤J¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó !!!", vbExclamation + vbOKOnly
'            Me.SSTab1.Tab = 0
'            Me.textPriorityDoc.SetFocus
'            GoTo EXITSUB
'         End If
'      End If
      '92.10.19 END
   End If
   ' °Ó¼ÐºØÃþ¬°Áp¦X°Ó¼Ð, ¨¾Å@°Ó¼Ð, Áp¦XªA°È¼Ð³¹, ¨¾Å@ªA°È¼Ð³¹®É¥¿°Ó¼Ð¸¹¼Æ¤£¥iªÅ¥Õ
   If IsEmptyText(textTM27) = True Then
      Select Case textTM08
         Case "2", "3", "5", "6":
            strTit = "ÀË®Ö¸ê®Æ"
            strMsg = "½Ð¿é¤J¥¿°Ó¼Ð¸¹¼Æ"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM08.SetFocus
            GoTo EXITSUB
      End Select
   End If
   'Add By Cheng 2002/06/05
'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'   'ÀË¬d­Y®×¥ó©Ê½è¬°"¥Ó½Ð"(101)®É, ·í¬O§_¥IÀu¥ýÅvÃÒ©ú¤å¥ó¿é¤J"Y"®É, ¤@©w­n¿é¤JÀu¥ýÅv¸ê®Æ
'   If m_CP10 = "101" And Me.textPriorityDoc.Text = "Y" Then
'      If frm880002.m_blnAddNew = False Then
'         MsgBox "½Ð¿é¤JÀu¥ýÅv¸ê®Æ!!!", vbExclamation + vbOKOnly
'         Me.SSTab1.Tab = 1
'         Me.cmdPriority.SetFocus
'         GoTo EXITSUB
'      End If
'   End If
    'Add By Cheng 2004/02/27
'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'    '­Y¬O§_¸É¥ó¦³¿ï¾Ü2¥B¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó¥¼¿é"N"
'    If InStr(Me.textAdd.Text, "2") > 0 And Me.textPriorityDoc.Text <> "N" Then
'        MsgBox "­Y¬O§_¸É¥óÄæ¦ì¦³¿ï¾Ü2, «h¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó¥²¶·¬°N!!!", vbExclamation + vbOKOnly
'        Me.textPriorityDoc.Text = "N"
'        Me.SSTab1.Tab = 1
'        Me.textAdd.SetFocus
'        textAdd_GotFocus
'        GoTo EXITSUB
'    End If
    'End
    '93.8.10 add by sonia
    '­Y¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó¿é"N"¥B¬O§_¸É¥ó¥¼¿ï¾Ü2
'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'    If Me.textPriorityDoc.Text <> "N" And InStr(Me.textAdd.Text, "2") <> 0 Then
'        MsgBox "­Y¬O§_ªþÀu¥ýÅvÃÒ©ú¤å¥ó¬° N ®É, «h¬O§_¸É¥óÄæ¦ì¥²¶·¦³¿ï¾Ü2!!!", vbExclamation + vbOKOnly
'        Me.SSTab1.Tab = 1
'        Me.textAdd.SetFocus
'        textAdd_GotFocus
'        GoTo EXITSUB
'    End If
    '93.8.10 End
    'Add By Cheng 2004/05/17
    'FCT°Ó¥Ó(101)µo¤å®É, ÀË¬d¥»®×¬O§_¦³¦¬¤å¥¼µo¤å¥¼¨ú®ø¦¬¤åªº¥D±iÀu¥ýÅv(108)¸ê®Æ
    m_blnOutGoingMsg108 = False
    If m_TM01 = "FCT" And m_CP10 = "101" Then
        StrSQLa = "Select Count(*) From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='108' And CP27 Is Null And CP57 Is Null "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If Val("" & rsA.Fields(0).Value) > 0 Then
            If MsgBox("¦¹®×¦³" & Val("" & rsA.Fields(0).Value) & "µ§¥D±iÀu¥ýÅv¦¬¤å¸ê®Æ, ½T©w¬O§_¦P®Éµo¤å???", vbExclamation + vbOKCancel) = vbOK Then
                m_blnOutGoingMsg108 = True
            End If
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
    'End
    
    'Added by Lydia 2021/09/02 ÀË¬dµe­±ªº TextBox, ComboBox ¬O§_§t¦³Unicode¤å¦r
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         GoTo EXITSUB
    End If
    
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textTM72_GotFocus()
    TextInverse Me.textTM72
End Sub

Private Sub textTM72_Validate(Cancel As Boolean)
    If Me.textTM72.Text <> "" Then
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            Me.textTM72_2.Text = PUB_GetSpecialPTName("2", Me.textTM72.Text)
            If Me.textTM72_2.Text = "" Then
                MsgBox "¯S®í°Ó¼Ð¥N½X¿é¤J¿ù»~!!!", vbExclamation + vbOKOnly
                Cancel = True
            End If
        Case Else
            Me.textTM72.Text = ""
            Me.textTM72_2.Text = ""
        End Select
    Else
        Me.textTM72.Text = "" 'Added by Lydia 2023/11/16
        Me.textTM72_2.Text = ""
    End If
    If Cancel = True Then TextInverse Me.textTM72
End Sub

'Added by Lydia 2023/11/14
Private Sub textTM72_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'add by nickc 2007/01/18
Private Sub textTM78_GotFocus()
InverseTextBox textTM78
End Sub
Private Sub textTM78_KeyPress(KeyAscii As Integer)
  KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM78_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM78_2 = Empty
   If IsEmptyText(textTM78) = False Then
      Me.textTM78.Text = ChangeCustomerL(Me.textTM78.Text)
      Dim oState As Boolean
      oState = True
      textTM78_2 = GetCustomerNameAndState(textTM78, "0", oState)
      If oState = False Then
         Cancel = True
         Exit Sub
      End If
      If textTM78_2 = Empty Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¥Ó½Ð¤H¥N½X<" & textTM78 & ">¤£¥¿½T"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   If Cancel = False Then
      If Me.textTM78.Text <> m_strCust2 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textTM78_GotFocus
   
End Sub
Private Sub textTM79_GotFocus()
InverseTextBox textTM79
End Sub
Private Sub textTM79_KeyPress(KeyAscii As Integer)
  KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM79_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM79_2 = Empty
   If IsEmptyText(textTM79) = False Then
      Me.textTM79.Text = ChangeCustomerL(Me.textTM79.Text)
      Dim oState As Boolean
      oState = True
      textTM79_2 = GetCustomerNameAndState(textTM79, "0", oState)
      If oState = False Then
         Cancel = True
         Exit Sub
      End If
      If textTM79_2 = Empty Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¥Ó½Ð¤H¥N½X<" & textTM79 & ">¤£¥¿½T"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   If Cancel = False Then
      If Me.textTM79.Text <> m_strCust3 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textTM79_GotFocus
End Sub
Private Sub textTM80_GotFocus()
InverseTextBox textTM80
End Sub
Private Sub textTM80_KeyPress(KeyAscii As Integer)
  KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM80_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM80_2 = Empty
   If IsEmptyText(textTM80) = False Then
      Me.textTM80.Text = ChangeCustomerL(Me.textTM80.Text)
      Dim oState As Boolean
      oState = True
      textTM80_2 = GetCustomerNameAndState(textTM80, "0", oState)
      If oState = False Then
         Cancel = True
         Exit Sub
      End If
      If textTM80_2 = Empty Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¥Ó½Ð¤H¥N½X<" & textTM80 & ">¤£¥¿½T"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   If Cancel = False Then
      If Me.textTM80.Text <> m_strCust4 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textTM80_GotFocus
End Sub
Private Sub textTM81_GotFocus()
InverseTextBox textTM81
End Sub
Private Sub textTM81_KeyPress(KeyAscii As Integer)
  KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub textTM81_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM81_2 = Empty
   If IsEmptyText(textTM81) = False Then
      Me.textTM81.Text = ChangeCustomerL(Me.textTM81.Text)
      Dim oState As Boolean
      oState = True
      textTM81_2 = GetCustomerNameAndState(textTM81, "0", oState)
      If oState = False Then
         Cancel = True
         Exit Sub
      End If
      If textTM81_2 = Empty Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¥Ó½Ð¤H¥N½X<" & textTM81 & ">¤£¥¿½T"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
   If Cancel = False Then
      If Me.textTM81.Text <> m_strCust5 Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then Cancel = True
      End If
   End If
   If Cancel = True Then textTM81_GotFocus
End Sub

' ¶Ê¼f´Á­­
Private Sub textUargeDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textUargeDate) = False Then
      If CheckIsTaiwanDate(textUargeDate, False) = False Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¶Ê¼f´Á­­¤é´Á¤£¥¿½T"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
   End If
End Sub

Private Sub textUargeDate_GotFocus()
   InverseTextBox textUargeDate
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'Private Sub textPriorityDoc_GotFocus()
'   InverseTextBox textPriorityDoc
'End Sub

'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'Private Sub textAddDate_GotFocus()
'   InverseTextBox textAddDate
'End Sub

Private Sub textAdd_GotFocus()
   InverseTextBox textAdd
End Sub

Private Sub textMail_GotFocus()
   InverseTextBox textMail
End Sub

Private Sub textDN_GotFocus()
   InverseTextBox textDN
End Sub
'Modify By Cheng 2002/09/18
'Private Sub textCP09S_GotFocus()
'   InverseTextBox textCP09S
'End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
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

'edit by nick 2004/08/13 ¤£­n¤F
'Private Sub textTM11_GotFocus()
'   InverseTextBox textTM11
'End Sub

'edit by nick ¤£­n¤F
'Private Sub textTM12_GotFocus()
'   InverseTextBox textTM12
'End Sub

Private Sub textTM21_GotFocus()
   InverseTextBox textTM21
End Sub

Private Sub textTM22_GotFocus()
   InverseTextBox textTM22
End Sub

Private Sub textTM23_GotFocus()
   InverseTextBox textTM23
End Sub

Private Sub textTM27_GotFocus()
   InverseTextBox textTM27
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

'edit by nick 2004/08/13 ¤£­n¤F
'Private Sub textTM67_GotFocus()
'   InverseTextBox textTM67
'End Sub

' ¦C¦L©w½Z«e±N¨Ò¥~Äæ¦ì¥[¤J¨ì¦C¦L©w½Z¨Ò¥~Äæ¦ìÀÉ®×¤¤
Private Sub InsExpField()
Dim nCount As Integer
Dim nIndex As Integer
Dim strSql As String
Dim strTemp As String
'Add By Cheng 2003/01/29
Dim strTemp1 As String
'Add By Cheng 2003/02/20
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strDebitNote As String 'Add By Sindy 2017/4/12
   
    'Add  By Cheng 2003/01/23
'edit by nick 2004/08/13 ­ì frm030203_02 ¤w¦³
'    '§PÂ_¬O§_¦³Àu¥ýÅv¸ê®Æ
'    strSQLA = "Select Count(*) From PriDate Where PD01='" & m_TM01 & "' And PD02='" & m_TM02 & "' And PD03='" & m_TM03 & "' And PD04='" & m_TM04 & "' "
'    rsA.CursorLocation = adUseClient
'    rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'    If rsA.Fields(0).Value > 0 Then
'        m_blnPriDate = True
'    Else
'        m_blnPriDate = False
'    End If
'    If rsA.State <> adStateClosed Then rsA.Close
'    Set rsA = Nothing
   ' ¬O§_¸É¥ó
   strTemp = Empty
   ' ¨Ì®×¥ó©Ê½è¤£¦P
   Select Case m_CP10
        'Add By Cheng 2003/01/13
        '¥[®×¥ó©Ê½è¬°¥Ó½Ð
      ' ¥Ó½Ð
      Case "101":
'edit by nick 2004/08/13 ¤w¸g§ï¨ì¥Ó½Ð®×¸¹¦C¦L
'         nCount = GetSubStringCount(textAdd)
'         For nIndex = 1 To nCount
'            'Modify By Cheng 2003/01/29
''            strTemp = GetSubString(textAdd, nIndex)
'            strTemp1 = GetSubString(textAdd, nIndex)
'            'Modify By Cheng 2003/12/09
'            '¨Ï¥Î·s³W©w
''            '­Y¥Ó½Ð¤é¤p©ó20031128
''            If DBDATE(Val(Me.textTM11.Text)) < 20031128 Then
''                Select Case strTemp1
''                   Case "1":
''                      If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
''                        'Modify By Cheng 2003/02/20
''    '                  strTemp = strTemp & "* Power of Attorney."
''                      strTemp = strTemp & "    * Power of Attorney."
''                   Case "2":
''                      If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
''                        'Modify By Cheng 2003/02/20
''    '                  strTemp = strTemp & "* Intend to Use Declaration."
''                      strTemp = strTemp & "    * Intent-to-Use Declaration."
''                   Case "3":
''                      If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
''                        'Modify By Cheng 2003/02/20
''    '                  strTemp = strTemp & "* A Certified copy of the home application (priority document)."
''                      strTemp = strTemp & "    * A certified copy of the home application (priority document) must " & vbCrLf & _
''                                                        "      be submitted to the IPO within three months from the day of filing " & vbCrLf & _
''                                                        "      the subject application."
''                   Case "4":
''                      If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
''                        'Modify By Cheng 2003/02/20
''    '                  strTemp = strTemp & "* Disclaim the excusive right of use of " & m_TM67 & "."
''                        'Modify By Cheng 2003/03/10
''    '                  strTemp = strTemp & "    * Disclaim the excusive right of use of " & m_TM67 & "."
''                      strTemp = strTemp & "    * A certified copy of Certificate of Incorporation from Register of Companies in Hong Kong."
''                End Select
''            '­Y¥Ó½Ð¤é¤j©óµ¥©ó20031128
''            Else
'                Select Case strTemp1
'                   Case "1":
'                      If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
'                      strTemp = strTemp & "    * Power of Attorney."
'                   Case "2":
'                      If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
'                      strTemp = strTemp & "    * A certified copy of the home application (priority document) must " & vbCrLf & _
'                                                        "      be submitted to the IPO within three months from the day of filing " & vbCrLf & _
'                                                        "      the subject application."
'                   Case "3":
'                      If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
'                      strTemp = strTemp & "    * A certified copy of Certificate of Incorporation from Register of Companies."
'                End Select
''            End If
'         Next nIndex
'        'Modify By Cheng 2003/02/20
''         If strTemp <> Empty Then: strTemp = "The remaining documents we need for the renewal application are : " & Chr(13) & Chr(10) & strTemp
'        'Modify By Cheng 2003/02/26
''         If strTemp <> Empty Then: strTemp = "    The remaining documents we need for the referenced application are : " & Chr(13) & Chr(10) & strTemp
'         '92.4.3 MODIFY BY SONIA
'         'If strTemp <> Empty Then: strTemp = vbCrLf & "    The remaining documents we need for the referenced application are : " & Chr(13) & Chr(10) & strTemp
'         If strTemp <> Empty Then: strTemp = vbCrLf & "    The remaining document(s) we need for the referenced application is/are : " & Chr(13) & Chr(10) & strTemp
'         '92.4.3 end
      ' ©µ®i
      Case "102":
         nCount = GetSubStringCount(textAdd)
         For nIndex = 1 To nCount
            'Modify By Cheng 2003/01/29
'            strTemp = GetSubString(textAdd, nIndex)
            strTemp1 = GetSubString(textAdd, nIndex)
            Select Case strTemp1
               Case "1":
                  If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                    'Modify By Cheng 2003/03/12
'                  strTemp = strTemp & "* Power of Attorney."
                  strTemp = strTemp & "    * Power of Attorney."
               Case "2":
                  If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                    'Modify By Cheng 2003/03/12
'                  strTemp = strTemp & "* the original Registration Certificate."
                  strTemp = strTemp & "    * the original Registration Certificate."
               Case "3":
                  If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                    'Modify By Cheng 2003/03/12
'                  strTemp = strTemp & "* proof of use of mark on the designated goods."
                  'Modify By Sindy 2017/5/10
'                  strTemp = strTemp & "    * proof of use of mark on the designated goods/services in Taiwan. If " & vbCrLf & _
'                                                    "      the registrant cannot provide dated evidence of use, a Use" & vbCrLf & _
'                                                    "      Declaration will be accepted by the Authorities as a substitute. A Use " & vbCrLf & _
'                                                    "      Declaration form is enclosed."
                  strTemp = strTemp & "    * An official document evidencing the change of the Registrant's name."
                  '2017/5/10 END
            End Select
         Next nIndex
        'Modify By Cheng 2003/03/12
'         If strTemp <> Empty Then: strTemp = "The remaining documents we need for the renewal application are : " & Chr(13) & Chr(10) & strTemp
         'Modify By Sindy 2017/5/10
         If InStr(textAdd, 3) > 0 Then
            If strTemp <> Empty Then: strTemp = "The remaining document(s) we need for the renewal application and the simultaneous recordal of the change of the Registrant' name are : " & Chr(13) & Chr(10) & strTemp & vbCrLf & vbCrLf & _
            "    We will keep you duly informed of any developments of the referenced case and look forward to receiving the required document(s)."
         Else
         '2017/5/10 END
            If strTemp <> Empty Then
               strTemp = "The remaining document(s) we need for the renewal application are : " & Chr(13) & Chr(10) & strTemp & vbCrLf & vbCrLf & _
               "¡@¡@We will keep you duly informed of further progress of the referenced matter."
            'Modify By Sindy 2018/9/3
            Else
               'strTemp = strTemp & vbCrLf & "    We will keep you duly informed of further progress of the referenced matter."
               strTemp = strTemp & "We will keep you duly informed of further progress of the referenced matter."
            '2018/9/3 END
            End If
         End If
      Case Else:
   End Select
   
   'Add By Sindy 2022/3/10 ³]©w¯S©w¤á¤§¯S§O³qª¾¨ç©w½Z
   If GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "2" And m_CP10 = "102" Then '©µ®i­^¤å©w½Z
      If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, m_ET03, , "01") = True Then
         '©w½Z
         EndLetter "01", m_CP09, m_ET03, strUserNum
         Exit Sub
      End If
   End If
   '2022/3/10 END
      
   Select Case m_CP10
        'Add By Cheng 2003/01/13
      ' ¥Ó½Ð
      Case "101":
'edit by nick 2004/08/13 ¤w¸g§ï¨ì¥Ó½Ð®×¸¹¦C¦L
'            '­Y¦³¿é¤J¥Ó½Ð®×¸¹
'            'edit by nick 2004/08/12 ­×§ï¦¨¤£ºÞ¥Ó½Ð®×¸¹¦³µL¿é¤J¬Ò¥i¥H¦³¨Ò¥~Äæ¦ì
'            'If Me.textTM12.Text <> "" Then
'                ' ©w½Z»y¤å
'                Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
'                ' ¤¤¤å
'                Case "1":
'                   ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                   EndLetter "02", m_CP09, "01", strUserNum
'                ' ­^¤å
'                Case "2":
'                    'Modify By Cheng 2003/02/20
'                    '§PÂ_¬O§_¦³Àu¥ýÅv¥X¤£¦Pªº©w½Z
''                   ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
''                   EndLetter "02", m_CP09, "02", strUserNum
''                   ' ¬O§_¸É¥ó
''                   If IsEmptyText(strTemp) = False Then
''                      ' ¬O§_¸É¥ó
''                      strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                               "VALUES ('" & "02" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
''                               "','¬O§_¸É¥ó','" & strTemp & "')"
''                      cnnConnection.Execute strSQL
''                   End If
''                   ' ¬O§_¦C¦LÂ½Ä¶¨ç
''                   If textPrtTrans <> "N" Then
''                      ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
''                      EndLetter "02", m_CP09, "03", strUserNum
''                   End If
'                'Modify By Cheng 2003/12/09
'                '¨Ï¥Î·s³W©w
''                '­Y¥Ó½Ð¤é¤p©ó20031128
''                If DBDATE(Val(Me.textTM11.Text)) < 20031128 Then
''                       ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
''                       EndLetter "02", m_CP09, IIf(m_blnPriDate, "02", "04"), strUserNum
''                       ' ¬O§_¸É¥ó
''                       If IsEmptyText(strTemp) = False Then
''                          ' ¬O§_¸É¥ó
''                          strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                   "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "02", "04") & "','" & strUserNum & _
''                                   "','¬O§_¸É¥ó','" & strTemp & "')"
''                          cnnConnection.Execute strSQL
''                       End If
''                       ' ¬O§_¦C¦LÂ½Ä¶¨ç
''                       If textPrtTrans <> "N" Then
''                          ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
''    '                      EndLetter "02", m_CP09, "03", strUserNum
''                             'Áp¦X°Ó¼Ð
''                            If Me.textTM08.Text = "2" Or Me.textTM08.Text = "5" Then
''                                EndLetter "02", m_CP09, IIf(m_blnPriDate, "06", "07"), strUserNum
''                                'Add By Cheng 2003/02/26
''                                '­Y¦³©ñ±ó±M¥ÎÅv
''                                'Modify By Cheng 2003/03/11
''    '                            If m_TM67 <> "" Then
''                                If Me.textTM67.Text <> "" Then
''                                    ' ©ñ±ó±M¥ÎÅv
''                                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "06", "07") & "','" & strUserNum & _
''                                             "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & Me.textTM67.Text & "')"
''                                    cnnConnection.Execute strSQL
''                                End If
''                            '«DÁp¦X°Ó¼Ð
''                            Else
''                                  EndLetter "02", m_CP09, IIf(m_blnPriDate, "03", "05"), strUserNum
''                                'Add By Cheng 2003/02/26
''                                '­Y¦³©ñ±ó±M¥ÎÅv
''                                'Modify By Cheng 2003/03/11
''    '                            If m_TM67 <> "" Then
''                                If Me.textTM67.Text <> "" Then
''                                    ' ©ñ±ó±M¥ÎÅv
''                                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "03", "05") & "','" & strUserNum & _
''                                             "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & Me.textTM67.Text & "')"
''                                    cnnConnection.Execute strSQL
''                                End If
''                            End If
''                       End If
''                    '­Y¥Ó½Ð¤é¤j©óµ¥©ó20031128
''                    Else
'                       ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                       'edit by nick 2004/08/13 Àu¥ýÅv¤w¸g²M°£
'                       'EndLetter "02", m_CP09, IIf(m_blnPriDate, "10", "12"), strUserNum
'                       EndLetter "02", m_CP09, "12", strUserNum
'                       ' ¬O§_¸É¥ó
'                       If IsEmptyText(strTemp) = False Then
'                          ' ¬O§_¸É¥ó
'                          'edit by nick 2004/08/13 Àu¥ýÅv¤w¸g²M°£
''                          strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                   "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "10", "12") & "','" & strUserNum & _
''                                   "','¬O§_¸É¥ó','" & strTemp & "')"
'                          strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                   "VALUES ('" & "02" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & _
'                                   "','¬O§_¸É¥ó','" & strTemp & "')"
'                          cnnConnection.Execute strSQL
'                       End If
'                       ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                       If textPrtTrans <> "N" Then
'                          ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                            'edit by nick 2004/08/13 Àu¥ýÅv¤w¸g²M°£
'                            'EndLetter "02", m_CP09, IIf(m_blnPriDate, "11", "13"), strUserNum
'                            EndLetter "02", m_CP09, "13", strUserNum
'                            'Add By Cheng 2003/02/26
'                            '­Y¦³©ñ±ó±M¥ÎÅv
''edit by nick 2004/08/13 ¤£­n¤F
''                            If Me.textTM67.Text <> "" Then
''                                ' ©ñ±ó±M¥ÎÅv
''                                strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                         "VALUES ('" & "02" & "','" & m_CP09 & "','" & IIf(m_blnPriDate, "11", "13") & "','" & strUserNum & _
''                                         "','©ñ±ó±M¥ÎÅv','" & vbCrLf & "The following part disclaimed : " & Me.textTM67.Text & "')"
''                                cnnConnection.Execute strSQL
''                            End If
'                       End If
''                    End If
'                ' ¤é¤å
'                Case "3":
'                   ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                   '93.6.2 MODIFY BY SONIA
'                   'EndLetter "02", m_CP09, "04", strUserNum
'                   EndLetter "02", m_CP09, "08", strUserNum
'                   ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                   If textPrtTrans <> "N" Then
'                      ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
'                        EndLetter "02", m_CP09, "09", strUserNum
'                   End If
'                   '93.6.2 END
'                End Select
'            'End If
      ' ©µ®i
      Case "102":
         ' ©w½Z»y¤å
         Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
            ' ¤¤¤å
            Case "1":
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "01", m_CP09, "01", strUserNum
            ' ­^¤å
            Case "2":
               ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
               EndLetter "01", m_CP09, "02", strUserNum
               'Add By Sindy 2012/11/26 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
               'Modify By Sindy 2017/4/12¡iFCT 01 102  02 ©µ®iµo¤å¡j
               m_MySt(1) = m_TM01: m_MySt(2) = m_TM02: m_MySt(3) = m_TM03: m_MySt(4) = m_TM04: m_Rule = m_CP09
               strDebitNote = ExceptFieldData2("FCT¯S®í½Ð´Ú¤å¦r¹ï·Ó")
               If bolEmail = True And bolPlusPaper = False Then
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                           "','¨Ò¥~¤º¤å','Enclosed please find scanned copies of the application as filed together with the filing receipt thereof for your reference. " & IIf(strDebitNote = "", "Our debit note for services rendered has also been attached for your kind settlement.", strDebitNote) & "')"
                  cnnConnection.Execute strSql
               Else '¶l¥ó
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                           "','¨Ò¥~¤º¤å','" & IIf(strDebitNote = "", "Enclosed please find our debit note for services rendered for your kind settlement.", strDebitNote) & " Copies of the application as filed together with the filing receipt thereof will be mailed to you with the confirmation copy of this letter for your records.')"
                  cnnConnection.Execute strSql
               End If
               '2012/11/26 End
               ' ¬O§_¸É¥ó
               If IsEmptyText(strTemp) = False Then
                  ' ¬O§_¸É¥ó
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                           "','¬O§_¸É¥ó','" & ChgSQL(strTemp) & "')"
                  cnnConnection.Execute strSql
               End If
            ' ¤é¤å
            Case "3":
               'Add By Sindy 2016/12/6 ÀË¬d¬O§_¦³ÅÜ§ó¨Æ¶µ
               'ÅÜ§ó¥Ó½Ð¤H:m_strCE04
               'ÅÜ§ó¦a§}:m_strCE23CE24CE25
               If PUB_FCTchkChangeEventData(m_CP09, "CE04", m_strCE04) = True Then
                  Call PUB_FCTchkChangeEventData(m_CP09, "CE23||CE24||CE25", m_strCE23CE24CE25)
               End If
               If m_strCE04 <> "" Or m_strCE23CE24CE25 <> "" Then
                  ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                  EndLetter "01", m_CP09, "05", strUserNum
                  ' ÅÜ§ó¨Æ¶µ
                  If m_strCE04 <> "" And m_strCE23CE24CE25 <> "" Then
                     'Modified by Morgan 2023/3/15
                     'strExc(1) = "¡]°Ó¼Ð“¸ªÌÇUªí¥Ü¤ÎÇZ¦í©ÒŒi§ó¥Ó½Ðþß§tÇeÇsþòÇiÇU¡^"
                     strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ1")
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "01" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & _
                              "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                     cnnConnection.Execute strSql
                  ElseIf m_strCE04 <> "" Then
                     'Modified by Morgan 2023/3/15
                     'strExc(1) = "¡]°Ó¼Ð“¸ªÌÇUªí¥ÜŒi§ó¥Ó½Ðþß§tÇeÇsþòÇiÇU¡^"
                     strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ2")
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "01" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & _
                              "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                     cnnConnection.Execute strSql
                  ElseIf m_strCE23CE24CE25 <> "" Then
                     'Modified by Morgan 2023/3/15
                     'strExc(1) = "¡]°Ó¼Ð“¸ªÌÇU¦í©ÒŒi§ó¥Ó½Ðþß§tÇeÇsþòÇiÇU¡^"
                     strExc(1) = PUB_GetUniText(Me.Name, "ÅÜ§ó¨Æ¶µ3")
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "01" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & _
                              "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                     cnnConnection.Execute strSql
                  End If
                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
                  If textPrtTrans <> "N" Then
                     ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                     EndLetter "01", m_CP09, "06", strUserNum
                     ' ¥XÄ@¤Hªº¦WºÙ
                     If m_strCE04 = "" Then
                        'Modify By Sindy 2017/2/3
'                        '¡¼¥XÄ@¤HÇU¦W†ï
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "01" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
'                                 "','¥XÄ@¤Hªº¦WºÙ','¥XÄ@¤HÇU¦W†ï')"
'                        cnnConnection.Execute strSql
                     Else
                        'Modify By Sindy 2017/2/3
                        '¡½¥XÄ@¤HÇU¦W†ï
                        'Removed by Morgan 2023/3/15 ©w½Z¨S¥Î¨ì
                        'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        '         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
                        '         "','¥XÄ@¤Hªº¦WºÙ','¥XÄ@¤HÇU¦W†ï ')"
                        'cnnConnection.Execute strSql
                     End If
                     ' ¥XÄ@¤Hªº¦í©Ò
                     If m_strCE23CE24CE25 = "" Then
                        'Modify By Sindy 2017/2/3
'                        '¡¼¥XÄ@¤HÇU¦í©Ò
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "01" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
'                                 "','¥XÄ@¤Hªº¦í©Ò','¥XÄ@¤HÇU¦í©Ò')"
'                        cnnConnection.Execute strSql
                     Else
                        'Modify By Sindy 2017/2/3
                        '¡½¥XÄ@¤HÇU¦í©Ò
                        'Removed by Morgan 2023/3/15 ©w½Z¨S¥Î¨ì
                        'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        '         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
                        '         "','¥XÄ@¤Hªº¦í©Ò','¥XÄ@¤HÇU¦í©Ò ')"
                        'cnnConnection.Execute strSql
                     End If
                     ' ÅÜ§ó¨Æ¶µ
                     If m_strCE04 <> "" And m_strCE23CE24CE25 <> "" Then
                        'Modified by Morgan 2023/3/15
                        'strExc(1) = "³Æ¦Ò¡G§ó·sµn“÷³\¥i³qª¾®ÑÇRþÝÆêþù¡B°Ó¼Ð“¸ªÌÇUªí¥Ü¤ÎÇZ¦í©ÒŒi§ó³\¥iÇyÇiª`°Oþêþù¤UþèÆê¡C "
                        strExc(1) = PUB_GetUniText(Me.Name, "³Æ¦Ò1")
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "01" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
                                 "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                        cnnConnection.Execute strSql
                     ElseIf m_strCE04 <> "" Then
                        'Modified by Morgan 2023/3/15
                        'strExc(1) = "³Æ¦Ò¡G§ó·sµn“÷³\¥i³qª¾®ÑÇRþÝÆêþù¡B°Ó¼Ð“¸ªÌÇUªí¥ÜŒi§ó³\¥iÇyÇiª`°Oþêþù¤UþèÆê¡C "
                        strExc(1) = PUB_GetUniText(Me.Name, "³Æ¦Ò2")
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "01" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
                                 "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                        cnnConnection.Execute strSql
                     ElseIf m_strCE23CE24CE25 <> "" Then
                        'Modified by Morgan 2023/3/15
                        'strExc(1) = "³Æ¦Ò¡G§ó·sµn“÷³\¥i³qª¾®ÑÇRþÝÆêþù¡B°Ó¼Ð“¸ªÌÇU¦í©ÒŒi§ó³\¥iÇyÇiª`°Oþêþù¤UþèÆê¡C "
                        strExc(1) = PUB_GetUniText(Me.Name, "³Æ¦Ò3")
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "01" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & _
                                 "','ÅÜ§ó¨Æ¶µ','" & strExc(1) & "')"
                        cnnConnection.Execute strSql
                     End If
                  End If
               Else
               '2016/12/6 END
                  ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                  EndLetter "01", m_CP09, "03", strUserNum
                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
                  If textPrtTrans <> "N" Then
                     ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
                     EndLetter "01", m_CP09, "04", strUserNum
                  End If
               End If
         End Select
      ' ¸É¥¿
      Case "201":
         ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
         EndLetter "01", m_CP09, "04", strUserNum
         'Add By Sindy 2012/11/26 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
         'Modify By Sindy 2017/4/13¡iFCT 01 000  04 ¨çª¾¤w¸É¤å¥ó.½Ð´Ú¡j
         m_MySt(1) = m_TM01: m_MySt(2) = m_TM02: m_MySt(3) = m_TM03: m_MySt(4) = m_TM04: m_Rule = m_CP09
         strDebitNote = ExceptFieldData2("FCT¯S®í½Ð´Ú¤å¦r¹ï·Ó")
         If bolEmail = True And bolPlusPaper = False Then
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & _
                     "','¨Ò¥~¤º¤å','Enclosed please find a scanned copy of our request for your records. " & IIf(strDebitNote = "", "Our debit note for services rendered is also attached for your kind settlement.", strDebitNote) & "')"
            cnnConnection.Execute strSql
         Else '¶l¥ó
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & _
                     "','¨Ò¥~¤º¤å','" & IIf(strDebitNote = "", "Enclosed please find our debit note for services rendered for your kind settlement.", strDebitNote) & " Copies of our request will be mailed to you with the confirmation copy of this letter for your records.')"
            cnnConnection.Execute strSql
         End If
         '2012/11/26 End
   End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ¦C¦L©w½Z
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   'Add by Morgan 2008/6/11
   Dim ET03 As String, ET03_1 As String, stContent As String
   Dim stLang As String, strFilePath As String, strFN01 As String, strFN02 As String 'Added by Lydia 2023/05/03
   
   'Add By Sindy 2012/11/23 ±q¤U­±µ{¦¡©¹¤WMove¦Ü¦¹
   bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, m_CP10 = "102", , bolPlusPaper) 'ÀË¬d¬O§_¥HE-Mail³qª¾
   '2012/11/23 End
   stLang = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) 'Added by Lydia 2023/05/03
   
   ' ¥ý©I¥s©w½Zµ{¦¡ªº²M°£­ì©w½Z¸ê®Æªº¨ç¦¡¥h²M°£¤§«e´Ý¯d¦b¨Ò¥~Äæ¦ìÀÉ¤¤ªº¸ê®Æ
   InsExpField
   
   Select Case m_CP10
        'Add By Cheng 2003/01/13
      ' ¥Ó½Ð
      Case "101":
'edit by nick 2004/08/13 ¤w¸g§ï¨ì¥Ó½Ð®×¸¹¦C¦L
'        '­Y¦³¿é¤J¥Ó½Ð®×¸¹
'        'edit by nick 2004/08/05 ­×§ï¦¨¤£ºÞ¥Ó½Ð®×¸¹¦³µL¿é¤J¬Ò¥i¥H¦L
'        'If Me.textTM12.Text <> "" Then
'            ' ©w½Z»y¤å
'            Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
'               ' ¤¤¤å
'               Case "1":
'                  ' ¦C¦L©w½Z
'                  NowPrint m_CP09, "02", "01", False, strUserNum, 0
'               ' ­^¤å
'               Case "2":
'                    'Modify By Cheng 2003/02/20
'                    '­Y¦³Àu¥ýÅv¥X¤£¦P©w½Z
''                  ' ¦C¦L©w½Z
''                  NowPrint m_CP09, "02", "02", False, strUserNum, 0
''                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
''                  If textPrtTrans <> "N" Then
''                     ' ¦C¦L©w½Z
''                     NowPrint m_CP09, "02", "03", False, strUserNum, 0
''                  End If
'                'Modify By Cheng 2003/12/09
'                '¨Ï¥Î·s³W©w
''                '­Y¥Ó½Ð¤é¤p©ó20031128
''                If DBDATE(Val(Me.textTM11.Text)) < 20031128 Then
''                    ' ¦C¦L©w½Z
''                    NowPrint m_CP09, "02", IIf(m_blnPriDate, "02", "04"), False, strUserNum, 0
''                    ' ¬O§_¦C¦LÂ½Ä¶¨ç
''                    If textPrtTrans <> "N" Then
''                       ' ¦C¦L©w½Z
''                       'Áp¦X°Ó¼Ð
''                       If Me.textTM08.Text = "2" Or Me.textTM08.Text = "5" Then
''                           NowPrint m_CP09, "02", IIf(m_blnPriDate, "06", "07"), False, strUserNum, 0
''                      '«DÁp¦X°Ó¼Ð
''                      Else
''                           NowPrint m_CP09, "02", IIf(m_blnPriDate, "03", "05"), False, strUserNum, 0
''                      End If
''                    End If
''                '­Y¥Ó½Ð¤é¤j©óµ¥©ó20031128
''                Else
'                    ' ¦C¦L©w½Z
'                    'edit by nick 2004/08/13 ¤w¸g¨ú®øÀu¥ýÅv
'                    'NowPrint m_CP09, "02", IIf(m_blnPriDate, "10", "12"), False, strUserNum, 0
'                    NowPrint m_CP09, "02", "12", False, strUserNum, 0
'                    ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                    If textPrtTrans <> "N" Then
'                       ' ¦C¦L©w½Z
'                       'edit by nick 2004/08/13 ¤w¸g¨ú®øÀu¥ýÅv
'                        'NowPrint m_CP09, "02", IIf(m_blnPriDate, "11", "13"), False, strUserNum, 0
'                        NowPrint m_CP09, "02", "13", False, strUserNum, 0
'                    End If
''                End If
'               ' ¤é¤å
'               Case "3":
'                  ' ¦C¦L©w½Z
'                  '93.6.2 MODIFY BY SONIA
'                  'NowPrint m_CP09, "02", "04", False, strUserNum, 0
'                  NowPrint m_CP09, "02", "08", False, strUserNum, 0
'                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
'                  If textPrtTrans <> "N" Then
'                      ' ¦C¦L©w½Z
'                      NowPrint m_CP09, "02", "09", False, strUserNum, 0
'                  End If
'                  '93.6.2 END
'            End Select
'        'edit by nick 2004/08/05 ­×§ï¦¨¤£ºÞ¥Ó½Ð®×¸¹¦³µL¿é¤J¬Ò¥i¥H¦L
'        'End If
      ' ©µ®i
      Case "102":
         ' ©w½Z»y¤å
         'Modified by Lydia 2023/05/03 §ï¦¨ÅÜ¼Æ
         Select Case stLang
            ' ¤¤¤å
            Case "1":
               ET03 = "01"
            ' ­^¤å
            Case "2":
               'Modify By Sindy 2022/3/10
               If PUB_SpecApplData_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, ET03, , "01") = False Then
               '2022/3/10 END
                  ET03 = "02"
               End If
            ' ¤é¤å
            Case "3":
               'Add By Sindy 2016/12/6 ÀË¬d¬O§_¦³ÅÜ§ó¨Æ¶µ
               If m_strCE04 <> "" Or m_strCE23CE24CE25 <> "" Then
                  ET03 = "05"
                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
                  If textPrtTrans <> "N" Then
                     ET03_1 = "06"
                  End If
               Else
               '2016/12/6 END
                  ET03 = "03"
                  ' ¬O§_¦C¦LÂ½Ä¶¨ç
                  If textPrtTrans <> "N" Then
                     ET03_1 = "04"
                  End If
               End If
         End Select
      ' ¸É¥¿
      Case "201":
         ET03 = "04"
         ' ¦C¦L©w½Z
         
   End Select
   
   
   If ET03 <> "" Then
      'Add by Morgan 2008/6/11
'      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, m_CP10 = "102", , bolPlusPaper)
      'If bolEmail Then 'Mark by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW
         'Add by Morgan 2009/10/20 +§PÂ_¬O§_EMail¦P®É±H¯È¥»
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'end 2009/10/20
         'Added by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW: ­^¤å²Õ¤À¦¨«H¨ç©MÂ½Ä¶¨â­ÓÀÉ®×
         If stLang <> "3" Then
            strFilePath = Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04)
            If Pub_GetFCTeFileName(strFilePath, m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, , strFN01, strFN02) = False Then
              Exit Sub
            End If
            NowPrint m_CP09, "01", ET03, True, strUserNum, , , , , iCopy, , True
            If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                Sleep 100
            End If
            If ET03_1 <> "" Then
               NowPrint m_CP09, "01", ET03_1, True, strUserNum, , , , , iCopy, , True
               If PUB_PrintWord2File(g_WordAp, strFilePath, strFN02) = True Then
                   Sleep 100
               End If
            End If
         Else  '¤é¤å²Õ:¤£§ïÅÜ¦sÀÉ¼Ò¦¡
         'end 2023/05/03
            'Added by Lydia 2024/11/14 ¦]¤é¥»¥N²z¤H¯S§O­n¨D¡A»Ý±N³qª¾«H¨ç»PÄ¶¤åµ¥¤À¶}¡A¨Ã¥B²Î¤@¦WºÙ¦p¤U(¼Ò²Õ¨ú±o)¡F­ì¥»ªºÀÉ®×(®×¸¹_¤é´Á=³qª¾¨ç+Ä¶¤å)¤´­n²£¥Í¡A¥H§K¤é«á¤S¦³¥N²z¤H­n¨D¦X¨Ö
            strFilePath = Pub_GetEFilePath_All(m_TM01, m_TM02, m_TM03, m_TM04)
            If Pub_GetFCTeFileName(strFilePath, m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, , strFN01, strFN02) = False Then
              Exit Sub
            End If
            NowPrint m_CP09, "01", ET03, True, strUserNum, , , , , iCopy, , True
            If PUB_PrintWord2File(g_WordAp, strFilePath, strFN01) = True Then
                Sleep 100
            End If
            If ET03_1 <> "" Then
               NowPrint m_CP09, "01", ET03_1, True, strUserNum, , , , , iCopy, , True
               If PUB_PrintWord2File(g_WordAp, strFilePath, strFN02) = True Then
                   Sleep 100
               End If
            End If
            'end 2024/11/14
            If ET03_1 <> "" Then
               NowPrint m_CP09, "01", ET03, False, strUserNum, , , , , iCopy
               NowPrint m_CP09, "01", ET03_1, False, strUserNum, , , , , iCopy
               NowPrint m_CP09, "01", ET03, False, strUserNum, , , True, stContent, , , , True
               NowPrint m_CP09, "01", ET03_1, False, strUserNum, , stContent, , , , , True, True
            Else
               NowPrint m_CP09, "01", ET03, False, strUserNum, , , , , iCopy, , True, True
            End If
         End If 'Added by Lydia 2023/05/03
         MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(m_TM01) & " ]¡I"
      'Mark by Lydia 2023/05/03
      'Else
      ''end 2008/6/11
       '  NowPrint m_CP09, "01", ET03, False, strUserNum, 0
       '  If ET03_1 <> "" Then
       '     NowPrint m_CP09, "01", ET03_1, False, strUserNum, 0
       '  End If
      'End If
      'end 2023/05/03
   End If
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim dblCountFee As Double
Dim intMoney As Long  '­¿¼Æ

TxtValidate = False
'add by nick 2004/08/13 µo¤å³W¶O¡A¥Ó½Ð°ê®a¥xÆW¤~ÀË¬d
If Me.textCP84.Enabled = True Then
   Cancel = False
   textCP84_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textCP84.Enabled = True And m_TM10 = "000" Then
   If Val(textCP84.Text) <> Val(m_CP84) Then
      MsgBox "µo¤å³W¶O[" & Trim(Val(m_CP84)) & "] »P¹ê»Úµo¤å³W¶O[" & Trim(Val(textCP84.Text)) & "]¤£¦P", , "Äµ§i¡I"
      textCP84_GotFocus
      Exit Function
   End If
   'Add By Sindy 2015/8/26
   If Trim(textTM09) = "" Then
      MsgBox "°Ó«~Ãþ§O¤£¥iªÅ¥Õ¡I"
      SSTab1.Tab = 1
      textTM09_GotFocus
      textTM09.SetFocus
      Exit Function
   End If
   '2015/8/26 END
   'Add By Sindy 2014/2/20 ÀË¬d©µ®i¸óÃþ³W¶O
   If m_CP10 = "102" Then
      textTM09 = IIf(Right(textTM09, 1) = ",", Mid(textTM09, 1, Len(textTM09) - 1), textTM09)
      textTM09 = IIf(Left(textTM09, 1) = ",", Mid(textTM09, 2, Len(textTM09)), textTM09)
      tmpArr = Split(textTM09, ",")
      '¦³¸óÃþ¤~»Ý­nÀË¬d
      If (Val(UBound(tmpArr)) + 1) > 1 Then
         intMoney = 1
'         If m_CP07 <> "" Then
'            '­Y¨t²Î¤éªº¬Q¤Ñ¬°«D¤u§@¤Ñ, «h¥H¨t²Î¤éªº«e¤@­Ó¤u§@¤Ñ°µ¤ñ¸û
'            If ChkWorkDay(DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1))))) = False Then
'               If (Val(DBDATE(m_CP07)) - 19110000) < (Val(CompWorkDay(1, DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1)))), 1)) - 19110000) And (Val(DBDATE(m_CP07)) - 19110000) <> 0 Then
'                  intMoney = 2
'               End If
'            Else
'               If (Val(DBDATE(m_CP07)) - 19110000) < Val(GetTaiwanTodayDate) And (Val(DBDATE(m_CP07)) - 19110000) <> 0 Then
'                  intMoney = 2
'               End If
'            End If
'         End If
         'Modify By Sindy 2015/4/24 ex.FCT-23444 ­ì¥ÎCP07ÀË¬d,§ï¥ÎTM22ÀË¬d
         '­Y¨t²Î¤éªº¬Q¤Ñ¬°«D¤u§@¤Ñ, «h¥H¨t²Î¤éªº«e¤@­Ó¤u§@¤Ñ°µ¤ñ¸û
         If ChkWorkDay(DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1))))) = False Then
            If (Val(DBDATE(m_TM22)) - 19110000) < (Val(CompWorkDay(1, DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1)))), 1)) - 19110000) And (Val(DBDATE(m_TM22)) - 19110000) <> 0 Then
               intMoney = 2
            End If
         Else
            If (Val(DBDATE(m_TM22)) - 19110000) < Val(GetTaiwanTodayDate) And (Val(DBDATE(m_TM22)) - 19110000) <> 0 Then
               intMoney = 2
            End If
         End If
         '2015/4/24 END
         dblCountFee = Val(4000 * (Val(UBound(tmpArr)) + 1) * intMoney)
         If Val(textCP84) <> dblCountFee Then
            MsgBox "³W¶O¤£²Å¡A¥»®×Ãþ§O¼Æ¦@" & (Val(UBound(tmpArr)) + 1) & "Ãþ¡A¨CÃþ4,000¤¸¡AÀ³¬°" & Format(dblCountFee, "#,##0") & "¤¸" & IIf(intMoney = 2, "¡]§t¹O´Á¥[­¿¡^", "") & "¡C" & vbCrLf & _
                   "­Y«D©Ò¦³Ãþ§O³£­n©µ®i¡A¥²¶·±N°Ó«~Ãþ§OÄæ¤§Ãþ§O¼Æ§ï¥¿½T!!"
            SSTab1.Tab = 1
            textTM09.SetFocus
            Exit Function
         End If
      End If
   End If
   '2014/2/20 END
End If

If Me.textAdd.Enabled = True Then
   Cancel = False
   textAdd_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'If Me.textAddDate.Enabled = True Then
'   Cancel = False
'   textAddDate_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If

If Me.textCP09_S.Enabled = True Then
   Cancel = False
   textCP09_S_Validate Cancel
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

'Add By Sindy 98/03/11
If Me.textCP113.Enabled = True Then
   Cancel = False
   textCP113_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'98/03/11 End

If Me.textDN.Enabled = True Then
   Cancel = False
   textDN_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textMail.Enabled = True Then
   Cancel = False
   textMail_Validate Cancel
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

'edit by nick 2004/08/13  ²¾¨ì frm030203_02
'If Me.textPriorityDoc.Enabled = True Then
'   Cancel = False
'   textPriorityDoc_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If

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
If Me.textTM72.Enabled = True Then
   Cancel = False
   textTM72_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Lydia 2023/11/16
If Me.cboTM08.Enabled = True Then
   Cancel = False
   cboTM08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.cboTM72.Enabled = True Then
   Cancel = False
   cboTM72_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'end 2023/11/16
   
If Me.textUargeDate.Enabled = True Then
   Cancel = False
   textUargeDate_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'add by nickc 2006/01/27
'edit by nickc 2006/02/07
If m_TM01 = "FCT" Then
    If Me.lstNameAgent.Enabled = True Then
        Cancel = False
        lstNameAgent_Validate Cancel
        If Cancel = True Then
            lstNameAgent.SetFocus
            Exit Function
        End If
    End If
End If

TxtValidate = True
End Function

'Add By Sindy 98/03/11
Private Sub textCP113_GotFocus()
   TextInverse textCP113
End Sub
Private Sub textCP113_Validate(Cancel As Boolean)
   If textCP113 <> "" Then
      If Not IsNumeric(textCP113) Then
         MsgBox "½Ð¿é¤J¼Æ¦r¡I", vbExclamation
         textCP113.SetFocus
         textCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   If GetPrjNation1(textTMKey) = "000" Then
      Cancel = Not PUB_CheckCP113(textCP113, m_TM01, m_CP10, m_CP14)
   End If
End Sub
'98/03/11 End

'Add By Sindy 2012/12/20
Private Sub textCP118_GotFocus()
   TextInverse textCP118
   CloseIme
End Sub

'Add By Sindy 2012/12/20
Private Sub textCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub


'Added by Lydia 2023/11/16
Private Sub cboTM72_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboTM72_Validate(Cancel As Boolean)
Dim intX As Integer, intQ As Integer
   
   intX = -1
   If Trim(cboTM72.Text) <> "" And cboTM72.Tag <> cboTM72.Text Then
        For intQ = 0 To cboTM72.ListCount - 1
           If InStr(cboTM72.List(intQ), Trim(cboTM72.Text)) > 0 Then
              intX = intQ
              Exit For
           End If
        Next intQ
        If intX = -1 Then
           cboTM72.SetFocus
           cboTM72.Tag = cboTM72.Text
           Cancel = True
           Exit Sub
        Else
           cboTM72.ListIndex = intX
        End If
   End If
   textTM72 = Trim(Left(cboTM72, 1))
   cboTM72.Tag = cboTM72.Text
End Sub

Private Sub cboTM08_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboTM08_Validate(Cancel As Boolean)
Dim intX As Integer, intQ As Integer
   
   intX = -1
   If Trim(cboTM08.Text) <> "" And cboTM08.Tag <> cboTM08.Text Then
        For intQ = 0 To cboTM08.ListCount - 1
           If InStr(cboTM08.List(intQ), Trim(cboTM08.Text)) > 0 Then
              intX = intQ
              Exit For
           End If
        Next intQ
        If intX = -1 Then
            cboTM08.SetFocus
            cboTM08.Tag = cboTM08.Text
            Cancel = True
            Exit Sub
        Else
            cboTM08.ListIndex = intX
        End If
   End If
   textTM08 = Trim(Left(cboTM08.Text, 1))
   cboTM08.Tag = cboTM08.Text
End Sub
'end 2023/11/16
