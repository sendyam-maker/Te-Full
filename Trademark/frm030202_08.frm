VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030202_08 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "µo¤å(²¾Âà)"
   ClientHeight    =   6470
   ClientLeft      =   4740
   ClientTop       =   2100
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6470
   ScaleWidth      =   8960
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8160
      TabIndex        =   30
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   330
      Index           =   0
      Left            =   6105
      TabIndex        =   28
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "¦^«eµe­±(&U)"
      Height          =   330
      Left            =   6930
      TabIndex        =   29
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H008080FF&
      Caption         =   "ÅÜ§ó¨Æ¶µ(&R)"
      Height          =   330
      Left            =   4890
      Style           =   1  '¹Ï¤ù¥~Æ[
      TabIndex        =   27
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "¬ÛÃö¨÷¸¹(&F)"
      Height          =   330
      Left            =   3660
      TabIndex        =   26
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "¦P®Éµo¤å(&N)"
      Height          =   330
      Index           =   1
      Left            =   2430
      TabIndex        =   25
      Top             =   15
      Width           =   1200
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   958
      Width           =   2085
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   3900
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   264
      Left            =   6720
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   360
      Width           =   2085
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   264
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   664
      Width           =   2085
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   930
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   654
      Width           =   2085
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   930
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   360
      Width           =   2085
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   930
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   948
      Width           =   2085
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   3900
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   948
      Width           =   1935
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   285
      Left            =   3900
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   654
      Width           =   1935
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3950
      Left            =   60
      TabIndex        =   58
      Top             =   2520
      Width           =   8900
      _ExtentX        =   15699
      _ExtentY        =   6967
      _Version        =   393216
      TabHeight       =   420
      TabCaption(0)   =   "°ò¥»¸ê®Æ"
      TabPicture(0)   =   "frm030202_08.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(10)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label23"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label22"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label14(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label25"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label28"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label36"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label37"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label17"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label39"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblNameAgent"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label15"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label16"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label18(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label19"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(12)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label55"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label43"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCP89_2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCP90_2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCP91_2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCP92_2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCP56_2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCP64"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lstNameAgent"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textTM15_S"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textUargeDate"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCP27"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textDN"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCP18"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textAdd"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCP84"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text7"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textPrint"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textCP56"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textCP89"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textCP90"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCP91"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textCP92"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCP113"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCP118"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).ControlCount=   44
      TabCaption(1)   =   "¥Nªí¤H-1"
      TabPicture(1)   =   "frm030202_08.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textTM105"
      Tab(1).Control(1)=   "textTM104"
      Tab(1).Control(2)=   "textTM103"
      Tab(1).Control(3)=   "Combo2(5)"
      Tab(1).Control(4)=   "textTM99"
      Tab(1).Control(5)=   "textTM98"
      Tab(1).Control(6)=   "textTM97"
      Tab(1).Control(7)=   "Combo2(3)"
      Tab(1).Control(8)=   "textTM52"
      Tab(1).Control(9)=   "textTM51"
      Tab(1).Control(10)=   "textTM50"
      Tab(1).Control(11)=   "Combo2(1)"
      Tab(1).Control(12)=   "textTM102"
      Tab(1).Control(13)=   "textTM101"
      Tab(1).Control(14)=   "textTM100"
      Tab(1).Control(15)=   "Combo2(4)"
      Tab(1).Control(16)=   "textTM96"
      Tab(1).Control(17)=   "textTM95"
      Tab(1).Control(18)=   "textTM94"
      Tab(1).Control(19)=   "Combo2(2)"
      Tab(1).Control(20)=   "textTM49"
      Tab(1).Control(21)=   "textTM48"
      Tab(1).Control(22)=   "textTM47"
      Tab(1).Control(23)=   "Combo2(0)"
      Tab(1).Control(24)=   "Label18(2)"
      Tab(1).Control(25)=   "Label14(1)"
      Tab(1).Control(26)=   "Label5(3)"
      Tab(1).Control(27)=   "Label5(4)"
      Tab(1).Control(28)=   "Label5(5)"
      Tab(1).Control(29)=   "Label5(6)"
      Tab(1).Control(30)=   "Label5(7)"
      Tab(1).Control(31)=   "Label5(8)"
      Tab(1).Control(32)=   "Label5(1)"
      Tab(1).Control(33)=   "Label5(2)"
      Tab(1).Control(34)=   "Label5(9)"
      Tab(1).Control(35)=   "Label5(10)"
      Tab(1).Control(36)=   "Label5(11)"
      Tab(1).Control(37)=   "Label5(12)"
      Tab(1).Control(38)=   "Label14(2)"
      Tab(1).Control(39)=   "Label18(1)"
      Tab(1).Control(40)=   "Label5(13)"
      Tab(1).Control(41)=   "Label5(14)"
      Tab(1).Control(42)=   "Label5(15)"
      Tab(1).Control(43)=   "Label5(16)"
      Tab(1).Control(44)=   "Label5(17)"
      Tab(1).Control(45)=   "Label5(18)"
      Tab(1).Control(46)=   "Label14(3)"
      Tab(1).Control(47)=   "Label18(3)"
      Tab(1).ControlCount=   48
      TabCaption(2)   =   "¥Nªí¤H-2"
      TabPicture(2)   =   "frm030202_08.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label18(6)"
      Tab(2).Control(1)=   "Label14(7)"
      Tab(2).Control(2)=   "Label5(32)"
      Tab(2).Control(3)=   "Label5(33)"
      Tab(2).Control(4)=   "Label5(34)"
      Tab(2).Control(5)=   "Label5(35)"
      Tab(2).Control(6)=   "Label5(36)"
      Tab(2).Control(7)=   "Label5(37)"
      Tab(2).Control(8)=   "Label18(7)"
      Tab(2).Control(9)=   "Label14(8)"
      Tab(2).Control(10)=   "Label5(38)"
      Tab(2).Control(11)=   "Label5(39)"
      Tab(2).Control(12)=   "Label5(40)"
      Tab(2).Control(13)=   "Label5(41)"
      Tab(2).Control(14)=   "Label5(42)"
      Tab(2).Control(15)=   "Label5(43)"
      Tab(2).Control(16)=   "TextTM106"
      Tab(2).Control(17)=   "TextTM107"
      Tab(2).Control(18)=   "TextTM109"
      Tab(2).Control(19)=   "TextTM110"
      Tab(2).Control(20)=   "Combo2(7)"
      Tab(2).Control(21)=   "Combo2(6)"
      Tab(2).Control(22)=   "TextTM108"
      Tab(2).Control(23)=   "TextTM111"
      Tab(2).Control(24)=   "TextTM112"
      Tab(2).Control(25)=   "TextTM113"
      Tab(2).Control(26)=   "TextTM115"
      Tab(2).Control(27)=   "TextTM116"
      Tab(2).Control(28)=   "Combo2(9)"
      Tab(2).Control(29)=   "Combo2(8)"
      Tab(2).Control(30)=   "TextTM114"
      Tab(2).Control(31)=   "TextTM117"
      Tab(2).ControlCount=   32
      Begin VB.TextBox textCP118 
         Height          =   270
         Left            =   5550
         MaxLength       =   1
         TabIndex        =   24
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   -69090
         MaxLength       =   4
         TabIndex        =   102
         Top             =   285
         Width           =   600
      End
      Begin VB.TextBox Text2 
         Height          =   264
         Left            =   -74160
         MaxLength       =   9
         TabIndex        =   101
         Top             =   2250
         Width           =   1092
      End
      Begin VB.TextBox textTM81_2 
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   2250
         Width           =   2505
      End
      Begin VB.TextBox Text3 
         Height          =   264
         Left            =   -69750
         MaxLength       =   9
         TabIndex        =   99
         Top             =   1980
         Width           =   1092
      End
      Begin VB.TextBox textTM80_2 
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -68670
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   1980
         Width           =   2505
      End
      Begin VB.TextBox Text4 
         Height          =   264
         Left            =   -74160
         MaxLength       =   9
         TabIndex        =   97
         Top             =   1980
         Width           =   1092
      End
      Begin VB.TextBox textTM79_2 
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   1980
         Width           =   2505
      End
      Begin VB.TextBox Text5 
         Height          =   264
         Left            =   -69750
         MaxLength       =   9
         TabIndex        =   95
         Top             =   1710
         Width           =   1092
      End
      Begin VB.TextBox textTM78_2 
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -68670
         Locked          =   -1  'True
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   1710
         Width           =   2505
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "²Ó©úÅé"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   40
         ItemData        =   "frm030202_08.frx":0054
         Left            =   -73515
         List            =   "frm030202_08.frx":005E
         Sorted          =   -1  'True
         Style           =   1  '¶µ¥Ø¥]§t®Ö¨ú¤è¶ô
         TabIndex        =   93
         Top             =   1320
         Width           =   1260
      End
      Begin VB.TextBox Text6 
         Height          =   288
         Left            =   -74790
         MaxLength       =   1
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1530
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  '¾a¥k¹ï»ô
         Height          =   285
         Left            =   -71430
         TabIndex        =   91
         Top             =   285
         Width           =   1425
      End
      Begin VB.TextBox textTM72 
         Height          =   264
         Left            =   -68370
         MaxLength       =   1
         TabIndex        =   90
         Top             =   600
         Width           =   372
      End
      Begin VB.TextBox textTM72_2 
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -67980
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   600
         Width           =   1785
      End
      Begin VB.TextBox Text9 
         Height          =   264
         Left            =   -68370
         MaxLength       =   1
         TabIndex        =   88
         Top             =   912
         Width           =   492
      End
      Begin VB.TextBox textTM05_1 
         Height          =   792
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  '««ª½±²¶b
         TabIndex        =   87
         Top             =   2610
         Width           =   7272
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   -66720
         MaxLength       =   1
         TabIndex        =   86
         Top             =   1290
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textPrtTrans 
         Height          =   264
         Left            =   -68370
         MaxLength       =   10
         TabIndex        =   85
         Top             =   1182
         Width           =   372
      End
      Begin VB.TextBox textCP09_S 
         Height          =   264
         Left            =   -70590
         MaxLength       =   1
         TabIndex        =   84
         Top             =   990
         Width           =   465
      End
      Begin VB.TextBox textCP09_S1 
         Height          =   264
         Left            =   -70050
         MaxLength       =   6
         TabIndex        =   83
         Top             =   990
         Width           =   975
      End
      Begin VB.TextBox textCP09_S2 
         Height          =   264
         Left            =   -68970
         MaxLength       =   1
         TabIndex        =   82
         Top             =   990
         Width           =   345
      End
      Begin VB.TextBox textCP09_S3 
         Height          =   264
         Left            =   -68520
         MaxLength       =   2
         TabIndex        =   81
         Top             =   990
         Width           =   465
      End
      Begin VB.TextBox textAdd_2 
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1425
         Width           =   6072
      End
      Begin VB.TextBox Text10 
         Height          =   264
         Left            =   -73320
         MaxLength       =   10
         TabIndex        =   79
         Top             =   1455
         Width           =   852
      End
      Begin VB.TextBox textTM09 
         Height          =   264
         Left            =   -73560
         MaxLength       =   395
         TabIndex        =   78
         Top             =   360
         Width           =   7272
      End
      Begin VB.TextBox textTM32 
         Height          =   264
         Left            =   -73560
         MaxLength       =   300
         TabIndex        =   77
         Top             =   660
         Width           =   7272
      End
      Begin VB.TextBox textMail 
         Height          =   264
         Left            =   -73560
         MaxLength       =   1
         TabIndex        =   76
         Top             =   960
         Width           =   492
      End
      Begin VB.TextBox Text11 
         Height          =   264
         Left            =   -73560
         MaxLength       =   7
         TabIndex        =   75
         Top             =   300
         Width           =   1092
      End
      Begin VB.TextBox Text12 
         Height          =   264
         Left            =   -73560
         MaxLength       =   7
         TabIndex        =   74
         Top             =   600
         Width           =   1092
      End
      Begin VB.TextBox textTM22 
         Height          =   264
         Left            =   -72120
         MaxLength       =   7
         TabIndex        =   73
         Top             =   912
         Width           =   852
      End
      Begin VB.TextBox textTM21 
         Height          =   264
         Left            =   -73296
         MaxLength       =   7
         TabIndex        =   72
         Top             =   912
         Width           =   852
      End
      Begin VB.TextBox Text13 
         Height          =   264
         Left            =   -73296
         MaxLength       =   1
         TabIndex        =   71
         Top             =   1182
         Width           =   372
      End
      Begin VB.TextBox textTM23_2 
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1710
         Width           =   2505
      End
      Begin VB.TextBox textTM07 
         Height          =   264
         Left            =   -73560
         MaxLength       =   40
         TabIndex        =   69
         Top             =   3105
         Width           =   7272
      End
      Begin VB.TextBox textTM06 
         Height          =   264
         Left            =   -73560
         MaxLength       =   60
         TabIndex        =   68
         Top             =   2835
         Width           =   7272
      End
      Begin VB.TextBox textTM05 
         Height          =   264
         Left            =   -73560
         MaxLength       =   40
         TabIndex        =   67
         Top             =   2610
         Width           =   7272
      End
      Begin VB.TextBox Text14 
         Height          =   264
         Left            =   -71550
         MaxLength       =   1
         TabIndex        =   66
         Top             =   600
         Width           =   372
      End
      Begin VB.TextBox textTM27 
         Height          =   288
         Left            =   -66285
         MaxLength       =   20
         TabIndex        =   65
         Top             =   3075
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.TextBox Text15 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   264
         Left            =   -67830
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   300
         Width           =   1635
      End
      Begin VB.TextBox Text16 
         Height          =   396
         Left            =   -73545
         MaxLength       =   2000
         TabIndex        =   63
         Top             =   1890
         Width           =   7272
      End
      Begin VB.TextBox textTM58 
         Height          =   276
         Left            =   -73545
         MaxLength       =   2000
         TabIndex        =   62
         Top             =   2325
         Width           =   7272
      End
      Begin VB.TextBox Text17 
         Height          =   264
         Left            =   -74160
         MaxLength       =   9
         TabIndex        =   61
         Top             =   1710
         Width           =   1092
      End
      Begin VB.TextBox textTM08_2 
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -71160
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   600
         Width           =   1755
      End
      Begin VB.TextBox textCP113 
         Height          =   285
         Left            =   5610
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox textCP92 
         Height          =   285
         Left            =   1170
         MaxLength       =   9
         TabIndex        =   13
         Top             =   2730
         Width           =   1092
      End
      Begin VB.TextBox textCP91 
         Height          =   285
         Left            =   1170
         MaxLength       =   9
         TabIndex        =   12
         Top             =   2460
         Width           =   1092
      End
      Begin VB.TextBox textCP90 
         Height          =   285
         Left            =   1170
         MaxLength       =   9
         TabIndex        =   11
         Top             =   2190
         Width           =   1092
      End
      Begin VB.TextBox textCP89 
         Height          =   285
         Left            =   1170
         MaxLength       =   9
         TabIndex        =   10
         Top             =   1920
         Width           =   1092
      End
      Begin VB.TextBox textCP56 
         Height          =   285
         Left            =   1170
         MaxLength       =   9
         TabIndex        =   9
         Top             =   1650
         Width           =   1092
      End
      Begin VB.TextBox textPrint 
         Height          =   285
         Left            =   3210
         MaxLength       =   1
         TabIndex        =   4
         Top             =   630
         Width           =   372
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   6300
         MaxLength       =   1
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   630
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox textCP84 
         Alignment       =   1  '¾a¥k¹ï»ô
         Height          =   285
         Left            =   3210
         TabIndex        =   1
         Top             =   345
         Width           =   1425
      End
      Begin VB.TextBox textAdd 
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   7
         Top             =   900
         Width           =   852
      End
      Begin VB.TextBox textCP18 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   285
         Left            =   7530
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   330
         Width           =   1005
      End
      Begin VB.TextBox textDN 
         Height          =   285
         Left            =   1170
         MaxLength       =   1
         TabIndex        =   3
         Top             =   630
         Width           =   492
      End
      Begin VB.TextBox textCP27 
         Height          =   285
         Left            =   1170
         MaxLength       =   7
         TabIndex        =   0
         Top             =   360
         Width           =   1092
      End
      Begin VB.TextBox textUargeDate 
         Height          =   285
         Left            =   5130
         MaxLength       =   7
         TabIndex        =   5
         Top             =   630
         Width           =   1092
      End
      Begin VB.TextBox textTM15_S 
         Height          =   495
         Left            =   1830
         MaxLength       =   200
         ScrollBars      =   2  '««ª½±²¶b
         TabIndex        =   8
         Top             =   1170
         Width           =   5625
      End
      Begin MSForms.TextBox TextTM117 
         Height          =   285
         Left            =   -69660
         TabIndex        =   260
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
      Begin MSForms.TextBox TextTM114 
         Height          =   285
         Left            =   -74070
         TabIndex        =   259
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
         Index           =   8
         Left            =   -74070
         TabIndex        =   22
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
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   9
         Left            =   -69660
         TabIndex        =   23
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
      Begin MSForms.TextBox TextTM116 
         Height          =   285
         Left            =   -69660
         TabIndex        =   258
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
         Left            =   -69660
         TabIndex        =   257
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
         Left            =   -74070
         TabIndex        =   256
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
      Begin MSForms.TextBox TextTM112 
         Height          =   285
         Left            =   -74070
         TabIndex        =   255
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
      Begin MSForms.TextBox TextTM111 
         Height          =   285
         Left            =   -69660
         TabIndex        =   254
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
      Begin MSForms.TextBox TextTM108 
         Height          =   285
         Left            =   -74070
         TabIndex        =   253
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
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   6
         Left            =   -74070
         TabIndex        =   20
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
         Index           =   7
         Left            =   -69660
         TabIndex        =   21
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
      Begin MSForms.TextBox TextTM110 
         Height          =   285
         Left            =   -69660
         TabIndex        =   252
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
         Left            =   -69660
         TabIndex        =   251
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
         Left            =   -74070
         TabIndex        =   250
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
      Begin MSForms.TextBox TextTM106 
         Height          =   285
         Left            =   -74070
         TabIndex        =   249
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
      Begin MSForms.TextBox textTM105 
         Height          =   285
         Left            =   -69720
         TabIndex        =   248
         Top             =   3525
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
         Left            =   -69720
         TabIndex        =   247
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
      Begin MSForms.TextBox textTM103 
         Height          =   285
         Left            =   -69720
         TabIndex        =   246
         Top             =   2940
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
         Left            =   -69720
         TabIndex        =   19
         Top             =   2640
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
      Begin MSForms.TextBox textTM99 
         Height          =   285
         Left            =   -69720
         TabIndex        =   245
         Top             =   2340
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
         Left            =   -69720
         TabIndex        =   244
         Top             =   2055
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
      Begin MSForms.TextBox textTM97 
         Height          =   285
         Left            =   -69720
         TabIndex        =   243
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
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   3
         Left            =   -69720
         TabIndex        =   17
         Top             =   1470
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
      Begin MSForms.TextBox textTM52 
         Height          =   285
         Left            =   -69720
         TabIndex        =   242
         Top             =   1185
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
         Left            =   -69720
         TabIndex        =   241
         Top             =   885
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
      Begin MSForms.TextBox textTM50 
         Height          =   285
         Left            =   -69720
         TabIndex        =   240
         Top             =   600
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
         Left            =   -69720
         TabIndex        =   15
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
      Begin MSForms.TextBox textTM102 
         Height          =   285
         Left            =   -74130
         TabIndex        =   239
         Top             =   3525
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
         Left            =   -74130
         TabIndex        =   238
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
      Begin MSForms.TextBox textTM100 
         Height          =   285
         Left            =   -74130
         TabIndex        =   237
         Top             =   2940
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
         Left            =   -74130
         TabIndex        =   18
         Top             =   2640
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
      Begin MSForms.TextBox textTM96 
         Height          =   285
         Left            =   -74130
         TabIndex        =   236
         Top             =   2340
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
         Left            =   -74130
         TabIndex        =   235
         Top             =   2055
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
      Begin MSForms.TextBox textTM94 
         Height          =   285
         Left            =   -74130
         TabIndex        =   234
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
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   2
         Left            =   -74130
         TabIndex        =   16
         Top             =   1470
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
      Begin MSForms.TextBox textTM49 
         Height          =   285
         Left            =   -74130
         TabIndex        =   233
         Top             =   1185
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
         Left            =   -74130
         TabIndex        =   232
         Top             =   885
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
      Begin MSForms.TextBox textTM47 
         Height          =   285
         Left            =   -74130
         TabIndex        =   231
         Top             =   600
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
         Left            =   -74130
         TabIndex        =   14
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
      Begin MSForms.ListBox lstNameAgent 
         Height          =   315
         Left            =   7500
         TabIndex        =   230
         Top             =   660
         Width           =   1260
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2222;556"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   435
         Left            =   1170
         TabIndex        =   229
         Top             =   3060
         Width           =   7545
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13309;767"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP56_2 
         Height          =   285
         Left            =   2280
         TabIndex        =   228
         TabStop         =   0   'False
         Top             =   1650
         Width           =   6525
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "11509;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP92_2 
         Height          =   285
         Left            =   2280
         TabIndex        =   227
         TabStop         =   0   'False
         Top             =   2730
         Width           =   6525
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "11509;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP91_2 
         Height          =   285
         Left            =   2280
         TabIndex        =   226
         TabStop         =   0   'False
         Top             =   2460
         Width           =   6525
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "11509;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP90_2 
         Height          =   285
         Left            =   2280
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   2190
         Width           =   6525
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "11509;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP89_2 
         Height          =   285
         Left            =   2280
         TabIndex        =   224
         TabStop         =   0   'False
         Top             =   1920
         Width           =   6525
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "11509;503"
         BorderColor     =   16777215
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
         Left            =   4380
         TabIndex        =   214
         Top             =   3540
         Width           =   2085
      End
      Begin VB.Label Label55 
         Caption         =   $"frm030202_08.frx":0072
         ForeColor       =   &H000000C0&
         Height          =   410
         Left            =   180
         TabIndex        =   213
         Top             =   3510
         Width           =   4070
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   43
         Left            =   -70095
         TabIndex        =   212
         Top             =   1269
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   42
         Left            =   -70095
         TabIndex        =   211
         Top             =   971
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   41
         Left            =   -70095
         TabIndex        =   210
         Top             =   673
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   40
         Left            =   -74475
         TabIndex        =   209
         Top             =   1269
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   39
         Left            =   -74475
         TabIndex        =   208
         Top             =   971
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   38
         Left            =   -74475
         TabIndex        =   207
         Top             =   673
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H7"
         Height          =   180
         Index           =   8
         Left            =   -74760
         TabIndex        =   206
         Top             =   375
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H8"
         Height          =   180
         Index           =   7
         Left            =   -70380
         TabIndex        =   205
         Top             =   375
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   37
         Left            =   -70095
         TabIndex        =   204
         Top             =   2467
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   36
         Left            =   -70095
         TabIndex        =   203
         Top             =   2163
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   35
         Left            =   -70095
         TabIndex        =   202
         Top             =   1865
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   34
         Left            =   -74475
         TabIndex        =   201
         Top             =   2467
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   33
         Left            =   -74475
         TabIndex        =   200
         Top             =   2163
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   32
         Left            =   -74475
         TabIndex        =   199
         Top             =   1865
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H9"
         Height          =   180
         Index           =   7
         Left            =   -74760
         TabIndex        =   198
         Top             =   1567
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H10"
         Height          =   180
         Index           =   6
         Left            =   -70380
         TabIndex        =   197
         Top             =   1567
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H2"
         Height          =   180
         Index           =   2
         Left            =   -70410
         TabIndex        =   196
         Top             =   345
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H1"
         Height          =   180
         Index           =   1
         Left            =   -74790
         TabIndex        =   195
         Top             =   345
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   3
         Left            =   -74505
         TabIndex        =   194
         Top             =   638
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   4
         Left            =   -74505
         TabIndex        =   193
         Top             =   931
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   5
         Left            =   -74505
         TabIndex        =   192
         Top             =   1224
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   6
         Left            =   -70125
         TabIndex        =   191
         Top             =   638
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   7
         Left            =   -70125
         TabIndex        =   190
         Top             =   931
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   8
         Left            =   -70125
         TabIndex        =   189
         Top             =   1224
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   1
         Left            =   -70125
         TabIndex        =   188
         Top             =   2396
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   2
         Left            =   -70125
         TabIndex        =   187
         Top             =   2103
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   9
         Left            =   -70125
         TabIndex        =   186
         Top             =   1810
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   10
         Left            =   -74505
         TabIndex        =   185
         Top             =   2396
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   11
         Left            =   -74505
         TabIndex        =   184
         Top             =   2103
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   12
         Left            =   -74505
         TabIndex        =   183
         Top             =   1810
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H3"
         Height          =   180
         Index           =   2
         Left            =   -74790
         TabIndex        =   182
         Top             =   1517
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H4"
         Height          =   180
         Index           =   1
         Left            =   -70410
         TabIndex        =   181
         Top             =   1517
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   13
         Left            =   -70125
         TabIndex        =   180
         Top             =   3577
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   14
         Left            =   -70125
         TabIndex        =   179
         Top             =   3275
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   15
         Left            =   -70125
         TabIndex        =   178
         Top             =   2982
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   16
         Left            =   -74505
         TabIndex        =   177
         Top             =   3577
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   17
         Left            =   -74505
         TabIndex        =   176
         Top             =   3275
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   18
         Left            =   -74505
         TabIndex        =   175
         Top             =   2982
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H5"
         Height          =   180
         Index           =   3
         Left            =   -74790
         TabIndex        =   174
         Top             =   2689
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H6"
         Height          =   180
         Index           =   3
         Left            =   -70410
         TabIndex        =   173
         Top             =   2689
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¤u§@®É¼Æ:"
         Height          =   180
         Index           =   5
         Left            =   -69885
         TabIndex        =   172
         Top             =   345
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H10"
         Height          =   180
         Index           =   5
         Left            =   -70395
         TabIndex        =   171
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H9"
         Height          =   180
         Index           =   5
         Left            =   -74775
         TabIndex        =   170
         Top             =   1455
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   30
         Left            =   -74490
         TabIndex        =   169
         Top             =   1710
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   29
         Left            =   -74490
         TabIndex        =   168
         Top             =   1965
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   28
         Left            =   -74490
         TabIndex        =   167
         Top             =   2220
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   27
         Left            =   -70110
         TabIndex        =   166
         Top             =   1710
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   26
         Left            =   -70110
         TabIndex        =   165
         Top             =   1965
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   25
         Left            =   -70110
         TabIndex        =   164
         Top             =   2220
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H8"
         Height          =   180
         Index           =   4
         Left            =   -70395
         TabIndex        =   163
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¥Nªí¤H7"
         Height          =   180
         Index           =   4
         Left            =   -74775
         TabIndex        =   162
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   24
         Left            =   -74490
         TabIndex        =   161
         Top             =   615
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   23
         Left            =   -74490
         TabIndex        =   160
         Top             =   870
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   22
         Left            =   -74490
         TabIndex        =   159
         Top             =   1125
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤¤):"
         Height          =   180
         Index           =   21
         Left            =   -70110
         TabIndex        =   158
         Top             =   615
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(­^):"
         Height          =   180
         Index           =   20
         Left            =   -70110
         TabIndex        =   157
         Top             =   870
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(¤é):"
         Height          =   180
         Index           =   19
         Left            =   -70110
         TabIndex        =   156
         Top             =   1125
         Width           =   345
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H5 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   155
         Top             =   2292
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H4 :"
         Height          =   180
         Left            =   -70470
         TabIndex        =   154
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H3 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   153
         Top             =   2022
         Width           =   720
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H2 :"
         Height          =   180
         Left            =   -70470
         TabIndex        =   152
         Top             =   1755
         Width           =   720
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "¥X¦W¥N²z¤H"
         Height          =   180
         Left            =   -74475
         TabIndex        =   151
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "µo¤å³W¶O¡G"
         Height          =   180
         Left            =   -72360
         TabIndex        =   150
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¯S®í°Ó¼Ð :"
         Height          =   180
         Index           =   7
         Left            =   -69270
         TabIndex        =   149
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_¿é¤JD/N :"
         Height          =   180
         Left            =   -69720
         TabIndex        =   148
         Top             =   954
         Width           =   1095
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "(Y:¿é¤J)"
         Height          =   180
         Left            =   -67770
         TabIndex        =   147
         Top             =   954
         Width           =   645
      End
      Begin VB.Label Label42 
         Caption         =   "®×¥ó¦WºÙ :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   146
         Top             =   2655
         Width           =   1335
      End
      Begin VB.Label Label33 
         Caption         =   "(N:¤£ºâ)"
         Height          =   255
         Left            =   -67200
         TabIndex        =   145
         Top             =   1380
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_¦C¦LÂ½Ä¶¨ç :"
         Height          =   180
         Left            =   -69750
         TabIndex        =   144
         Top             =   1224
         Width           =   1350
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "(N:¤£¦L)"
         Height          =   180
         Left            =   -67950
         TabIndex        =   143
         Top             =   1224
         Width           =   645
      End
      Begin VB.Line Line2 
         X1              =   -70320
         X2              =   -68250
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label38 
         Caption         =   "¬O§_¸É¥ó(¥i½Æ¿ï) :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   142
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "°Ó«~Ãþ§O :"
         Height          =   252
         Index           =   14
         Left            =   -74880
         TabIndex        =   141
         Top             =   360
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "°Ó«~²Õ¸s :"
         Height          =   252
         Index           =   13
         Left            =   -74880
         TabIndex        =   140
         Top             =   660
         Width           =   852
      End
      Begin VB.Label Label41 
         Caption         =   "(Y:¶l±H)"
         Height          =   252
         Left            =   -72960
         TabIndex        =   139
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label40 
         Caption         =   "¬O§_¶l±H¥Ó½Ð :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   138
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¶Ê¼f´Á­­ :"
         Height          =   180
         Index           =   6
         Left            =   -74910
         TabIndex        =   137
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label44 
         Caption         =   "¥N²z¤H :"
         Height          =   252
         Left            =   120
         TabIndex        =   136
         Top             =   -360
         Width           =   972
      End
      Begin VB.Line Line1 
         X1              =   -72360
         X2              =   -72240
         Y1              =   1008
         Y2              =   1008
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "©µ®i«á±M¥Î´Á­­ :"
         Height          =   180
         Index           =   31
         Left            =   -74910
         TabIndex        =   135
         Top             =   954
         Width           =   1350
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "¦C¦L©w½Z :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   134
         Top             =   1224
         Width           =   810
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "(N:¤£¦L)"
         Height          =   180
         Left            =   -72840
         TabIndex        =   133
         Top             =   1224
         Width           =   645
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð¤H1 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   132
         Top             =   1752
         Width           =   720
      End
      Begin VB.Label Label48 
         Caption         =   "®×¥ó¤é¤å¦WºÙ :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   131
         Top             =   3210
         Width           =   1455
      End
      Begin VB.Label Label49 
         Caption         =   "®×¥ó­^¤å¦WºÙ :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   130
         Top             =   2925
         Width           =   1215
      End
      Begin VB.Label Label50 
         Caption         =   "®×¥ó¤¤¤å¦WºÙ :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   129
         Top             =   2655
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "°Ó¼ÐºØÃþ :"
         Height          =   180
         Index           =   8
         Left            =   -72420
         TabIndex        =   128
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "¥¿°Ó¼Ð¸¹¼Æ:"
         Height          =   255
         Index           =   15
         Left            =   -66480
         TabIndex        =   127
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ÂI¼Æ :"
         Height          =   180
         Index           =   16
         Left            =   -68340
         TabIndex        =   126
         Top             =   345
         Width           =   450
      End
      Begin VB.Label Label51 
         Caption         =   "¬O§_ºâ®×¥ó¼Æ :"
         Height          =   255
         Left            =   -67410
         TabIndex        =   125
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label52 
         Caption         =   "¬d¦W¥»©Ò®×¸¹ :"
         Height          =   255
         Left            =   -71910
         TabIndex        =   124
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label53 
         Caption         =   "¶i«×³Æµù :"
         Height          =   255
         Left            =   -74865
         TabIndex        =   123
         Top             =   1890
         Width           =   975
      End
      Begin VB.Label Label54 
         Caption         =   "®×¥ó³Æµù :"
         Height          =   255
         Left            =   -74835
         TabIndex        =   122
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "¤u§@®É¼Æ:"
         Height          =   180
         Index           =   12
         Left            =   4815
         TabIndex        =   121
         Top             =   390
         Width           =   765
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "¨üÅý¤H5 :"
         Height          =   180
         Left            =   90
         TabIndex        =   120
         Top             =   2730
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "¨üÅý¤H4 :"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   119
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "¨üÅý¤H3 :"
         Height          =   180
         Left            =   90
         TabIndex        =   118
         Top             =   2205
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "¨üÅý¤H2 :"
         Height          =   180
         Left            =   90
         TabIndex        =   117
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "¥X¦W¥N²z¤H"
         Height          =   180
         Left            =   6600
         TabIndex        =   116
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "µo¤å³W¶O¡G"
         Height          =   180
         Left            =   2310
         TabIndex        =   115
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_¸É¥ó(¥i½Æ¿ï) :"
         Height          =   180
         Left            =   60
         TabIndex        =   114
         Top             =   960
         Width           =   1470
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "(1:¨üÅý¤H©e¥ôª¬ 2:²¾Âà«´¬ù®Ñ 3:¨üÅý¤Hªk¤HÃÒ©ú 4:µù¥UÃÒ)"
         Height          =   180
         Left            =   2460
         TabIndex        =   113
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "(Y:¿é¤J)"
         Height          =   180
         Left            =   1650
         TabIndex        =   112
         Top             =   690
         Width           =   645
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_¿é¤JD/N :"
         Height          =   180
         Left            =   60
         TabIndex        =   111
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "¶i«×³Æµù :"
         Height          =   180
         Left            =   90
         TabIndex        =   110
         Top             =   3060
         Width           =   810
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "µo¤å¤é :"
         Height          =   180
         Left            =   60
         TabIndex        =   109
         Top             =   390
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "¶Ê¼f´Á­­ :"
         Height          =   180
         Index           =   0
         Left            =   4320
         TabIndex        =   108
         Top             =   690
         Width           =   810
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "¦C¦L©w½Z :"
         Height          =   180
         Left            =   2370
         TabIndex        =   107
         Top             =   690
         Width           =   810
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:¤£¦L)"
         Height          =   180
         Left            =   3630
         TabIndex        =   106
         Top             =   690
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ÂI¡@¡@¼Æ :"
         Height          =   180
         Index           =   10
         Left            =   6630
         TabIndex        =   105
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "¤£¤@¨Ö²¾Âàµù¥U¸¹¼Æ :"
         Height          =   180
         Left            =   60
         TabIndex        =   104
         Top             =   1230
         Width           =   1710
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "¨üÅý¤H1 :"
         Height          =   180
         Left            =   90
         TabIndex        =   103
         Top             =   1680
         Width           =   720
      End
   End
   Begin MSForms.TextBox textTM81 
      Height          =   285
      Left            =   930
      TabIndex        =   223
      Top             =   1830
      Width           =   2085
      VariousPropertyBits=   671105055
      Size            =   "3678;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM80 
      Height          =   285
      Left            =   3900
      TabIndex        =   222
      Top             =   1830
      Width           =   1935
      VariousPropertyBits=   671105055
      Size            =   "3413;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM79 
      Height          =   285
      Left            =   6720
      TabIndex        =   221
      Top             =   1536
      Width           =   2085
      VariousPropertyBits=   671105055
      Size            =   "3678;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM78 
      Height          =   285
      Left            =   3900
      TabIndex        =   220
      Top             =   1536
      Width           =   1935
      VariousPropertyBits=   671105055
      Size            =   "3413;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   930
      TabIndex        =   219
      Top             =   1536
      Width           =   2085
      VariousPropertyBits=   671105055
      Size            =   "3678;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44 
      Height          =   285
      Left            =   6720
      TabIndex        =   218
      Top             =   1242
      Width           =   2085
      VariousPropertyBits=   671105055
      Size            =   "3678;503"
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   930
      TabIndex        =   217
      Top             =   2160
      Width           =   7875
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13891;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   3900
      TabIndex        =   216
      Top             =   1242
      Width           =   1935
      VariousPropertyBits=   671105055
      Size            =   "3413;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   930
      TabIndex        =   215
      TabStop         =   0   'False
      Top             =   1242
      Width           =   2085
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3678;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H5 :"
      Height          =   180
      Left            =   3060
      TabIndex        =   57
      Top             =   1882
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H4 :"
      Height          =   180
      Left            =   90
      TabIndex        =   56
      Top             =   1882
      Width           =   720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H3 :"
      Height          =   180
      Left            =   5880
      TabIndex        =   55
      Top             =   1588
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H2 :"
      Height          =   180
      Left            =   3060
      TabIndex        =   54
      Top             =   1588
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "¥N²z¤H :"
      Height          =   180
      Left            =   5880
      TabIndex        =   53
      Top             =   1294
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤H1 :"
      Height          =   180
      Left            =   90
      TabIndex        =   52
      Top             =   1588
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "°Ó¼ÐºØÃþ :"
      Height          =   180
      Index           =   4
      Left            =   5880
      TabIndex        =   51
      Top             =   1000
      Width           =   810
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "©Ó¿ì¤H :"
      Height          =   180
      Left            =   90
      TabIndex        =   50
      Top             =   1294
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "·~°È°Ï§O :"
      Height          =   180
      Index           =   2
      Left            =   3060
      TabIndex        =   49
      Top             =   360
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "µoÃÒ¤é :"
      Height          =   180
      Index           =   3
      Left            =   5880
      TabIndex        =   48
      Top             =   360
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð®×¸¹ :"
      Height          =   180
      Left            =   5880
      TabIndex        =   47
      Top             =   706
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥»©Ò®×¸¹ :"
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   46
      Top             =   706
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¦¬¤å¸¹ :"
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   45
      Top             =   360
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó©Ê½è :"
      Height          =   180
      Index           =   6
      Left            =   90
      TabIndex        =   44
      Top             =   1000
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "©¼©Ò®×¸¹ :"
      Height          =   180
      Index           =   9
      Left            =   3060
      TabIndex        =   43
      Top             =   1000
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "´¼Åv¤H­û :"
      Height          =   180
      Index           =   11
      Left            =   3060
      TabIndex        =   42
      Top             =   1294
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¼f©w¸¹¼Æ :"
      Height          =   180
      Left            =   3060
      TabIndex        =   41
      Top             =   706
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó¦WºÙ :"
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   40
      Top             =   2220
      Width           =   810
   End
End
Attribute VB_Name = "frm030202_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/02 §ï¦¨Form2.0 ;cmbTM05¡BtextCP13¡BtextCP14¡BtextCP64¡BtextTM44¡BtextTM23¡BtextTM78~81¡BtextCP56_2¡BtextCP89_2¡BtextCP90_2¡BtextCP91_2¡BtextCP92_2¡BlstNameAgent¡BCombo2(index)¡BtextTM47~52¡BtextTM94~117
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
'©Ó¿ì¤H Add By Sindy 98/03/11
Dim m_CP14 As String
Dim m_CP82 As String 'Added by Lydia 2018/08/10 µo¤å®É¶¡

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
'Add By Cheng 2003/03/07
Dim m_CP55 As String '­ì²¾Âà¤H
Dim m_TM23 As String '­ì¥Ó½Ð¤H
'add by nickc 2007/01/29
Dim m_CP93 As String
Dim m_CP94 As String
Dim m_CP95 As String
Dim m_CP96 As String
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
'Add By Sindy 2010/4/15
Dim m_TM08 As String
Dim m_TM58 As String
'2010/4/15 End

'Add By Cheng 2003/10/06
Public m_blnClkChgButton As Boolean '¬O§_«ö¤UÅÜ§ó¨Æ¶µ«ö¶s
'add by nick 2004/08/13
Dim m_CP84 As String       'µo¤å³W¶O
'add by nickc 2006/01/26
Dim m_CP110 As String
'add by nickc 2008/02/22
Dim m_CP44 As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP09s As String, m_CP123s As String 'Add by Sindy 98/3/24 ¦¬¤å¸¹,¬O§_ºâµo¤å«Ç®×¥ó
Dim m_CP130s As String 'Add by Sindy 2009/4/24 µo¤å-¥DºÞ¾÷Ãö
Dim m_IsSend As Boolean 'Add By Sindy 2012/8/10 ¬O§_¸gµo¤å«Çµo¤å
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer


Private Sub cmdCancel_Click()
   frm030202_01.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
   Unload frm030202_01
   Unload Me
   'frm030202_01.Show
End Sub

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
   frm030202_05.SetParent "frm030202_08"
   'Me.Hide
   frm030202_05.Show
   frm030202_05.QueryData
'    m_blnClkChgButton = True
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bolIsCaseNum As Boolean 'Add By Sindy 2018/8/3
   
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
                  'Modify By Sindy 2018/8/3 + bolIsCaseNum:¬O§_ºâµo¤å«Ç¥ó¼Æ
                  If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27, , True, bolIsCaseNum) = False Then
                     Exit Sub
                  End If
                  'end 2016/5/16
                  
                  'Add By Sindy 2018/8/3 ¦]¦³¤@¤å¦h®×ªº°ÝÃD¡A¦ý­Y¹q¤l°e¥ó§¡¤£¸gµo¤å«Ç
                  '                      ´NµLªk§PÂ_¬O§_­n¦L©w½Z, §ï§PÂ_¬O§_ºâµo¤å«Ç¥ó¼Æ
                  'Modify By Sindy 2018/11/23 if ®³±¼ Or GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "3" §PÂ_
                  'Modify Sindy 2019/1/17 Mark,¹q¤l°e¥ó¤@©w­n¥X©w½Z,¤£¥Î¯S§O§PÂ_
'                  If bolIsCaseNum = True Then
                     m_IsSend = True
'                  Else
'                     m_IsSend = False
'                  End If
'                  'ªü½¬»¡­n¼W¥[§PÂ_¬O­^¤å©w½Z¤~»Ý­n¸ß°Ý
'                  'Modify By Sindy 2018/11/23 ¤é¤å©w½Z¤]­n¸ß°Ý
'                  If bolIsCaseNum = True And Trim(textPrint.Text) = "" And _
'                     (GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "2" Or GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "3") Then
'                     If MsgBox("¬O§_»Ý­n¦C¦L©w½Z¡H", vbExclamation + vbYesNo) = vbNo Then
'                        textPrint.Text = "N"
'                     End If
'                  End If
'                  '2018/8/3 End
                  
                  'add by sonia 2016/3/31
                  strExc(0) = Trim(InputBox("½Ð¿é¤J´¼¼z§½¦¬¤å¤å¸¹!!"))
                  If strExc(0) = "" Then
                     Exit Sub
                  Else
                     textCP64 = "´¼¼z§½¦¬¤å¤å¸¹:" & strExc(0) & ";" & Trim(textCP64)
                  End If
                  'end 2016/3/31
               Else
                  'Add by Sindy 2009/4/24
                  If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
                     Exit Sub
                  Else
                     'Add By Sindy 2012/8/6 ¦]¦³¤@¤å¦h®×ªº°ÝÃD¡A©Ò¥H­Y¸gµo¤å«Ç¥B§@·~µe­±¤W¬°¦C¦L©w½Z®É,¸ß°Ý¨Ï¥ÎªÌ
                     '                      ¤é¤å¤@©w­n¦C¦L©w½Z
                     'Modify By Sindy 2018/11/23 if ®³±¼ Or (m_CP123s = "N" And GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "3") §PÂ_
                     If m_CP123s = "Y" Then
                        m_IsSend = True
                     Else
                        m_IsSend = False
                     End If
                     'Modify By Sindy 2012/8/31 ªü½¬»¡­n¼W¥[§PÂ_¬O­^¤å©w½Z¤~»Ý­n¸ß°Ý
                     'Modify By Sindy 2018/11/23 ¤é¤å©w½Z¤]­n¸ß°Ý
                     If m_CP123s = "Y" And Trim(textPrint.Text) = "" And _
                        (GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "2" Or GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "3") Then
                        If MsgBox("¬O§_»Ý­n¦C¦L©w½Z¡H" & vbCrLf & "(¤@¤å¦h®×»Ý¨ì<©w½Z¸ê®ÆºûÅ@>²£¥X©w½Z)", vbExclamation + vbYesNo) = vbNo Then
                           textPrint.Text = "N"
                        End If
                     End If
                     '2012/8/6 End
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
            'Add By Cheng 2003/01/15
            '§PÂ_¬O¦C¦L©w½Z
            If Me.textPrint.Text <> "N" Then
               If m_IsSend = True Then 'Modify By Sindy 2012/8/3 +if ¦]¦³¤@¤å¦h®×ª¬ªp¡A©Ò¥H¼W¥[§PÂ_¸gµo¤å«Ç®É¤~»Ý¥X©w½Z
                  PrintLetter
               End If
            End If
            ' ³]©w·Æ¹«´å¼Ð¬°¹w³]
            Screen.MousePointer = vbDefault
            
            'Add By Sindy 2012/4/5 CFT,FCT©Ò¦³®×¥ó©Ê½èµo¤å®É,ÀË¬d¥Nªí¹Ï¬O§_¦s¦b
            'Mark by Amy 2018/07/31 ¦]ChkIsExistImg¤£¨Ï¥Î,»PSindy½T»{FCT¤£¼uMsg¬G®³±¼
            'Call ChkIsExistImg(m_TM01, m_TM02, m_TM03, m_TM04)

            'Added by Lyddia 2018/08/10 ¼W¥[­«·sµo¤å§PÂ_
            strExc(1) = m_CP82
            If Val(m_CP82) > 0 Then
                 If MsgBox("­«·sµo¤å¬O§_¤W¶ÇÀÉ®×¨ì¨÷©v°Ï¡H", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                     strExc(1) = ""
                 End If
            End If
            If Val(strExc(1)) = 0 Then
            'end 2018/08/10
                'Added by Lydia 2018/07/19 FCTµo¤å¦Û°Ê±N¤U¸üªºPDFÀÉ,¤W¶Ç¨ì¨÷©v°Ï
                If Pub_AutoSavePdf_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_CP10) = False Then
                End If
                'end 2018/07/19
            End If 'end 2018/08/10
            
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
      'edit by nick 2004/11/17  ¦]¬°½Ð´Ú¤w¸g¦³²£¥Í¤F
      '        pub_AddressListSN = pub_AddressListSN + 1
      '        PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
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
   textTM08.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   'add by nickc 2007/01/29
   textTM78.BackColor = &H8000000F
   textTM79.BackColor = &H8000000F
   textTM80.BackColor = &H8000000F
   textTM81.BackColor = &H8000000F
   
   textTM44.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   textCP56_2.BackColor = &H8000000F
   'add by nickc 2007/01/29
   textCP89_2.BackColor = &H8000000F
   textCP90_2.BackColor = &H8000000F
   textCP91_2.BackColor = &H8000000F
   textCP92_2.BackColor = &H8000000F
   
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
'    m_blnClkChgButton = False
   'Add by nickc 2006/01/26
   '¥xÆW¥[¥X¦W¥N²z¤H²M³æ¨Ñ¤Ä¿ï,­ì¬O§_¥X¦WÄæ¦ì¤£Åã¥Ü
   Text7.Visible = False
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Added by Lydia 2021/09/02 ¦pªG¤@¶}©l±NListBox©Ô¨ì»Ý­nªº¤j¤p¡A¦r«¬·|¦Û°Ê©ñ¤j¡F©Ò¥Hµe­±¹w³]¬°¤@¦C°ª«×¡AForm_Load¤~©ñ¤j¨ì»Ý­nªº¤j¤p
   lstNameAgent.Height = 885
   lstNameAgent.Width = 1260
   Me.SSTab1.Tab = 0
   
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
      m_TM119 = CheckStr(rsTmp.Fields("TM119"))
      m_TM120 = CheckStr(rsTmp.Fields("TM120"))
      ' ¼f©w¸¹¼Æ
      If IsNull(rsTmp.Fields("TM15")) = False Then: textTM15 = rsTmp.Fields("TM15")
      ' ¥Ó½Ð®×¸¹
      If IsNull(rsTmp.Fields("TM12")) = False Then: textTM12 = rsTmp.Fields("TM12")
      ' µoÃÒ¤é
      If IsNull(rsTmp.Fields("TM20")) = False Then: textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
      ' ®×¥ó¤¤¤å¦WºÙ
      If IsNull(rsTmp.Fields("TM05")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM05")
      ' ®×¥ó­^¤å¦WºÙ
      If IsNull(rsTmp.Fields("TM06")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM06")
      ' ®×¥ó¤é¤å¦WºÙ
      If IsNull(rsTmp.Fields("TM07")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM07")
      ' Åã¥Ü®×¥ó¦WºÙ
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      
      'Add By Sindy 2010/4/15
      '®×¥ó³Æµù
      m_TM58 = "" & rsTmp.Fields("TM58")
      
      ' °Ó¼ÐºØÃþ
      m_TM08 = "" & rsTmp.Fields("TM08") 'Add By Sindy 2010/4/15
      textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      ' ¥Ó½Ð°ê®a
      If IsNull(rsTmp.Fields("TM10")) = False Then: m_TM10 = rsTmp.Fields("TM10")
      ' ¥Ó½Ð¤H
      If IsNull(rsTmp.Fields("TM23")) = False Then: textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      'add by nickc 2007/01/29
      If IsNull(rsTmp.Fields("TM78")) = False Then: textTM78 = GetCustomerName("" & rsTmp.Fields("TM78"), 0)
      If IsNull(rsTmp.Fields("TM79")) = False Then: textTM79 = GetCustomerName("" & rsTmp.Fields("TM79"), 0)
      If IsNull(rsTmp.Fields("TM80")) = False Then: textTM80 = GetCustomerName("" & rsTmp.Fields("TM80"), 0)
      If IsNull(rsTmp.Fields("TM81")) = False Then: textTM81 = GetCustomerName("" & rsTmp.Fields("TM81"), 0)
        'Add By Cheng 2003/03/07
        '°O¿ý¥Ó½Ð¤H
        m_TM23 = "" & rsTmp("TM23").Value
        'add by nickc 2007/01/29
        m_TM78 = "" & rsTmp("TM78").Value
        m_TM79 = "" & rsTmp("TM79").Value
        m_TM80 = "" & rsTmp("TM80").Value
        m_TM81 = "" & rsTmp("TM81").Value
        
      ' FC¥N²z¤H
      If IsNull(rsTmp.Fields("TM44")) = False Then: textTM44 = GetFAgentName(rsTmp.Fields("TM44"))
      ' ©¼©Ò®×¸¹
      If IsNull(rsTmp.Fields("TM45")) = False Then: textTM45 = rsTmp.Fields("TM45")
      
      'Add By Sindy 2010/3/3
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
      textTM108 = Empty
      If IsNull(rsTmp.Fields("TM108")) = False Then: textTM108 = rsTmp.Fields("TM108")
      SetTMSPFieldOldData "TM108", textTM108, 0
      TextTM109 = Empty
      If IsNull(rsTmp.Fields("TM109")) = False Then: TextTM109 = rsTmp.Fields("TM109")
      SetTMSPFieldOldData "TM109", TextTM109, 0
      TextTM110 = Empty
      If IsNull(rsTmp.Fields("TM110")) = False Then: TextTM110 = rsTmp.Fields("TM110")
      SetTMSPFieldOldData "TM110", TextTM110, 0
      textTM111 = Empty
      If IsNull(rsTmp.Fields("TM111")) = False Then: textTM111 = rsTmp.Fields("TM111")
      SetTMSPFieldOldData "TM111", textTM111, 0
      TextTM112 = Empty
      If IsNull(rsTmp.Fields("TM112")) = False Then: TextTM112 = rsTmp.Fields("TM112")
      SetTMSPFieldOldData "TM112", TextTM112, 0
      TextTM113 = Empty
      If IsNull(rsTmp.Fields("TM113")) = False Then: TextTM113 = rsTmp.Fields("TM113")
      SetTMSPFieldOldData "TM113", TextTM113, 0
      textTM114 = Empty
      If IsNull(rsTmp.Fields("TM114")) = False Then: textTM114 = rsTmp.Fields("TM114")
      SetTMSPFieldOldData "TM114", textTM114, 0
      TextTM115 = Empty
      If IsNull(rsTmp.Fields("TM115")) = False Then: TextTM115 = rsTmp.Fields("TM115")
      SetTMSPFieldOldData "TM115", TextTM115, 0
      TextTM116 = Empty
      If IsNull(rsTmp.Fields("TM116")) = False Then: TextTM116 = rsTmp.Fields("TM116")
      SetTMSPFieldOldData "TM116", TextTM116, 0
      textTM117 = Empty
      If IsNull(rsTmp.Fields("TM117")) = False Then: textTM117 = rsTmp.Fields("TM117")
      SetTMSPFieldOldData "TM117", textTM117, 0
      '2010/3/3 End
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' ¨ú±o®×¥ó¶i«×ÀÉªºÄæ¦ì¤º®e
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim strSubSQL As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSubTmp As New ADODB.Recordset
   Dim strDate As String
   Dim strCP27 As String
   Dim strCP44 As String
   Dim strCP45 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
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
      m_CP82 = "" & rsTmp.Fields("CP82")  'Added by Lydia 2018/08/10 µo¤å®É¶¡
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
      If IsNull(rsTmp.Fields("CP13")) = False Then: textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      
      'Add By Sindy 98/03/11
      '¤u§@®É¼Æ
      textCP113 = "" & rsTmp.Fields("CP113")
      SetCPFieldOldData "CP113", textCP113, 1
      '©Ó¿ì¤H
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
      ' ÂI¼Æ
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then: textCP18 = rsTmp.Fields("CP18")
        
      ' ²¾Âà¥Ó½Ð¤H
      textCP56 = Empty
      If IsNull(rsTmp.Fields("CP56")) = False Then
         textCP56 = rsTmp.Fields("CP56")
         textCP56_Validate False
      End If
      SetCPFieldOldData "CP56", textCP56, 0
      
      'add by nickc 2007/01/29
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
      
      'Modify By Sindy 2012/12/27
'      'Add By Cheng 2003/03/07
'      '°O¿ý²¾Âà¤H
'      m_CP55 = "" & rsTmp("CP55").Value
'      'add by nickc 2007/01/29
'      m_CP93 = "" & rsTmp("CP93").Value
'      m_CP94 = "" & rsTmp("CP94").Value
'      m_CP95 = "" & rsTmp("CP95").Value
'      m_CP96 = "" & rsTmp("CP96").Value
      m_CP55 = Empty
      If IsNull(rsTmp.Fields("CP55")) = False Then
         m_CP55 = rsTmp.Fields("CP55")
      End If
      SetCPFieldOldData "CP55", m_CP55, 0
      m_CP93 = Empty
      If IsNull(rsTmp.Fields("CP93")) = False Then
         m_CP93 = rsTmp.Fields("CP93")
      End If
      SetCPFieldOldData "CP93", m_CP93, 0
      m_CP94 = Empty
      If IsNull(rsTmp.Fields("CP94")) = False Then
         m_CP94 = rsTmp.Fields("CP94")
      End If
      SetCPFieldOldData "CP94", m_CP94, 0
      m_CP95 = Empty
      If IsNull(rsTmp.Fields("CP95")) = False Then
         m_CP95 = rsTmp.Fields("CP95")
      End If
      SetCPFieldOldData "CP95", m_CP95, 0
      m_CP96 = Empty
      If IsNull(rsTmp.Fields("CP96")) = False Then
         m_CP96 = rsTmp.Fields("CP96")
      End If
      SetCPFieldOldData "CP96", m_CP96, 0
      '2012/12/27 End
      
      ' ¶i«×³Æµù
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then: textCP64 = rsTmp.Fields("CP64")
      SetCPFieldOldData "CP64", textCP64, 0
      
      'Add By Sindy 2010/3/3
      '¥Nªí¤H
      Dim i As Integer, j As Integer
      For i = 0 To 9
         Combo2(i).AddItem ""
      Next
      If rsTmp.Fields("CP56").Value <> "" Then
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("CP56").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(0).AddItem rsTmp.Fields("CP56").Value & "-" & j & strExc(0)
               Combo2(1).AddItem rsTmp.Fields("CP56").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      If rsTmp.Fields("CP89").Value <> "" Then
         '­^->¤¤->¤é
         strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("CP89").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(2).AddItem rsTmp.Fields("CP89").Value & "-" & j & strExc(0)
               Combo2(3).AddItem rsTmp.Fields("CP89").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      If rsTmp.Fields("CP90").Value <> "" Then
         '­^->¤¤->¤é
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("CP90").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(4).AddItem rsTmp.Fields("CP90").Value & "-" & j & strExc(0)
               Combo2(5).AddItem rsTmp.Fields("CP90").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      If rsTmp.Fields("CP91").Value <> "" Then
         '­^->¤¤->¤é
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("CP91").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(6).AddItem rsTmp.Fields("CP91").Value & "-" & j & strExc(0)
               Combo2(7).AddItem rsTmp.Fields("CP91").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      If rsTmp.Fields("CP92").Value <> "" Then
         '­^->¤¤->¤é
         strExc(0) = "SELECT nvl(CU40,nvl(cu39,cu41)),nvl(CU43,nvl(cu42,cu44)),nvl(CU46,nvl(cu45,cu47)),nvl(CU49,nvl(cu48,cu50)),nvl(CU52,nvl(cu51,cu53)),nvl(CU55,nvl(cu54,cu56)) FROM CUSTOMER WHERE " & ChgCustomer(rsTmp.Fields("CP92").Value)
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            For j = 1 To 6
               If IsNull(RsTemp.Fields(j - 1)) Then
                  strExc(0) = ""
               Else
                  strExc(0) = "-" & RsTemp.Fields(j - 1)
               End If
               Combo2(8).AddItem rsTmp.Fields("CP92").Value & "-" & j & strExc(0)
               Combo2(9).AddItem rsTmp.Fields("CP92").Value & "-" & j & strExc(0)
            Next
         End If
      End If
      '2010/3/3 End
      
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
   'SetCPFieldOldData "CP110", m_CP110, 0
   'Modify By Sindy 2010/9/20
   If m_CP110 = "" Then m_CP110 = CheckStr(rsTmp.Fields("cp110"))
   SetCPFieldOldData "CP110", CheckStr(rsTmp.Fields("cp110")), 0
   '2010/9/20 End
   
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' Åª¨ú¸ê®Æ®w
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   'add by nickc 2006/01/26
   Dim tm(1 To 4) As String
   ' ¥ý²M°£°Ó¼Ð°ò¥»ÀÉ©ÎªA°È·~°È°ò¥»ÀÉÄæ¦ì¦ê¦C
   ClearTMSPFieldList
   ' ¥ý²M°£®×¥ó¶i«×ÀÉÄæ¦ì¦ê¦C
   ClearCPFieldList
   
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
   
   ' ¥»©Ò®×¸¹
'   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)
   
   'add by nickc 2006/01/26
   tm(1) = m_TM01
   tm(2) = m_TM02
   tm(3) = m_TM03
   tm(4) = m_TM04
   'Modify By Sindy 2010/9/20 ¹w³]¥X¦W¥N²z¤H,²¾¨ì¤U­±Åª§¹CP¦A°µ
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110
   '2010/9/20 End
   
   'Åª¨ú°Ó¼Ð°ò¥»ÀÉ
   QueryTradeMark
   
   ' Åª¨ú®×¥ó¶i«×ÀÉ
   QueryCaseProgress
   'Modified by Lydia 2021/09/02 + Form 2.0 = True
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True 'Modify By Sindy 2010/9/20
   
   ' ¨ú±o¶Ê¼f´Á­­ªº¤é´Á
   textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
   Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 °O¿ýµo¤å¤é¡A¦³­×§ïµo¤å¤é¦A­«·s­pºâ¶Ê¼f´Á­­
   
   'Add By Sindy 2012/12/20 ¥~°Ó000¥xÆW®×©Ò¦³®×¥ó©Ê½è¥[¹q¤l°e¥ó¥\¯à
   If m_TM01 = "FCT" And m_TM10 = "000" Then
      Label43.Visible = True
      textCP118.Visible = True
   Else
      Label43.Visible = False
      textCP118.Visible = False
   End If
   '2012/12/20 End
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030202_08 = Nothing
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
      Me.SSTab1.Tab = 0
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
      Select Case strTemp
        'Modify By Cheng 2003/03/13
'         Case "1", "2", "3", "4", "5":
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
   
EXITSUB:
End Sub
'add by nickc 2007/01/29
Private Sub textCP56_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' ¨üÅý¤H
Private Sub textCP56_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCP56_2 = Empty
   If IsEmptyText(textCP56) = False Then
      'edit by 2004/07/22 nick  ÀË¬d¸Ó¥Ó½Ð¤H©Î¥N²z¤Hª¬ºA¡A­Y¬°¤£¦A¨Ï¥Î«h°±¦b­ì¦a
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
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¨üÅý¤H¥N½X<" & textCP56 & ">¤£¦s¦b"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP56_GotFocus
      End If
   End If
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

      ' ¨ú±o¶Ê¼f´Á­­ªº¤é´Á
      If Me.textCP27.Tag <> Me.textCP27.Text Then 'Added by Lydia 2019/11/08 °O¿ýµo¤å¤é¡A¦³­×§ïµo¤å¤é¦A­«·s­pºâ¶Ê¼f´Á­­
          textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
      End If
      Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 °O¿ýµo¤å¤é¡A¦³­×§ïµo¤å¤é¦A­«·s­pºâ¶Ê¼f´Á­­
   End If
EXITSUB:
End Sub

'edit by nickc 2006/01/26
'Private Sub textCP64_2_GotFocus()
'   TextInverse textCP64_2
'End Sub


'add by nick 2004/08/13
Private Sub textCP84_GotFocus()
   Me.SSTab1.Tab = 0
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
'add by nickc 2007/01/29
Private Sub textCP89_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textCP89
End Sub
Private Sub textCP89_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'add by nickc 2007/01/29
Private Sub textCP89_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCP89_2 = Empty
   If IsEmptyText(textCP89) = False Then
      Dim oState As Boolean
      oState = True
      textCP89_2 = GetCustomerNameAndState(textCP89, "0", oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If textCP89_2 = Empty Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¨üÅý¤H¥N½X<" & textCP89 & ">¤£¦s¦b"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP89_GotFocus
      End If
   End If
End Sub
Private Sub textCP90_GotFocus()
   Me.SSTab1.Tab = 0
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
      Dim oState As Boolean
      oState = True
      textCP90_2 = GetCustomerNameAndState(textCP90, "0", oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If textCP90_2 = Empty Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¨üÅý¤H¥N½X<" & textCP90 & ">¤£¦s¦b"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP90_GotFocus
      End If
   End If
End Sub
Private Sub textCP91_GotFocus()
   Me.SSTab1.Tab = 0
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
      Dim oState As Boolean
      oState = True
      textCP91_2 = GetCustomerNameAndState(textCP91, "0", oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If textCP91_2 = Empty Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¨üÅý¤H¥N½X<" & textCP91 & ">¤£¦s¦b"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP91_GotFocus
      End If
   End If
End Sub
Private Sub textCP92_GotFocus()
   Me.SSTab1.Tab = 0
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
      Dim oState As Boolean
      oState = True
      textCP92_2 = GetCustomerNameAndState(textCP92, "0", oState)
      If oState = False Then
            Cancel = True
            Exit Sub
      End If
      If textCP92_2 = Empty Then
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "¨üÅý¤H¥N½X<" & textCP92 & ">¤£¦s¦b"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP92_GotFocus
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
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "¸ê®ÆÀË®Ö"
            strMsg = "¥u¥i¿é¤JªÅ¥Õ©ÎN"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

' §ó·sÄæ¦ìªº¤º®e
Private Sub OnUpdateField()
   ' Add By Sindy 98/03/11
   SetCPFieldNewData "CP113", textCP113
   ' µo¤å¤é
   SetCPFieldNewData "CP27", DBDATE(textCP27)
    'Add By Cheng 2003/03/07
    '­Y¦³¿é¤J²¾Âà¥Ó½Ð¤H
    If Me.textCP56.Text <> "" Then
        '­Y²¾Âà¤H»P­ì¥Ó½Ð¤H¤£¦P®É
        '2009/10/19 modify by sonia À³¬°§PÂ_²¾Âà¥Ó½Ð¤H»P­ì¥Ó½Ð¤H¤£¦P®É
        'If ChangeCustomerL(m_CP55) <> ChangeCustomerL(m_TM23) Then
        If ChangeCustomerL(textCP56) <> ChangeCustomerL(m_TM23) Then
            '§ó·s¶i«×ÀÉ²¾Âà¤H
            SetCPFieldNewData "CP55", ChangeCustomerL(m_TM23)
        End If
    End If
   ' ²¾Âà¥Ó½Ð¤H
   If IsEmptyText(textCP56) = False Then
      SetCPFieldNewData "CP56", textCP56 & String(9 - Len(textCP56), "0")
   Else
      SetCPFieldNewData "CP56", textCP56
   End If
   'add by nickc 2007/01/29
    If Me.textCP89.Text <> "" Then
        '­Y²¾Âà¤H»P­ì¥Ó½Ð¤H¤£¦P®É
        '2009/10/19 modify by sonia À³¬°§PÂ_²¾Âà¥Ó½Ð¤H»P­ì¥Ó½Ð¤H¤£¦P®É
        'If ChangeCustomerL(m_CP93) <> ChangeCustomerL(m_TM78) Then
        If ChangeCustomerL(textCP89) <> ChangeCustomerL(m_TM78) Then
            '§ó·s¶i«×ÀÉ²¾Âà¤H
            SetCPFieldNewData "CP93", ChangeCustomerL(m_TM78)
        End If
    End If
   ' ²¾Âà¥Ó½Ð¤H
   If IsEmptyText(textCP89) = False Then
      SetCPFieldNewData "CP89", textCP89 & String(9 - Len(textCP89), "0")
   Else
      SetCPFieldNewData "CP89", textCP89
   End If
    If Me.textCP90.Text <> "" Then
        '­Y²¾Âà¤H»P­ì¥Ó½Ð¤H¤£¦P®É
        '2009/10/19 modify by sonia À³¬°§PÂ_²¾Âà¥Ó½Ð¤H»P­ì¥Ó½Ð¤H¤£¦P®É
        'If ChangeCustomerL(m_CP94) <> ChangeCustomerL(m_TM79) Then
        If ChangeCustomerL(textCP90) <> ChangeCustomerL(m_TM79) Then
            '§ó·s¶i«×ÀÉ²¾Âà¤H
            SetCPFieldNewData "CP94", ChangeCustomerL(m_TM79)
        End If
    End If
   ' ²¾Âà¥Ó½Ð¤H
   If IsEmptyText(textCP90) = False Then
      SetCPFieldNewData "CP90", textCP90 & String(9 - Len(textCP90), "0")
   Else
      SetCPFieldNewData "CP90", textCP90
   End If
    If Me.textCP91.Text <> "" Then
        '­Y²¾Âà¤H»P­ì¥Ó½Ð¤H¤£¦P®É
        '2009/10/19 modify by sonia À³¬°§PÂ_²¾Âà¥Ó½Ð¤H»P­ì¥Ó½Ð¤H¤£¦P®É
        'If ChangeCustomerL(m_CP95) <> ChangeCustomerL(m_TM80) Then
        If ChangeCustomerL(textCP91) <> ChangeCustomerL(m_TM80) Then
            '§ó·s¶i«×ÀÉ²¾Âà¤H
            SetCPFieldNewData "CP95", ChangeCustomerL(m_TM80)
        End If
    End If
   ' ²¾Âà¥Ó½Ð¤H
   If IsEmptyText(textCP91) = False Then
      SetCPFieldNewData "CP91", textCP91 & String(9 - Len(textCP91), "0")
   Else
      SetCPFieldNewData "CP91", textCP91
   End If
    If Me.textCP92.Text <> "" Then
        '­Y²¾Âà¤H»P­ì¥Ó½Ð¤H¤£¦P®É
        '2009/10/19 modify by sonia À³¬°§PÂ_²¾Âà¥Ó½Ð¤H»P­ì¥Ó½Ð¤H¤£¦P®É
        'If ChangeCustomerL(m_CP96) <> ChangeCustomerL(m_TM81) Then
        If ChangeCustomerL(textCP92) <> ChangeCustomerL(m_TM81) Then
            '§ó·s¶i«×ÀÉ²¾Âà¤H
            SetCPFieldNewData "CP96", ChangeCustomerL(m_TM81)
        End If
    End If
   ' ²¾Âà¥Ó½Ð¤H
   If IsEmptyText(textCP92) = False Then
      SetCPFieldNewData "CP92", textCP92 & String(9 - Len(textCP92), "0")
   Else
      SetCPFieldNewData "CP92", textCP92
   End If
   
   ' ¶i«×³Æµù
   '910801 Sieg 602
    'Modify By Cheng 2003/06/03
    If Me.textTM15_S.Text <> "" Then
        textCP64 = textCP64 & "¤£¤@¨Ö²¾Âàµù¥U¸¹¼Æ : " & textTM15_S
    End If
'edit by nickc     2006/01/26
'   If textCP64_2 <> "" Then
'      If textCP64 = "" Then
'         textCP64 = textCP64_2
'      Else
'         textCP64 = textCP64 & "," & textCP64_2
'      End If
'   End If
   
   SetCPFieldNewData "CP64", textCP64
   
   'add by nickc 2006/01/26
   SetCPFieldNewData "CP110", m_CP110
   'add by nickc 2006/02/10
   SetCPFieldNewData "CP22", Text7
   
   'Add By Sindy 2012/12/20
   ' ¬O§_¹q¤l°e¥ó
   SetCPFieldNewData "CP118", textCP118
   
   'Add By Sindy 2010/3/3
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
   SetTMSPFieldNewData "TM108", textTM108
   SetTMSPFieldNewData "TM109", TextTM109
   SetTMSPFieldNewData "TM110", TextTM110
   SetTMSPFieldNewData "TM111", textTM111
   SetTMSPFieldNewData "TM112", TextTM112
   SetTMSPFieldNewData "TM113", TextTM113
   SetTMSPFieldNewData "TM114", textTM114
   SetTMSPFieldNewData "TM115", TextTM115
   SetTMSPFieldNewData "TM116", TextTM116
   SetTMSPFieldNewData "TM117", textTM117
   '2010/3/3 End
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
            ' 91.03.25 modify by louis (³æ¤Þ¸¹)
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
'Add By Cheng 2003/03/07
Dim StrSQLa As String
Dim rsTmp As New ADODB.Recordset  '2011/5/12 add by sonia
      
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' §ó·s®×¥ó¶i«×ÀÉ
   OnUpdateCaseProperty
   
   'Add By Sindy 2010/3/3
   '§ó·s°Ó¼Ð°ò¥»ÀÉ
   OnUpdateTradeMark
   
    'Add By Cheng 2003/03/07
    '­Y¦³¿é¤J²¾Âà¥Ó½Ð¤H
    If Me.textCP56.Text <> "" Then
        '§ó·s°ò¥»ÀÉ¥Ó½Ð¤H¤Î¬ÛÃö¸ê®Æ
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM23='" & ChangeCustomerL(Me.textCP56.Text) & "' " & _
                            " ,TM24='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "1")) & "' " & _
                            " ,TM25='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "2")) & "' " & _
                            " ,TM26='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "3")) & "' " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
        End Select
    End If
    
    'add by nickc 2007/01/29
    If Me.textCP89.Text <> "" Then
        '§ó·s°ò¥»ÀÉ¥Ó½Ð¤H¤Î¬ÛÃö¸ê®Æ
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM78='" & ChangeCustomerL(Me.textCP89.Text) & "' " & _
                            " ,TM82='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP89.Text), "1")) & "' " & _
                            " ,TM86='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP89.Text), "2")) & "' " & _
                            " ,TM90='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP89.Text), "3")) & "' " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
        End Select
   '2009/12/4 add by sonia T-158071
   Else
        '§ó·s°ò¥»ÀÉ¥Ó½Ð¤H¤Î¬ÛÃö¸ê®Æ
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM78=NULL,TM82=NULL,TM86=NULL,TM90=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
        End Select
    '2009/12/4 END
    End If
    
    If Me.textCP90.Text <> "" Then
        '§ó·s°ò¥»ÀÉ¥Ó½Ð¤H¤Î¬ÛÃö¸ê®Æ
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM79='" & ChangeCustomerL(Me.textCP90.Text) & "' " & _
                            " ,TM83='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP90.Text), "1")) & "' " & _
                            " ,TM87='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP90.Text), "2")) & "' " & _
                            " ,TM91='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP90.Text), "3")) & "' " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
        End Select
   '2009/12/4 add by sonia T-158071
   Else
        '§ó·s°ò¥»ÀÉ¥Ó½Ð¤H¤Î¬ÛÃö¸ê®Æ
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM79=NULL,TM83=NULL,TM87=NULL,TM91=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
        End Select
    '2009/12/4 END
    End If
    
    If Me.textCP91.Text <> "" Then
        '§ó·s°ò¥»ÀÉ¥Ó½Ð¤H¤Î¬ÛÃö¸ê®Æ
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM80='" & ChangeCustomerL(Me.textCP91.Text) & "' " & _
                            " ,TM84='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP91.Text), "1")) & "' " & _
                            " ,TM88='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP91.Text), "2")) & "' " & _
                            " ,TM92='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP91.Text), "3")) & "' " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
        End Select
   '2009/12/4 add by sonia T-158071
   Else
        '§ó·s°ò¥»ÀÉ¥Ó½Ð¤H¤Î¬ÛÃö¸ê®Æ
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM80=NULL,TM84=NULL,TM88=NULL,TM92=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
        End Select
    '2009/12/4 END
    End If
    
    If Me.textCP92.Text <> "" Then
        '§ó·s°ò¥»ÀÉ¥Ó½Ð¤H¤Î¬ÛÃö¸ê®Æ
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM81='" & ChangeCustomerL(Me.textCP92.Text) & "' " & _
                            " ,TM85='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP92.Text), "1")) & "' " & _
                            " ,TM89='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP92.Text), "2")) & "' " & _
                            " ,TM93='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP92.Text), "3")) & "' " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
        End Select
   '2009/12/4 add by sonia T-158071
   Else
        '§ó·s°ò¥»ÀÉ¥Ó½Ð¤H¤Î¬ÛÃö¸ê®Æ
        Select Case m_TM01
        Case "T", "TF", "FCT"
            StrSQLa = "Update TradeMark Set TM81=NULL,TM85=NULL,TM89=NULL,TM93=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
        End Select
    '2009/12/4 END
    End If
    
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
   '                        DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strUserNum & "'," & strNP22 & ")"
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
   End If
    
    'Add By Sindy 2010/7/8 ÀË¬d°Ó«~¸ê®Æ»P°ò¥»ÀÉ°Ó«~Ãþ§O¬O§_¤@­P
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
    
   '2011/5/12 add by sonia §ó·s¤U¤@µ{§Çªº´¼Åv¤H­û
   'If rsTmp.State = 1 Then rsTmp.Close
   'strSql = "SELECT * FROM Customer Where Cu01 = '" & Mid(ChangeCustomerL(textCP56), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP56), 9, 1) & "' "
   'rsTmp.CursorLocation = adUseClient
   'rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'If rsTmp.RecordCount > 0 Then
   '     pub_ChgSalesTargetIsNp m_TM01, m_TM02, m_TM03, m_TM04, PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
   'End If
   pub_ChgSalesTargetIsNp m_TM01, m_TM02, m_TM03, m_TM04, PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
   '2013/4/18 end
   '2011/5/12 end
   
   'Add By Sindy 2012/7/19 ­Y¨üÅý¤H¬°¿ÕµØ¤½¥qªÌ¡A®×¥ó³Æµù­YµL"¤£¾P¨÷"¦r¼Ë,«h­n¥[¤J
   If (textCP56 <> "" And InStr(strTmNovartisCust, Left(textCP56, 6)) > 0) Or _
      (textCP89 <> "" And InStr(strTmNovartisCust, Left(textCP89, 6)) > 0) Or _
      (textCP90 <> "" And InStr(strTmNovartisCust, Left(textCP90, 6)) > 0) Or _
      (textCP91 <> "" And InStr(strTmNovartisCust, Left(textCP91, 6)) > 0) Or _
      (textCP92 <> "" And InStr(strTmNovartisCust, Left(textCP92, 6)) > 0) Then
      strSql = "update trademark" & _
               " set tm58=decode(tm58,null,'" & ChangeTStringToTDateString(strSrvDate(2)) & "¤£¾P¨÷','" & ChangeTStringToTDateString(strSrvDate(2)) & "¤£¾P¨÷,'||tm58)" & _
               " Where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "'" & _
               " and (instr(tm58,'¤£¾P¨÷')=0 or tm58 is null)"
      cnnConnection.Execute strSql
   End If
   '2012/7/19 end
   
   'Add By Sindy 2012/9/26 ÀË¬d¬O§_¬°¤@¥Ó½Ð®Ñ¦h¥ó¨Ã§ó·s¸ê®Æ
   'Modify By Sindy 2013/4/9 ©w½Z»y¤å¬O­^¤å®É¤~°µ¤@¥Ó½Ð®Ñ¦h¥ó
   'Modify By Sindy 2014/6/24 mark : ¤£ºÞ©w½Z»y¤å
   'If GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "2" Then
   '2013/4/9 End
   '2014/6/24 END
      Call PUB_UpdateCP148(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, textCP27)
   'End If
   
   '911107 nick transation
   cnnConnection.CommitTrans
   
   'Add by nickc 2008/02/22 ÀË¬d¥N²z¤HEmail(»Ý¦Ò¼{¥i¯à¬°FF®×¥ó)
   PUB_CheckEMail m_CP44, m_CP116
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

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   'Add By Sindy 2012/4/17
   If m_blnClkChgButton = False Then
      MsgBox "½Ð¿é¤JÅÜ§ó¨Æ¶µ!!!", vbExclamation + vbOKOnly
      Me.cmdMod.SetFocus
      GoTo EXITSUB
   End If
   
   ' µo¤å¤é
   If IsEmptyText(textCP27) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "½Ð¿é¤Jµo¤å¤é"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' ¨üÅý¤H¤£¥iªÅ¥Õ
   'edit by nickc 2007/01/29
   'If IsEmptyText(textCP56) = True Then
   If IsEmptyText(textCP56) = True And IsEmptyText(textCP89) = True And IsEmptyText(textCP90) = True And IsEmptyText(textCP91) = True And IsEmptyText(textCP92) = True Then
      strTit = "ÀË®Ö¸ê®Æ"
      strMsg = "½Ð¿é¤J¨üÅý¤H"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      textCP56.SetFocus
      GoTo EXITSUB
   End If
   
   'Add By Sindy 2011/01/06
   '¥~°Ó(S)¥Ó½Ð¤H1©ÎFC¥N²z¤H¦Ü¤Ö­n¿é¤J¤@­Ó
   '¨ä¥Lªº¤@©w­n¿é¤J¥Ó½Ð¤H1
   If m_TM01 = "S" Then
        If textTM23 = "" And m_TM44 = "" Then
            MsgBox "¥Ó½Ð¤H1©ÎFC¥N²z¤H¦Ü¤Ö­n¿é¤J¤@­Ó!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   Else
        If textTM23 = "" Then
            MsgBox "¥Ó½Ð¤H1¤£¥iªÅ¥Õ!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   End If
   
   'Added by Lydia 2021/09/02 ÀË¬dµe­±ªº TextBox, ComboBox ¬O§_§t¦³Unicode¤å¦r
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
   End If
    
   CheckDataValid = True
EXITSUB:
End Function

' ¤£¤@¨Ö²¾Âàµù¥U¸¹¼Æ
Private Sub textTM15_S_Validate(Cancel As Boolean)
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
   If IsEmptyText(textTM15_S) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textTM15_S)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM15_S, nIndex)
      strSql = "SELECT * FROM TradeMark " & _
               "WHERE TM15 = '" & textTM15_S & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         rsTmp.Close
         Cancel = True
         strTit = "ÀË®Ö¸ê®Æ"
         strMsg = "µù¥U¸¹¼Æ<" & strTemp & ">¤£¦s¦b"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 0
         GoTo EXITSUB
      End If
      rsTmp.Close
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTM15_S, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTM15_S, nCount) Then
               Cancel = True
               strTit = "ÀË®Ö¸ê®Æ"
               strMsg = "µù¥U¸¹¼Æ<" & strTemp & ">¤£¥i­«ÂÐ"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               Me.SSTab1.Tab = 0
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   
EXITSUB:
   Set rsTmp = Nothing
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
         Me.SSTab1.Tab = 0
      End If
   End If
End Sub

Private Sub textCP27_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textCP27
End Sub

Private Sub textUargeDate_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textUargeDate
End Sub

Private Sub textDN_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textDN
End Sub

Private Sub textAdd_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textAdd
End Sub

Private Sub textPrint_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textPrint
End Sub

Private Sub textTM15_S_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textTM15_S
End Sub

Private Sub textCP56_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textCP56
End Sub

Private Sub textCP64_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textCP64
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

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
End If


If Me.textAdd.Enabled = True Then
   Cancel = False
   textAdd_Validate Cancel
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

If Me.textCP56.Enabled = True Then
   Cancel = False
   textCP56_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'add by nickc 2007/01/29
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

'add by nickc 2006/01/27
'edit by nickc 2006/02/07
If m_TM01 = "FCT" Then
    If Me.lstNameAgent.Enabled = True Then
        Cancel = False
        lstNameAgent_Validate Cancel
        If Cancel = True Then
            Me.SSTab1.Tab = 0
            lstNameAgent.SetFocus
            Exit Function
        End If
    End If
End If
TxtValidate = True
End Function

' ¦C¦L©w½Z«e±N¨Ò¥~Äæ¦ì¥[¤J¨ì¦C¦L©w½Z¨Ò¥~Äæ¦ìÀÉ®×¤¤
Private Sub InsExpField()
Dim nCount As Integer
Dim nIndex As Integer
Dim strSql As String
Dim strTemp As String
Dim strTemp1 As String
Dim strDebitNote As String 'Add By Sindy 2017/4/13
    
   'Modify By Sindy 2017/4/13¡iFCT 01 501  00 ­^->¨çª¾¤w´£²¾Âà¡j
   m_MySt(1) = m_TM01: m_MySt(2) = m_TM02: m_MySt(3) = m_TM03: m_MySt(4) = m_TM04: m_Rule = m_CP09
   strDebitNote = ExceptFieldData2("FCT¯S®í½Ð´Ú¤å¦r¹ï·Ó")
   '2017/4/13 END
   
    ' ¬O§_¸É¥ó
    strTemp = Empty
    ' ¨Ì®×¥ó©Ê½è¤£¦P
    Select Case m_CP10
    ' ²¾Âà
    Case "501":
       nCount = GetSubStringCount(textAdd)
       For nIndex = 1 To nCount
          strTemp1 = GetSubString(textAdd, nIndex)
          Select Case strTemp1
             Case "1": '¨üÅý¤H©e¥ôª¬
                If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                strTemp = strTemp & "    * Power of Attorney of the Assignee."
             Case "2": '²¾Âà«´¬ù®Ñ
                If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                strTemp = strTemp & "    * Deed of Assignment respectively signed by the Assignee and Assignor."
             Case "3": '¨üÅý¤Hªk¤HÃÒ©ú
                If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                strTemp = strTemp & "    * A notarized Certificate of Corporation of the Assignee."
             Case "4": 'ÃÒ¥UÃÒ
                If strTemp <> Empty Then: strTemp = strTemp & Chr(13) & Chr(10)
                strTemp = strTemp & "    * Original Certificates of Registration."
          End Select
       Next nIndex
       '92.4.3 modify by sonia
       'If strTemp <> Empty Then: strTemp = vbCrLf & "    The remaining documents for the referenced assignment application follow : " & Chr(13) & Chr(10) & strTemp
       If strTemp <> Empty Then: strTemp = vbCrLf & "    The remaining document(s) for the referenced assignment application follow : " & Chr(13) & Chr(10) & strTemp
       '92.4.3 end
       'add by nick 2004/08/17
       Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
       '­^¤å
       Case "2":
            ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
            EndLetter "01", m_CP09, "00", strUserNum
            ' ¬O§_¸É¥ó
            If IsEmptyText(strTemp) = False Then
               ' ¬O§_¸É¥ó
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('01','" & m_CP09 & "','00','" & strUserNum & _
                        "','¬O§_¸É¥ó','" & strTemp & "')"
               cnnConnection.Execute strSql
            End If
            'Add By Sindy 2012/11/26 eMail Only©w½Z : ¥H¹q¤l¶l¥ó³qª¾¡A¨Ã¥B¤£±H¯È¥»
            If bolEmail = True And bolPlusPaper = False Then
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('01','" & m_CP09 & "','00','" & strUserNum & _
                        "','¨Ò¥~¤º¤å','Enclosed herewith are scanned copies of the assignment application and the filing receipt for your reference. " & IIf(strDebitNote = "", "Our debit note is also enclosed herewith for your kind settlement.", strDebitNote) & "')"
               cnnConnection.Execute strSql
            Else '¶l¥ó
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('01','" & m_CP09 & "','00','" & strUserNum & _
                        "','¨Ò¥~¤º¤å','" & IIf(strDebitNote = "", "Enclosed please find our debit note for services rendered for your kind settlement.", strDebitNote) & " Copies of the assignment application and the filing receipt will be mailed to you with the confirmation copy of this letter for your records.')"
               cnnConnection.Execute strSql
            End If
            '2012/11/26 End
        Case "3":
            ' ²M°£©w½Z¨Ò¥~Äæ¦ìÀÉ­ì¦³¸ê®Æ
            EndLetter "01", m_CP09, "01", strUserNum
            'add by nick 2004/12/16
            'Modify By Sindy 2010/4/15
            'EndLetter "01", m_CP09, "02", strUserNum
            EndLetter "01", m_CP09, "03", strUserNum
            'Add By Sindy 2010/4/15
            If m_TM08 = "2" Or m_TM08 = "3" Or m_TM08 = "5" Then
               strTemp = "¡½°Ó¼Ð¡@¡¼°Ó¼Ð(92¦~§ï¥¿«eÇ±¡ÐÇÏÇµÇÚ¡ÐÇ«)"
            End If
            If m_TM08 = "1" Or m_TM08 = "4" Or m_TM08 = "6" Then
               If InStr(m_TM58, "­ì¬°ªA°È¼Ð³¹,92/11/28­×ªk§ï¬°¥¿°Ó¼Ð") > 0 Or _
                  InStr(m_TM58, "­ì¬°Áp¦XªA°È¼Ð³¹,92/11/28­×ªk§ï¬°¥¿°Ó¼Ð") > 0 Then
                  strTemp = "¡¼°Ó¼Ð¡@¡½°Ó¼Ð(92¦~§ï¥¿«eÇ±¡ÐÇÏÇµÇÚ¡ÐÇ«)"
               Else
                  strTemp = "¡½°Ó¼Ð¡@¡¼°Ó¼Ð(92¦~§ï¥¿«eÇ±¡ÐÇÏÇµÇÚ¡ÐÇ«)"
               End If
            End If
            If m_TM08 = "7" Then
               strTemp = strTemp & "¡@¡½µý©ú¼Ð³¹" & vbCrLf
            Else
               strTemp = strTemp & "¡@¡¼µý©ú¼Ð³¹" & vbCrLf
            End If
            If m_TM08 = "8" Then
               strTemp = strTemp & "¡@¡@¡½’NÊ^¼Ð³¹"
            Else
               strTemp = strTemp & "¡@¡@¡¼’NÊ^¼Ð³¹"
            End If
            If m_TM08 = "9" Then
               strTemp = strTemp & "¡@¡½’NÊ^°Ó¼Ð"
            Else
               strTemp = strTemp & "¡@¡¼’NÊ^°Ó¼Ð"
            End If
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('01','" & m_CP09 & "','03','" & strUserNum & _
                        "','°Ó¼ÐºØÃþ','" & strTemp & "')"
            cnnConnection.Execute strSql
            '2010/4/15 End
        Case Else
        End Select
    End Select
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ¦C¦L©w½Z
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
   'Add by Morgan 2008/6/12
   Dim ET03 As String, ET03_1 As String, stContent As String
   Dim stLang As String, strFilePath As String, strFN01 As String, strFN02 As String 'Added by Lydia 2023/05/03
   
   'Add By Sindy 2012/11/23 ±q¤U­±µ{¦¡©¹¤WMove¦Ü¦¹
   bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper) 'ÀË¬d¬O§_¥HE-Mail³qª¾
   '2012/11/23 End
   stLang = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) 'Added by Lydia 2023/05/03
   
   ' ¥ý©I¥s©w½Zµ{¦¡ªº²M°£­ì©w½Z¸ê®Æªº¨ç¦¡¥h²M°£¤§«e´Ý¯d¦b¨Ò¥~Äæ¦ìÀÉ¤¤ªº¸ê®Æ
   InsExpField
   'add by nick 2004/08/17
   'Modified by Lydia 2023/05/03 §ï¦¨ÅÜ¼Æ
   Select Case stLang
      Case "2":
         ET03 = "00"
      Case "3":
         ET03 = "01"
         'add by nick 2004/12/16   ¥Ó½Ð®Ñ¤À¶}
         'Modify By Sindy 2010/4/15
         'ET03_1 = "02"
         ET03_1 = "03"
      Case Else
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
         Me.SSTab1.Tab = 0
         MsgBox "¹q¤lÀÉ¤w¦s©ó [ " & PUB_GetEFilePath(m_TM01) & " ]¡I"
      'Mark by Lydia 2023/05/03 ³ø§i«È¤á¤§¸ê®Æ²Î¤@¦sÀÉFCT_WORKFLOW
      'Else
      'end 2008/6/12
      '   NowPrint m_CP09, "01", ET03, False, strUserNum
      '   If ET03_1 <> "" Then
      '      NowPrint m_CP09, "01", ET03_1, False, strUserNum
      '   End If
      'End If
      'end 2023/05/03
   End If
End Sub

'Add By Sindy 98/03/11
Private Sub textCP113_GotFocus()
   Me.SSTab1.Tab = 0
   TextInverse textCP113
End Sub
Private Sub textCP113_Validate(Cancel As Boolean)
   If textCP113 <> "" Then
      If Not IsNumeric(textCP113) Then
         MsgBox "½Ð¿é¤J¼Æ¦r¡I", vbExclamation
         Me.SSTab1.Tab = 0
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

'Add By Sindy 2010/3/3
Private Sub Combo2_Click(Index As Integer)
Dim i As Integer, strTmp As String
   
   If (Combo2(Index).Text = "") Then
      If Index <= 1 Then
        For i = 0 To 2
           Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = ""
        Next i
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
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Index <= 1 Then
        For i = 0 To 2
           If Not IsNull(RsTemp.Fields(i)) Then
              Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = RsTemp.Fields(i)
           Else
              Me.Controls("textTM" & Format(47 + i + 3 * Index, "#")).Text = ""
           End If
        Next
      Else
        For i = 0 To 2
           If Not IsNull(RsTemp.Fields(i)) Then
              Me.Controls("textTM" & Format(94 + i + 3 * Index, "#")).Text = RsTemp.Fields(i)
           Else
              Me.Controls("textTM" & Format(94 + i + 3 * Index, "#")).Text = ""
           End If
        Next
      End If
   End If
End Sub

'Add By Sindy 2010/3/3
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
