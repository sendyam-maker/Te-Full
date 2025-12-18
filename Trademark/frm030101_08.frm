VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030101_08 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(移轉)"
   ClientHeight    =   6230
   ClientLeft      =   5330
   ClientTop       =   1610
   ClientWidth     =   9130
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6230
   ScaleWidth      =   9130
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   7050
      Locked          =   -1  'True
      TabIndex        =   231
      TabStop         =   0   'False
      Top             =   780
      Width           =   1905
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   230
      TabStop         =   0   'False
      Top             =   780
      Width           =   1905
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1905
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1905
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   480
      Width           =   1905
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   780
      Width           =   1905
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   7050
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   480
      Width           =   1905
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   480
      Width           =   1905
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   7050
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1905
   End
   Begin VB.TextBox textTM27 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1905
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   400
      Index           =   1
      Left            =   2520
      TabIndex        =   14
      Top             =   15
      Width           =   1092
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   3660
      TabIndex        =   15
      Top             =   15
      Width           =   1092
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H008080FF&
      Caption         =   "變更事項(&R)"
      Height          =   400
      Left            =   4800
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   15
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6990
      TabIndex        =   18
      Top             =   15
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6030
      TabIndex        =   17
      Top             =   15
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8130
      TabIndex        =   19
      Top             =   15
      Width           =   912
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3810
      Left            =   90
      TabIndex        =   42
      Top             =   2370
      Width           =   8895
      _ExtentX        =   15681
      _ExtentY        =   6703
      _Version        =   393216
      Tab             =   2
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm030101_08.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtCP113"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "textCP89"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "textCP90"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "textCP91"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "textCP92"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "textCP27"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "textUargeDate"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "textPrint"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "textCP18"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textCF09"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textCP44"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "textTM15_S"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textCP56"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textCP89_2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textCP90_2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textCP91_2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCP92_2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textCP44_2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textCP56_2"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textCP64"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblCP113(18)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label15"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label16"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label18(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label19"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label28"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label25"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label14(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label4"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label22"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label23"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label1(10)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label1(12)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label11"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label8"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label7"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).ControlCount=   36
      TabCaption(1)   =   "代表人-1"
      TabPicture(1)   =   "frm030101_08.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textTM105"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "textTM104"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "textTM103"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Combo2(5)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "textTM99"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "textTM98"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "textTM97"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Combo2(3)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "textTM52"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "textTM51"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "textTM50"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Combo2(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "textTM102"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "textTM101"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "textTM100"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Combo2(4)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "textTM96"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "textTM95"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "textTM94"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Combo2(2)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "textTM49"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "textTM48"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "textTM47"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Combo2(0)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label18(3)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label14(3)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label5(18)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label5(17)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label5(16)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label5(15)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Label5(14)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Label5(13)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Label18(1)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label14(2)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Label5(12)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Label5(11)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Label5(10)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Label5(9)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Label5(2)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Label5(1)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Label5(8)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Label5(7)"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Label5(6)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Label5(5)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "Label5(4)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "Label5(3)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Label14(1)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "Label18(2)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).ControlCount=   48
      TabCaption(2)   =   "代表人-2"
      TabPicture(2)   =   "frm030101_08.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label5(43)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label5(42)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label5(41)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label5(40)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label5(39)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label5(38)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label14(8)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Label18(7)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label5(37)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Label5(36)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label5(35)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Label5(34)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label5(33)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label5(32)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Label14(7)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Label18(6)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "TextTM106"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "TextTM107"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "TextTM109"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "TextTM110"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Combo2(7)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Combo2(6)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "TextTM108"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "TextTM111"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "TextTM112"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "TextTM113"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "TextTM115"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "TextTM116"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Combo2(9)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Combo2(8)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "TextTM114"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "TextTM117"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).ControlCount=   32
      Begin VB.TextBox txtCP113 
         Height          =   270
         Left            =   -71040
         MaxLength       =   4
         TabIndex        =   5
         Top             =   897
         Width           =   540
      End
      Begin VB.TextBox textCP89 
         Height          =   285
         Left            =   -73860
         MaxLength       =   9
         TabIndex        =   9
         Top             =   1995
         Width           =   1092
      End
      Begin VB.TextBox textCP90 
         Height          =   285
         Left            =   -73860
         MaxLength       =   9
         TabIndex        =   10
         Top             =   2280
         Width           =   1092
      End
      Begin VB.TextBox textCP91 
         Height          =   285
         Left            =   -73860
         MaxLength       =   9
         TabIndex        =   11
         Top             =   2565
         Width           =   1092
      End
      Begin VB.TextBox textCP92 
         Height          =   285
         Left            =   -73860
         MaxLength       =   9
         TabIndex        =   12
         Top             =   2850
         Width           =   1092
      End
      Begin VB.TextBox textCP27 
         Height          =   264
         Left            =   -73860
         MaxLength       =   8
         TabIndex        =   0
         Top             =   300
         Width           =   1092
      End
      Begin VB.TextBox textUargeDate 
         Height          =   264
         Left            =   -71490
         MaxLength       =   8
         TabIndex        =   1
         Top             =   300
         Width           =   1092
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   -73860
         MaxLength       =   1
         TabIndex        =   4
         Top             =   900
         Width           =   372
      End
      Begin VB.TextBox textCP18 
         Height          =   264
         Left            =   -69240
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   2
         Top             =   300
         Width           =   825
      End
      Begin VB.TextBox textCF09 
         Height          =   264
         Left            =   -69150
         MaxLength       =   3
         TabIndex        =   6
         Top             =   900
         Width           =   612
      End
      Begin VB.ComboBox textCP44 
         Height          =   300
         Left            =   -73860
         TabIndex        =   3
         Top             =   570
         Width           =   1452
      End
      Begin VB.TextBox textTM15_S 
         Height          =   510
         Left            =   -73140
         MaxLength       =   200
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   7
         Top             =   1170
         Width           =   6915
      End
      Begin VB.TextBox textCP56 
         Height          =   285
         Left            =   -73860
         MaxLength       =   9
         TabIndex        =   8
         Top             =   1710
         Width           =   1092
      End
      Begin VB.TextBox textTM08_2 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -71160
         Locked          =   -1  'True
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   600
         Width           =   1755
      End
      Begin VB.TextBox Text17 
         Height          =   264
         Left            =   -74160
         MaxLength       =   9
         TabIndex        =   123
         Top             =   1710
         Width           =   1092
      End
      Begin VB.TextBox textTM58 
         Height          =   276
         Left            =   -73545
         MaxLength       =   2000
         TabIndex        =   122
         Top             =   2325
         Width           =   7272
      End
      Begin VB.TextBox Text16 
         Height          =   396
         Left            =   -73545
         MaxLength       =   2000
         TabIndex        =   121
         Top             =   1890
         Width           =   7272
      End
      Begin VB.TextBox Text15 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -67830
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   300
         Width           =   1635
      End
      Begin VB.TextBox Text14 
         Height          =   264
         Left            =   -71550
         MaxLength       =   1
         TabIndex        =   119
         Top             =   600
         Width           =   372
      End
      Begin VB.TextBox textTM05 
         Height          =   264
         Left            =   -73560
         MaxLength       =   40
         TabIndex        =   118
         Top             =   2610
         Width           =   7272
      End
      Begin VB.TextBox textTM06 
         Height          =   264
         Left            =   -73560
         MaxLength       =   60
         TabIndex        =   117
         Top             =   2835
         Width           =   7272
      End
      Begin VB.TextBox textTM07 
         Height          =   264
         Left            =   -73560
         MaxLength       =   40
         TabIndex        =   116
         Top             =   3105
         Width           =   7272
      End
      Begin VB.TextBox textTM23_2 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   1710
         Width           =   2505
      End
      Begin VB.TextBox Text13 
         Height          =   264
         Left            =   -73296
         MaxLength       =   1
         TabIndex        =   114
         Top             =   1182
         Width           =   372
      End
      Begin VB.TextBox textTM21 
         Height          =   264
         Left            =   -73296
         MaxLength       =   7
         TabIndex        =   113
         Top             =   912
         Width           =   852
      End
      Begin VB.TextBox textTM22 
         Height          =   264
         Left            =   -72120
         MaxLength       =   7
         TabIndex        =   112
         Top             =   912
         Width           =   852
      End
      Begin VB.TextBox Text12 
         Height          =   264
         Left            =   -73560
         MaxLength       =   7
         TabIndex        =   111
         Top             =   600
         Width           =   1092
      End
      Begin VB.TextBox Text11 
         Height          =   264
         Left            =   -73560
         MaxLength       =   7
         TabIndex        =   110
         Top             =   300
         Width           =   1092
      End
      Begin VB.TextBox textMail 
         Height          =   264
         Left            =   -73560
         MaxLength       =   1
         TabIndex        =   109
         Top             =   960
         Width           =   492
      End
      Begin VB.TextBox textTM32 
         Height          =   264
         Left            =   -73560
         MaxLength       =   300
         TabIndex        =   108
         Top             =   660
         Width           =   7272
      End
      Begin VB.TextBox textTM09 
         Height          =   264
         Left            =   -73560
         MaxLength       =   395
         TabIndex        =   107
         Top             =   360
         Width           =   7272
      End
      Begin VB.TextBox Text10 
         Height          =   264
         Left            =   -73320
         MaxLength       =   10
         TabIndex        =   106
         Top             =   1455
         Width           =   852
      End
      Begin VB.TextBox textAdd_2 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   1425
         Width           =   6072
      End
      Begin VB.TextBox textCP09_S3 
         Height          =   264
         Left            =   -68520
         MaxLength       =   2
         TabIndex        =   104
         Top             =   990
         Width           =   465
      End
      Begin VB.TextBox textCP09_S2 
         Height          =   264
         Left            =   -68970
         MaxLength       =   1
         TabIndex        =   103
         Top             =   990
         Width           =   345
      End
      Begin VB.TextBox textCP09_S1 
         Height          =   264
         Left            =   -70050
         MaxLength       =   6
         TabIndex        =   102
         Top             =   990
         Width           =   975
      End
      Begin VB.TextBox textCP09_S 
         Height          =   264
         Left            =   -70590
         MaxLength       =   1
         TabIndex        =   101
         Top             =   990
         Width           =   465
      End
      Begin VB.TextBox textPrtTrans 
         Height          =   264
         Left            =   -68370
         MaxLength       =   10
         TabIndex        =   100
         Top             =   1182
         Width           =   372
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   -66720
         MaxLength       =   1
         TabIndex        =   99
         Top             =   1290
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textTM05_1 
         Height          =   792
         Left            =   -73560
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   98
         Top             =   2610
         Width           =   7272
      End
      Begin VB.TextBox Text9 
         Height          =   264
         Left            =   -68370
         MaxLength       =   1
         TabIndex        =   97
         Top             =   912
         Width           =   492
      End
      Begin VB.TextBox textTM72_2 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -67980
         Locked          =   -1  'True
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   600
         Width           =   1785
      End
      Begin VB.TextBox textTM72 
         Height          =   264
         Left            =   -68370
         MaxLength       =   1
         TabIndex        =   95
         Top             =   600
         Width           =   372
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Left            =   -71430
         TabIndex        =   94
         Top             =   285
         Width           =   1425
      End
      Begin VB.TextBox Text6 
         Height          =   288
         Left            =   -74790
         MaxLength       =   1
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1530
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         ItemData        =   "frm030101_08.frx":0054
         Left            =   -73515
         List            =   "frm030101_08.frx":005E
         Sorted          =   -1  'True
         Style           =   1  '項目包含核取方塊
         TabIndex        =   92
         Top             =   1320
         Width           =   1260
      End
      Begin VB.TextBox textTM78_2 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -68670
         Locked          =   -1  'True
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   1710
         Width           =   2505
      End
      Begin VB.TextBox Text5 
         Height          =   264
         Left            =   -69750
         MaxLength       =   9
         TabIndex        =   90
         Top             =   1710
         Width           =   1092
      End
      Begin VB.TextBox textTM79_2 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1980
         Width           =   2505
      End
      Begin VB.TextBox Text4 
         Height          =   264
         Left            =   -74160
         MaxLength       =   9
         TabIndex        =   88
         Top             =   1980
         Width           =   1092
      End
      Begin VB.TextBox textTM80_2 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -68670
         Locked          =   -1  'True
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   1980
         Width           =   2505
      End
      Begin VB.TextBox Text3 
         Height          =   264
         Left            =   -69750
         MaxLength       =   9
         TabIndex        =   86
         Top             =   1980
         Width           =   1092
      End
      Begin VB.TextBox textTM81_2 
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -73080
         Locked          =   -1  'True
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   2250
         Width           =   2505
      End
      Begin VB.TextBox Text2 
         Height          =   264
         Left            =   -74160
         MaxLength       =   9
         TabIndex        =   84
         Top             =   2250
         Width           =   1092
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   -69090
         MaxLength       =   4
         TabIndex        =   83
         Top             =   285
         Width           =   600
      End
      Begin MSForms.TextBox textCP89_2 
         Height          =   285
         Left            =   -72750
         TabIndex        =   236
         TabStop         =   0   'False
         Top             =   1995
         Width           =   6525
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "11509;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP90_2 
         Height          =   285
         Left            =   -72750
         TabIndex        =   235
         TabStop         =   0   'False
         Top             =   2280
         Width           =   6525
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "11509;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP91_2 
         Height          =   285
         Left            =   -72750
         TabIndex        =   234
         TabStop         =   0   'False
         Top             =   2565
         Width           =   6525
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "11509;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP92_2 
         Height          =   285
         Left            =   -72750
         TabIndex        =   233
         TabStop         =   0   'False
         Top             =   2850
         Width           =   6525
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "11509;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   264
         Left            =   -72390
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   600
         Width           =   6150
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
      Begin MSForms.TextBox textCP56_2 
         Height          =   285
         Left            =   -72750
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   1710
         Width           =   6525
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "11509;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM105 
         Height          =   285
         Left            =   -69705
         TabIndex        =   82
         Top             =   3450
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM104 
         Height          =   285
         Left            =   -69705
         TabIndex        =   81
         Top             =   3165
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM103 
         Height          =   285
         Left            =   -69705
         TabIndex        =   80
         Top             =   2880
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   5
         Left            =   -69705
         TabIndex        =   79
         Top             =   2580
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM99 
         Height          =   285
         Left            =   -69705
         TabIndex        =   78
         Top             =   2265
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM98 
         Height          =   285
         Left            =   -69705
         TabIndex        =   77
         Top             =   1987
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM97 
         Height          =   285
         Left            =   -69705
         TabIndex        =   76
         Top             =   1710
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   3
         Left            =   -69705
         TabIndex        =   75
         Top             =   1410
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM52 
         Height          =   285
         Left            =   -69705
         TabIndex        =   74
         Top             =   1140
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM51 
         Height          =   285
         Left            =   -69705
         TabIndex        =   73
         Top             =   855
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM50 
         Height          =   285
         Left            =   -69705
         TabIndex        =   72
         Top             =   570
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   1
         Left            =   -69705
         TabIndex        =   71
         Top             =   270
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM102 
         Height          =   285
         Left            =   -74115
         TabIndex        =   70
         Top             =   3435
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM101 
         Height          =   285
         Left            =   -74115
         TabIndex        =   69
         Top             =   3158
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM100 
         Height          =   285
         Left            =   -74115
         TabIndex        =   68
         Top             =   2880
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   4
         Left            =   -74115
         TabIndex        =   67
         Top             =   2565
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM96 
         Height          =   285
         Left            =   -74115
         TabIndex        =   66
         Top             =   2265
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM95 
         Height          =   285
         Left            =   -74115
         TabIndex        =   65
         Top             =   1987
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM94 
         Height          =   285
         Left            =   -74115
         TabIndex        =   64
         Top             =   1710
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   2
         Left            =   -74115
         TabIndex        =   63
         Top             =   1395
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM49 
         Height          =   285
         Left            =   -74115
         TabIndex        =   62
         Top             =   1140
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM48 
         Height          =   285
         Left            =   -74115
         TabIndex        =   61
         Top             =   855
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM47 
         Height          =   285
         Left            =   -74115
         TabIndex        =   60
         Top             =   570
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   0
         Left            =   -74115
         TabIndex        =   59
         Top             =   270
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM117 
         Height          =   285
         Left            =   5325
         TabIndex        =   58
         Top             =   2385
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM114 
         Height          =   285
         Left            =   915
         TabIndex        =   57
         Top             =   2385
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   8
         Left            =   915
         TabIndex        =   56
         Top             =   1485
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   9
         Left            =   5325
         TabIndex        =   55
         Top             =   1500
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM116 
         Height          =   285
         Left            =   5325
         TabIndex        =   54
         Top             =   2092
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM115 
         Height          =   285
         Left            =   5325
         TabIndex        =   53
         Top             =   1800
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM113 
         Height          =   285
         Left            =   915
         TabIndex        =   52
         Top             =   2092
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM112 
         Height          =   285
         Left            =   915
         TabIndex        =   51
         Top             =   1800
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM111 
         Height          =   285
         Left            =   5325
         TabIndex        =   50
         Top             =   1200
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM108 
         Height          =   285
         Left            =   915
         TabIndex        =   49
         Top             =   1200
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   6
         Left            =   915
         TabIndex        =   48
         Top             =   300
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   7
         Left            =   5325
         TabIndex        =   47
         Top             =   300
         Width           =   3525
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6218;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM110 
         Height          =   285
         Left            =   5325
         TabIndex        =   46
         Top             =   900
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM109 
         Height          =   285
         Left            =   5325
         TabIndex        =   45
         Top             =   600
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM107 
         Height          =   285
         Left            =   915
         TabIndex        =   44
         Top             =   900
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   60
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox TextTM106 
         Height          =   285
         Left            =   915
         TabIndex        =   43
         Top             =   600
         Width           =   3525
         VariousPropertyBits=   671105051
         BackColor       =   -2147483644
         MaxLength       =   40
         Size            =   "6218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   555
         Left            =   -73860
         TabIndex        =   13
         Top             =   3150
         Width           =   7545
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13309;979"
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
         Left            =   -72000
         TabIndex        =   241
         Top             =   942
         Width           =   765
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "受讓人2 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   240
         Top             =   2051
         Width           =   720
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "受讓人3 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   239
         Top             =   2317
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "受讓人4 :"
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   238
         Top             =   2583
         Width           =   720
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "受讓人5 :"
         Height          =   180
         Left            =   -74880
         TabIndex        =   237
         Top             =   2850
         Width           =   720
      End
      Begin VB.Label Label28 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   229
         Top             =   3180
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "發文日 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   228
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "催審期限 :"
         Height          =   255
         Index           =   0
         Left            =   -72480
         TabIndex        =   227
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "代理人 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   226
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   225
         Top             =   930
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不印)"
         Height          =   255
         Left            =   -73440
         TabIndex        =   224
         Top             =   930
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "點數 :"
         Height          =   255
         Index           =   10
         Left            =   -69750
         TabIndex        =   223
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "大約"
         Height          =   255
         Index           =   12
         Left            =   -69660
         TabIndex        =   222
         Top             =   930
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "可接獲回音"
         Height          =   255
         Left            =   -68490
         TabIndex        =   221
         Top             =   930
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "不一併移轉註冊號數 :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   220
         Top             =   1230
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "受讓人1 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   219
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label54 
         Caption         =   "案件備註 :"
         Height          =   255
         Left            =   -74835
         TabIndex        =   216
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label53 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   -74865
         TabIndex        =   215
         Top             =   1890
         Width           =   975
      End
      Begin VB.Label Label52 
         Caption         =   "查名本所案號 :"
         Height          =   255
         Left            =   -71910
         TabIndex        =   214
         Top             =   990
         Width           =   1275
      End
      Begin VB.Label Label51 
         Caption         =   "是否算案件數 :"
         Height          =   255
         Left            =   -67410
         TabIndex        =   213
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "點數 :"
         Height          =   180
         Index           =   17
         Left            =   -68340
         TabIndex        =   212
         Top             =   345
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "正商標號數:"
         Height          =   255
         Index           =   16
         Left            =   -66480
         TabIndex        =   211
         Top             =   3120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "商標種類 :"
         Height          =   180
         Index           =   15
         Left            =   -72420
         TabIndex        =   210
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label50 
         Caption         =   "案件中文名稱 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   209
         Top             =   2655
         Width           =   1335
      End
      Begin VB.Label Label49 
         Caption         =   "案件英文名稱 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   208
         Top             =   2925
         Width           =   1215
      End
      Begin VB.Label Label48 
         Caption         =   "案件日文名稱 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   207
         Top             =   3210
         Width           =   1455
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "申請人1 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   206
         Top             =   1752
         Width           =   720
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印)"
         Height          =   180
         Left            =   -72840
         TabIndex        =   205
         Top             =   1224
         Width           =   645
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "列印定稿 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   204
         Top             =   1224
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "延展後專用期限 :"
         Height          =   180
         Index           =   31
         Left            =   -74910
         TabIndex        =   203
         Top             =   954
         Width           =   1350
      End
      Begin VB.Line Line1 
         X1              =   -72360
         X2              =   -72240
         Y1              =   1008
         Y2              =   1008
      End
      Begin VB.Label Label44 
         Caption         =   "代理人 :"
         Height          =   252
         Left            =   120
         TabIndex        =   202
         Top             =   -360
         Width           =   972
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "催審期限 :"
         Height          =   180
         Index           =   6
         Left            =   -74910
         TabIndex        =   201
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "發文日 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   200
         Top             =   345
         Width           =   630
      End
      Begin VB.Label Label40 
         Caption         =   "是否郵寄申請 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   199
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label41 
         Caption         =   "(Y:郵寄)"
         Height          =   252
         Left            =   -72960
         TabIndex        =   198
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "商品組群 :"
         Height          =   252
         Index           =   13
         Left            =   -74880
         TabIndex        =   197
         Top             =   660
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "商品類別 :"
         Height          =   252
         Index           =   14
         Left            =   -74880
         TabIndex        =   196
         Top             =   360
         Width           =   852
      End
      Begin VB.Label Label38 
         Caption         =   "是否補件(可複選) :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   195
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Line Line2 
         X1              =   -70320
         X2              =   -68250
         Y1              =   1110
         Y2              =   1110
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印)"
         Height          =   180
         Left            =   -67950
         TabIndex        =   194
         Top             =   1224
         Width           =   645
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "是否列印翻譯函 :"
         Height          =   180
         Left            =   -69750
         TabIndex        =   193
         Top             =   1224
         Width           =   1350
      End
      Begin VB.Label Label33 
         Caption         =   "(N:不算)"
         Height          =   255
         Left            =   -67200
         TabIndex        =   192
         Top             =   1380
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label42 
         Caption         =   "案件名稱 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   191
         Top             =   2655
         Width           =   1335
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "(Y:輸入)"
         Height          =   180
         Left            =   -67770
         TabIndex        =   190
         Top             =   954
         Width           =   645
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "是否輸入D/N :"
         Height          =   180
         Left            =   -69720
         TabIndex        =   189
         Top             =   954
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊商標 :"
         Height          =   180
         Index           =   7
         Left            =   -69270
         TabIndex        =   188
         Top             =   645
         Width           =   810
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Left            =   -72360
         TabIndex        =   187
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   -74475
         TabIndex        =   186
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "申請人2 :"
         Height          =   180
         Left            =   -70470
         TabIndex        =   185
         Top             =   1755
         Width           =   720
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "申請人3 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   184
         Top             =   2022
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "申請人4 :"
         Height          =   180
         Left            =   -70470
         TabIndex        =   183
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "申請人5 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   182
         Top             =   2292
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   19
         Left            =   -70110
         TabIndex        =   181
         Top             =   1125
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   20
         Left            =   -70110
         TabIndex        =   180
         Top             =   870
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   21
         Left            =   -70110
         TabIndex        =   179
         Top             =   615
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   22
         Left            =   -74490
         TabIndex        =   178
         Top             =   1125
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   23
         Left            =   -74490
         TabIndex        =   177
         Top             =   870
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   24
         Left            =   -74490
         TabIndex        =   176
         Top             =   615
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人7"
         Height          =   180
         Index           =   4
         Left            =   -74775
         TabIndex        =   175
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人8"
         Height          =   180
         Index           =   4
         Left            =   -70395
         TabIndex        =   174
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   25
         Left            =   -70110
         TabIndex        =   173
         Top             =   2220
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   26
         Left            =   -70110
         TabIndex        =   172
         Top             =   1965
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   27
         Left            =   -70110
         TabIndex        =   171
         Top             =   1710
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   28
         Left            =   -74490
         TabIndex        =   170
         Top             =   2220
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   29
         Left            =   -74490
         TabIndex        =   169
         Top             =   1965
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   30
         Left            =   -74490
         TabIndex        =   168
         Top             =   1710
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人9"
         Height          =   180
         Index           =   5
         Left            =   -74775
         TabIndex        =   167
         Top             =   1455
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人10"
         Height          =   180
         Index           =   5
         Left            =   -70395
         TabIndex        =   166
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作時數:"
         Height          =   180
         Index           =   5
         Left            =   -69885
         TabIndex        =   165
         Top             =   345
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人6"
         Height          =   180
         Index           =   3
         Left            =   -70410
         TabIndex        =   164
         Top             =   2640
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人5"
         Height          =   180
         Index           =   3
         Left            =   -74790
         TabIndex        =   163
         Top             =   2640
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   18
         Left            =   -74505
         TabIndex        =   162
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   17
         Left            =   -74505
         TabIndex        =   161
         Top             =   3165
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   16
         Left            =   -74505
         TabIndex        =   160
         Top             =   3450
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   15
         Left            =   -70125
         TabIndex        =   159
         Top             =   2880
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   14
         Left            =   -70125
         TabIndex        =   158
         Top             =   3165
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   13
         Left            =   -70125
         TabIndex        =   157
         Top             =   3450
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人4"
         Height          =   180
         Index           =   1
         Left            =   -70410
         TabIndex        =   156
         Top             =   1470
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人3"
         Height          =   180
         Index           =   2
         Left            =   -74790
         TabIndex        =   155
         Top             =   1470
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   12
         Left            =   -74505
         TabIndex        =   154
         Top             =   1710
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   11
         Left            =   -74505
         TabIndex        =   153
         Top             =   1987
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   10
         Left            =   -74505
         TabIndex        =   152
         Top             =   2265
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   9
         Left            =   -70125
         TabIndex        =   151
         Top             =   1710
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   2
         Left            =   -70125
         TabIndex        =   150
         Top             =   1987
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   1
         Left            =   -70125
         TabIndex        =   149
         Top             =   2265
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   8
         Left            =   -70125
         TabIndex        =   148
         Top             =   1140
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   7
         Left            =   -70125
         TabIndex        =   147
         Top             =   855
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   6
         Left            =   -70125
         TabIndex        =   146
         Top             =   570
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   5
         Left            =   -74505
         TabIndex        =   145
         Top             =   1140
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   4
         Left            =   -74505
         TabIndex        =   144
         Top             =   855
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   3
         Left            =   -74505
         TabIndex        =   143
         Top             =   570
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人1"
         Height          =   180
         Index           =   1
         Left            =   -74790
         TabIndex        =   142
         Top             =   345
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人2"
         Height          =   180
         Index           =   2
         Left            =   -70410
         TabIndex        =   141
         Top             =   345
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人10"
         Height          =   180
         Index           =   6
         Left            =   4620
         TabIndex        =   140
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人9"
         Height          =   180
         Index           =   7
         Left            =   240
         TabIndex        =   139
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   32
         Left            =   525
         TabIndex        =   138
         Top             =   1800
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   33
         Left            =   525
         TabIndex        =   137
         Top             =   2092
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   34
         Left            =   525
         TabIndex        =   136
         Top             =   2385
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   35
         Left            =   4905
         TabIndex        =   135
         Top             =   1800
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   36
         Left            =   4905
         TabIndex        =   134
         Top             =   2092
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   37
         Left            =   4905
         TabIndex        =   133
         Top             =   2385
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人8"
         Height          =   180
         Index           =   7
         Left            =   4620
         TabIndex        =   132
         Top             =   375
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人7"
         Height          =   180
         Index           =   8
         Left            =   240
         TabIndex        =   131
         Top             =   375
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   38
         Left            =   525
         TabIndex        =   130
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   39
         Left            =   525
         TabIndex        =   129
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   40
         Left            =   525
         TabIndex        =   128
         Top             =   1200
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   41
         Left            =   4905
         TabIndex        =   127
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   42
         Left            =   4905
         TabIndex        =   126
         Top             =   900
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   43
         Left            =   4905
         TabIndex        =   125
         Top             =   1200
         Width           =   345
      End
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1140
      TabIndex        =   40
      Top             =   2010
      Width           =   7875
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "13891;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1170
      TabIndex        =   244
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1905
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3360;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   4200
      TabIndex        =   243
      Top             =   1380
      Width           =   1905
      VariousPropertyBits=   671105055
      Size            =   "3360;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1170
      TabIndex        =   242
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1905
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "3360;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   255
      Left            =   6150
      TabIndex        =   232
      Top             =   780
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "案件名稱 :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   41
      Top             =   2010
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   255
      Left            =   3150
      TabIndex        =   39
      Top             =   780
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   3150
      TabIndex        =   38
      Top             =   1380
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   255
      Index           =   9
      Left            =   3150
      TabIndex        =   37
      Top             =   1080
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   36
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   35
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   780
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "發證日 :"
      Height          =   255
      Index           =   3
      Left            =   6150
      TabIndex        =   33
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "業務區別 :"
      Height          =   255
      Index           =   2
      Left            =   3150
      TabIndex        =   32
      Top             =   480
      Width           =   885
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "正商標號數:"
      Height          =   255
      Index           =   8
      Left            =   3150
      TabIndex        =   30
      Top             =   1680
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   255
      Index           =   4
      Left            =   6150
      TabIndex        =   29
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "frm030101_08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/08/11 改成Form2.0 ; textCP13、textCP14、textTM23、cmbTM05、Combo2(index)、textCP56_2、textCP89_2、textCP90_2、textCP91_2、textCP92_2、textCP64、textTM47~52、textTM94~117
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
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
'Add By Cheng 2003/03/07
Dim m_CP55 As String '原移轉人
Dim m_TM23 As String '原申請人
Dim m_TM78 As String, m_TM79 As String, m_TM80 As String, m_TM81 As String
'Add By Cheng 2003/10/06
Public m_blnClkChgButton As Boolean '是否按下變更事項按鈕
'add by nickc 2008/02/22
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP13 As String 'Add By Sindy 2014/9/11
Dim m_990CP09 As String 'Add By Sindy 2016/12/20
Dim m_CP07 As String 'Add By Sindy 2019/6/11


Private Sub cmdCancel_Click()
   frm030101_01.Show
   Unload Me
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
   frm030101_05.SetParent "frm030101_08"
   'Me.Hide
   frm030101_05.Show
   frm030101_05.QueryData
'    m_blnClkChgButton = True
End Sub

Private Sub cmdOK_Click(Index As Integer)
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
               If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = False Then
                  'Add By Sindy 2024/8/19
                  If frm030101_01.bolIsEMPFlow = True Then
                     Unload frm030101_01
                     frm090202_4.Show
                     Unload Me
                     Exit Sub
                  End If
                  '2024/8/19 End
               End If
               frm030101_01.Show
               ' 90.12.07 modify by louis
         '      frm030101_01.Clear
               'Add By Cheng 2002/01/10
               frm030101_01.Clear1
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
   'frm880002.strPriority1 = strPriority1
   'frm880002.strPriority2 = strPriority2
   'frm880002.strPriority3 = strPriority3
   'frm880002.Show vbModal
   'strPriority1 = frm880002.strPriority1
   'strPriority2 = frm880002.strPriority2
   'strPriority3 = frm880002.strPriority3
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
'
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
   textTM08.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM27.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   textCP56_2.BackColor = &H8000000F
   'Add By Sindy 2016/6/17
   textCP89_2.BackColor = &H8000000F
   textCP90_2.BackColor = &H8000000F
   textCP91_2.BackColor = &H8000000F
   textCP92_2.BackColor = &H8000000F
   '2016/6/17 END
   
   'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
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
   
   SSTab1.Tab = 0 'Added by Lydia 2021/06/04
   
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
         textTM20 = DBDATE(rsTmp.Fields("TM20"))
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
         textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
        'Add By Cheng 2003/03/07
        '記錄申請人
        m_TM23 = "" & rsTmp("TM23").Value
      ' 正商標號數
      If IsNull(rsTmp.Fields("TM27")) = False Then
         textTM27 = rsTmp.Fields("TM27")
      End If
      
      'Add By Sindy 2010/3/3
      ' 代表人1(中)
      textTM47 = Empty
      If IsNull(rsTmp.Fields("TM47")) = False Then: textTM47 = rsTmp.Fields("TM47")
      SetTMSPFieldOldData "TM47", textTM47, 0
      ' 代表人1(英)
      textTM48 = Empty
      If IsNull(rsTmp.Fields("TM48")) = False Then: textTM48 = rsTmp.Fields("TM48")
      SetTMSPFieldOldData "TM48", textTM48, 0
      ' 代表人1(日)
      textTM49 = Empty
      If IsNull(rsTmp.Fields("TM49")) = False Then: textTM49 = rsTmp.Fields("TM49")
      SetTMSPFieldOldData "TM49", textTM49, 0
      ' 代表人2(中)
      textTM50 = Empty
      If IsNull(rsTmp.Fields("TM50")) = False Then: textTM50 = rsTmp.Fields("TM50")
      SetTMSPFieldOldData "TM50", textTM50, 0
      ' 代表人2(英)
      textTM51 = Empty
      If IsNull(rsTmp.Fields("TM51")) = False Then: textTM51 = rsTmp.Fields("TM51")
      SetTMSPFieldOldData "TM51", textTM51, 0
      ' 代表人2(日)
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
      '2010/3/3 End
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
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
        'Add By Cheng 2003/03/07
        '記錄申請人
        m_TM23 = "" & rsTmp("SP08").Value
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("SP12")) = False Then
         textTM20 = DBDATE(rsTmp.Fields("SP12"))
      End If
      
      'Add By Sindy 2010/3/3
      SSTab1.TabVisible(1) = False
      SSTab1.TabVisible(2) = False
      '2010/3/3 End
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
      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      
      'Add By Sindy 2019/6/11
      '法定期限
      m_CP07 = Empty
      If IsNull(rsTmp.Fields("CP07")) = False Then: m_CP07 = rsTmp.Fields("CP07")
      '2019/6/11 End
      
      ' 案件性質
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
      m_CP13 = "" 'Add By Sindy 2014/9/11
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
         m_CP13 = rsTmp.Fields("CP13") 'Add By Sindy 2014/9/11
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         '92.10.6 ADD BY SONIA
         m_CP14 = rsTmp.Fields("CP14")
         '92'10'6 END
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 發文日(預設為系統日)
      'edit by nickc 2006/03/17
      'textCP27 = DBDATE(Date)
      textCP27 = strSrvDate(1)
      strCP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         strCP27 = rsTmp.Fields("CP27")
      End If
      SetCPFieldOldData "CP27", strCP27, 1
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
      SetCPFieldOldData "CP18", textCP18, 0
      ' 彼所案號
      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         strCP45 = rsTmp.Fields("CP45")
      End If
      SetCPFieldOldData "CP45", strCP45, 0
      'Add By Cheng 2003/03/07
      '記錄移轉人
      m_CP55 = "" & rsTmp("CP55").Value
      SetCPFieldOldData "CP55", m_CP55, 0 'Add By Sindy 2012/12/27
      
      ' 移轉申請人1
      textCP56 = Empty
      If IsNull(rsTmp.Fields("CP56")) = False Then
         textCP56 = rsTmp.Fields("CP56")
         textCP56_Validate False
      End If
      SetCPFieldOldData "CP56", textCP56, 0
      
      'Add By Sindy 2016/6/17
      ' 移轉申請人2
      textCP89 = Empty
      If IsNull(rsTmp.Fields("CP89")) = False Then
         textCP89 = rsTmp.Fields("CP89")
         textCP89_Validate False
      End If
      SetCPFieldOldData "CP89", textCP89, 0
      ' 移轉申請人3
      textCP90 = Empty
      If IsNull(rsTmp.Fields("CP90")) = False Then
         textCP90 = rsTmp.Fields("CP90")
         textCP90_Validate False
      End If
      SetCPFieldOldData "CP90", textCP90, 0
      ' 移轉申請人4
      textCP91 = Empty
      If IsNull(rsTmp.Fields("CP91")) = False Then
         textCP91 = rsTmp.Fields("CP91")
         textCP91_Validate False
      End If
      SetCPFieldOldData "CP91", textCP91, 0
      ' 移轉申請人5
      textCP92 = Empty
      If IsNull(rsTmp.Fields("CP92")) = False Then
         textCP92 = rsTmp.Fields("CP92")
         textCP92_Validate False
      End If
      SetCPFieldOldData "CP92", textCP92, 0
      '2016/6/17 END
      
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
      
      'Add By Sindy 2010/3/3
      SSTab1.TabVisible(1) = True
      SSTab1.TabVisible(2) = True
      '代表人
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
         '英->中->日
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
         '英->中->日
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
         '英->中->日
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
         '英->中->日
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
      
      ' 代理人
      ClearAgentList
      'Add By Sindy 2013/5/23 若是原先有，也要加入
      If textCP44.Text <> "" Then
'         If InStr(textCP44, "-") > 0 Then
'            If ClsPDGetContact(textCP44, strCP44) Then
'               AddAgent textCP44, strCP44
'            End If
'         Else
            strCP44 = GetFAgentName(textCP44)
            AddAgent textCP44, strCP44
'         End If
      End If
      '2013/5/23 End
      '2009/2/3 modify by sonia B類收文之文件簽證711及申請英文證明304不要列入
      '2010/9/7 Modify by Sindy 文件簽證711及申請英文證明304不要列入
      strSubSQL = "SELECT CP44, MAX(CP27) AS CP27 FROM CASEPROGRESS " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null " & _
                        "AND CP10 NOT IN ('711','304') " & _
                  "GROUP BY CP44 " & _
                  "ORDER BY CP27 DESC "
      rsSubTmp.CursorLocation = adUseClient
      rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSubTmp.RecordCount > 0 Then
         rsSubTmp.MoveFirst
         ' 依序將代理人加入到系統串列中
         Do While rsSubTmp.EOF = False
            If IsNull(rsSubTmp.Fields("CP44")) = False Then
               strCP44 = GetFAgentName(rsSubTmp.Fields("CP44"))
               AddAgent rsSubTmp.Fields("CP44"), GetFAgentName(rsSubTmp.Fields("CP44"))
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
   Set rsSubTmp = Nothing
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
   
   ' 本所案號
'   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)

   ' 收文號
   textCP09 = m_CP09
   
   ' 取得國家代碼
   m_TM10 = GetNationNo(m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 取得案件進度檔的欄位
   QueryCaseProgress
   
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
      If IsNull(rsTmp.Fields("CF09")) = False Then
         textCF09 = rsTmp.Fields("CF09")
      End If
   End If
   rsTmp.Close
   
   ' 計算催審期限
   strDay = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
   If IsEmptyText(strDay) = False Then
      textUargeDate = strDay
   End If
   Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Add By Cheng 2002/07/19
   Set frm030101_08 = Nothing
End Sub

Private Sub textCP44_Click()
   textCP44_2 = m_AgentList(textCP44.ListIndex).aiName
End Sub
'add by nickc 2007/04/02 阿蓮--聯絡單
Private Sub textCP44_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP56_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textCP56
End Sub

Private Sub textCP56_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 受讓人1
Private Sub textCP56_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCP56_2 = Empty
   If IsEmptyText(textCP56) = False Then
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
         strMsg = "受讓人1代碼<" & textCP56 & ">不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP56_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2016/6/17
Private Sub textCP89_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textCP89
End Sub
Private Sub textCP89_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
' 受讓人2
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
         strTit = "檢核資料"
         strMsg = "受讓人2代碼<" & textCP89 & ">不存在"
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
' 受讓人3
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
         strTit = "檢核資料"
         strMsg = "受讓人3代碼<" & textCP90 & ">不存在"
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
' 受讓人4
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
         strTit = "檢核資料"
         strMsg = "受讓人4代碼<" & textCP91 & ">不存在"
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
' 受讓人5
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
         strTit = "檢核資料"
         strMsg = "受讓人5代碼<" & textCP92 & ">不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP92_GotFocus
      End If
   End If
End Sub
'2016/6/17 END

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
      'edit by nickc 2006/03/17
      'If Val(DBDATE(textCP27)) > Val(DBDATE(Date)) Then
      If Val(DBDATE(textCP27)) > Val(strSrvDate(1)) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "發文日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' 計算催審期限
      If Me.textCP27.Tag <> Me.textCP27.Text Then 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
            strDay = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
            If IsEmptyText(strDay) = False Then
               textUargeDate = strDay
            End If
      End If
      Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
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
   Dim strTempName As String   '2010/11/24 add by sonia
   
   Cancel = False
   'Add By Cheng 2002/03/08
   If m_TM10 <> 台灣國家代號 Then
      If Len(Me.textCP44.Text) <= 0 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Me.SSTab1.Tab = 0
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

' 更新欄位的內容
Private Sub OnUpdateField()
   ' 點數
   SetCPFieldNewData "CP18", textCP18
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
   '若有輸入移轉申請人1
   If Me.textCP56.Text <> "" Then
       '若移轉人與原申請人不同時
       '2009/10/19 modify by sonia 應為判斷移轉申請人與原申請人不同時
       'If ChangeCustomerL(m_CP55) <> ChangeCustomerL(m_TM23) Then
       If ChangeCustomerL(textCP56) <> ChangeCustomerL(m_TM23) Then
           '更新進度檔移轉人
           SetCPFieldNewData "CP55", ChangeCustomerL(m_TM23)
       End If
   End If
   ' 受讓人1
   If IsEmptyText(textCP56) = False Then
      SetCPFieldNewData "CP56", textCP56 & String(9 - Len(textCP56), "0")
   Else
      SetCPFieldNewData "CP56", textCP56
   End If
   'Add By Sindy 2016/6/17
   If Me.textCP89.Text <> "" Then
       '若移轉人與原申請人不同時
       If ChangeCustomerL(textCP89) <> ChangeCustomerL(m_TM78) Then
          '更新進度檔移轉人
          SetCPFieldNewData "CP93", ChangeCustomerL(m_TM78)
       End If
   End If
   ' 移轉申請人2
   If IsEmptyText(textCP89) = False Then
      SetCPFieldNewData "CP89", textCP89 & String(9 - Len(textCP89), "0")
   Else
      SetCPFieldNewData "CP89", textCP89
   End If
   If Me.textCP90.Text <> "" Then
       '若移轉人與原申請人不同時
       If ChangeCustomerL(textCP90) <> ChangeCustomerL(m_TM79) Then
          '更新進度檔移轉人
          SetCPFieldNewData "CP94", ChangeCustomerL(m_TM79)
       End If
   End If
   ' 移轉申請人3
   If IsEmptyText(textCP90) = False Then
      SetCPFieldNewData "CP90", textCP90 & String(9 - Len(textCP90), "0")
   Else
      SetCPFieldNewData "CP90", textCP90
   End If
   If Me.textCP91.Text <> "" Then
      '若移轉人與原申請人不同時
      If ChangeCustomerL(textCP91) <> ChangeCustomerL(m_TM80) Then
         '更新進度檔移轉人
         SetCPFieldNewData "CP95", ChangeCustomerL(m_TM80)
      End If
   End If
   ' 移轉申請人4
   If IsEmptyText(textCP91) = False Then
      SetCPFieldNewData "CP91", textCP91 & String(9 - Len(textCP91), "0")
   Else
      SetCPFieldNewData "CP91", textCP91
   End If
   If Me.textCP92.Text <> "" Then
      '若移轉人與原申請人不同時
      If ChangeCustomerL(textCP92) <> ChangeCustomerL(m_TM81) Then
         '更新進度檔移轉人
         SetCPFieldNewData "CP96", ChangeCustomerL(m_TM81)
      End If
   End If
   ' 移轉申請人5
   If IsEmptyText(textCP92) = False Then
      SetCPFieldNewData "CP92", textCP92 & String(9 - Len(textCP92), "0")
   Else
      SetCPFieldNewData "CP92", textCP92
   End If
   '2016/6/17 END
   
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
   ' 進度備註
    'Modify By Cheng 2003/06/03
'   SetCPFieldNewData "CP64", textCP64 & " 不一併移轉註冊號數 : " & textTM15_S
   If Me.textTM15_S.Text <> "" Then SetCPFieldNewData "CP64", textCP64 & " 不一併移轉註冊號數 : " & textTM15_S
   
   'Add By Sindy 2010/3/3
   ' 代表人1(中)
   SetTMSPFieldNewData "TM47", textTM47
   ' 代表人1(英)
   SetTMSPFieldNewData "TM48", textTM48
   ' 代表人1(日)
   SetTMSPFieldNewData "TM49", textTM49
   ' 代表人2(中)
   SetTMSPFieldNewData "TM50", textTM50
   ' 代表人2(英)
   SetTMSPFieldNewData "TM51", textTM51
   ' 代表人2(日)
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
   '2010/3/3 End
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
Dim strNP08 As String
Dim strNP07 As String
Dim strNP22 As String
'Add By Cheng 2003/03/07
Dim StrSQLa As String
Dim strNP10 As String 'Add By Sindy 2014/9/11
   
'911106 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans
   
   'Add By Sindy 2010/3/3
   '更新商標基本檔
   OnUpdateTradeMark
   
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
   
   'Add By Cheng 2003/03/07
   '若有輸入移轉申請人1
   If Me.textCP56.Text <> "" Then
       '更新基本檔申請人及相關資料
       Select Case m_TM01
       'Modify By Cheng 2003/06/03
'        Case "T", "TF", "FCT"
       Case "T", "TF", "FCT", "CFT"
           StrSQLa = "Update TradeMark Set TM23='" & ChangeCustomerL(Me.textCP56.Text) & "' " & _
                           " ,TM24='" & ChgSQL(PUB_GetCustEachAdd(ChangeCustomerL(Me.textCP56.Text), "1")) & "' " & _
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
   'Add By Sindy 2016/6/17
   '若有輸入移轉申請人2
   If Me.textCP89.Text <> "" Then
       '更新基本檔申請人及相關資料
       Select Case m_TM01
       Case "T", "TF", "FCT", "CFT"
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
   'add by sonia 2017/8/1 CFT-016974
   Else
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT", "CFT"
            StrSQLa = "Update TradeMark Set TM78=NULL,TM82=NULL,TM86=NULL,TM90=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP58=NULL " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
    'end 2017/8/1
   End If
   '若有輸入移轉申請人3
   If Me.textCP90.Text <> "" Then
       '更新基本檔申請人及相關資料
       Select Case m_TM01
       Case "T", "TF", "FCT", "CFT"
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
   'add by sonia 2017/8/1 CFT-016974
   Else
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT", "CFT"
            StrSQLa = "Update TradeMark Set TM79=NULL,TM83=NULL,TM87=NULL,TM91=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP59=NULL " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
    'end 2017/8/1
   End If
   '若有輸入移轉申請人4
   If Me.textCP91.Text <> "" Then
       '更新基本檔申請人及相關資料
       Select Case m_TM01
       Case "T", "TF", "FCT", "CFT"
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
   'add by sonia 2017/8/1 CFT-016974
   Else
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT", "CFT"
            StrSQLa = "Update TradeMark Set TM80=NULL,TM84=NULL,TM88=NULL,TM92=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP65=NULL " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
    'end 2017/8/1
   End If
   '若有輸入移轉申請人5
   If Me.textCP92.Text <> "" Then
       '更新基本檔申請人及相關資料
       Select Case m_TM01
       Case "T", "TF", "FCT", "CFT"
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
   'add by sonia 2017/8/1 CFT-016974
   Else
        '更新基本檔申請人及相關資料
        Select Case m_TM01
        Case "T", "TF", "FCT", "CFT"
            StrSQLa = "Update TradeMark Set TM81=NULL,TM85=NULL,TM89=NULL,TM93=NULL " & _
                            " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        Case Else
            StrSQLa = "Update ServicePractice Set SP66=NULL " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute StrSQLa
        End Select
    'end 2017/8/1
   End If
   '2016/6/17 END
   
   ' 若有輸入催審期限時, 新增一筆催審的記錄到下一程序檔
   If IsEmptyText(textUargeDate) = False Then
      strNP07 = "305"
      strNP22 = GetNextProgressNo()
      '92.10.6 MODIFY BY SONIA
      'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
      '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
      '                  DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strUserNum & "'," & strNP22 & ")"
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
      Select Case strNP07
         Case "102", "105", "702", "708", "305", "998", "997":
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
        'Modify By Cheng 2003/09/02
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
         Select Case strNP07
            Case "102", "105", "702", "708", "305", "998", "997":
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
   'add by nickc 2007/01/30 要更新下一程序智權人員
   '2013/4/18 modify by sonia 改用共用函數
   'If rsTmp.State = 1 Then rsTmp.Close
   'strSql = "SELECT * FROM Customer Where Cu01 = '" & Mid(ChangeCustomerL(textCP56), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textCP56), 9, 1) & "' "
   'rsTmp.CursorLocation = adUseClient
   'rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'If rsTmp.RecordCount > 0 Then
   '     pub_ChgSalesTargetIsNp m_TM01, m_TM02, m_TM03, m_TM04, CheckStr(rsTmp.Fields("cu13"))
   'End If
   'rsTmp.Close
   'Set rsTmp = Nothing
   pub_ChgSalesTargetIsNp m_TM01, m_TM02, m_TM03, m_TM04, PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
   '2013/4/18 end
   
   'Added by Lydia 2024/07/09 判斷案件國家收費表內有設定提申期限(天)CF11，要加掛提申(998)期限；
   Call Pub_GetCF11to998(m_TM10, m_TM01, m_TM02, m_TM03, m_TM04, m_CP07, m_CP09, m_CP10, m_CP14, textCP27)
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
   
   'ADD BY SONIA 2021/3/10 若申請為旅狐國際及部分關係企業者，案件備註若無"不銷卷"字樣,則要加入(內商2015/11/24有控制)
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
   'END 2021/3/10
   
'911106 nick transation
    cnnConnection.CommitTrans
   
     'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
    PUB_CheckEMail m_CP44New, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22
    
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 列印定稿
   If textPrint <> "N" Then
      PrintLetter
   End If
   
    Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
   OnSaveData = False
End Function

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
   Cancel = False
   ' 無資料時不做任何檢查
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
         strTit = "檢核資料"
         strMsg = "註冊號數<" & strTemp & ">不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM15_S_GotFocus
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
               strTit = "檢核資料"
               strMsg = "註冊號數<" & strTemp & ">不可重覆"
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
   
   CheckDataValid = False
   'add by nickc 2008/05/01
   If IsDebt(m_TM10, textCP09) Then
        strTit = "警告！禁止發文！"
        strMsg = "未收款且無 預定收款日 請轉告智權同仁！！"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        Me.SSTab1.Tab = 0
        GoTo EXITSUB
   End If
   
   'Add By Sindy 2012/4/17
   If m_blnClkChgButton = False Then
      MsgBox "請輸入變更事項!!!", vbExclamation + vbOKOnly
      Me.cmdMod.SetFocus
      GoTo EXITSUB
   End If
   
   ' 發文日
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   'add by nickc 2006/03/17 加入驗證
   Dim Cancel As Boolean
   Cancel = False
   textCP27_Validate Cancel
   If Cancel = True Then GoTo EXITSUB
   
   ' 代理人
   If IsEmptyText(textCP44) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入代理人"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      textCP44.SetFocus
      GoTo EXITSUB
   End If
   
   ' 受讓人不可空白
   'Add By Sindy 2016/6/17
   If IsEmptyText(textCP56) = True And IsEmptyText(textCP89) = True And IsEmptyText(textCP90) = True And IsEmptyText(textCP91) = True And IsEmptyText(textCP92) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入受讓人"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Me.SSTab1.Tab = 0
      textCP56.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textUargeDate_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textUargeDate
End Sub

Private Sub textPrint_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textPrint
End Sub

Private Sub textTM15_S_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textTM15_S
End Sub

Private Sub textCP18_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textCP18
End Sub

Private Sub textCP27_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textCP27
End Sub

Private Sub textCP44_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textCP44
End Sub

Private Sub textCP64_GotFocus()
   Me.SSTab1.Tab = 0
   InverseTextBox textCP64
End Sub

Private Sub textCF09_GotFocus()
   Me.SSTab1.Tab = 0
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
            ' 回音
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                     "','回音','" & textCF09 & "')"
            cnnConnection.Execute strSql
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
   
   'Add By Sindy 2016/6/17
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
   '2016/6/17 END
   
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
