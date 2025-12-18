VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_05 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(查名, 申請, 延展, 補換發註冊證, 英文證明, 中文證明, 領土延伸, 刊登廣告, 分割, 商業司查詢)"
   ClientHeight    =   5748
   ClientLeft      =   4320
   ClientTop       =   2712
   ClientWidth     =   9132
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9132
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1392
      Locked          =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   936
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5160
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   1236
      Width           =   3945
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1392
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   1236
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1392
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   336
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1392
      Locked          =   -1  'True
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   636
      Width           =   2532
   End
   Begin VB.TextBox textTM12_S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   936
      Width           =   3945
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5160
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   636
      Width           =   3945
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   325
      Left            =   8292
      TabIndex        =   54
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   325
      Index           =   0
      Left            =   6300
      TabIndex        =   52
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   325
      Left            =   7095
      TabIndex        =   53
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H008080FF&
      Caption         =   "變更事項(&R)"
      Height          =   325
      Left            =   5100
      Style           =   1  '圖片外觀
      TabIndex        =   51
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   325
      Left            =   3885
      TabIndex        =   50
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   325
      Index           =   1
      Left            =   2685
      TabIndex        =   49
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3930
      Left            =   150
      TabIndex        =   55
      Top             =   1830
      Width           =   8895
      _ExtentX        =   15685
      _ExtentY        =   6922
      _Version        =   393216
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料1"
      TabPicture(0)   =   "frm020102_05.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(12)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label16"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label15"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(10)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label23"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label22"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label4"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label14"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label25"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label38"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label39"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label27"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label40"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label41"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label42"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label43"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblPayToday"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblCP113(18)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textTM05"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textTM07"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textTM23_2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCP44_2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textTM05_1"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textTM78_2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textTM79_2"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textTM80_2"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textTM81_2"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textCP44"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textTM23"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textCF09"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCP26"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textPetition"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCP18"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textTM32"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textTM06"
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
      Tab(0).Control(50)=   "textCP84"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "textTM78"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "textTM79"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "textTM80"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "textTM81"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "cmdGoods"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "textTM09"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "textCP118"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "cmdCountry"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtPayToday"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtCP113"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).ControlCount=   61
      TabCaption(1)   =   "基本資料2"
      TabPicture(1)   =   "frm020102_05.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textTM08_2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "textTM27"
      Tab(1).Control(2)=   "textTM08"
      Tab(1).Control(3)=   "cmdPriority"
      Tab(1).Control(4)=   "textPriorityDate"
      Tab(1).Control(5)=   "textCP22"
      Tab(1).Control(6)=   "textTM72_2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "textTM72"
      Tab(1).Control(8)=   "textTM02_2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "textTM02"
      Tab(1).Control(10)=   "textTM04"
      Tab(1).Control(11)=   "textTM03"
      Tab(1).Control(12)=   "textTM01"
      Tab(1).Control(13)=   "textMediaDate"
      Tab(1).Control(14)=   "textMediaName"
      Tab(1).Control(15)=   "textMediaType"
      Tab(1).Control(16)=   "textTM12"
      Tab(1).Control(17)=   "textTM11"
      Tab(1).Control(18)=   "textMail"
      Tab(1).Control(19)=   "Label44"
      Tab(1).Control(20)=   "cboTM72"
      Tab(1).Control(21)=   "cboTM08"
      Tab(1).Control(22)=   "textTM67"
      Tab(1).Control(23)=   "lstNameAgent"
      Tab(1).Control(24)=   "textCP64"
      Tab(1).Control(25)=   "textTM58"
      Tab(1).Control(26)=   "Label1(8)"
      Tab(1).Control(27)=   "Label1(4)"
      Tab(1).Control(28)=   "Label26"
      Tab(1).Control(29)=   "Label1(13)"
      Tab(1).Control(30)=   "Label30"
      Tab(1).Control(31)=   "Label31"
      Tab(1).Control(32)=   "lblNameAgent"
      Tab(1).Control(33)=   "Label1(14)"
      Tab(1).Control(34)=   "Label36"
      Tab(1).Control(35)=   "Label33"
      Tab(1).Control(36)=   "Label32"
      Tab(1).Control(37)=   "Label20"
      Tab(1).Control(38)=   "Label19"
      Tab(1).Control(39)=   "Label13"
      Tab(1).Control(40)=   "Label12"
      Tab(1).Control(41)=   "Label17"
      Tab(1).Control(42)=   "Label18"
      Tab(1).Control(43)=   "Label21"
      Tab(1).Control(44)=   "Label28"
      Tab(1).Control(45)=   "Label29"
      Tab(1).ControlCount=   46
      TabCaption(2)   =   "刊登廣告明細"
      TabPicture(2)   =   "frm020102_05.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdList"
      Tab(2).Control(1)=   "cmdAddItem"
      Tab(2).Control(2)=   "cmdModItem"
      Tab(2).Control(3)=   "cmdDelItem"
      Tab(2).Control(4)=   "textTM15S"
      Tab(2).Control(5)=   "Label35(3)"
      Tab(2).Control(6)=   "Label35(2)"
      Tab(2).Control(7)=   "Label35(1)"
      Tab(2).Control(8)=   "Label35(0)"
      Tab(2).Control(9)=   "nick911015(3)"
      Tab(2).Control(10)=   "nick911015(2)"
      Tab(2).Control(11)=   "nick911015(1)"
      Tab(2).Control(12)=   "nick911015(0)"
      Tab(2).Control(13)=   "Label34"
      Tab(2).ControlCount=   14
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   2172
         Left            =   -74904
         TabIndex        =   146
         Top             =   1656
         Width           =   8736
         _ExtentX        =   15409
         _ExtentY        =   3831
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtCP113 
         Height          =   270
         Left            =   7950
         MaxLength       =   4
         TabIndex        =   3
         Top             =   555
         Width           =   540
      End
      Begin VB.TextBox txtPayToday 
         Height          =   264
         Left            =   8085
         MaxLength       =   1
         TabIndex        =   142
         Top             =   1152
         Width           =   255
      End
      Begin VB.CommandButton cmdCountry 
         Caption         =   "指定國家"
         Height          =   300
         Left            =   2160
         TabIndex        =   140
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox textCP118 
         Height          =   270
         Left            =   7890
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1416
         Width           =   375
      End
      Begin VB.TextBox textTM09 
         Height          =   264
         Left            =   1200
         MaxLength       =   395
         TabIndex        =   22
         Top             =   3330
         Width           =   7515
      End
      Begin VB.CommandButton cmdGoods 
         Caption         =   "商品名稱"
         Height          =   325
         Left            =   7860
         TabIndex        =   21
         Top             =   3030
         Width           =   990
      End
      Begin VB.TextBox textTM81 
         Height          =   264
         Left            =   870
         MaxLength       =   9
         TabIndex        =   20
         Top             =   3060
         Width           =   885
      End
      Begin VB.TextBox textTM80 
         Height          =   264
         Left            =   5250
         MaxLength       =   9
         TabIndex        =   19
         Top             =   2790
         Width           =   885
      End
      Begin VB.TextBox textTM79 
         Height          =   264
         Left            =   870
         MaxLength       =   9
         TabIndex        =   18
         Top             =   2790
         Width           =   885
      End
      Begin VB.TextBox textTM78 
         Height          =   264
         Left            =   5250
         MaxLength       =   9
         TabIndex        =   17
         Top             =   2484
         Width           =   885
      End
      Begin VB.TextBox textTM08_2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -67656
         Locked          =   -1  'True
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   1224
         Visible         =   0   'False
         Width           =   1404
      End
      Begin VB.TextBox textTM27 
         Height          =   264
         Left            =   -69432
         MaxLength       =   20
         TabIndex        =   27
         Top             =   660
         Width           =   2292
      End
      Begin VB.TextBox textTM08 
         BackColor       =   &H00FFFFC0&
         Height          =   264
         Left            =   -68112
         MaxLength       =   1
         TabIndex        =   32
         Top             =   1248
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdPriority 
         Caption         =   "輸入(&V)"
         Height          =   252
         Left            =   -69990
         TabIndex        =   34
         Top             =   1560
         Width           =   1332
      End
      Begin VB.TextBox textPriorityDate 
         Height          =   264
         Left            =   -69450
         MaxLength       =   9
         TabIndex        =   29
         Top             =   930
         Width           =   1092
      End
      Begin VB.TextBox textCP22 
         Height          =   264
         Left            =   -66645
         MaxLength       =   1
         TabIndex        =   123
         Top             =   1770
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textCP84 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Left            =   5445
         TabIndex        =   1
         Top             =   270
         Width           =   1425
      End
      Begin VB.TextBox textTM72_2 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H80000007&
         Height          =   264
         Left            =   -70656
         Locked          =   -1  'True
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   3528
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox textTM72 
         BackColor       =   &H00FFFFC0&
         Height          =   264
         Left            =   -71136
         MaxLength       =   1
         TabIndex        =   44
         Top             =   3552
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "新增(&A)"
         Height          =   324
         Left            =   -69168
         TabIndex        =   46
         Top             =   360
         Width           =   972
      End
      Begin VB.CommandButton cmdModItem 
         Caption         =   "修改(&M)"
         Height          =   324
         Left            =   -68184
         TabIndex        =   47
         Top             =   360
         Width           =   972
      End
      Begin VB.CommandButton cmdDelItem 
         Caption         =   "刪除(&D)"
         Height          =   324
         Left            =   -67200
         TabIndex        =   48
         Top             =   360
         Width           =   972
      End
      Begin VB.TextBox textTM15S 
         Height          =   264
         Left            =   -74064
         MaxLength       =   20
         TabIndex        =   45
         Top             =   360
         Width           =   1092
      End
      Begin VB.TextBox textTM02_2 
         Height          =   264
         Left            =   -72120
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1860
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox textTM02 
         Height          =   264
         Left            =   -72948
         MaxLength       =   6
         TabIndex        =   36
         Top             =   1860
         Width           =   1095
      End
      Begin VB.TextBox textTM04 
         Height          =   264
         Left            =   -71640
         MaxLength       =   2
         TabIndex        =   39
         Top             =   1860
         Width           =   492
      End
      Begin VB.TextBox textTM03 
         Height          =   264
         Left            =   -71868
         MaxLength       =   1
         TabIndex        =   38
         Top             =   1860
         Width           =   255
      End
      Begin VB.TextBox textTM01 
         Height          =   264
         Left            =   -73560
         MaxLength       =   3
         TabIndex        =   35
         Top             =   1860
         Width           =   612
      End
      Begin VB.TextBox textMediaDate 
         Enabled         =   0   'False
         Height          =   264
         Left            =   -73080
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   30
         Top             =   1260
         Width           =   1092
      End
      Begin VB.TextBox textMediaName 
         Height          =   264
         Left            =   -73560
         MaxLength       =   12
         TabIndex        =   33
         Top             =   1545
         Width           =   2292
      End
      Begin VB.TextBox textMediaType 
         Height          =   264
         Left            =   -73080
         MaxLength       =   1
         TabIndex        =   28
         Top             =   960
         Width           =   372
      End
      Begin VB.TextBox textTM12 
         Height          =   264
         Left            =   -73560
         MaxLength       =   9
         TabIndex        =   26
         Top             =   660
         Width           =   2292
      End
      Begin VB.TextBox textTM11 
         Height          =   264
         Left            =   -69450
         MaxLength       =   8
         TabIndex        =   25
         Top             =   360
         Width           =   1092
      End
      Begin VB.TextBox textCP27 
         Height          =   264
         Left            =   1020
         MaxLength       =   8
         TabIndex        =   0
         Top             =   300
         Width           =   1092
      End
      Begin VB.TextBox textUargeDate 
         Height          =   264
         Left            =   1020
         MaxLength       =   8
         TabIndex        =   2
         Top             =   570
         Width           =   1092
      End
      Begin VB.TextBox textTM22 
         Height          =   264
         Left            =   2490
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1140
         Width           =   1092
      End
      Begin VB.TextBox textTM21 
         Height          =   264
         Left            =   1020
         MaxLength       =   8
         TabIndex        =   6
         Top             =   1140
         Width           =   1092
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   1020
         MaxLength       =   1
         TabIndex        =   9
         Top             =   1395
         Width           =   372
      End
      Begin VB.TextBox textTM06 
         Height          =   264
         Left            =   1440
         MaxLength       =   60
         TabIndex        =   14
         Top             =   1956
         Width           =   7272
      End
      Begin VB.TextBox textTM32 
         Height          =   264
         Left            =   1200
         MaxLength       =   699
         TabIndex        =   23
         Top             =   3600
         Width           =   7512
      End
      Begin VB.TextBox textCP18 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   5460
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   600
         Width           =   1005
      End
      Begin VB.TextBox textPetition 
         Height          =   264
         Left            =   7380
         MaxLength       =   8
         TabIndex        =   5
         Top             =   855
         Width           =   1092
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   4920
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1152
         Width           =   372
      End
      Begin VB.TextBox textCF09 
         Height          =   264
         Left            =   4920
         MaxLength       =   12
         TabIndex        =   10
         Top             =   1416
         Width           =   612
      End
      Begin VB.TextBox textMail 
         Height          =   264
         Left            =   -73560
         MaxLength       =   1
         TabIndex        =   24
         Top             =   360
         Width           =   372
      End
      Begin VB.TextBox textTM23 
         Height          =   264
         Left            =   870
         MaxLength       =   9
         TabIndex        =   16
         Top             =   2484
         Width           =   885
      End
      Begin VB.ComboBox textCP44 
         Height          =   276
         Left            =   1020
         TabIndex        =   4
         Top             =   825
         Width           =   1356
      End
      Begin VB.Label Label44 
         Caption         =   "放棄專用權以逗號區隔"
         ForeColor       =   &H000000FF&
         Height          =   204
         Left            =   -71064
         TabIndex        =   147
         Top             =   1944
         Width           =   1812
      End
      Begin MSForms.ComboBox cboTM72 
         Height          =   300
         Left            =   -73560
         TabIndex        =   43
         Top             =   3456
         Width           =   2124
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3746;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboTM08 
         Height          =   288
         Left            =   -70008
         TabIndex        =   31
         Top             =   1224
         Width           =   1932
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3408;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM67 
         Height          =   270
         Left            =   -73560
         TabIndex        =   40
         Top             =   2160
         Width           =   7275
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "12827;466"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   825
         Left            =   -68130
         TabIndex        =   145
         Top             =   1500
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;1455"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   672
         Left            =   -73560
         TabIndex        =   41
         Top             =   2460
         Width           =   7272
         VariousPropertyBits=   679493659
         MaxLength       =   2000
         Size            =   "12827;1185"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM58 
         Height          =   300
         Left            =   -73560
         TabIndex        =   42
         Top             =   3150
         Width           =   7272
         VariousPropertyBits=   679493659
         MaxLength       =   2000
         Size            =   "12827;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM81_2 
         Height          =   300
         Left            =   1770
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   3060
         Width           =   2655
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         Size            =   "4683;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM80_2 
         Height          =   300
         Left            =   6150
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   2790
         Width           =   2655
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         Size            =   "4683;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM79_2 
         Height          =   300
         Left            =   1770
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   2790
         Width           =   2655
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         Size            =   "4683;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM78_2 
         Height          =   300
         Left            =   6150
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   2484
         Width           =   2655
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         Size            =   "4683;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05_1 
         Height          =   792
         Left            =   1440
         TabIndex        =   12
         Top             =   1680
         Width           =   7272
         VariousPropertyBits=   679493659
         MaxLength       =   140
         ScrollBars      =   2
         Size            =   "12827;1397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   300
         Left            =   2415
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   825
         Width           =   3855
         VariousPropertyBits=   679493663
         MaxLength       =   20
         Size            =   "6800;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM23_2 
         Height          =   300
         Left            =   1770
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2484
         Width           =   2655
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         Size            =   "4683;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM07 
         Height          =   264
         Left            =   1440
         TabIndex        =   15
         Top             =   2220
         Width           =   7272
         VariousPropertyBits=   679493661
         MaxLength       =   40
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM05 
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1692
         Width           =   7272
         VariousPropertyBits=   679493661
         MaxLength       =   40
         Size            =   "12827;503"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCP113 
         AutoSize        =   -1  'True
         Caption         =   "工作時數:"
         Height          =   180
         Index           =   18
         Left            =   7050
         TabIndex        =   144
         Top             =   600
         Width           =   765
      End
      Begin VB.Label lblPayToday 
         AutoSize        =   -1  'True
         Caption         =   "電子送件是否當日扣款:         (Y/N)"
         Height          =   180
         Left            =   6150
         TabIndex        =   143
         Top             =   1170
         Width           =   2655
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "是否電子送件:          (Y: 是)"
         Height          =   180
         Left            =   6720
         TabIndex        =   139
         Top             =   1470
         Width           =   2085
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "申請人5 :"
         Height          =   180
         Left            =   120
         TabIndex        =   138
         Top             =   3105
         Width           =   720
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "申請人4 :"
         Height          =   180
         Left            =   4500
         TabIndex        =   136
         Top             =   2820
         Width           =   720
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "申請人3 :"
         Height          =   180
         Left            =   120
         TabIndex        =   134
         Top             =   2835
         Width           =   720
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "申請人2 :"
         Height          =   180
         Left            =   4500
         TabIndex        =   132
         Top             =   2526
         Width           =   720
      End
      Begin VB.Label Label1 
         Caption         =   "正商標號數:"
         Height          =   255
         Index           =   8
         Left            =   -71040
         TabIndex        =   130
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "商標種類 :"
         Height          =   255
         Index           =   4
         Left            =   -71040
         TabIndex        =   129
         Top             =   1230
         Width           =   855
      End
      Begin VB.Label Label26 
         Caption         =   "優先權資料 :"
         Height          =   255
         Left            =   -71040
         TabIndex        =   128
         Top             =   1559
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "補優先權文件期限:"
         Height          =   255
         Index           =   13
         Left            =   -71040
         TabIndex        =   127
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label30 
         Caption         =   "是否出名 :"
         Height          =   255
         Left            =   -67770
         TabIndex        =   125
         Top             =   1815
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "(N:不出名)"
         Height          =   255
         Left            =   -66165
         TabIndex        =   124
         Top             =   1770
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   -69090
         TabIndex        =   122
         Top             =   1860
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Left            =   4440
         TabIndex        =   121
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "特殊商標 :"
         Height          =   255
         Index           =   14
         Left            =   -74880
         TabIndex        =   120
         Top             =   3525
         Width           =   855
      End
      Begin VB.Label Label38 
         Caption         =   "案件名稱 :"
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label35 
         Caption         =   "日文名稱:"
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   116
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "英文名稱:"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   115
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "中文名稱:"
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   114
         Top             =   930
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "本所案號:"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   113
         Top             =   690
         Width           =   855
      End
      Begin MSForms.Label nick911015 
         Height          =   300
         Index           =   3
         Left            =   -73920
         TabIndex        =   112
         Top             =   1455
         Width           =   7695
         VariousPropertyBits=   27
         Size            =   "13573;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label nick911015 
         Height          =   300
         Index           =   2
         Left            =   -73920
         TabIndex        =   111
         Top             =   1230
         Width           =   7695
         VariousPropertyBits=   27
         Size            =   "13573;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label nick911015 
         Height          =   300
         Index           =   1
         Left            =   -73920
         TabIndex        =   110
         Top             =   945
         Width           =   7695
         VariousPropertyBits=   27
         Size            =   "13573;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label nick911015 
         Height          =   300
         Index           =   0
         Left            =   -73920
         TabIndex        =   109
         Top             =   690
         Width           =   7695
         VariousPropertyBits=   27
         Size            =   "13573;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label34 
         Caption         =   "審定號 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   108
         Top             =   360
         Width           =   828
      End
      Begin VB.Label Label36 
         Caption         =   "查名本所案號 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   107
         Top             =   1860
         Width           =   1245
      End
      Begin VB.Label Label33 
         Caption         =   "下次刊登廣告期限 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   106
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label Label32 
         Caption         =   "雜誌社, 報社 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   105
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "(1:雜誌 2:報紙)"
         Height          =   255
         Left            =   -72600
         TabIndex        =   104
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label19 
         Caption         =   "大陸刊登廣告媒體 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   103
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "申請案號 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   102
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "申請日 :"
         Height          =   255
         Left            =   -71040
         TabIndex        =   101
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "發文日 :"
         Height          =   255
         Left            =   120
         TabIndex        =   81
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "催審期限 :"
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "代理人 :"
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   870
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   2250
         X2              =   2370
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label Label5 
         Caption         =   "專用期限 :"
         Height          =   225
         Left            =   120
         TabIndex        =   78
         Top             =   1170
         Width           =   1455
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
         Height          =   180
         Left            =   1410
         TabIndex        =   76
         Top             =   1470
         Width           =   2745
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "申請人1 :"
         Height          =   180
         Left            =   120
         TabIndex        =   75
         Top             =   2526
         Width           =   720
      End
      Begin VB.Label Label7 
         Caption         =   "案件日文名稱 :"
         Height          =   252
         Left            =   120
         TabIndex        =   74
         Top             =   2232
         Width           =   1452
      End
      Begin VB.Label Label8 
         Caption         =   "案件英文名稱 :"
         Height          =   252
         Left            =   120
         TabIndex        =   73
         Top             =   1944
         Width           =   1212
      End
      Begin VB.Label Label9 
         Caption         =   "案件中文名稱 :"
         Height          =   252
         Left            =   120
         TabIndex        =   72
         Top             =   1680
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "商品類別 :"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   71
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "商品組群 :"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   70
         Top             =   3630
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "點數 :"
         Height          =   255
         Index           =   10
         Left            =   4470
         TabIndex        =   69
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "提申期限 :"
         Height          =   255
         Left            =   6510
         TabIndex        =   68
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "(N:不算)"
         Height          =   255
         Left            =   5370
         TabIndex        =   67
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "是否算案件數 :"
         Height          =   255
         Left            =   3660
         TabIndex        =   66
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "大約"
         Height          =   255
         Index           =   12
         Left            =   4440
         TabIndex        =   65
         Top             =   1470
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "可接獲回音"
         Height          =   255
         Left            =   5580
         TabIndex        =   64
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "(Y:郵寄)"
         Height          =   252
         Left            =   -73080
         TabIndex        =   63
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label18 
         Caption         =   "是否郵寄申請 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   62
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label21 
         Caption         =   "放棄專用權 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   61
         Top             =   2190
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   60
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "案件備註 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   59
         Top             =   3135
         Width           =   975
      End
   End
   Begin MSForms.TextBox textTM44 
      Height          =   285
      Left            =   5160
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   330
      Width           =   3945
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6959;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5160
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   1536
      Width           =   3945
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6959;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1392
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   1536
      Width           =   2532
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "4466;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   252
      Left            =   312
      TabIndex        =   100
      Top             =   936
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4155
      TabIndex        =   99
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   255
      Index           =   9
      Left            =   4155
      TabIndex        =   98
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   312
      TabIndex        =   97
      Top             =   1236
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   312
      TabIndex        =   96
      Top             =   336
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   312
      TabIndex        =   95
      Top             =   636
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   255
      Left            =   4155
      TabIndex        =   94
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "發證日 :"
      Height          =   255
      Index           =   3
      Left            =   4155
      TabIndex        =   93
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "FC代理人 :"
      Height          =   255
      Index           =   2
      Left            =   4155
      TabIndex        =   92
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   312
      TabIndex        =   91
      Top             =   1536
      Width           =   852
   End
   Begin VB.Label Label37 
      Caption         =   "TS商品類別 不可 輸在""案件備註""欄!!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   90
      TabIndex        =   117
      Top             =   0
      Width           =   1965
   End
End
Attribute VB_Name = "frm020102_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/13 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/22 Form2.0已修改 textTM44/textCP14/textCP13/textCP44_2/textTM05_1/textTM05/textTM07/textTM23_2(申請人名).../textCP64/textCP58/nick911015()/lstNameAgent/grdList/textTM67(111/8/8 Lydia)
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
'Add By Sindy 2009/04/30
Dim m_DC01 As String
Dim m_DC02 As String
Dim m_DC03 As String
Dim m_DC04 As String
Dim m_DC05 As String
Dim m_DC06 As String
Dim m_DC07 As String
Dim m_DC08 As String
'Dim m_DcTM15 As String
'Dim m_DcTM16 As String
'Add By Cheng 2002/06/14
Dim m_TM08 As String '商標種類代號
' 收文號
Dim m_CP09 As String
Dim m_CP31 As String 'Add By Sindy 2011/7/12
' 智權人員
Dim m_CP13 As String
Dim m_CP12 As String
' 申請國家
Dim m_TM10 As String
' 案件性質代號
Dim m_CP10 As String
' 發證日  add by sonia 2023/11/17
Dim m_TM20 As String
' 原專用期限起日
Dim m_TM21 As String
' 原專用期限止日
Dim m_TM22 As String
' 申請人
Dim m_TM23 As String
'add by nickc 2007/01/02
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
' 法定期限  2007/8/7 ADD BY SONIA
Dim m_CP07 As String
' 申請國家的延展年度
Dim m_NA14 As String
' 申請國家的延展時間(月)   2007/8/7 ADD BY SONIA
Dim m_NA15 As String
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 宣告欄位內容結構
Private Type DBFIELDITEM
   fiName As String
   fiData As String
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
'add by nickc 2007/01/02
Dim m_strCust2 As String '申請人2
Dim m_strCust3 As String '申請人3
Dim m_strCust4 As String '申請人4
Dim m_strCust5 As String '申請人5
Public m_CU10 As String
Dim m_NewCPList() As FIELDITEM
Dim m_NewCPListCount As Integer
'Add By Cheng 2002/11/18
Dim m_FA10 As String 'FC代理人
'Add By Cheng 2002/12/12
Dim m_CP14 As String '原承辦人
'Add By Cheng 2003/10/06
Public m_blnClkChgButton As Boolean '是否按下變更事項按鈕
'add by nick 2004/08/12
Dim m_CP84 As String       '發文規費
'add by nick 2004/09/27
Public m_CU103 As String         '公司負責人英文名稱
'add by nick 2004/10/05
Public m_CU05 As String         '客戶英文名稱
Public m_CU88 As String         '客戶英文名稱
Public m_CU89 As String         '客戶英文名稱
Public m_CU90 As String         '客戶英文名稱
'add by nickc 2005/11/18
Dim m_TM24 As String
Dim m_tm25 As String
Dim m_tm26 As String
'Add By Sindy 2012/2/7
Public m_CU39 As String         '代表人1（中）
Public m_CU40 As String         '代表人1（英）
Public m_CU41 As String         '代表人1（日）
'2012/2/7 End

'add by nickc 2007/01/02
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

'add by nickc 2006/01/20
Public m_CU112 As String        '客戶中文地址郵遞區號
'add by nickc 2006/01/27
Dim m_CP110 As String
'add by nickc 2006/06/02
Dim strCountry As String '指定國家
Dim strLicenceCountry As String '勾選指定國家
'add by nickc 2006/06/20
Dim IsHaveGoods As Boolean
Public ChkTG As Boolean
'2006/11/13 ADD BY SONIA
Dim m_textUargeDate As String
'add by nickc 2006/11/17
Dim m_textPrint As String
'add by nickc 2007/08/10
Dim SeekCu05(1 To 5) As String
Dim SeekCu88(1 To 5) As String
Dim SeekCu89(1 To 5) As String
Dim SeekCu90(1 To 5) As String
Dim SeekCu103(1 To 5) As String
Dim SeekCu112(1 To 5) As String
'add by nickc 2008/02/22
'Add By Sindy 2012/2/7
Dim SeekCu39(1 To 5) As String
Dim SeekCu40(1 To 5) As String
Dim SeekCu41(1 To 5) As String
'2012/2/7 End
'Add By Sindy 2012/10/31
Dim SeekCu10(1 To 5) As String
'2012/2/7 End
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP09s As String, m_CP123s As String 'Add by Sindy 98/3/24 收文號,是否算發文室案件
Dim m_CP130s As String 'Add by Sindy 2009/4/24 發文-主管機關
'Dim strCP45 As String
Dim m_TM09 As String, tmpArr As Variant 'Add By Sindy 2014/2/20
Dim m_CP60 As String 'Add By Sindy 2014/3/10
Dim m_QSP As Boolean 'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
Dim tmpGoods1205 As String 'Add by Amy 2014/10/16 部分核駁商品及服務
Dim m_990CP09 As String 'Add By Sindy 2016/12/16
Dim mTQD11 As String 'Added by Lydia 2019/12/12 內商T案申請:查名單近似本所案經核可後，設定" 是否出名 "
Dim strLD18 As String 'Add By Sindy 2019/12/20 信函總收文號
Dim m_strCF10 As String 'Add By Sindy 2020/8/12 取得主管機關
Dim m_AgentName As String 'Add By Amy 2021/12/23
Dim bolFixed  As Boolean 'Added by Lydia 2023/10/13
Dim strPTM As String, strSPT As String 'Added by Lydia 2023/11/16 暫存商標種類及特殊商標的Combo.ItemData
Dim m_CP43 As String 'Added by Lydia 2024/11/21 相關收文號

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

'Modify By Sindy 2012/6/21 原在確定按鍵裡,提出來為獨立按鍵
Private Sub cmdCountry_Click()
   'add by nickc 2006/06/01 加入馬德里續展子案項目
   strCountry = ""
   strLicenceCountry = "" '預設都先不勾
   If m_TM01 = "TF" And m_CP10 = "102" Then
      CheckOC3
      '2006/10/16 MODIFY BY SONIA 母案延展抓出所有未閉卷子案,含領土延伸之子案
      '                           領土延伸延展只抓出該領土延伸之未閉卷子案
      'strSQL = "select * from trademark where tm01='" & m_TM01 & "' and substr(tm02,1,5)=substr('" & m_TM02 & "',1,5) and tm03<>'0' and (tm16='1' or tm16 is null or tm16='') order by tm04 "
      'Modify By Sindy 2012/6/21 Mark
'      If Mid(m_TM02, 6, 1) = "0" Then
         strSql = "select tm10 from trademark where tm01='" & m_TM01 & "' and substr(tm02,1,5)=substr('" & m_TM02 & "',1,5) and tm03<>'0' and (tm16='1' or tm16 is null or tm16='') AND TM29 IS NULL order by tm04 "
'      Else
'         strSql = "select * from trademark where tm01='" & m_TM01 & "' and TM02='" & m_TM02 & "' and tm03<>'0' and (tm16='1' or tm16 is null or tm16='') AND TM29 IS NULL order by tm04 "
'      End If
      '2012/6/21 End
      '2006/10/16 END
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount <> 0 Then
         AdoRecordSet3.MoveFirst
         Do While Not AdoRecordSet3.EOF
            strCountry = strCountry & "" & AdoRecordSet3.Fields("TM10") & ","
            AdoRecordSet3.MoveNext
         Loop
         Do While strLicenceCountry = ""
         frm880008.strCountry = strCountry
         frm880008.strLicenceCountry = strLicenceCountry
         frm880008.Caption = "馬德里子案續展選擇"
         frm880008.IsByTM = True
         'add by sonia 2014/10/31
         frm880008.Command1(0).Visible = False
         frm880008.Command1(1).Visible = False
         'end 2014/10/31
         frm880008.Show vbModal
         strCountry = frm880008.strCountry
         strLicenceCountry = frm880008.strLicenceCountry
         If frm880008.IsByTM = False Then
            Unload frm880008
            Exit Sub
         Else
            Unload frm880008
         End If
         Dim yoy As Variant
         Dim ioi As Integer
         yoy = Split(strLicenceCountry, ",")
         strLicenceCountry = ""
         For ioi = 0 To UBound(yoy)
            If Trim(yoy(ioi)) <> "" Then
               If strLicenceCountry <> "" Then
                  strLicenceCountry = strLicenceCountry & ","
               End If
               strLicenceCountry = strLicenceCountry & "'" & yoy(ioi) & "'"
            End If
         Next ioi
         If strLicenceCountry = "" Then MsgBox "最少要勾選一個", vbInformation, "操作錯誤！"
         Loop
      Else
         MsgBox "查無任何子案可以續展，請補輸後再繼續！", vbExclamation
         Exit Sub
      End If
      CheckOC3
   End If
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

Private Sub cmdGoods_Click()
frm03010303_04.Hide
Set frm03010303_04.UpForm = Me
frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
frm03010303_04.AllClass = textTM09.Text
frm03010303_04.cmdOK(2).Visible = True
Me.Hide
frm03010303_04.QueryData
frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
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
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
'Dim strCP31 As String     'Add By Sindy 2009/04/23
'Add By Sindy 2009/10/21
Dim bDiviCSon As Boolean
Dim intTemp As Integer
Dim strTemp As String
'2009/10/21 End
'Add By Sindy 2011/1/26
Dim strApplID As String
Dim rsAddrNotAlike As New ADODB.Recordset
'2011/1/26
Dim strNewCP64 As String 'Add by Amy 2020/02/05 進度備註

   'Modify By Sindy 2010/11/19 把「確定」及「同時發文」按鈕程式碼合併
   Select Case Index
      Case 0, 1
         If CheckDataValid = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            'add by nickc 2006/06/20  加入 101 時，若申請國家非台灣，則必須要有輸入，台灣只是提醒
            If m_CP10 = "101" And (m_TM01 = "T" Or m_TM01 = "TF") Then
              Dim arrTM09 As Variant
              Dim iTm09 As Integer
              IsHaveGoods = True
              arrTM09 = Split(textTM09.Text, ",")
              For iTm09 = 0 To UBound(arrTM09)
                  CheckOC3
                  'Modify By Sindy 2013/6/21
                  'strSql = "select * from tmgoods where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05='" & arrTM09(iTm09) & "' and length(rtrim(tg06||tg07||tg08))>0"
                  strSql = "select * from tmgoods where tg01='" & m_TM01 & "' and tg02='" & m_TM02 & "' and tg03='" & m_TM03 & "' and tg04='" & m_TM04 & "' and tg05='" & arrTM09(iTm09) & "' and (length(rtrim(tg06))>0 or length(rtrim(tg07))>0 or length(rtrim(tg08))>0)"
                  '2013/6/21 END
                  AdoRecordSet3.CursorLocation = adUseClient
                  AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If AdoRecordSet3.RecordCount = 0 Then
                      IsHaveGoods = False
                      Exit For
                  End If
              Next iTm09
              If m_TM10 <> "000" And IsHaveGoods = False Then
                  MsgBox "請輸入商品名稱！", vbExclamation
                  Exit Sub
              End If
              If IsHaveGoods = False Then
                  If MsgBox("商品名稱未建立完全，是否要補資料？", vbInformation + vbYesNo) = vbYes Then
                      Exit Sub
                  End If
              End If
            End If
            'add 2006/06/20 end
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
                  'Modified by Lydia 2024/07/03 改傳入變數;
                  'GetCu103ByCustomer Me, textTM23.Text
                  Call Pub_GetDataFrm020102(textTM23.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                  
                  'edit by nickc 2006/01/20
                  'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Then
                  'Modify By Sindy 2012/2/7
                  'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Or (m_CU39 & m_CU40 & m_CU41) = "" Then
                        'Modified by Lydia 2024/07/03
                        'Set frm020102_22.oNextForm = Me
                        Call frm020102_22.SetParent(Me, Me.textTM23.Text)
                        frm020102_22.Label4.Caption = textTM23.Text & " " & textTM23_2 'Add By Sindy 2014/7/30
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
                  
                  If textTM78.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                  'Modified by Lydia 2024/07/03 改傳入變數;
                  'GetCu103ByCustomer Me, textTM78.Text
                  Call Pub_GetDataFrm020102(textTM78.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                  
                  'Modify By Sindy 2012/2/7
                  'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Or (m_CU39 & m_CU40 & m_CU41) = "" Then
                        'Modified by Lydia 2024/07/03
                        'Set frm020102_22.oNextForm = Me
                        Call frm020102_22.SetParent(Me, Me.textTM78.Text)
                        frm020102_22.Label4.Caption = textTM78.Text & " " & textTM78_2 'Add By Sindy 2014/7/30
                        frm020102_22.Show vbModal
                        SeekCu05(2) = m_CU05
                        SeekCu88(2) = m_CU88
                        SeekCu89(2) = m_CU89
                        SeekCu90(2) = m_CU90
                        SeekCu103(2) = m_CU103
                        SeekCu112(2) = m_CU112
                        'Add By Sindy 2012/2/7
                        SeekCu39(2) = m_CU39
                        SeekCu40(2) = m_CU40
                        SeekCu41(2) = m_CU41
                        '2012/2/7 End
                        'Add By Sindy 2012/10/31
                        SeekCu10(2) = m_CU10
                        '2012/10/31 End
                  End If
                  End If
                  If textTM79.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                  'Modified by Lydia 2024/07/03 改傳入變數;
                  'GetCu103ByCustomer Me, textTM79.Text
                  Call Pub_GetDataFrm020102(textTM79.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                  
                  'Modify By Sindy 2012/2/7
                  'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Or (m_CU39 & m_CU40 & m_CU41) = "" Then
                        'Modified by Lydia 2024/07/03
                        'Set frm020102_22.oNextForm = Me
                        Call frm020102_22.SetParent(Me, Me.textTM79.Text)
                        frm020102_22.Label4.Caption = textTM79.Text & " " & textTM79_2 'Add By Sindy 2014/7/30
                        frm020102_22.Show vbModal
                        SeekCu05(3) = m_CU05
                        SeekCu88(3) = m_CU88
                        SeekCu89(3) = m_CU89
                        SeekCu90(3) = m_CU90
                        SeekCu103(3) = m_CU103
                        SeekCu112(3) = m_CU112
                        'Add By Sindy 2012/2/7
                        SeekCu39(3) = m_CU39
                        SeekCu40(3) = m_CU40
                        SeekCu41(3) = m_CU41
                        '2012/2/7 End
                        'Add By Sindy 2012/10/31
                        SeekCu10(3) = m_CU10
                        '2012/10/31 End
                  End If
                  End If
                  If textTM80.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                  'Modified by Lydia 2024/07/03 改傳入變數;
                  'GetCu103ByCustomer Me, textTM80.Text
                  Call Pub_GetDataFrm020102(textTM80.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                  
                  'Modify By Sindy 2012/2/7
                  'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Or (m_CU39 & m_CU40 & m_CU41) = "" Then
                        'Modified by Lydia 2024/07/03
                        'Set frm020102_22.oNextForm = Me
                        Call frm020102_22.SetParent(Me, Me.textTM80.Text)
                        frm020102_22.Label4.Caption = textTM80.Text & " " & textTM80_2 'Add By Sindy 2014/7/30
                        frm020102_22.Show vbModal
                        SeekCu05(4) = m_CU05
                        SeekCu88(4) = m_CU88
                        SeekCu89(4) = m_CU89
                        SeekCu90(4) = m_CU90
                        SeekCu103(4) = m_CU103
                        SeekCu112(4) = m_CU112
                        'Add By Sindy 2012/2/7
                        SeekCu39(4) = m_CU39
                        SeekCu40(4) = m_CU40
                        SeekCu41(4) = m_CU41
                        '2012/2/7 End
                        'Add By Sindy 2012/10/31
                        SeekCu10(4) = m_CU10
                        '2012/10/31 End
                  End If
                  End If
                  If textTM81.Text <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
                  'Modified by Lydia 2024/07/03 改傳入變數;
                  'GetCu103ByCustomer Me, textTM81.Text
                  Call Pub_GetDataFrm020102(textTM81.Text, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
                  
                  'Modify By Sindy 2012/2/7
                  'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Or (m_CU39 & m_CU40 & m_CU41) = "" Then
                        'Modified by Lydia 2024/07/03
                        'Set frm020102_22.oNextForm = Me
                        Call frm020102_22.SetParent(Me, Me.textTM81.Text)
                        frm020102_22.Label4.Caption = textTM81.Text & " " & textTM81_2 'Add By Sindy 2014/7/30
                        frm020102_22.Show vbModal
                        SeekCu05(5) = m_CU05
                        SeekCu88(5) = m_CU88
                        SeekCu89(5) = m_CU89
                        SeekCu90(5) = m_CU90
                        SeekCu103(5) = m_CU103
                        SeekCu112(5) = m_CU112
                        'Add By Sindy 2012/2/7
                        SeekCu39(5) = m_CU39
                        SeekCu40(5) = m_CU40
                        SeekCu41(5) = m_CU41
                        '2012/2/7 End
                        'Add By Sindy 2012/10/31
                        SeekCu10(5) = m_CU10
                        '2012/10/31 End
                  End If
                  End If
            End If
            
'            'add by nickc 2006/06/01 加入馬德里續展子案項目
'            strCountry = ""
'            strLicenceCountry = "" '預設都先不勾
'            If m_TM01 = "TF" And m_CP10 = "102" Then
'               CheckOC3
'               '2006/10/16 MODIFY BY SONIA 母案延展抓出所有未閉卷子案,含領土延伸之子案
'               '                           領土延伸延展只抓出該領土延伸之未閉卷子案
'               'strSQL = "select * from trademark where tm01='" & m_TM01 & "' and substr(tm02,1,5)=substr('" & m_TM02 & "',1,5) and tm03<>'0' and (tm16='1' or tm16 is null or tm16='') order by tm04 "
'               If Mid(m_TM02, 6, 1) = "0" Then
'                  strSql = "select * from trademark where tm01='" & m_TM01 & "' and substr(tm02,1,5)=substr('" & m_TM02 & "',1,5) and tm03<>'0' and (tm16='1' or tm16 is null or tm16='') AND TM29 IS NULL order by tm04 "
'               Else
'                  strSql = "select * from trademark where tm01='" & m_TM01 & "' and TM02='" & m_TM02 & "' and tm03<>'0' and (tm16='1' or tm16 is null or tm16='') AND TM29 IS NULL order by tm04 "
'               End If
'               '2006/10/16 END
'               AdoRecordSet3.CursorLocation = adUseClient
'               AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'               If AdoRecordSet3.RecordCount <> 0 Then
'                  AdoRecordSet3.MoveFirst
'                  Do While Not AdoRecordSet3.EOF
'                      strCountry = strCountry & "" & AdoRecordSet3.Fields("TM10") & ","
'                      AdoRecordSet3.MoveNext
'                  Loop
'                  Do While strLicenceCountry = ""
'                  frm880008.strCountry = strCountry
'                  frm880008.strLicenceCountry = strLicenceCountry
'                  frm880008.Caption = "馬德里子案續展選擇"
'                  frm880008.IsByTM = True
'                  frm880008.Show vbModal
'                  strCountry = frm880008.strCountry
'                  strLicenceCountry = frm880008.strLicenceCountry
'                  If frm880008.IsByTM = False Then
'                      Unload frm880008
'                      Exit Sub
'                  Else
'                      Unload frm880008
'                  End If
'                  Dim yoy As Variant
'                  Dim ioi As Integer
'                  yoy = Split(strLicenceCountry, ",")
'                  strLicenceCountry = ""
'                  For ioi = 0 To UBound(yoy)
'                      If Trim(yoy(ioi)) <> "" Then
'                          If strLicenceCountry <> "" Then
'                              strLicenceCountry = strLicenceCountry & ","
'                          End If
'                          strLicenceCountry = strLicenceCountry & "'" & yoy(ioi) & "'"
'                      End If
'                  Next ioi
'                  If strLicenceCountry = "" Then MsgBox "最少要勾選一個", vbInformation, "操作錯誤！"
'                  Loop
'              'add by nickc 2006/08/29 若沒有子案可以續展，請程序補輸
'               Else
'                  MsgBox "查無任何子案可以續展，請補輸後再繼續！", vbExclamation
'                  Exit Sub
'               End If
'               CheckOC3
'            End If
            'Add By Sindy 2012/6/21
            If m_TM01 = "TF" And m_CP10 = "102" Then
               If strLicenceCountry = "" Then
                  MsgBox "指定國家最少要勾選一個！", vbInformation, "操作錯誤！"
                  SSTab1.Tab = 0
                  Exit Sub
               End If
            End If
            '2012/6/21 End
            
            'Add by Sindy 98/3/24
            If m_TM10 = "000" Then
               m_CP09s = m_CP09
               'Modify By Sindy 2009/04/23
               '分割子案的CP123存Null(未經發文室)
      '         strCP31 = ""
      '         If m_CP10 = "308" Then
      '            StrSQLa = "Select CP31 From CaseProgress Where CP09='" & m_CP09 & "' "
      '            rsA.CursorLocation = adUseClient
      '            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      '            If rsA.RecordCount > 0 Then
      '               If IsNull(rsA.Fields("CP31")) = False Then
      '                  strCP31 = Trim(rsA.Fields("CP31"))
      '               End If
      '            End If
      '            If rsA.State <> adStateClosed Then rsA.Close
      '            Set rsA = Nothing
      '         End If
      '         If strCP31 = "Y" Then '為子案
      '            '取得母案
      '            strExc(0) = "SELECT * FROM DivisionCase WHERE DC01='" & m_TM01 & "' AND DC02='" & m_TM02 & "' AND DC03='" & m_TM03 & "' AND DC04='" & m_TM04 & "' "
      '            intI = 1
      '            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      '            If intI = 1 Then
      '               m_DC01 = "" & Trim(RsTemp("DC05"))
      '               m_DC02 = "" & Trim(RsTemp("DC06"))
      '               m_DC03 = "" & Trim(RsTemp("DC07"))
      '               m_DC04 = "" & Trim(RsTemp("DC08"))
      '            End If
      '            '取得母案之審定號及目前准駁
      '            strExc(0) = "SELECT TM15,TM16 FROM TradeMark WHERE TM01='" & m_DC01 & "' AND TM02='" & m_DC02 & "' AND TM03='" & m_DC03 & "' AND TM04='" & m_DC04 & "' "
      '            intI = 1
      '            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      '            If intI = 1 Then
      '               m_DcTM15 = "" & Trim(RsTemp("TM15"))
      '               m_DcTM16 = "" & Trim(RsTemp("TM16"))
      '            End If
      '            m_CP09s = m_CP09
      '            '註冊後分割
      '            If m_DcTM15 <> "" And m_DcTM16 = "1" Then
      '               m_CP123s = ""
      '            '申請中分割
      '            Else
      '               'Add By Sindy 2009/06/04
      '               '取得主管機關名稱
      '               strExc(0) = "SELECT * FROM CaseFee WHERE CF01='" & m_TM01 & "' AND CF02='000' AND CF03='" & m_CP10 & "' "
      '               intI = 1
      '               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      '               If intI = 1 Then
      '                  m_CP130s = "" & RsTemp("CF10")
      '               End If
      '               '2009/06/04 End
      '               m_CP123s = "N"
      '            End If
      '         Else
      '            'Add by Sindy 2009/4/24
      '            If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
      '                Exit Sub
      '   '         Else
      '   '            m_CP123s = GetCPMSendYn(m_TM01, m_CP10, 1)
      '            End If
      '         End If
      '         '2009/04/23 End
               'Modify By Sindy 2009/10/21
               '1.改用分割案件關係檔判斷是否為子案
               '2.子案均不經發文室
               '3.母案分割發文：提醒操作者有幾件子案，顯示子案案號
               '4.子案分割發文：提醒還有幾件子案尚未發文，顯示未發文案號
               bDiviCSon = False
               If m_CP10 = "308" Then
                  'Modify By Sindy 2011/10/20 因FCT-029340為分割再分割,所以不能先判斷是否為子案,應先判斷是否為母案
                  'StrSQLa = "SELECT * FROM DivisionCase WHERE DC01='" & m_TM01 & "' AND DC02='" & m_TM02 & "' AND DC03='" & m_TM03 & "' AND DC04='" & m_TM04 & "' "
                  StrSQLa = "SELECT * FROM DivisionCase WHERE DC05='" & m_TM01 & "' AND DC06='" & m_TM02 & "' AND DC07='" & m_TM03 & "' AND DC08='" & m_TM04 & "' "
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  'If rsA.RecordCount > 0 Then
                  If rsA.RecordCount <= 0 Then '為子案
                     '取得母案號
                     strSql = "SELECT * FROM DivisionCase WHERE DC01='" & m_TM01 & "' AND DC02='" & m_TM02 & "' AND DC03='" & m_TM03 & "' AND DC04='" & m_TM04 & "' "
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        m_DC05 = "" & Trim(RsTemp.Fields("DC05"))
                        m_DC06 = "" & Trim(RsTemp.Fields("DC06"))
                        m_DC07 = "" & Trim(RsTemp.Fields("DC07"))
                        m_DC08 = "" & Trim(RsTemp.Fields("DC08"))
                     End If
                     bDiviCSon = True
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
               End If
               If bDiviCSon = True Then '為子案
                  m_CP123s = ""
                  '取得尚未發文的分割資訊
                  strExc(0) = "SELECT CP01,CP02,CP03,CP04,CP27 FROM DivisionCase,CaseProgress" & _
                                    " WHERE DC05='" & m_DC05 & "' AND DC06='" & m_DC06 & "' AND DC07='" & m_DC07 & "' AND DC08='" & m_DC08 & "'" & _
                                    " AND DC01=CP01(+) AND DC02=CP02(+) AND DC03=CP03(+) AND DC04=CP04(+)" & _
                                    " AND CP10='308'" & _
                                    " Union All" & _
                                    " SELECT CP01,CP02,CP03,CP04,CP27 FROM CaseProgress" & _
                                    " WHERE CP01='" & m_DC05 & "' AND CP02='" & m_DC06 & "' AND CP03='" & m_DC07 & "' AND CP04='" & m_DC08 & "'" & _
                                    " AND CP10='308'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  intTemp = 0: strTemp = ""
                  If intI = 1 Then
                     With RsTemp
                        RsTemp.MoveFirst
                        Do While Not RsTemp.EOF
                           If (IsNull(RsTemp("CP27")) Or Val("" & RsTemp("CP27")) <= 0) And _
                              (RsTemp("CP01") = m_TM01 And _
                               RsTemp("CP02") = m_TM02 And _
                               RsTemp("CP03") = m_TM03 And _
                               RsTemp("CP04") = m_TM04) = False Then
                              '未發文
                              intTemp = intTemp + 1
                              If strTemp <> "" Then strTemp = strTemp & "及"
                              strTemp = strTemp & Trim(RsTemp("CP01")) & "-" & Trim(RsTemp("CP02")) & "-" & Trim(RsTemp("CP03")) & "-" & Trim(RsTemp("CP04"))
                           End If
                           RsTemp.MoveNext
                        Loop
                     End With
                     If intTemp > 0 Then
                        MsgBox "尚有" & intTemp & "件未發文，案號為" & strTemp, vbInformation
                     End If
                  End If
               Else '為母案
                  '取得子案資訊
                  strExc(0) = "SELECT * FROM DivisionCase WHERE DC05='" & m_TM01 & "' AND DC06='" & m_TM02 & "' AND DC07='" & m_TM03 & "' AND DC08='" & m_TM04 & "' "
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  intTemp = 0: strTemp = ""
                  If intI = 1 Then
                     intTemp = RsTemp.RecordCount
                     With RsTemp
                        RsTemp.MoveFirst
                        Do While Not RsTemp.EOF
                           If strTemp <> "" Then strTemp = strTemp & "及"
                           strTemp = strTemp & Trim(RsTemp("DC01")) & "-" & Trim(RsTemp("DC02")) & "-" & Trim(RsTemp("DC03")) & "-" & Trim(RsTemp("DC04"))
                           RsTemp.MoveNext
                        Loop
                     End With
                     If intTemp > 0 Then
                        MsgBox "子案有" & intTemp & "件，案號為" & strTemp, vbInformation
                     End If
                  End If
                  '2009/10/21 End
                  
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
                     'Add by Sindy 2009/4/24
                     If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
                         Exit Sub
            '         Else
            '            m_CP123s = GetCPMSendYn(m_TM01, m_CP10, 1)
                     End If
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
            If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !" & Err.Description, vbCritical: Screen.MousePointer = vbDefault: Exit Sub
            'Add By Cheng 2002/11/08
            ' 列印定稿
            If textPrint <> "N" Then
               PrintLetter
            'Add By Sindy 2021/2/25
            End If
            If textPrint = "N" Then
                If strLD18 <> "" Then
                   Call PUB_TCaseAskIsPost(strLD18)
                End If
            '2021/2/25 END
            End If
            
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
            
            ' 設定滑鼠游標為預設
            Screen.MousePointer = vbDefault
            
            'Add By Sindy 2011/1/26 檢查相同國家若有舊案申請地址與客戶目前申請地址不同者
            strApplID = ""
            If Trim(textTM23) <> "" Then '申請人1
               If strApplID <> "" Then strApplID = strApplID & ","
               strApplID = strApplID & "'" & Trim(textTM23) & "'"
            End If
            If Trim(textTM78) <> "" Then '申請人2
               If strApplID <> "" Then strApplID = strApplID & ","
               strApplID = strApplID & "'" & Trim(textTM78) & "'"
            End If
            If Trim(textTM79) <> "" Then '申請人3
               If strApplID <> "" Then strApplID = strApplID & ","
               strApplID = strApplID & "'" & Trim(textTM79) & "'"
            End If
            If Trim(textTM80) <> "" Then '申請人4
               If strApplID <> "" Then strApplID = strApplID & ","
               strApplID = strApplID & "'" & Trim(textTM80) & "'"
            End If
            If Trim(textTM81) <> "" Then '申請人5
               If strApplID <> "" Then strApplID = strApplID & ","
               strApplID = strApplID & "'" & Trim(textTM81) & "'"
            End If
            If ChkOCaseAndCAddrNotAlike(strApplID, m_TM10, m_TM01, m_CP10, rsAddrNotAlike, False) = True Then
               Set frm880018.fmParent = Me
               Set frm880018.RsTemp = rsAddrNotAlike
               frm880018.m_Appl1 = Trim(Me.textTM23.Text)
               frm880018.m_Appl2 = Trim(Me.textTM78.Text)
               frm880018.m_Appl3 = Trim(Me.textTM79.Text)
               frm880018.m_Appl4 = Trim(Me.textTM80.Text)
               frm880018.m_Appl5 = Trim(Me.textTM81.Text)
               frm880018.Show vbModal
            End If
            '2011/1/26 End
            
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
            
            If Index = 0 Then '確定鍵
               '********* 901123 nick   清畫面
               'frm020102_01.radio(0).Value = True
               'frm020102_01.textCP09.Enabled = True
               'frm020102_01.textCP09.Text = ""
               'frm020102_01.textTM01.Enabled = False
               'frm020102_01.textTM01.Text = "" modify by sonia
               'frm020102_01.textTM02.Enabled = False
               'frm020102_01.textTM02.Text = ""
               'frm020102_01.textTM02_2.Enabled = False
               'frm020102_01.textTM02_2.Text = ""
               'frm020102_01.textTM03.Enabled = False
               'frm020102_01.textTM03.Text = ""
               'frm020102_01.textTM04.Enabled = False
               'frm020102_01.textTM04.Text = ""
               'frm020102_01.grdList.Clear
               'frm020102_01.grdList.Rows = 2
               '*********************************
               
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
               
               'frm020102_01.RefreshData
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

Private Sub cmdPriority_Click()
   ' 修改優先權資料
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify by Sindy 2017/10/12 + , m_Priority(6)
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
'
'      ' 設定滑鼠游標為等待狀態
'      Screen.MousePointer = vbHourglass
'      ' 更新欄位輸入的內容
'      OnUpdateField
'      ' 存檔
'        'Modify By Cheng 2002/11/11
''      'OnSaveData
'      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
'        ' 列印定稿
'        If textPrint <> "N" Then
'           PrintLetter
'        End If
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

'add by sonia 2018/9/27 原已取消,因中間接進來延展案要預設延展後專用期間T-217280
Private Sub Form_Activate()
   'Add By Cheng 2003/10/06
   '若有按下變更事項按鈕 , 則重新讀取資料
   'If m_blnClkChgButton = True Then
   'add by sonia 2018/9/26 預設延展後專用期限,中間接進來延展案補完基本檔再重讀原專用期間
   If Val(textTM21) = 0 And m_CP10 = "102" And m_TM01 <> "TM" Then
      QueryTradeMark
      textTM21 = Val(Get102TM21TM22("TM21"))
      textTM22 = Val(Get102TM21TM22("TM22"))
      If m_TM10 = 台灣國家代號 Then
         textTM21 = ChangeWStringToTString(textTM21)
         textTM22 = ChangeWStringToTString(textTM22)
      End If
      'end 2018/9/26
      'add by sonia 2018/11/27 T-217498 重新預設否則會被上面QueryTradeMark又蓋掉
      If Trim(textPrint) = "" Then
           textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
      End If
      'end 2018/11/27
   End If
   
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08_2.BackColor = &H8000000F
   textTM12_S.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23_2.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   textTM72_2.BackColor = &H8000000F
   
   'add by nickc 2007/02/15
   textTM78_2.BackColor = &H8000000F
   textTM79_2.BackColor = &H8000000F
   textTM80_2.BackColor = &H8000000F
   textTM81_2.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textTM44.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   
   ' 91.09.02 modify by louis
   ' 初始化GridList
   InitialGrdList
   MoveFormToCenter Me
   textMail = "Y"
'    m_blnClkChgButton = False
   'Add by nickc 2006/01/27
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Add by Amy 2021/12/23 一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 700
   lstNameAgent.Width = 1300
    
   'add by nickc 2006/06/20
   IsHaveGoods = False
   
   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
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
      ' 91.09.02 marked by louis
      ' 查名總收文號
      'Case 99: textCP09S = strData
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
Private Sub ClearNewCPFieldList()
   If m_NewCPListCount > 0 Then
      Erase m_NewCPList
   End If
   m_NewCPListCount = 0
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetNewCPFieldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_NewCPListCount - 1
      If m_NewCPList(nPos).fiName = strFieldName Then
         bFind = True
         m_NewCPList(nPos).fiOldData = strFieldData
         m_NewCPList(nPos).fiNewData = strFieldData
         m_NewCPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_NewCPList(m_NewCPListCount + 1)
      m_NewCPList(m_NewCPListCount).fiName = strFieldName
      m_NewCPList(m_NewCPListCount).fiOldData = strFieldData
      m_NewCPList(m_NewCPListCount).fiNewData = strFieldData
      m_NewCPList(m_NewCPListCount).fiType = nFieldType
      m_NewCPListCount = m_NewCPListCount + 1
   End If
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
   Dim strTemp As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
      
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
      End If
      ' 申請日
      strTemp = Empty
      If IsNull(rsTmp.Fields("TM11")) = False Then
         strTemp = rsTmp.Fields("TM11")
         textTM11 = TAIWANDATE(rsTmp.Fields("TM11"))
      Else
      '92.2.5 MODIFY by sonia
      '   textTM11 = TAIWANDATE(SystemDate())
         '93.9.30 modify by sonia
         'If rsTmp.Fields("TM10") = "000" And m_CP10 = "101" Then
         '2008/10/21 MODIFY BY SONIA 取消分割案
         'If rsTmp.Fields("TM10") = "000" And (m_CP10 = "101" Or m_CP10 = "308") Then
         If rsTmp.Fields("TM10") = "000" And m_CP10 = "101" Then
         '93.9.30
            textTM11 = TAIWANDATE(SystemDate())
         End If
      '92.2.5 end
      End If
      '2008/10/24 add by sonia 分割子案申請日預設母案申請日
      StrSQLa = "Select * From DivisionCase,TradeMark Where DC01='" & m_TM01 & "' And DC02='" & m_TM02 & "' And DC03='" & m_TM03 & "' And DC04='" & m_TM04 & "' and DC05=TM01(+) and DC06=TM02(+) and DC07=TM03(+) and DC08=TM04(+) "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         If IsNull(rsA.Fields("TM11")) = False Then
            textTM11 = TAIWANDATE(rsA.Fields("TM11"))
         End If
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      '2008/10/24 end
      SetTMSPFieldOldData "TM11", strTemp, 1
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12_S = rsTmp.Fields("TM12")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
      End If
'      ' 案件中文名稱
'      textTM05 = Empty
'      If IsNull(rsTmp.Fields("TM05")) = False Then
'         textTM05 = rsTmp.Fields("TM05")
'      End If
'      SetTMSPFieldOldData "TM05", textTM05, 0
      ' 案件中文名稱
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
      m_TM09 = "" 'Add By Sindy 2014/2/20
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
         m_TM09 = rsTmp.Fields("TM09") 'Add By Sindy 2014/2/20
      End If
      SetTMSPFieldOldData "TM09", textTM09, 0
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         m_NA14 = GetNationExtentYear(m_TM10)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      SetTMSPFieldOldData "TM12", textTM12, 0
      ' 發證日  add by sonia 2023/11/17
      If IsNull(rsTmp.Fields("TM20")) = False Then
         m_TM20 = TAIWANDATE(rsTmp.Fields("TM20"))
      End If
      'end 2023/11/17
      ' 專用期限起日
      If IsNull(rsTmp.Fields("TM21")) = False Then
         m_TM21 = TAIWANDATE(rsTmp.Fields("TM21"))
      End If
      ' 專用期限止日
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_TM22 = TAIWANDATE(rsTmp.Fields("TM22"))
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = ChangeCustomerL(rsTmp.Fields("TM23"))
         textTM23 = ChangeCustomerL(rsTmp.Fields("TM23"))
         ' 90.07.19 modify (帶出申請人名稱)
         If IsEmptyText(textTM23) = False Then
            'Modify By Cheng 2002/09/23
'            textTM23_Validate False
            textTM23_2 = GetCustomerName(textTM23, 0)
         End If
      End If
      SetTMSPFieldOldData "TM23", textTM23, 0
      'Add By Cheng 2002/08/23
      m_strCust1 = "" & Me.textTM23.Text
      Me.textTM23.Tag = Me.textTM23.Text 'Add by Amy 2018/10/30
      
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
      'Add By Cheng 2002/06/14
      '取得商標種類代號
      m_TM08 = "" & rsTmp.Fields("TM08").Value
        'Add By Cheng 2002/11/18
        m_FA10 = IIf("" & rsTmp.Fields("TM44").Value = "", "", GetFagentNation("" & rsTmp.Fields("TM44").Value))
      ' 特殊商標
      textTM72 = Empty: Me.textTM72_2.Text = ""
      If IsNull(rsTmp.Fields("TM72")) = False Then
         textTM72 = "" & rsTmp.Fields("TM72")
         textTM72_2 = PUB_GetSpecialPTName("2", "" & rsTmp.Fields("TM72").Value)
      End If
      SetTMSPFieldOldData "TM72", textTM72, 0
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("tm77"))
      m_textPrint = textPrint
      SetTMSPFieldOldData "TM77", textPrint, 0
      'Added by Lydia 2023/11/16 內外商之分案及商標基本資料維護之商標種類、特殊商標欄位增加下拉功能
      Pub_SetTMcombo "1", cboTM08, textTM08, IIf(m_TM10 <> "000", True, False), strPTM '商標種類
      Pub_SetTMcombo "2", cboTM72, textTM72, IIf(m_TM10 <> "000", True, False), strSPT '特殊商標種類
      'end 2023/11/16
      
      'add by nickc 2007/01/02
      ' 申請人2
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = ChangeCustomerL(rsTmp.Fields("TM78"))
         textTM78 = ChangeCustomerL(rsTmp.Fields("TM78"))
         ' (帶出申請人名稱)
         If IsEmptyText(textTM78) = False Then
            textTM78_2 = GetCustomerName(textTM78, 0)
         End If
      End If
      SetTMSPFieldOldData "TM78", textTM78, 0
      m_strCust2 = "" & Me.textTM78.Text
      Me.textTM78.Tag = Me.textTM78.Text 'Add by Amy 2018/10/30
      ' 中文地址
      m_TM82 = ""
      If IsNull(rsTmp.Fields("TM82")) = False Then
         m_TM82 = rsTmp.Fields("TM82")
      End If
      SetTMSPFieldOldData "TM82", m_TM82, 0
      ' 英文地址
      m_TM86 = ""
      If IsNull(rsTmp.Fields("TM86")) = False Then
         m_TM86 = rsTmp.Fields("TM86")
      End If
      SetTMSPFieldOldData "TM86", m_TM86, 0
      ' 日文地址
      m_TM90 = ""
      If IsNull(rsTmp.Fields("TM90")) = False Then
         m_TM90 = rsTmp.Fields("TM90")
      End If
      SetTMSPFieldOldData "TM90", m_TM90, 0
      ' 申請人3
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = ChangeCustomerL(rsTmp.Fields("TM79"))
         textTM79 = ChangeCustomerL(rsTmp.Fields("TM79"))
         ' (帶出申請人名稱)
         If IsEmptyText(textTM79) = False Then
            textTM79_2 = GetCustomerName(textTM79, 0)
         End If
      End If
      SetTMSPFieldOldData "TM79", textTM79, 0
      m_strCust3 = "" & Me.textTM79.Text
      Me.textTM79.Tag = Me.textTM79.Text 'Add by Amy 2018/10/30
      ' 中文地址
      m_TM83 = ""
      If IsNull(rsTmp.Fields("TM83")) = False Then
         m_TM83 = rsTmp.Fields("TM83")
      End If
      SetTMSPFieldOldData "TM83", m_TM83, 0
      ' 英文地址
      m_TM87 = ""
      If IsNull(rsTmp.Fields("TM87")) = False Then
         m_TM87 = rsTmp.Fields("TM87")
      End If
      SetTMSPFieldOldData "TM87", m_TM87, 0
      ' 日文地址
      m_TM91 = ""
      If IsNull(rsTmp.Fields("TM91")) = False Then
         m_TM91 = rsTmp.Fields("TM91")
      End If
      SetTMSPFieldOldData "TM91", m_TM91, 0
      ' 申請人4
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = ChangeCustomerL(rsTmp.Fields("TM80"))
         textTM80 = ChangeCustomerL(rsTmp.Fields("TM80"))
         ' (帶出申請人名稱)
         If IsEmptyText(textTM80) = False Then
            textTM80_2 = GetCustomerName(textTM80, 0)
         End If
      End If
      SetTMSPFieldOldData "TM80", textTM80, 0
      m_strCust4 = "" & Me.textTM80.Text
      Me.textTM80.Tag = Me.textTM80.Text 'Add by Amy 2018/10/30
      ' 中文地址
      m_TM84 = ""
      If IsNull(rsTmp.Fields("TM84")) = False Then
         m_TM84 = rsTmp.Fields("TM84")
      End If
      SetTMSPFieldOldData "TM84", m_TM84, 0
      ' 英文地址
      m_TM88 = ""
      If IsNull(rsTmp.Fields("TM88")) = False Then
         m_TM88 = rsTmp.Fields("TM88")
      End If
      SetTMSPFieldOldData "TM88", m_TM88, 0
      ' 日文地址
      m_TM92 = ""
      If IsNull(rsTmp.Fields("TM92")) = False Then
         m_TM92 = rsTmp.Fields("TM92")
      End If
      SetTMSPFieldOldData "TM92", m_TM92, 0
      ' 申請人5
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = ChangeCustomerL(rsTmp.Fields("TM81"))
         textTM81 = ChangeCustomerL(rsTmp.Fields("TM81"))
         ' (帶出申請人名稱)
         If IsEmptyText(textTM81) = False Then
            textTM81_2 = GetCustomerName(textTM81, 0)
         End If
      End If
      SetTMSPFieldOldData "TM81", textTM81, 0
      m_strCust5 = "" & Me.textTM81.Text
      Me.textTM81.Tag = Me.textTM81.Text 'Add by Amy 2018/10/30
      ' 中文地址
      m_TM85 = ""
      If IsNull(rsTmp.Fields("TM85")) = False Then
         m_TM85 = rsTmp.Fields("TM85")
      End If
      SetTMSPFieldOldData "TM85", m_TM85, 0
      ' 英文地址
      m_TM89 = ""
      If IsNull(rsTmp.Fields("TM89")) = False Then
         m_TM89 = rsTmp.Fields("TM89")
      End If
      SetTMSPFieldOldData "TM89", m_TM89, 0
      ' 日文地址
      m_TM93 = ""
      If IsNull(rsTmp.Fields("TM93")) = False Then
         m_TM93 = rsTmp.Fields("TM93")
      End If
      SetTMSPFieldOldData "TM93", m_TM93, 0
      
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得服務業務基本檔的欄位內容
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   
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
      Select Case m_TM01
      Case "TS"
          textTM05_1 = Empty
          textTM05_1 = "" & rsTmp.Fields("SP05")
          SetTMSPFieldOldData "SP05", textTM05_1, 0
      Case Else
          textTM05 = Empty
      '      If IsNull(rsTmp.Fields("SP05")) = False Then
          textTM05 = "" & rsTmp.Fields("SP05")
      '      End If
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
      ' 申請日
      strTemp = Empty
      If IsNull(rsTmp.Fields("SP10")) = False Then
         strTemp = rsTmp.Fields("SP10")
         textTM11 = TAIWANDATE(rsTmp.Fields("SP10"))
      '92.2.5 cancel by sonia
      'Else
      '   textTM11 = TAIWANDATE(SystemDate())
      '92.2.5 end
      End If
      SetTMSPFieldOldData "SP10", strTemp, 1
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = ChangeCustomerL(rsTmp.Fields("SP08"))
         textTM23 = ChangeCustomerL(rsTmp.Fields("SP08"))
         ' 90.07.19 modify (帶出申請人名稱)
         If IsEmptyText(textTM23) = False Then
            'Modify By Cheng 2002/09/23
'            textTM23_Validate False
            textTM23_2 = GetCustomerName(textTM23, 0)
         End If
      End If
      SetTMSPFieldOldData "SP08", textTM23, 0
      'Add By Cheng 2002/08/23
      m_strCust1 = "" & Me.textTM23.Text
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
         m_NA14 = GetNationExtentYear(m_TM10)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12_S = rsTmp.Fields("SP11")
         textTM12 = rsTmp.Fields("SP11")
      End If
      SetTMSPFieldOldData "SP11", textTM12, 0
      ' 專用期限起日
      If IsNull(rsTmp.Fields("SP20")) = False Then
         m_TM21 = TAIWANDATE(rsTmp.Fields("SP20"))
      End If
      ' 專用期限止日
      If IsNull(rsTmp.Fields("SP21")) = False Then
         m_TM22 = TAIWANDATE(rsTmp.Fields("SP21"))
      End If
      ' 案件備註
      textTM58 = Empty
      If IsNull(rsTmp.Fields("SP18")) = False Then
         textTM58 = rsTmp.Fields("SP18")
      End If
      SetTMSPFieldOldData "SP18", textTM58, 0
        'Add By Cheng 2002/11/18
        m_FA10 = IIf("" & rsTmp.Fields("SP26").Value = "", "", GetFagentNation("" & rsTmp.Fields("SP26").Value))
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("sp72"))
      m_textPrint = textPrint
      SetTMSPFieldOldData "SP72", textPrint, 0
      'add by nickc 2007/01/02
      ' 申請人2
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("SP58")) = False Then
         m_TM78 = ChangeCustomerL(rsTmp.Fields("SP58"))
         textTM78 = ChangeCustomerL(rsTmp.Fields("SP58"))
         ' (帶出申請人名稱)
         If IsEmptyText(textTM78) = False Then
            textTM78_2 = GetCustomerName(textTM78, 0)
         End If
      End If
      SetTMSPFieldOldData "SP58", textTM78, 0
      m_strCust2 = "" & Me.textTM78.Text
      ' 申請人3
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("SP59")) = False Then
         m_TM79 = ChangeCustomerL(rsTmp.Fields("SP59"))
         textTM79 = ChangeCustomerL(rsTmp.Fields("SP59"))
         ' (帶出申請人名稱)
         If IsEmptyText(textTM79) = False Then
            textTM79_2 = GetCustomerName(textTM79, 0)
         End If
      End If
      SetTMSPFieldOldData "SP59", textTM79, 0
      m_strCust3 = "" & Me.textTM79.Text
      ' 申請人4
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("SP65")) = False Then
         m_TM80 = ChangeCustomerL(rsTmp.Fields("SP65"))
         textTM80 = ChangeCustomerL(rsTmp.Fields("SP65"))
         ' (帶出申請人名稱)
         If IsEmptyText(textTM80) = False Then
            textTM80_2 = GetCustomerName(textTM80, 0)
         End If
      End If
      SetTMSPFieldOldData "SP65", textTM80, 0
      m_strCust4 = "" & Me.textTM80.Text
      ' 申請人5
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("SP66")) = False Then
         m_TM81 = ChangeCustomerL(rsTmp.Fields("SP66"))
         textTM81 = ChangeCustomerL(rsTmp.Fields("SP66"))
         ' (帶出申請人名稱)
         If IsEmptyText(textTM81) = False Then
            textTM81_2 = GetCustomerName(textTM81, 0)
         End If
      End If
      SetTMSPFieldOldData "SP66", textTM81, 0
      m_strCust5 = "" & Me.textTM81.Text
      textTM09 = Empty
      If IsNull(rsTmp.Fields("SP73")) = False Then
         textTM09 = rsTmp.Fields("SP73")
      End If
      SetTMSPFieldOldData "SP73", textTM09, 0
      textTM32 = Empty
      If IsNull(rsTmp.Fields("SP74")) = False Then
         textTM32 = rsTmp.Fields("SP74")
      End If
      SetTMSPFieldOldData "SP74", textTM32, 0
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
Dim strTemp As String
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
      '2007/8/7 ADD BY SONIA 法定期限
      m_CP07 = ""
      If IsNull(rsTmp.Fields("CP07")) = False Then
         m_CP07 = rsTmp.Fields("CP07")
      End If
      '2007/8/7 END
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
        '取得承辦人代號
        m_CP14 = "" & rsTmp.Fields("CP14").Value
        
      ' 發文日(預設為系統日)
      textCP27 = TAIWANDATE(SystemDate())
      strCP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         strCP27 = DBDATE(rsTmp.Fields("CP27"))
      End If
      SetCPFieldOldData "CP27", strCP27, 1
      
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = ChangeCustomerL(rsTmp.Fields("CP44"))
      End If
      SetCPFieldOldData "CP44", textCP44, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      
      'Add By Sindy 2011/7/12
      m_CP31 = Empty
      If IsNull(rsTmp.Fields("CP31")) = False Then
         m_CP31 = rsTmp.Fields("CP31")
      End If
      
      'Add By Sindy 2014/3/10
      m_CP60 = Empty
      If IsNull(rsTmp.Fields("CP60")) = False Then
         m_CP60 = rsTmp.Fields("CP60")
      End If
      
      ' 是否出名
      textCP22 = Empty
      If IsNull(rsTmp.Fields("CP22")) = False Then
         textCP22 = rsTmp.Fields("CP22")
      End If
      SetCPFieldOldData "CP22", textCP22, 0
      ' 是否算案件數
      textCP26 = Empty
      If IsNull(rsTmp.Fields("CP26")) = False Then
         textCP26 = rsTmp.Fields("CP26")
      End If
      SetCPFieldOldData "CP26", textCP26, 0
      
      'Add By Sindy 2011/3/9
      ' 是否電子送件
      textCP118 = Empty
      If IsNull(rsTmp.Fields("CP118")) = False Then
         textCP118 = rsTmp.Fields("CP118")
      End If
      SetCPFieldOldData "CP118", textCP118, 0
      'ADD BY SONIA 2014/11/6 電子送件案預設發文日為承辦人發文日CP85
      If textCP118 = "Y" Then
         textCP27 = TAIWANDATE("" & rsTmp.Fields("CP85"))
      End If
      'END  2014/11/6

      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
    
      ' 彼所案號
      textTM45 = Empty
'      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         textTM45 = rsTmp.Fields("CP45")
'         strCP45 = rsTmp.Fields("CP45")
      End If
'      SetCPFieldOldData "CP45", strCP45, 0
      SetCPFieldOldData "CP45", textTM45, 0
      ' 案件性質為延展時, 更新案件進度檔的授權期間欄位
      If m_CP10 = "102" Then
         ' 授權期間(起)
         strTemp = Empty
         If IsNull(rsTmp.Fields("CP53")) = False Then
            strTemp = rsTmp.Fields("CP53")
         End If
         SetCPFieldOldData "CP53", strTemp, 1
         ' 授權期間(迄)
         strTemp = Empty
         If IsNull(rsTmp.Fields("CP54")) = False Then
            strTemp = rsTmp.Fields("CP54")
         End If
         SetCPFieldOldData "CP54", strTemp, 1
      End If
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
      If m_CP110 = "" And m_CP10 = "102" And m_TM10 = "000" Then m_CP110 = "94007,81040" 'Add By Sindy 2016/8/31 延展(102)時,出名代理人預設為94007.林景郁和81040.閻啟泰
      SetCPFieldOldData "CP110", CheckStr(rsTmp.Fields("cp110")), 0
      '2010/9/20 End
      
      'add by sonia 2017/8/31 台灣分割子案預設母案分割之出名代理人T-210392
      If m_TM10 = "000" And m_CP110 = "" And m_CP10 = "308" Then '申請案且無輸入補優先權文件期限
         CheckOC3
         strSql = "Select cp110 From DivisionCase,CaseProgress Where DC01='" & m_TM01 & "' and DC02='" & m_TM02 & "' and DC03='" & m_TM03 & "' and DC04='" & m_TM04 & "' And DC05=CP01(+) And DC06=CP02(+) And DC07=CP03(+) And DC08=CP04(+) And CP10='308' "
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If AdoRecordSet3.RecordCount > 0 Then
            m_CP110 = CheckStr(AdoRecordSet3.Fields("cp110"))
         End If
         CheckOC3
      End If
      'end 2017/8/31
      
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
               If PUB_GetAgentName(m_TM01, "" & rsSubTmp("CP44"), strTempName) Then
                  strCP44 = strTempName
               Else
                  strCP44 = ""
               End If
               'Modify By Cheng 2002/07/09
'               AddAgent rsSubTmp.Fields("CP44"), GetFAgentName(rsSubTmp.Fields("CP44"))
               AddAgent rsSubTmp.Fields("CP44"), strCP44
            End If
            rsSubTmp.MoveNext
         Loop
      End If
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
    rsSubTmp.Close
      'add by nick 2004/08/12 發文規費
      If IsNull(rsTmp.Fields("CP17")) = False And textCP84.Enabled = True Then
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
   
   m_TM21 = Empty
   m_TM22 = Empty
   
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
      m_CP43 = "" & rsTmp.Fields("CP43") 'Added by Lydia 2024/11/21
   End If
   rsTmp.Close
    'Add By Cheng 2003/11/10
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF", "TS"
        Me.Label9.Visible = False
        Me.Label8.Visible = False
        Me.Label7.Visible = False
        Me.textTM05.Visible = False
        Me.textTM05.Enabled = False
        Me.textTM06.Visible = False
        Me.textTM06.Enabled = False
        Me.textTM07.Visible = False
        Me.textTM07.Enabled = False
    Case Else
        Me.Label38.Visible = False
        Me.textTM05_1.Visible = False
        Me.textTM05_1.Enabled = False
    End Select
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
   
   'Add By Cheng 2002/06/14
   m_TM08 = ""
   'Add By Cheng 2002/07/17
   m_NA14 = ""
    'Add By Cheng 2002/11/18
    m_FA10 = ""
   'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
   m_QSP = False
   
   ' 取得基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "T", "TF", "FCT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
         'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
         m_QSP = True
   End Select
   
   'Add By Sindy 2020/8/12 取得主管機關, 只有主管機關是"經濟部智慧財產局"時, 才做自動扣款
   Call GetCaseFeeByNick(m_TM01, m_TM10, m_CP10, m_strCF10)
   If m_strCF10 <> "經濟部智慧財產局" And txtPayToday.Visible = True And txtPayToday <> "" Then
      txtPayToday = ""
   End If
   '2020/8/12 END
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
      'Added by Lydia 2019/12/12 內商T案申請:查名單近似本所案經核可後，設定" 是否出名 "
      strExc(1) = Pub_ChkTQD11(m_CP09, mTQD11)
      If mTQD11 <> "" Then
            strExc(2) = m_CP110
            If Left(mTQD11, 1) = "2" Then '2=不出名
                  MsgBox "查名結果為近似本所案經核可後，設定為不出名代理！", vbInformation, "是否出名"
                  strExc(2) = "N"
            ElseIf Left(mTQD11, 1) = "1" Then '1=第三人
                  MsgBox "查名結果為近似本所案經核可後，設定為第三人出名！", vbInformation, "是否出名"
            End If
            'Modify by Amy 2021/12/23 改Form2.0,bForm2設True
            PUB_SetOurAgent lstNameAgent, tm(), strExc(2), m_CP10, True '傳入已設定的出名代理人/N=不出名不預設
      Else
       'end 2019/12/12
             '2010/5/6 MODIFY BY SONIA 新申請案預設出名代理人
             'PUB_SetOurAgent lstNameAgent, tm(), m_CP110
             'Modify by Amy 2021/12/23 改Form2.0,bForm2設True
             PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True
            '2010/5/6 END
      End If
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2012/7/26 End
   
   m_CU10 = Empty
   'add by nickc 2007/01/02
   Set rsTmp = New ADODB.Recordset
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open "select * from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM23.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM23.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <> 0 Then
       m_CU10 = CheckStr(rsTmp.Fields("CU10"))
   End If
   rsTmp.Close
   
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
      'add by nickc 2005/08/31 提申期限
      If textPetition.Text = "" And textCP27.Text <> "" And IsNull(rsTmp.Fields("CF11").Value) = False Then
         textPetition.Text = ChangeWStringToTString(CompDate(2, rsTmp.Fields("CF11").Value, ChangeTStringToWString(textCP27)))
         '2007/8/7 ADD BY SONIA 非台灣延展案且未到可辦日(法定期限-延展時間),以延展可辦日+10天為提申期限
         'Modify By Sindy 2009/06/16 增加109.被異議續展
         'If m_TM10 <> "000" And m_CP10 = "102" Then
         If m_TM10 <> "000" And (m_CP10 = "102" Or m_CP10 = "109") Then
            m_NA15 = GetDelayTime(m_TM10)
            If m_NA15 > 0 And CompDate(1, m_NA15 * -1, m_CP07) > ChangeTStringToWString(textCP27.Text) Then
               textPetition.Text = ChangeWStringToTString(CompDate(2, 10, CompDate(1, m_NA15 * -1, m_CP07)))
            End If
         End If
         'Add By Sindy 2019/6/11 檢查期限是否正確
         textPetition.Text = PUB_T997998LimitDate(textPetition.Text, m_CP07, 2)
         '2019/6/11 END
      End If
      '2006/10/23 ADD BY SONIA 催審期限textUargeDate
      textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
      textCP27.Tag = textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
      'Mark by Amy 2015/09/10 避免tm22 null一進來造成error
'      'Add by Amy 2015/05/26 +T大陸續展 催審期限管制
'      If m_CP10 = "102" And m_TM10 = "020" Then
'         Call SetUargeDate_102
'      End If
'      'end 2015/05/26
      '2006/10/23 END
      '2006/11/2 ADD BY SONIA 發證後之分割不掛催審期限
      If m_CP10 = "308" And textTM20 <> "" Then
         textUargeDate = ""
      End If
      '2006/11/2 END
      m_textUargeDate = textUargeDate  '2006/11/13 ADD BY SONIA
   End If
   rsTmp.Close
   
   ' 設定欄位
   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         EnableTextBox textTM08, True
         EnableTextBox textTM09, True
         EnableTextBox textTM27, True
         EnableTextBox textTM32, True
         EnableTextBox textTM67, True
      Case Else:
         EnableTextBox textTM08, False
         EnableTextBox textTM09, False
         EnableTextBox textTM27, False
         EnableTextBox textTM32, False
         EnableTextBox textTM67, False
   End Select
   
   ' 案件性質為延展時才可輸入延展後專用期限
   If m_CP10 = "102" Then
      EnableTextBox textTM21, True
      EnableTextBox textTM22, True
   Else
      EnableTextBox textTM21, False
      EnableTextBox textTM22, False
   End If
   
   ' 案件性質為申請, 申請國家為台灣
   '2008/10/21 MODIFY BY SONIA 取消分割案
   'If (m_CP10 = "101" Or m_CP10 = "308") And m_TM10 < "010" Then
   If m_CP10 = "101" And m_TM10 < "010" Then
      EnableTextBox textTM11, True
      EnableTextBox textTM12, True
   Else
      EnableTextBox textTM11, False
      EnableTextBox textTM12, False
   End If
   'add by nickc 2007/03/03 若是查名
    If textTM08 <> "7" And textTM08 <> "8" Then
        If (m_TM01 = "T" And m_TM10 = "000" And m_CP10 = "101") Or (m_TM01 = "TS" And m_CU10 <> "000") Then
            EnableTextBox textTM09, True
            EnableTextBox textTM32, True
        ElseIf (m_TM01 = "T" And m_TM10 = "000" And m_CP10 <> "101") Or (m_TM01 = "T" And m_TM10 = "020") Or (m_TM01 = "TF") Or (m_TM01 = "TS") Then
            EnableTextBox textTM09, True
        End If
    End If

   ' 申請國家為大陸, 且案件性質為刊登廣告時, 才可輸入該兩個欄位
   If m_CP10 = "702" And m_TM10 = "020" Then
      EnableTextBox textMediaType, True
      ' 下次刊登廣告期限
      textMediaDate = ""
      strSql = "SELECT MIN(NP09) FROM NEXTPROGRESS " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = '" & m_CP10 & "' AND NP06 IS NULL"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then
            textMediaDate = rsTmp.Fields(0)
         End If
      End If
      rsTmp.Close
      'EnableTextBox textMediaDate, True
   Else
      EnableTextBox textMediaType, False
      'EnableTextBox textMediaDate, False
   End If
   
   ' 讀取優先權資料
   m_Pa(1) = m_TM01
   m_Pa(2) = m_TM02
   m_Pa(3) = m_TM03
   m_Pa(4) = m_TM04
   'edit by nickc 2007/02/06 不用 dll 了 objPublicData.ReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify by Sindy 2017/10/12 + , m_Priority(6)
   ClsPDReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)
   
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
   
   'Add by Sindy 2011/3/9 101.申請
   'Add By Sindy 2011/10/28 T內商000台灣案所有案件性質加電子送件功能
   'Modify by Amy 2020/01/13 +是否電子送件
   lblPayToday.Visible = False
   txtPayToday.Visible = False
   If m_TM01 = "T" And m_TM10 = "000" Then
   'If m_CP10 = "101" Then
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
   '2011/3/9 End
   
   'Add By Sindy 2012/6/21 指定國家
   cmdCountry.Enabled = False
   If m_TM01 = "TF" And m_CP10 = "102" Then
      cmdCountry.Enabled = True
      Call cmdCountry_Click 'Add By Sindy 2025/2/18
   End If
   '2012/6/21 End
   
   Call PUB_TCaseEFeeRemind(m_CP09) 'Add By Sindy 2016/5/9 內商電子收文請款提醒訊息
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
    'Add By Cheng 2002/07/18
   Set frm020102_05 = Nothing
End Sub

' 91.09.02 marked by louis
' 大陸查名總收文號
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
'         strTit = "資料檢核"
'         strMsg = "大陸查名總收文號不可為本案之收文號"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP09S_GotFocus
'         GoTo EXITSUB
'      End If
'
'      strSQL = "SELECT * FROM CaseProgress " & _
'               "WHERE CP01 = 'TS' AND " & _
'                     "CP09 = '" & textCP09S & "' "
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount <= 0 Then
'         rsTmp.Close
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "大陸查名總收文號資料不存在"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP09S_GotFocus
'         GoTo EXITSUB
'      End If
'      rsTmp.Close
'   End If
'EXITSUB:
'   Set rsTmp = Nothing
'End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 And grdList.row < grdList.Rows Then
      textTM15S = grdList.TextMatrix(grdList.row, 1)
   End If
   grdList_ShowSelection
End Sub

' 91.09.02 modify by louis
'edit by nickc 2006/01/27
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
      Me.SSTab1.Tab = 1
      lstNameAgent.SetFocus
   End If
   '2015/7/22 END
End Sub

'Add By Sindy 2011/3/9
Private Sub textCP118_GotFocus()
   TextInverse textCP118
   CloseIme
End Sub

'Add By Sindy 2011/3/9
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
      
      '2006/11/13 ADD BY SONIA
      'Modify by Amy 2015/05/26 if中增加 Or判斷 for T大陸續展催審期限控管
      'Modify by Amy 2016/03/10 +TF馬德里商標續展同 T 大陸案
      'Modified by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
      'If textUargeDate = m_textUargeDate Or (m_CP10 = "102" And (m_TM10 = "020" Or m_TM10 = "238") And textUargeDate = "") Then
      'If textUargeDate = m_textUargeDate Or (m_CP10 = "102" And (m_TM10 = "020" Or m_TM10 = "238") And textUargeDate = "") Then
      If textCP27.Tag <> textCP27.Text Or (m_CP10 = "102" And (m_TM10 = "020" Or m_TM10 = "238") And textUargeDate = "") Then
            textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
            If m_CP10 = "102" And (m_TM10 = "020" Or m_TM10 = "238") Then
                Cancel = Not (SetUargeDate_102)
            End If
            m_textUargeDate = textUargeDate
      End If
      '2006/11/13 END
      textCP27.Tag = textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
      
      '2007/7/23 add by sonia 國內商標申請,分割案發文日更新至申請日
      '2008/10/21 modify by sonia 取消分割案 T-161630註冊後分割發證書發現
      'If (m_CP10 = "101" Or m_CP10 = "308") And m_TM10 < "010" And IsEmptyText(textMail) = False Then
      If m_CP10 = "101" And m_TM10 < "010" And IsEmptyText(textMail) = False Then
         textTM11 = textCP27
      End If
      '2007/7/23 end
      
      'Added by Lydia 2015/11/24 管控台灣延展案102,發文日不可小於"延展期滿前6個月"
      'modify by sonia 2016/7/5 改為發文日不可小於"延展期滿前6個月+1天"  T-093656(法定1051224不可於1050624發文)
      'If m_TM10 = 台灣國家代號 And m_CP10 = "102" And TransDate(textCP27, 2) < CompDate(1, -6, m_CP07) Then
      'Modified by Lydia 2017/06/01 延展期滿日期改用模組控制
      'If m_TM10 = 台灣國家代號 And m_CP10 = "102" And TransDate(textCP27, 2) < CompDate(2, 1, CompDate(1, -6, m_CP07)) Then
      If m_TM10 = 台灣國家代號 And m_CP10 = "102" And TransDate(textCP27, 2) < PUB_Get102DeadLine("3", m_CP07) Then
          Cancel = True
          strTit = "資料檢核"
          strMsg = "台灣延展案發文日不得早於延展期滿前6個月+1天!"
          nResponse = MsgBox(strMsg, vbOKOnly, strTit)
          textCP27_GotFocus
          GoTo EXITSUB
      End If
      'end 2015/11/24
   End If
EXITSUB:
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
   Dim nResponse
   Cancel = False
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

' 是否郵寄申請
Private Sub textMail_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否郵寄申請
Private Sub textMail_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textMail) = False Then
      Select Case textMail
         Case "", " ", "Y":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "請輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textMail_GotFocus
      End Select
   End If
End Sub

' 下次刊登廣告期限
Private Sub textMediaDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textMediaDate) = False Then
      ' 下次刊登廣告期限
      If CheckIsTaiwanDate(textMediaDate, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的下次刊登廣告期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textMediaDate_GotFocus
      End If
   End If
End Sub

Private Sub textMediaName_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If m_CP10 = "702" And m_TM10 = "020" Then
      If IsEmptyText(textMediaName) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入雜誌社, 報社"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Cancel = True
      End If
   End If
End Sub

' 提申期限
Private Sub textPetition_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPetition) = False Then
      ' 發文日日期不正確
      If CheckIsTaiwanDate(textPetition, False) = False Then
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
   Dim strTM15 As String
   Dim nIndex As Integer
   
   'add by nickc 2007/01/02
   Dim rsTmp As New ADODB.Recordset
   
   ' 是否算案件數
   SetCPFieldNewData "CP26", textCP26
   
   'Add By Sindy 2011/3/9
   ' 是否電子送件
   SetCPFieldNewData "CP118", textCP118
   
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
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
   ' 91.09.02 modify by louis
   ' 進度備註
   'If IsEmptyText(textMediaName) = False Then
   '   textCP64 = textCP64 & " 雜誌社/報社:" & textMediaName
   'End If
   'SetCPFieldNewData "CP64", textCP64
   strCP64 = textCP64
   If IsEmptyText(textMediaName) = False Then
      strCP64 = strCP64 & "," & "雜誌社/報社:" & textMediaName
   End If
'edit by nickc 2006/01/27
'   If IsEmptyText(textAgName) = False Then
'      strCP64 = strCP64 & "," & "本所出名代理人:" & textAgName
'   End If
   
   ' 是否出名
   SetCPFieldNewData "CP22", textCP22
   'add by nickc 2006/01/27
   SetCPFieldNewData "CP110", m_CP110

   ' 91.09.03 modify by louis 加入刊登廣告的審定號(非本所案件)
   For nIndex = 1 To grdList.Rows - 1
      Dim strCP01 As String
      Dim strCP02 As String
      Dim strCP03 As String
      Dim strCP04 As String
      If Not IsEmptyText(grdList.TextMatrix(nIndex, 1)) Then
         If ExistTM15(grdList.TextMatrix(nIndex, 1), strCP01, strCP02, strCP03, strCP04) Then
            grdList.TextMatrix(nIndex, 2) = "1"
            grdList.TextMatrix(nIndex, 3) = strCP01
            grdList.TextMatrix(nIndex, 4) = strCP02
            grdList.TextMatrix(nIndex, 5) = strCP03
            grdList.TextMatrix(nIndex, 6) = strCP04
            '911104 nick 邱小姐說不管是不是本所案件，都要存進度備註
            If Not IsEmptyText(strTM15) Then strTM15 = strTM15 & ","
            strTM15 = strTM15 & grdList.TextMatrix(nIndex, 1)
         Else
            grdList.TextMatrix(nIndex, 2) = "0"
            If Not IsEmptyText(strTM15) Then strTM15 = strTM15 & ","
            strTM15 = strTM15 & grdList.TextMatrix(nIndex, 1)
         End If
      End If
   Next nIndex
   If Not IsEmptyText(strTM15) Then
      strCP64 = strCP64 & "," & "同時刊登廣告之審定號:" & strTM15
   End If
    'Modify By Cheng 2003/09/05
    '取消
    'Begin
'    'Add By Cheng 2003/06/16
'    '若有輸入查名本所案號
'    If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
'        strCP64 = strCP64 & IIf(strCP64 <> "", ",", "") & "原查名本所案號：" & Me.textTM01.Text & "-" & Me.textTM02.Text & Me.textTM02_2.Text & "-" & Left(Me.textTM03.Text & "0", 1) & "-" & Left(Me.textTM04.Text & "00", 2)
'    End If
    'End
   SetCPFieldNewData "CP64", strCP64
   
   ' 案件性質為延展時, 更新案件進度檔的授權期間欄位
   If m_CP10 = "102" Then
      ' 授權期間(起)
      SetCPFieldNewData "CP53", DBDATE(textTM21)
      ' 授權期間(迄)
      SetCPFieldNewData "CP54", DBDATE(textTM22)
   End If
   
   Select Case m_TM01
      ' 系統類別為CFT的為更新商標基本檔
      Case "T", "TF", "FCT":
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
         ' 申請日
         SetTMSPFieldNewData "TM11", DBDATE(textTM11)
         ' 申請案號
         SetTMSPFieldNewData "TM12", textTM12
         ' 申請人
         If IsEmptyText(textTM23) = False Then
            SetTMSPFieldNewData "TM23", textTM23 & String(9 - Len(textTM23), "0")
         Else
            SetTMSPFieldNewData "TM23", textTM23
         End If
         'add by nickc 2005/11/18 若有修改申請人時，要更新基本檔的申請地址
         If m_TM23 & String(9 - Len(m_TM23), "0") <> textTM23 & String(9 - Len(textTM23), "0") Then
            'edit by nickc 2007/01/02 宣告搬到上面
            'Dim rsTmp As New ADODB.Recordset
            Set rsTmp = New ADODB.Recordset
            If rsTmp.State = 1 Then rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            'Modify By Sindy 2014/3/10 * ==> customer.*,nvl(cu04,nvl(cu05,cu06)) CName
            rsTmp.Open "select customer.*,nvl(cu04,nvl(cu05,cu06)) CName from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM23.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM23.Text), 9, 1) & "'", cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <> 0 Then
               'edit by nickc 2006/01/26
               'SetTMSPFieldNewData "TM24", CheckStr(rsTmp.Fields("cu23"))
               SetTMSPFieldNewData "TM24", CheckStr(rsTmp.Fields("cu112")) & CheckStr(rsTmp.Fields("cu23"))
               SetTMSPFieldNewData "TM25", CheckStr(rsTmp.Fields("CU24")) & IIf(CheckStr(rsTmp.Fields("cu25")) <> "", " " & CheckStr(rsTmp.Fields("cu25")), "") & IIf(CheckStr(rsTmp.Fields("cu26")) <> "", " " & CheckStr(rsTmp.Fields("cu26")), "") & IIf(CheckStr(rsTmp.Fields("cu27")) <> "", " " & CheckStr(rsTmp.Fields("cu27")), "") & IIf(CheckStr(rsTmp.Fields("cu28")) <> "", " " & CheckStr(rsTmp.Fields("cu28")), "")
               SetTMSPFieldNewData "TM26", CheckStr(rsTmp.Fields("cu29"))
               'Add By Sindy 2014/3/10
               If m_CP60 <> "" Then
                  strExc(1) = m_TM01
                  strExc(2) = m_TM02
                  strExc(3) = m_TM03
                  strExc(4) = m_TM04
                  strExc(5) = m_CP60
                  strExc(6) = ChangeCustomerL(textTM23)
                  strExc(7) = rsTmp.Fields("CName")
                  strExc(8) = ChangeCustomerL(m_TM23)
                  If Not ClsLawUpdAcc0k0(strExc(), True) Then
                     textTM23.SetFocus
                  End If
               End If
               '2014/3/10 END
            End If
         Else
            'edit by nickc 2006/01/26
            'SetTMSPFieldNewData "TM24", m_tm24
            If m_CU112 <> "" Then
                'Modify By Sindy 2011/2/22
                'SetTMSPFieldNewData "TM24", Pub_RplCu112(m_TM24, m_CU112)
                SetTMSPFieldNewData "TM24", Pub_RplCu112(m_TM24, m_CU112, m_TM23)
            Else
                SetTMSPFieldNewData "TM24", m_TM24
            End If
            SetTMSPFieldNewData "TM25", m_tm25
            SetTMSPFieldNewData "TM26", m_tm26
         End If
         ' 正商標號數
         SetTMSPFieldNewData "TM27", textTM27
         ' 商品組群
         SetTMSPFieldNewData "TM32", textTM32
         ' 案件備註
         SetTMSPFieldNewData "TM58", textTM58
         ' 放棄專用權
         SetTMSPFieldNewData "TM67", textTM67
         ' 特殊商標
         SetTMSPFieldNewData "TM72", textTM72
         'add by nickc 2006/11/17
         If Trim(textPrint) <> "N" Then
            SetTMSPFieldNewData "TM77", textPrint
         Else
            SetTMSPFieldNewData "TM77", m_textPrint
         End If
         'add by nickc 2007/01/02 申請人
         If IsEmptyText(textTM78) = False Then
            SetTMSPFieldNewData "TM78", textTM78 & String(9 - Len(textTM78), "0")
         Else
            SetTMSPFieldNewData "TM78", textTM78
         End If
         '若有修改申請人時，要更新基本檔的申請地址
         If m_TM78 & String(9 - Len(m_TM78), "0") <> textTM78 & String(9 - Len(textTM78), "0") Then
            Set rsTmp = New ADODB.Recordset
            If rsTmp.State = 1 Then rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open "select * from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM78.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM78.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <> 0 Then
                SetTMSPFieldNewData "TM82", CheckStr(rsTmp.Fields("cu112")) & CheckStr(rsTmp.Fields("cu23"))
                SetTMSPFieldNewData "TM86", CheckStr(rsTmp.Fields("CU24")) & IIf(CheckStr(rsTmp.Fields("cu25")) <> "", " " & CheckStr(rsTmp.Fields("cu25")), "") & IIf(CheckStr(rsTmp.Fields("cu26")) <> "", " " & CheckStr(rsTmp.Fields("cu26")), "") & IIf(CheckStr(rsTmp.Fields("cu27")) <> "", " " & CheckStr(rsTmp.Fields("cu27")), "") & IIf(CheckStr(rsTmp.Fields("cu28")) <> "", " " & CheckStr(rsTmp.Fields("cu28")), "")
                SetTMSPFieldNewData "TM90", CheckStr(rsTmp.Fields("cu29"))
            End If
         Else
            SetTMSPFieldNewData "TM82", m_TM82
            SetTMSPFieldNewData "TM86", m_TM86
            SetTMSPFieldNewData "TM90", m_TM90
         End If
         If IsEmptyText(textTM79) = False Then
            SetTMSPFieldNewData "TM79", textTM79 & String(9 - Len(textTM79), "0")
         Else
            SetTMSPFieldNewData "TM79", textTM79
         End If
         '若有修改申請人時，要更新基本檔的申請地址
         If m_TM79 & String(9 - Len(m_TM79), "0") <> textTM79 & String(9 - Len(textTM79), "0") Then
            Set rsTmp = New ADODB.Recordset
            If rsTmp.State = 1 Then rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open "select * from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM79.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM79.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <> 0 Then
                SetTMSPFieldNewData "TM83", CheckStr(rsTmp.Fields("cu112")) & CheckStr(rsTmp.Fields("cu23"))
                SetTMSPFieldNewData "TM87", CheckStr(rsTmp.Fields("CU24")) & IIf(CheckStr(rsTmp.Fields("cu25")) <> "", " " & CheckStr(rsTmp.Fields("cu25")), "") & IIf(CheckStr(rsTmp.Fields("cu26")) <> "", " " & CheckStr(rsTmp.Fields("cu26")), "") & IIf(CheckStr(rsTmp.Fields("cu27")) <> "", " " & CheckStr(rsTmp.Fields("cu27")), "") & IIf(CheckStr(rsTmp.Fields("cu28")) <> "", " " & CheckStr(rsTmp.Fields("cu28")), "")
                SetTMSPFieldNewData "TM91", CheckStr(rsTmp.Fields("cu29"))
            End If
         Else
            SetTMSPFieldNewData "TM83", m_TM83
            SetTMSPFieldNewData "TM87", m_TM87
            SetTMSPFieldNewData "TM91", m_TM91
         End If
         If IsEmptyText(textTM80) = False Then
            SetTMSPFieldNewData "TM80", textTM80 & String(9 - Len(textTM80), "0")
         Else
            SetTMSPFieldNewData "TM80", textTM80
         End If
         '若有修改申請人時，要更新基本檔的申請地址
         If m_TM80 & String(9 - Len(m_TM80), "0") <> textTM80 & String(9 - Len(textTM80), "0") Then
            Set rsTmp = New ADODB.Recordset
            If rsTmp.State = 1 Then rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open "select * from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM80.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM80.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <> 0 Then
                SetTMSPFieldNewData "TM84", CheckStr(rsTmp.Fields("cu112")) & CheckStr(rsTmp.Fields("cu23"))
                SetTMSPFieldNewData "TM88", CheckStr(rsTmp.Fields("CU24")) & IIf(CheckStr(rsTmp.Fields("cu25")) <> "", " " & CheckStr(rsTmp.Fields("cu25")), "") & IIf(CheckStr(rsTmp.Fields("cu26")) <> "", " " & CheckStr(rsTmp.Fields("cu26")), "") & IIf(CheckStr(rsTmp.Fields("cu27")) <> "", " " & CheckStr(rsTmp.Fields("cu27")), "") & IIf(CheckStr(rsTmp.Fields("cu28")) <> "", " " & CheckStr(rsTmp.Fields("cu28")), "")
                SetTMSPFieldNewData "TM92", CheckStr(rsTmp.Fields("cu29"))
            End If
         Else
            SetTMSPFieldNewData "TM84", m_TM84
            SetTMSPFieldNewData "TM88", m_TM88
            SetTMSPFieldNewData "TM92", m_TM92
         End If
         If IsEmptyText(textTM81) = False Then
            SetTMSPFieldNewData "TM81", textTM81 & String(9 - Len(textTM81), "0")
         Else
            SetTMSPFieldNewData "TM81", textTM81
         End If
         '若有修改申請人時，要更新基本檔的申請地址
         If m_TM81 & String(9 - Len(m_TM81), "0") <> textTM81 & String(9 - Len(textTM81), "0") Then
            Set rsTmp = New ADODB.Recordset
            If rsTmp.State = 1 Then rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open "select * from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM81.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM81.Text), 9, 1) & "' ", cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <> 0 Then
                SetTMSPFieldNewData "TM85", CheckStr(rsTmp.Fields("cu112")) & CheckStr(rsTmp.Fields("cu23"))
                SetTMSPFieldNewData "TM89", CheckStr(rsTmp.Fields("CU24")) & IIf(CheckStr(rsTmp.Fields("cu25")) <> "", " " & CheckStr(rsTmp.Fields("cu25")), "") & IIf(CheckStr(rsTmp.Fields("cu26")) <> "", " " & CheckStr(rsTmp.Fields("cu26")), "") & IIf(CheckStr(rsTmp.Fields("cu27")) <> "", " " & CheckStr(rsTmp.Fields("cu27")), "") & IIf(CheckStr(rsTmp.Fields("cu28")) <> "", " " & CheckStr(rsTmp.Fields("cu28")), "")
                SetTMSPFieldNewData "TM93", CheckStr(rsTmp.Fields("cu29"))
            End If
         Else
            SetTMSPFieldNewData "TM85", m_TM85
            SetTMSPFieldNewData "TM89", m_TM89
            SetTMSPFieldNewData "TM93", m_TM93
         End If
    Case Else:
        Select Case m_TM01
        Case "TS"
            ' 案件名稱
            SetTMSPFieldNewData "SP05", textTM05_1
        Case Else
            ' 案件中文名稱
            SetTMSPFieldNewData "SP05", textTM05
        End Select
         ' 案件英文名稱
         SetTMSPFieldNewData "SP06", textTM06
         ' 案件日文名稱
         SetTMSPFieldNewData "SP07", textTM07
         ' 申請人
         If IsEmptyText(textTM23) = False Then
            SetTMSPFieldNewData "SP08", textTM23 & String(9 - Len(textTM23), "0")
         Else
            SetTMSPFieldNewData "SP08", textTM23
         End If
         'Add By Sindy 2014/3/10
         If m_TM23 & String(9 - Len(m_TM23), "0") <> textTM23 & String(9 - Len(textTM23), "0") Then
            Set rsTmp = New ADODB.Recordset
            If rsTmp.State = 1 Then rsTmp.Close
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open "select customer.*,nvl(cu04,nvl(cu05,cu06)) CName from customer Where Cu01 = '" & Mid(ChangeCustomerL(textTM23.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM23.Text), 9, 1) & "'", cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <> 0 Then
               If m_CP60 <> "" Then
                  strExc(1) = m_TM01
                  strExc(2) = m_TM02
                  strExc(3) = m_TM03
                  strExc(4) = m_TM04
                  strExc(5) = m_CP60
                  strExc(6) = ChangeCustomerL(textTM23)
                  strExc(7) = rsTmp.Fields("CName")
                  strExc(8) = ChangeCustomerL(m_TM23)
                  If Not ClsLawUpdAcc0k0(strExc(), True) Then
                     textTM23.SetFocus
                  End If
               End If
            End If
         End If
         '2014/3/10 END
         'add by nickc 2007/01/02
         If IsEmptyText(textTM78) = False Then
            SetTMSPFieldNewData "SP58", textTM78 & String(9 - Len(textTM78), "0")
         Else
            SetTMSPFieldNewData "SP58", textTM78
         End If
         If IsEmptyText(textTM79) = False Then
            SetTMSPFieldNewData "SP59", textTM79 & String(9 - Len(textTM79), "0")
         Else
            SetTMSPFieldNewData "SP59", textTM79
         End If
         If IsEmptyText(textTM80) = False Then
            SetTMSPFieldNewData "SP65", textTM80 & String(9 - Len(textTM80), "0")
         Else
            SetTMSPFieldNewData "SP65", textTM80
         End If
         If IsEmptyText(textTM81) = False Then
            SetTMSPFieldNewData "SP66", textTM81 & String(9 - Len(textTM81), "0")
         Else
            SetTMSPFieldNewData "SP66", textTM81
         End If
         
         ' 申請日
         SetTMSPFieldNewData "SP10", DBDATE(textTM11)
         ' 申請案號
         SetTMSPFieldNewData "SP11", textTM12
         ' 案件備註
         SetTMSPFieldNewData "SP18", textTM58
         'add by nickc 2006/11/17
         If Trim(textPrint) <> "N" Then
            SetTMSPFieldNewData "SP72", textPrint
         Else
            SetTMSPFieldNewData "SP72", m_textPrint
         End If
         'add by nickc 2007/01/02
         SetTMSPFieldNewData "SP73", textTM09
         SetTMSPFieldNewData "SP74", textTM32
   End Select
      
End Sub

' 更新商標基本檔的相關欄位
'Modify By Cheng 2002/11/06
'Private Sub OnUpdateTradeMark()
Private Function OnUpdateTradeMark() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnUpdateTradeMark = True
      
   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
            Else
               ' 91.03.25 modify by louis (單引號轉換)
               'strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
               strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
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
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateTradeMark = False
End Function

' 更新服務業務基本檔的相關欄位
'Modify By Cheng 2002/11/06
'Private Sub OnUpdateServicePractice()
Private Function OnUpdateServicePractice() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnUpdateServicePractice = True
   
   ' 更新案件進度檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
            Else
               ' 91.03.25 modify by louis (單引號轉換)
               'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
               strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
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
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateServicePractice = False
End Function

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
               ' 91.03.25 modify by louis (單引號轉換)
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
Dim strTmp As String
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim nIndex As Integer
Dim bFirst As Boolean
Dim bDifference As Boolean
Dim strNP07 As String
Dim strNP08 As String
Dim strNP22 As String
Dim objCopyCP As ClsCopyCP
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCP06 As String
Dim strCP07 As String
Dim strCP09 As String 'Add By Sindy 2010/6/10
Dim strCP44 As String 'Add By Sindy 2010/11/5
Dim bolSysDt As Boolean 'Add By Sindy 2010/12/28
Dim i As Integer
Dim strCP05 As String, strCP27 As String 'Add by Amy 2015/04/15 大陸分割子案用
'add by sonia 2023/11/17
Dim ii As Integer
Dim str105DATE As String   '使用宣誓起算日
'end 2023/11/17

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
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新案件進度檔
    'Modify By Cheng 2002/11/06
'   OnUpdateCaseProgress
   If OnUpdateCaseProgress = False Then GoTo ErrorHandler
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新基本檔
   Select Case m_TM01
      ' 系統類別為CFT的為更新商標基本檔
      Case "T", "TF", "FCT":
        'Modify By Cheng 2002/11/06
'         OnUpdateTradeMark
         If OnUpdateTradeMark = False Then GoTo ErrorHandler
      Case Else:
        'Modify By Cheng 2002/11/06
'         OnUpdateServicePractice
         If OnUpdateServicePractice = False Then GoTo ErrorHandler
   End Select
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Add by Amy 2014/10/27 T大陸分割案控制
   If m_TM01 = "T" And m_TM10 = 大陸國家代號 And m_CP10 = "308" Then
        strExc(0) = "Select DC01,DC02,DC03,DC04,CP09 From DivisionCase,CaseProgress " & _
                          "Where DC05='" & m_TM01 & "' And DC06='" & m_TM02 & "' And DC07='" & m_TM03 & "' And DC08='" & m_TM04 & "' " & _
                          "And DC01=CP01(+) And DC02=CP02(+) And DC03=CP03(+) And DC04=CP04(+) And CP10='308' "
        
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            With RsTemp
                m_DC01 = .Fields("DC01")
                m_DC02 = .Fields("DC02")
                m_DC03 = .Fields("DC03")
                m_DC04 = .Fields("DC04")
                m_CP09s = .Fields("CP09") '子案總收文號
            End With
        End If
        '將母案商品服務轉至新案
        strExc(0) = "Select * From TMGoods Where TG01='" & m_TM01 & "' And TG02='" & m_TM02 & "' And TG03='" & m_TM03 & "' And TG04='" & m_TM04 & "' "
        
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            With RsTemp
                Do While .EOF = False
                    strExc(1) = "Select Count(*) From TMGoods Where TG01='" & m_DC01 & "' And TG02='" & m_DC02 & "' And TG03='" & m_DC03 & "' And TG04='" & m_DC04 & "' And TG05='" & .Fields("TG05") & "' "
                    If rsA.State <> adStateClosed Then rsA.Close
                    rsA.CursorLocation = adUseClient
                    rsA.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
                    If rsA.RecordCount > 0 And rsA.Fields(0) > 0 Then
                        strSql = "Update TMGoods set TG06='" & ChgSQL("" & .Fields("TG06")) & "',TG07='" & ChgSQL("" & .Fields("TG07")) & "',TG08='" & ChgSQL("" & .Fields("TG08")) & "'" & _
                                    ",TG12='" & strUserNum & "',TG13=to_number(to_char(sysdate,'YYYYMMDD')),TG14=to_number(to_char(sysdate,'HH24MI')) " & _
                                    "Where TG01='" & m_DC01 & "' And TG02='" & m_DC02 & "' And TG03='" & m_DC03 & "' And TG04='" & m_DC04 & "' And TG05='" & .Fields("TG05") & "' "
                    Else
                        strSql = "Insert Into TMGoods (TG01,TG02,TG03,TG04,TG05,TG06,TG07,TG08,TG09,TG10,TG11) Values(" & _
                                    "'" & m_DC01 & "' ,'" & m_DC02 & "' ,'" & m_DC03 & "' ,'" & m_DC04 & "' ,'" & .Fields("TG05") & "' ," & _
                                    "'" & ChgSQL("" & .Fields("TG06")) & "' ,'" & ChgSQL("" & .Fields("TG07")) & "' ,'" & ChgSQL("" & .Fields("TG08")) & "' ,'" & strUserNum & "' ,to_number(to_char(sysdate,'YYYYMMDD'))," & _
                                    "to_number(to_char(sysdate,'HH24MI')) ) "
                    End If
                    
                    cnnConnection.Execute strSql
                    .MoveNext
                Loop
            End With
        End If
        
        'Added by Lydia 2024/11/21 內商大陸之部份核駁商品異動，記錄各類的比對結果
        If m_CP43 <> "" Then
            strSql = "select tg01,tg02,tg03,tg04,tg05,tg06 from tmgoods where tg01='" & Mid(m_CP43, 1, 3) & "' and tg02='" & Mid(m_CP43, 4, 6) & "' and tg03='0' and tg04='00' order by tg05 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  If "" & RsTemp.Fields("tg06") <> "" Then
                     strSql = "Update TMGoods Set TG06='" & ChgSQL(RsTemp.Fields("tg06")) & "' Where TG01='" & m_TM01 & "' And TG02='" & m_TM02 & "' And TG03='" & m_TM03 & "' And TG04='" & m_TM04 & "' And TG05='" & RsTemp.Fields("tg05") & "' "
                     cnnConnection.Execute strSql, intI
                     If intI = 0 Then
                        strSql = "insert into Tmgoods (tg01,tg02,tg03,tg04,tg05,tg06,tg15,tg09,tg10,tg11) values ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & "" & RsTemp.Fields("tg05") & "','" & ChgSQL("" & RsTemp.Fields("tg06")) & "', null,'" & strUserNum & "',to_number(to_char(sysdate,'YYYYMMDD')),to_number(to_char(sysdate,'HH24MI'))) "
                        cnnConnection.Execute strSql
                     End If
                  End If
                  RsTemp.MoveNext
               Loop
            End If
            strSql = "delete from tmgoods where tg01='" & Mid(m_CP43, 1, 3) & "' and tg02='" & Mid(m_CP43, 4, 6) & "' and tg03='0' and tg04='00' "
            cnnConnection.Execute strSql
        Else
        'end 2024/11/21
           '母案商品服務改為1205程序進度備註(部分核駁商品)內容-多類先塞同一類(目前未遇到不知如何處理)
           tmpGoods1205 = Replace(GetGoods1205(m_TM01, m_TM02, m_TM03, m_TM04), "部分核駁商品：", "")
           If tmpGoods1205 <> "" Then
              strSql = "Update TMGoods Set TG06='" & tmpGoods1205 & "' Where TG01='" & m_TM01 & "' And TG02='" & m_TM02 & "' And TG03='" & m_TM03 & "' And TG04='" & m_TM04 & "' "
              cnnConnection.Execute strSql
           End If
        End If 'Added by Lydia 2024/11/21
        
        '更新子案新案申請日(同母案申請日)及上發文日
        'Modify by Amy 2015/04/15 +申請案號(為母案申請案號+A)
        strSql = "Update TradeMark Set TM11='" & DBDATE(textTM11) & "',TM12='" & textTM12 & "A' " & _
                    "Where TM01='" & m_DC01 & "' And TM02='" & m_DC02 & "' And TM03='" & m_DC03 & "' And TM04='" & m_DC04 & "' "
        cnnConnection.Execute strSql
        strSql = "Update CaseProgress Set CP27='" & DBDATE(textCP27) & "' " & _
                    "Where CP01='" & m_DC01 & "' And CP02='" & m_DC02 & "' And CP03='" & m_DC03 & "' And CP04='" & m_DC04 & "' "
        cnnConnection.Execute strSql
       
       'Add by Amy 2015/04/15 新增子案申請101假收文，自動發文
        strCP09 = AutoNo("B", 6)
        strCP05 = DBDATE("111111")
        strCP27 = DBDATE("111111")
        strSql = "INSERT INTO CaseProgress (CP01, CP02, CP03, CP04, CP05, CP09, CP10, CP12, CP13, CP14,  CP26, cp27,  CP43) " & _
                    "VALUES ('" & m_DC01 & "','" & m_DC02 & "','" & m_DC03 & "','" & m_DC04 & "'," & strCP05 & ",'" & strCP09 & "','" & "101" & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "','" & "N" & "'," & strCP27 & ",'" & m_CP09 & "')"
        cnnConnection.Execute strSql
        'Add by Amy 2015/04/15 新增子案申請之催審期限-發文日+180天
        strExc(0) = DBDATE(DateAdd("d", 180, Format(textCP27 + 19110000, "####/##/##")))
        'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
        'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                    "VALUES ('" & strCP09 & "','" & m_DC01 & "','" & m_DC02 & "','" & m_DC03 & "','" & m_DC04 & "',305," & _
                          strExc(1) & "," & strExc(0) & ",'" & m_CP14 & "'," & GetNextProgressNo & ")"
        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                    "VALUES ('" & strCP09 & "','" & m_DC01 & "','" & m_DC02 & "','" & m_DC03 & "','" & m_DC04 & "',305," & _
                          PUB_GetWorkDay1(strExc(0), True) & "," & strExc(0) & ",'" & m_CP14 & "'," & GetNextProgressNo & ")"
        cnnConnection.Execute strSql
        
   End If
   'end 2014/10/27
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有輸入催審期限時, 新增一筆催審的記錄到下一程序檔
   If IsEmptyText(textUargeDate) = False Then
      strNP07 = "305"
      strNP22 = GetNextProgressNo()
        'Modify By Cheng 2002/12/12
        '期限的智權人員欄位應掛承辦人非使用者
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                          DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strUserNum & "'," & strNP22 & ")"
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
         'Modify By Sindy 2009/06/16 增加109.被異議續展
         'Case "102", "105", "702", "708", "305", "998", "997"
         Case "102", "105", "702", "708", "305", "998", "997", "109"
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
            'strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP08)), Val(DBMONTH(strNP08)), Val(DBDAY(strNP08)) + Val(rsTmp.Fields("CF23")))))
            'NICK        *********
            'edit by nickc 2005/11/30
            'strNP08 = DBDATE(Val(DBDATE(strNP08)) + Val(rsTmp.Fields("CF23")))
            strNP08 = DBDATE(DateAdd("d", Val(rsTmp.Fields("CF23")), ChangeWStringToWDateString(DBDATE(strNP08))))
            ' *************
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
         ' 延展, 使用宣誓, 刊登廣告, 繳年費, 收達不印接洽結案單
'         '92.6.8 SONIA 加 言詞辯論, 準備程序
         Select Case strNP07
'            Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
            'Modify By Sindy 2009/06/16 增加109.被異議續展
            'Case "102", "105", "702", "708", "305", "998", "997"
            Case "102", "105", "702", "708", "305", "998", "997", "109"
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

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 有輸入提申期限時, 新增一筆提申的記錄到下一程序檔
   If IsEmptyText(textPetition) = False Then
      strNP07 = "998"
      'Add By Sindy 2010/12/28
      '非台灣案發文, 法定期限有值且為系統日或者過期時, 997.收達期限或998.提申期限都管制為系統日期
      If bolSysDt = True Then
         strNP08 = strSrvDate(1)
      Else
      '2010/12/28 End
         strNP08 = DBDATE(textPetition)
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
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 收達不印接洽結案單
'      '92.6.8 SONIA 加 言詞辯論, 準備程序
      Select Case strNP07
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
         'Modify By Sindy 2009/06/16 增加109.被異議續展
         'Case "102", "105", "702", "708", "305", "998", "997"
         Case "102", "105", "702", "708", "305", "998", "997", "109"
         Case Else:
            ' 列印國內案件接洽及結案記錄單
            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   End If
   
   ' 91.09.02 marked by louis
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 有輸入大陸查名, 商業司查詢總收文號時, 更新此收文號之本所案號為本案之本所案號
   'If IsEmptyText(textCP09S) = False Then
   '   strSQL = "UPDATE CaseProgress SET " & _
   '                  "CP01 = '" & m_TM01 & "', " & _
   '                  "CP02 = '" & m_TM02 & "', " & _
   '                  "CP03 = '" & m_TM03 & "', " & _
   '                  "CP04 = '" & m_TM04 & "' " & _
   '            "WHERE CP09 = '" & textCP09S & "' "
   '   cnnConnection.Execute strSQL
   'End If
   ' 91.09.02 modify by louis
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有輸入查名本所案號時, 更新該本所案號所有的案件進度檔其本所案號為本案之本所案號
   If IsEmptyText(textTM01) = False And IsEmptyText(Me.textTM02.Text) = False Then
      Dim strTM01 As String
      Dim strTM02 As String
      Dim strTM03 As String
      Dim strTM04 As String
      ' 組本所案號
      strTM01 = textTM01
      strTM02 = textTM02
      If (strTM01 = "TF") Then strTM02 = strTM02 & textTM02_2
      strTM03 = textTM03 & String(1 - Len(textTM03), "0")
      strTM04 = textTM04 & String(2 - Len(textTM04), "0")
       'add by nickc 2005/10/28 清未結餘的可結餘日期
       strSql = "UPDATE CaseProgress SET cp109=null " & _
                "WHERE CP01 = '" & strTM01 & "' AND " & _
                              "CP02 = '" & strTM02 & "' AND " & _
                              "CP03 = '" & strTM03 & "' AND " & _
                              "CP04 = '" & strTM04 & "'  and cp59 is null "
       cnnConnection.Execute strSql
       'edit by nickc 2006/07/18 加入 cp31=null
       strSql = "UPDATE CaseProgress SET cp31=null " & _
                "WHERE CP01 = '" & strTM01 & "' AND " & _
                              "CP02 = '" & strTM02 & "' AND " & _
                              "CP03 = '" & strTM03 & "' AND " & _
                              "CP04 = '" & strTM04 & "'  "
       cnnConnection.Execute strSql
      ' 組SQL語法
      strSql = "UPDATE CaseProgress SET CP01 = '" & m_TM01 & "', CP02 = '" & m_TM02 & "', " & _
                     "CP03 = '" & m_TM03 & "', CP04 = '" & m_TM04 & "', " & _
                     "CP64=CP64||Decode(CP64,Null,'','，')||'" & "原查名本所案號：" & Me.textTM01.Text & "-" & Me.textTM02.Text & Me.textTM02_2.Text & "-" & Left(Me.textTM03.Text & "0", 1) & "-" & Left(Me.textTM04.Text & "00", 2) & "' " & _
               "WHERE CP01 = '" & strTM01 & "' AND " & _
                     "CP02 = '" & strTM02 & "' AND " & _
                     "CP03 = '" & strTM03 & "' AND " & _
                     "CP03 = '" & strTM03 & "' "
      ' 執行更新的SQL指令
      cnnConnection.Execute strSql
      'Add By Cheng 2003/06/16
      strSql = "Update ServicePractice Set SP18=SP18||Decode(SP18,Null,'','，')||'轉入商標：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "' Where " & ChgService(strTM01 & strTM02 & strTM03 & strTM04)
      cnnConnection.Execute strSql
      '2005/4/18 ADD BY SONIA 1~4欄原查名本所案號,5~8欄新商標本所案號
      If PUB_UpdOther(Me.textTM01.Text, Me.textTM02.Text, Left(Me.textTM03.Text & "0", 1), Left(Me.textTM04.Text & "00", 2), m_TM01, m_TM02, m_TM03, m_TM04) = False Then
         GoTo ErrorHandler
      End If
      '2005/4/18 END
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 有輸入補優先權文件期限時, 新增下一程序為補文件的記錄
   If IsEmptyText(textPriorityDate) = False Then
      '2007/9/28 modify by sonia 改下一程序案件性質,所以不必放備註
      'strNP07 = "201"
      strNP07 = "208"
      strNP22 = GetNextProgressNo()
        'Modify By Cheng 2003/11/19
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22) " & _
'               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                       DBDATE(textPriorityDate) & "," & DBDATE(textPriorityDate) & ",'" & m_CP13 & "','" & "補優先權文件" & "'," & strNP22 & ")"
      '2007/9/28 modify by sonia 改下一程序案件性質,所以不必放備註
      'strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22) " & _
      '         "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
      '                 DBDATE(textPriorityDate) & "," & DBDATE(textPriorityDate) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & "補優先權文件" & "'," & strNP22 & ")"
      'Modify By Sindy 2013/11/22 改本所期限為法定期限減十天(日曆天)
      'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strNP08 = DBDATE(DateAdd("d", -10, ChangeWStringToWDateString(DBDATE(textPriorityDate))))
      strNP08 = PUB_GetWorkDay1(DateAdd("d", -10, ChangeWStringToWDateString(DBDATE(textPriorityDate))), True)
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                       strNP08 & "," & DBDATE(textPriorityDate) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
      '2013/11/22 END
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 收達不印接洽結案單
'      '92.6.8 SONIA 加 言詞辯論, 準備程序
      Select Case strNP07
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
         'Modify By Sindy 2009/06/16 增加109.被異議續展
         'Case "102", "105", "702", "708", "305", "998", "997"
         Case "102", "105", "702", "708", "305", "998", "997", "109"
         Case Else:
            ' 列印國內案件接洽及結案記錄單
            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   End If
   
   '92.3.8 CANCEL BY SONIA
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 有輸入下次刊登廣告期限時, 新增一筆刊登廣告的記錄到下一程序檔
   'If IsEmptyText(textMediaDate) = False Then
   '   strNP07 = "702"
   '   strNP22 = GetNextProgressNo()
   '   strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
   '            "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
   '                    DBDATE(textMediaDate) & "," & DBDATE(textMediaDate) & ",'" & m_CP13 & "'," & strNP22 & ")"
   '   cnnConnection.Execute strSQL
   '   ' 延展, 使用宣誓, 刊登廣告, 繳年費, 收達不印接洽結案單
   '92.6.8 SONIA 加 言詞辯論, 準備程序
   '   Select Case strNP07
   '      Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
   '      Case Else:
   '         ' 列印國內案件接洽及結案記錄單
   '         g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
   '   End Select
   'End If
   '92.3.8 END
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 儲存優先權資料
    'Modify By Cheng 2002/11/06
'   objPublicData.SavePriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)
   'edit by nickc 2007/02/06 不用 dll 了
   'If objPublicData.SavePriority(m_Pa, m_Priority(1), m_Priority(2), m_Priority(3)) = False Then GoTo ErrorHandler
   'Modify by Amy 2014/04/17 +, m_Priority(4), m_Priority(5)
   'Modify by Sindy 2017/10/12 + , m_Priority(6)
   If ClsPDSavePriority(m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5), m_Priority(6)) = False Then GoTo ErrorHandler
    'Add By Cheng 2003/11/12
    '若為商申案且有優先權資料, 則管制"主張優先權"(108)的期限
    If m_CP10 = "101" And m_Priority(1) <> "" Then
        '法定期限
        strCP07 = DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(Me.textCP27.Text))))
        '本所期限
         'Modify By Sindy 2014/10/6 台灣案之本所期限設定
         If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
            strCP06 = PUB_GetOurDeadline(DBDATE(strCP07))
         Else
         '2014/10/6 END
            strCP06 = DBDATE(DateAdd("d", -4, ChangeWStringToWDateString(DBDATE(strCP07))))
         End If
         strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
        StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='108' "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        '若有收文主張優先權, 更新進度檔
        If rsA.RecordCount > 0 Then
            StrSQLa = "Update CaseProgress Set CP06=" & strCP06 & ", CP07=" & strCP07 & " Where CP09='" & rsA("CP09").Value & "' "
            cnnConnection.Execute StrSQLa
        '若未收文主張優先權, 新增下一程序檔
        Else
            strNP07 = "108"
            strNP22 = GetNextProgressNo()
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                            "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
            cnnConnection.Execute strSql
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若該筆記錄是母案時, 同時對所有的子案做新增案件進度檔的工作
   If m_TM01 = "TF" And m_TM03 = "0" And m_TM04 = "00" Then
      Set objCopyCP = New ClsCopyCP
      'edit by nickc 2006/06/02
      If m_CP10 <> "102" Then
        objCopyCP.CopyCaseProgress m_CP09
      Else
        objCopyCP.CopyCaseProgress m_CP09, strLicenceCountry
      End If
      Set objCopyCP = Nothing
      If m_CP10 = "102" And Trim(Replace(strLicenceCountry, ",", "")) <> "" Then
            'add by nickc 2006/06/02 將未勾取的上閉卷，日期和原因都不用上，秀玲說的
            '2006/10/16 MODIFY BY SONIA 母案延展抓出所有未閉卷子案,含領土延伸之子案
            '                           領土延伸延展只抓出該領土延伸之未閉卷子案
            'strSQL = "update trademark set tm29='Y' where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03<>'0' and tm10 not in (" & strLicenceCountry & ")"
            'Modify By Sindy 2012/6/21 Mark
'            If Mid(m_TM02, 6, 1) = "0" Then
               strSql = "update trademark set tm29='Y' where tm01='" & m_TM01 & "' and SUBSTR(tm02,1,5)=SUBSTR('" & m_TM02 & "',1,5) and tm03<>'0' AND TM29 IS NULL and tm10 not in (" & strLicenceCountry & ")"
'            Else
'               strSql = "update trademark set tm29='Y' where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03<>'0' AND TM29 IS NULL and tm10 not in (" & strLicenceCountry & ")"
'            End If
            '2012/6/21 End
            '2006/10/16 END
            cnnConnection.Execute strSql
      End If
      'add by sonia 2023/11/17 領土延伸104，子案要檢查是否要掛使用宣誓期限TF-00083-3-0-00
      If m_CP10 = "104" Then
         Dim MyTFrs As New ADODB.Recordset
         Set MyTFrs = New ADODB.Recordset
         If MyTFrs.State = 1 Then MyTFrs.Close
         MyTFrs.CursorLocation = adUseClient
         MyTFrs.Open "select * from trademark where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm04<>'00' and tm29 is null ", cnnConnection, adOpenStatic, adLockReadOnly
         If MyTFrs.RecordCount <> 0 Then
            MyTFrs.MoveFirst
            Do While Not MyTFrs.EOF
               '菲律賓030,波多黎各112申請日+3年的期限，莫三比克318申請日+5年
               If Val("" & MyTFrs.Fields("TM11")) > 0 And (CheckStr(MyTFrs.Fields("tm10")) = "030" Or CheckStr(MyTFrs.Fields("tm10")) = "112" Or CheckStr(MyTFrs.Fields("tm10")) = "318") Then
                  strCP07 = DBDATE(DateAdd("yyyy", 3, ChangeWStringToWDateString(MyTFrs.Fields("TM11"))))
                  If CheckStr(MyTFrs.Fields("tm10")) = "318" Then
                     strCP07 = DBDATE(DateAdd("yyyy", 5, ChangeWStringToWDateString(MyTFrs.Fields("TM11"))))
                  End If
                  If Val(strCP07) >= Val(strSrvDate(1)) Then
                     strCP06 = CompDate(1, -2, strCP07)
                     strCP06 = PUB_GetWorkDay1(strCP06, True) '若本所期限非工作天則直接調整至最近的工作天
                     Set rsA = New ADODB.Recordset
                     StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(CheckStr(MyTFrs.Fields("tm01")) & CheckStr(MyTFrs.Fields("tm02")) & CheckStr(MyTFrs.Fields("tm03")) & CheckStr(MyTFrs.Fields("tm04"))) & _
                               " And np07=105 AND NP06 IS NULL"
                     rsA.CursorLocation = adUseClient
                     rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                     If rsA.RecordCount > 0 Then
                         strSql = "update NextProgress set NP01='" & m_CP09 & "',np08=" & DBDATE(strCP06) & ",np09=" & DBDATE(strCP07) & " where " & ChgNextProgress(CheckStr(MyTFrs.Fields("tm01")) & CheckStr(MyTFrs.Fields("tm02")) & CheckStr(MyTFrs.Fields("tm03")) & CheckStr(MyTFrs.Fields("tm04"))) & _
                                  " And np07=105 And NP06 IS NULL"
                         cnnConnection.Execute strSql
                     Else
                         strNP07 = "105"
                         strNP22 = GetNextProgressNo()
                         strNP08 = DBDATE(strCP06)
                         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                                 "VALUES ('" & m_CP09 & "','" & CheckStr(MyTFrs.Fields("tm01")) & "','" & CheckStr(MyTFrs.Fields("tm02")) & "','" & CheckStr(MyTFrs.Fields("tm03")) & "','" & CheckStr(MyTFrs.Fields("tm04")) & "'," & strNP07 & "," & _
                                 DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
                         cnnConnection.Execute strSql
                     End If
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                  End If
               End If
               '已有發證日才做以下
               If Val(m_TM20) > 0 Then
                  ii = 1
                  ' 取得使用宣誓年度
                  Set rsA = New ADODB.Recordset
                  Set rsA = Nothing
                  StrSQLa = "SELECT * FROM Nation WHERE NA01 = '" & CheckStr(MyTFrs.Fields("tm10")) & "' AND NA38 IS NOT NULL "
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     strCP07 = DBDATE(DateAdd("yyyy", Val("" & rsA.Fields("NA38")), ChangeWStringToWDateString(DBDATE(m_TM20))))
                     '104墨西哥管制三年使用宣誓期限，即註冊日起滿三年後之三個月內應提出使用宣誓
                     '110海地案為五年加三個月
                     If CheckStr(MyTFrs.Fields("tm10")) = "104" Or CheckStr(MyTFrs.Fields("tm10")) = "110" Then
                       strCP07 = CompDate(1, 3, strCP07)
                     End If
                     '檢查是否過期，未過期要計算下一次
                     If Val(strCP07) >= Val(strSrvDate(1)) Then
                         If Val(strCP07) >= Val(DBDATE(m_TM22)) Then
                             strCP07 = ""
                         End If
                     Else
                        If Not IsNull(rsA.Fields("NA39")) Then
                           '以發證日計算使用宣誓法定期限
                           str105DATE = Val(m_TM20)
ReDo:
                           strCP07 = DBDATE(Format(DateSerial(Val(DBYEAR(str105DATE)) + Val("" & rsA.Fields("NA38")) + Val("" & rsA.Fields("NA39") * ii), Val(DBMONTH(str105DATE)), Val(DBDAY(str105DATE)))))
                           If Val(strCP07) < Val(strSrvDate(1)) Then
                              ii = ii + 1
                              GoTo ReDo
                           Else
                              If Val(strCP07) >= Val(DBDATE(m_TM22)) Then
                                 '菲律賓2017/8/1新法延展核准後一年方再提出「延展使用宣誓」，故菲律賓不檢查專用期止日
                                 If CheckStr(MyTFrs.Fields("tm10")) <> "030" Then
                                    strCP07 = ""
                                 End If
                              End If
                           End If
                        '延展後之使用宣誓
                        ElseIf Not IsNull(rsA.Fields("NA78")) Then
                           str105DATE = Val(m_TM20)
ReDo105:
                           strCP07 = DBDATE(Format(DateSerial(Val(DBYEAR(str105DATE)) + Val("" & rsA.Fields("NA13")) + Val("" & rsA.Fields("NA78")) + Val("" & rsA.Fields("NA14")) * ii, Val(DBMONTH(str105DATE)), Val(DBDAY(str105DATE)))))
                           '110海地案為五年加三個月
                           If CheckStr(MyTFrs.Fields("tm10")) = "110" Then
                              strCP07 = CompDate(1, 3, strCP07)
                           End If
                           If Val(strCP07) < Val(strSrvDate(1)) Then
                              ii = ii + 1
                              GoTo ReDo105
                           Else
                              If Val(strCP07) >= Val(DBDATE(m_TM22)) Then
                                 '菲律賓2017/8/1新法延展核准後一年方再提出「延展使用宣誓」，故菲律賓不檢查專用期止日
                                 If CheckStr(MyTFrs.Fields("tm10")) <> "030" Then
                                    strCP07 = ""
                                 End If
                              End If
                           End If
                        End If
                     End If
                     If Val(strCP07) > 0 Then
                        strCP06 = CompDate(1, -2, strCP07)
                        strCP06 = PUB_GetWorkDay1(strCP06, True) '若本所期限非工作天則直接調整至最近的工作天
                        Set rsA = New ADODB.Recordset
                        StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(CheckStr(MyTFrs.Fields("tm01")) & CheckStr(MyTFrs.Fields("tm02")) & CheckStr(MyTFrs.Fields("tm03")) & CheckStr(MyTFrs.Fields("tm04"))) & _
                                  " And np07=105 AND NP06 IS NULL"
                        rsA.CursorLocation = adUseClient
                        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                        If rsA.RecordCount > 0 Then
                            strSql = "update NextProgress set NP01='" & m_CP09 & "',np08=" & DBDATE(strCP06) & ",np09=" & DBDATE(strCP07) & " where " & ChgNextProgress(CheckStr(MyTFrs.Fields("tm01")) & CheckStr(MyTFrs.Fields("tm02")) & CheckStr(MyTFrs.Fields("tm03")) & CheckStr(MyTFrs.Fields("tm04"))) & _
                                     " And np07=105 And NP06 IS NULL"
                            cnnConnection.Execute strSql
                        Else
                            strNP07 = "105"
                            strNP22 = GetNextProgressNo()
                            strNP08 = DBDATE(strCP06)
                            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                                    "VALUES ('" & m_CP09 & "','" & CheckStr(MyTFrs.Fields("tm01")) & "','" & CheckStr(MyTFrs.Fields("tm02")) & "','" & CheckStr(MyTFrs.Fields("tm03")) & "','" & CheckStr(MyTFrs.Fields("tm04")) & "'," & strNP07 & "," & _
                                    DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
                            cnnConnection.Execute strSql
                        End If
                     End If
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
               End If
               MyTFrs.MoveNext
            Loop
         End If
      End If
      'end 2023/11/17
   End If
   
   'add by nick 2004/08/12 更新實際發文規費
   If textCP84.Enabled = True Then
      strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
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
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 91.09.03 modify by louis
   ' 對所輸入的審定號檢查若存在則產生B類案件進度資料
    'Modify By Cheng 2002/11/06
'   OnCopyCPData m_CP09
   If OnCopyCPData(m_CP09) = False Then GoTo ErrorHandler
   
   'add by nick 2004/09/27 存公司負責人英文名稱
   'edit by nick 2004/10/07
   'If m_CU103 <> "" And m_TM01 <> "FCT" Then
   'edit by nickc 2006/01/20
   'If (m_CU103 <> "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) <> "") And m_TM01 <> "FCT" Then
   'edit by nickc 2007/08/10 改多申請人
   'If (m_CU103 <> "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) <> "" Or m_CU112 <> "") And m_TM01 <> "FCT" Then
   'Modify By Sindy 2012/2/27 +SeekCu39(1),SeekCu39(2),SeekCu39(3),SeekCu39(4),SeekCu39(5)
   '                          +SeekCu40(1),SeekCu40(2),SeekCu40(3),SeekCu40(4),SeekCu40(5)
   '                          +SeekCu41(1),SeekCu41(2),SeekCu41(3),SeekCu41(4),SeekCu41(5)
   'Modify By Sindy 2012/10/31 +SeekCu10(1),SeekCu10(2),SeekCu10(3),SeekCu10(4),SeekCu10(5)
   If (SeekCu103(1) <> "" Or (SeekCu05(1) & SeekCu88(1) & SeekCu89(1) & SeekCu90(1)) <> "" Or SeekCu112(1) <> "" Or (SeekCu39(1) & SeekCu40(1) & SeekCu41(1)) <> "" Or SeekCu10(1) <> "") And m_TM01 <> "FCT" Then
            'edit by nickc 2006/01/20
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "'  Where Cu01 = '" & Mid(ChangeCustomerL(textTM23.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM23.Text), 9, 1) & "' "
            'edit by nickc 2007/08/10 改多申請人
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "',cu112='" & ChgSQL(m_CU112) & "'  Where Cu01 = '" & Mid(ChangeCustomerL(textTM23.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM23.Text), 9, 1) & "' "
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(1)) & "',cu05='" & ChgSQL(SeekCu05(1)) & "',cu88='" & ChgSQL(SeekCu88(1)) & "',cu89='" & ChgSQL(SeekCu89(1)) & "',cu90='" & ChgSQL(SeekCu90(1)) & "',cu112='" & ChgSQL(SeekCu112(1)) & "',cu39='" & ChgSQL(SeekCu39(1)) & "',cu40='" & ChgSQL(SeekCu40(1)) & "',cu41='" & ChgSQL(SeekCu41(1)) & "',cu10='" & ChgSQL(SeekCu10(1)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textTM23.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM23.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(1)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textTM23.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM23.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   'add by nickc 2007/08/10 加多申請人也要
   If (SeekCu103(2) <> "" Or (SeekCu05(2) & SeekCu88(2) & SeekCu89(2) & SeekCu90(2)) <> "" Or SeekCu112(2) <> "" Or (SeekCu39(2) & SeekCu40(2) & SeekCu41(2)) <> "" Or SeekCu10(2) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(2)) & "',cu05='" & ChgSQL(SeekCu05(2)) & "',cu88='" & ChgSQL(SeekCu88(2)) & "',cu89='" & ChgSQL(SeekCu89(2)) & "',cu90='" & ChgSQL(SeekCu90(2)) & "',cu112='" & ChgSQL(SeekCu112(2)) & "',cu39='" & ChgSQL(SeekCu39(2)) & "',cu40='" & ChgSQL(SeekCu40(2)) & "',cu41='" & ChgSQL(SeekCu41(2)) & "',cu10='" & ChgSQL(SeekCu10(2)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textTM78.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM78.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(2)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textTM78.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM78.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(3) <> "" Or (SeekCu05(3) & SeekCu88(3) & SeekCu89(3) & SeekCu90(3)) <> "" Or SeekCu112(3) <> "" Or (SeekCu39(3) & SeekCu40(3) & SeekCu41(3)) <> "" Or SeekCu10(3) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(3)) & "',cu05='" & ChgSQL(SeekCu05(3)) & "',cu88='" & ChgSQL(SeekCu88(3)) & "',cu89='" & ChgSQL(SeekCu89(3)) & "',cu90='" & ChgSQL(SeekCu90(3)) & "',cu112='" & ChgSQL(SeekCu112(3)) & "',cu39='" & ChgSQL(SeekCu39(3)) & "',cu40='" & ChgSQL(SeekCu40(3)) & "',cu41='" & ChgSQL(SeekCu41(3)) & "',cu10='" & ChgSQL(SeekCu10(3)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textTM79.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM79.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(3)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textTM79.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM79.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(4) <> "" Or (SeekCu05(4) & SeekCu88(4) & SeekCu89(4) & SeekCu90(4)) <> "" Or SeekCu112(4) <> "" Or (SeekCu39(4) & SeekCu40(4) & SeekCu41(4)) <> "" Or SeekCu10(4) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(4)) & "',cu05='" & ChgSQL(SeekCu05(4)) & "',cu88='" & ChgSQL(SeekCu88(4)) & "',cu89='" & ChgSQL(SeekCu89(4)) & "',cu90='" & ChgSQL(SeekCu90(4)) & "',cu112='" & ChgSQL(SeekCu112(4)) & "',cu39='" & ChgSQL(SeekCu39(4)) & "',cu40='" & ChgSQL(SeekCu40(4)) & "',cu41='" & ChgSQL(SeekCu41(4)) & "',cu10='" & ChgSQL(SeekCu10(4)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textTM80.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM80.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(4)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textTM80.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM80.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(5) <> "" Or (SeekCu05(5) & SeekCu88(5) & SeekCu89(5) & SeekCu90(5)) <> "" Or SeekCu112(5) <> "" Or (SeekCu39(5) & SeekCu40(5) & SeekCu41(5)) <> "" Or SeekCu10(5) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(5)) & "',cu05='" & ChgSQL(SeekCu05(5)) & "',cu88='" & ChgSQL(SeekCu88(5)) & "',cu89='" & ChgSQL(SeekCu89(5)) & "',cu90='" & ChgSQL(SeekCu90(5)) & "',cu112='" & ChgSQL(SeekCu112(5)) & "',cu39='" & ChgSQL(SeekCu39(5)) & "',cu40='" & ChgSQL(SeekCu40(5)) & "',cu41='" & ChgSQL(SeekCu41(5)) & "',cu10='" & ChgSQL(SeekCu10(5)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textTM81.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM81.Text), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(5)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(textTM81.Text), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(textTM81.Text), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   
'   'Add By Sindy 2010/6/10
'   '由巨京代理之延展案，發文延展時，請同時假收文變更代理人
'   '變更案的收文與發文時間同延展案
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
'      If m_TM01 = "T" And m_TM10 = "020" And m_CP10 = "102" Then
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
   
   'Add by Sindy 98/3/24
   If m_TM10 = "000" Then
      'Modify By Sindy 2009/04/24
      'PUB_UpdateDispatch m_CP09s, m_CP123s
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
      
      'Add By Sindy 2014/2/20 更新商品檔未延展註記
      If m_CP10 = "102" Then
         '依基本檔的商品類別,逐筆檢查類別是否已存在商品檔裡,若沒有則新增資料
         If Trim(textTM09) <> "" Then 'Add By Sindy 2014/4/18 +if 證明標章沒有商品類別
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
            textTM09 = IIf(Right(textTM09, 1) = ",", Mid(textTM09, 1, Len(textTM09) - 1), textTM09)
            textTM09 = IIf(Left(textTM09, 1) = ",", Mid(textTM09, 2, Len(textTM09)), textTM09)
            '不存在畫面上的商品類別,在商品檔必須更新為未延展
            strSql = "Update TMGoods Set TG18='N'" & _
                     " Where TG01='" & m_TM01 & "' and TG02='" & m_TM02 & "' and TG03='" & m_TM03 & "' and TG04='" & m_TM04 & "'" & _
                     " and TG05 not in('" & Replace(textTM09, ",", "','") & "')"
            cnnConnection.Execute strSql
         End If
      End If
      '2014/2/20 END
   End If
   
   'Add by Sindy 2012/10/4 外->台,智權人員是葉雪貞及巨京,發文規費和收文規費不相同時,系統自動更改進度檔內規費費用及計算點數
   'Modified by Lydia 2015/10/16 + m_CP84
   Call PUB_TSendUpdateCP1718(m_CP09, textCP84, textPrint, m_TM10, m_CP13, m_CP84)
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   
   'Add By Sindy 2019/12/20 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = m_CP09
      PUB_AddLetterProgress strLD18, 0, IIf(textPrint = "N", False, True), "", False, m_TM23, m_CP10, m_TM44
   End If
   '2019/12/20 END
   Call PUB_UpdateLP19_T(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textCP27) 'Add by Sindy 2020/2/12 收據/回執設定
   
   'Add By Sindy 2016/12/16
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/16 END

   
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If
   Set rsTmp = Nothing
'Add By Cheng 2002/11/06
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

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號的系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM01) = False Then
      Select Case textTM01
         Case "TF":
            textTM02_2.Visible = True
            textTM02_2.Locked = False
            textTM02_2.TabStop = True
            textTM02.MaxLength = 5
         Case Else
            textTM02_2.Visible = False
            textTM02_2.Locked = True
            textTM02_2.TabStop = False
            textTM02.MaxLength = 6
      End Select
   Else
      textTM02_2.Visible = False
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02.MaxLength = 6
   End If
EXITSUB:
End Sub

Private Sub textTM04_LostFocus()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
    'Modify By Cheng 2003/06/16
    '存檔時再指定
'    'Add By Cheng 2002/12/11
'    '若有輸入查名本所案號
'    If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
'        Me.textCP64.Text = Me.textCP64.Text & "原查名本所案號：" & Me.textTM01.Text & "-" & Me.textTM02.Text & Me.textTM02_2.Text & "-" & Left(Me.textTM03.Text & "0", 1) & "-" & Left(Me.textTM04.Text & "00", 2) & "，"
'    End If
    'Add By Cheng 2003/06/16
    '若有輸入查名本所案號
    If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
        StrSQLa = "Select * From ServicePractice Where " & ChgService(Me.textTM01.Text & Me.textTM02.Text & Me.textTM02_2.Text & Left(Me.textTM03.Text & "0", 1) & Left(Me.textTM04.Text & "00", 2))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount <= 0 Then
            MsgBox "您輸入的查名本所案號錯誤，請重新輸入!!!", vbExclamation + vbOKOnly
            Me.textTM01.SetFocus
            textTM01_GotFocus
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
End Sub

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
        Me.SSTab1.Tab = 0
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
        Me.SSTab1.Tab = 0
      textTM05_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM05.IMEMode = 2
   If Cancel = False Then CloseIme
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
        Me.SSTab1.Tab = 0
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
        Me.SSTab1.Tab = 0
      textTM07_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textTM07.IMEMode = 2
   If Cancel = False Then CloseIme
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

' 申請日
Private Sub textTM11_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM11) = False Then
      If CheckIsTaiwanDate(textTM11, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請輸入正確的申請日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM11_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2010/8/31
Private Sub textTM12_Validate(Cancel As Boolean)
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   If IsEmptyText(textTM12) = False Then
      If textTM12.Enabled = True And textTM12.Locked = False Then
         '檢查申請案號所輸入的長度是否正確
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
   End If
End Sub

Private Sub textTM15S_GotFocus()
   InverseTextBox textTM15S
End Sub

Private Sub textTM15S_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM15S_LostFocus()
Dim nick911015Rs As New ADODB.Recordset
Dim nickstrsql As String
Set nick911015Rs = New ADODB.Recordset
nickstrsql = "select tm01||'-'||tm02||'-'||tm03||'-'||tm04,tm05,tm06,tm07  from trademark where tm15='" & textTM15S.Text & "' "
nick911015Rs.CursorLocation = adUseClient
nick911015Rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
If nick911015Rs.RecordCount <> 0 Then
    nick911015(0).Caption = CheckStr(nick911015Rs.Fields(0).Value)
    nick911015(1).Caption = CheckStr(nick911015Rs.Fields(1).Value)
    nick911015(2).Caption = CheckStr(nick911015Rs.Fields(2).Value)
    nick911015(3).Caption = CheckStr(nick911015Rs.Fields(3).Value)
Else
    nick911015(0).Caption = ""
    nick911015(1).Caption = ""
    nick911015(2).Caption = ""
    nick911015(3).Caption = ""
End If

End Sub

' 專用期限起日
Private Sub textTM21_Validate(Cancel As Boolean)
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM21) = False Then
        'Modify By Cheng 2003/08/18
        '若為大陸延展案時, 專用期限輸西元日期
        If m_TM10 <> 台灣國家代號 And m_CP10 = "102" Then
            If CheckIsDate(textTM21, False) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "請輸入正確的延展後專用期限起日"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM21_GotFocus
               GoTo EXITSUB
            End If
            If IsEmptyText(m_TM22) = False Then
               'Modify By Cheng 2003/09/01
               'strTemp = DBDATE(DateSerial(Val(DBYEAR(m_TM22)), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22)) + 1))
               'modify by sonia 2014/10/31
               'strTemp = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22))))
               'modify by sonia 2018/9/26 改用共用
               'If m_TM01 = "TF" Then
               '   strTemp = DBDATE(m_TM22)
               'Else
               '   strTemp = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22))))
               'End If
               strTemp = Val(Get102TM21TM22("TM21"))
               'end 2018/9/26
               'end 2014/10/31
               If DBDATE(textTM21) <> strTemp Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "延展後專用期限起日應為<" & strTemp & ">"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM21_GotFocus
                  GoTo EXITSUB
               End If
            End If
        '其他為民國日期
        Else
            If CheckIsTaiwanDate(textTM21, False) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "請輸入正確的延展後專用期限起日"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM21_GotFocus
               GoTo EXITSUB
            End If
            If IsEmptyText(m_TM22) = False Then
                'Modify By Cheng 2003/09/01
'               strTemp = DBDATE(DateSerial(Val(DBYEAR(m_TM22)), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22)) + 1))
               'modify by sonia 2014/10/31
               'strTemp = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22))))
               'modify by sonia 2018/9/26 改用共用
               'If m_TM01 = "TF" Then
               '   strTemp = DBDATE(m_TM22)
               'Else
               '   strTemp = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22))))
               'End If
               strTemp = Val(Get102TM21TM22("TM21"))
               'end 2018/9/26
               'end 2014/10/31
               If DBDATE(textTM21) <> strTemp Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "延展後專用期限起日應為<" & ACDate(strTemp) & ">"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM21_GotFocus
                  GoTo EXITSUB
               End If
            End If
        End If
      
   End If
EXITSUB:
End Sub

' 專用期限止日
Private Sub textTM22_Validate(Cancel As Boolean)
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTM22) = False Then
        'Modify By Cheng 2003/08/18
        '若為大陸延展案時, 專用期限輸西元日期
        If m_TM10 <> 台灣國家代號 And m_CP10 = "102" Then
            If CheckIsDate(textTM22, False) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "請輸入正確的延展後專用期限止日"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM22_GotFocus
               GoTo EXITSUB
            End If
            If IsEmptyText(m_TM22) = False And IsEmptyText(m_NA14) = False Then
                'Modify By Cheng 2003/09/01
'               strTemp = DBDATE(DateSerial(Val(DBYEAR(m_TM22)) + Val(m_NA14), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22))))
               'edit by nickc 2006/03/08 遇到 2/28 時，檢查 2/29 有的話已 2/29 為準
               'strTemp = DBDATE(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
               'modify by sonia 2018/9/26 改用共用
               'If Mid(ChangeWDateStringToWString(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22)))), 5) = "0228" Then
               '     If Mid(ChangeWDateStringToWString(DateAdd("d", 1, DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))), 5) = "0229" Then
               '         strTemp = DBDATE(DateAdd("d", 1, DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22)))))
               '     Else
               '         strTemp = DBDATE(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
               '     End If
               'Else
               '     strTemp = DBDATE(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
               'End If
               strTemp = Val(Get102TM21TM22("TM22"))
               'end 2018/9/26
               If DBDATE(textTM22) <> strTemp Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "延展後專用期限止日應為<" & strTemp & ">"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM22_GotFocus
                  GoTo EXITSUB
               End If
            End If
        '其他為民國日期
        Else
            If CheckIsTaiwanDate(textTM22, False) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "請輸入正確的延展後專用期限止日"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM22_GotFocus
               GoTo EXITSUB
            End If
            If IsEmptyText(m_TM22) = False And IsEmptyText(m_NA14) = False Then
                'Modify By Cheng 2003/09/01
'               strTemp = DBDATE(DateSerial(Val(DBYEAR(m_TM22)) + Val(m_NA14), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22))))
               'edit by nickc 2006/03/08 遇到 2/28 時，檢查 2/29 有的話已 2/29 為準
               'strTemp = DBDATE(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
               'modify by sonia 2018/9/26 改用共用
               'If Mid(ChangeWDateStringToWString(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22)))), 5) = "0228" Then
               '     If Mid(ChangeWDateStringToWString(DateAdd("d", 1, DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))), 5) = "0229" Then
               '         strTemp = DBDATE(DateAdd("d", 1, DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22)))))
               '     Else
               '         strTemp = DBDATE(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
               '     End If
               'Else
               '     strTemp = DBDATE(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
               'End If
               strTemp = Val(Get102TM21TM22("TM22"))
               'end 2018/9/26
               If DBDATE(textTM22) <> strTemp Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "延展後專用期限止日應為<" & ACDate(strTemp) & ">"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTM22_GotFocus
                  GoTo EXITSUB
               End If
            End If
        End If
   End If
EXITSUB:
End Sub

Private Sub textTM23_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 申請人
Private Sub textTM23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim bolShowAddr As Boolean 'Add by Amy 2018/10/30 是否彈修改地址畫面
   
   Cancel = False
   textTM23_2 = Empty
   If IsEmptyText(textTM23) = False Then
        'Add By Cheng 2004/04/20
        '申請人代號補滿9碼
        Me.textTM23.Text = ChangeCustomerL(Me.textTM23.Text)
        'End
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
      'Modify by Amy 2018/10/30
      'If Me.textTM23.Text <> m_strCust1 Then
      If Me.textTM23.Text <> Me.textTM23.Tag Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then
            Cancel = True
         Else
            bolShowAddr = True
         End If
      End If
   End If
   If Cancel = True Then textTM23_GotFocus
   'Add by Amy 2018/10/30
   If bolShowAddr = True Then
        Me.textTM23.Tag = Me.textTM23 '記錄每次修改,改回仍需存基本檔
        frm020102_23.Hide
        Set frm020102_23.UpForm = Me
        frm020102_23.m_CP09 = m_CP09
        frm020102_23.stModApply = "1;" & Me.textTM23
        frm020102_23.QueryData
        frm020102_23.Show vbModal
   End If
End Sub

' 商品組群
Private Sub textTM32_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String 'Integer
   
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textTM32) = True Then
      GoTo EXITSUB
   End If
   
   'Modify By Sindy 2024/4/18 商品組群欄人員貼上資料後將全形或半形的「；」分號，轉為半形的逗號存入TM32。
   textTM32 = Replace(Replace(textTM32, ";", ","), "；", ",")
   '2024/4/18 END
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

Private Sub textTM72_GotFocus()
    TextInverse Me.textTM72
End Sub
'Added by Lydia 2023/11/14
Private Sub textTM72_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM72_Validate(Cancel As Boolean)
    If Me.textTM72.Text <> "" Then
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            Me.textTM72_2.Text = PUB_GetSpecialPTName("2", Me.textTM72.Text)
            If Me.textTM72_2.Text = "" Then
                MsgBox "特殊商標代碼輸入錯誤!!!", vbExclamation + vbOKOnly
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
'add by nickc 2007/01/02
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
   Dim bolShowAddr As Boolean 'Add by Amy 2018/10/30 是否彈修改地址畫面
   
   Cancel = False
   textTM78_2 = Empty
   If IsEmptyText(textTM78) = False Then
     '申請人代號補滿9碼
      Me.textTM78.Text = ChangeCustomerL(Me.textTM78.Text)
      Dim oState As Boolean
      oState = True
      textTM78_2 = GetCustomerNameAndState(textTM78, 0, oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textTM78_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM78 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM78_GotFocus
      End If
   End If
   If Cancel = False Then
      'Modif by Amy 2018/10/30
      'If Me.textTM78.Text <> m_strCust2 Then
      If Me.textTM78.Text <> Me.textTM78.Tag Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then
            Cancel = True
         Else
            bolShowAddr = True
         End If
      End If
   End If
   If Cancel = True Then textTM78_GotFocus
   'Add by Amy 2018/10/30
   If bolShowAddr = True Then
        Me.textTM78.Tag = Me.textTM78 '記錄每次修改,改回仍需存基本檔
        frm020102_23.Hide
        Set frm020102_23.UpForm = Me
        frm020102_23.m_CP09 = m_CP09
        frm020102_23.stModApply = "2;" & Me.textTM78
        frm020102_23.QueryData
        frm020102_23.Show vbModal
   End If
End Sub

'add by nickc 2007/01/02
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
   Dim bolShowAddr As Boolean 'Add by Amy 2018/10/30 是否彈修改地址畫面
   
   Cancel = False
   textTM79_2 = Empty
   If IsEmptyText(textTM79) = False Then
     '申請人代號補滿9碼
      Me.textTM79.Text = ChangeCustomerL(Me.textTM79.Text)
      Dim oState As Boolean
      oState = True
      textTM79_2 = GetCustomerNameAndState(textTM79, 0, oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textTM79_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM79 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM79_GotFocus
      End If
   End If
   If Cancel = False Then
      'Modify by Amy 2018/10/30
      'If Me.textTM79.Text <> m_strCust3 Then
      If Me.textTM79.Text <> Me.textTM79.Tag Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then
            Cancel = True
         Else
            bolShowAddr = True
         End If
      End If
      'end 2018/10/30
   End If
   If Cancel = True Then textTM79_GotFocus
   'Add by Amy 2018/10/30
   If bolShowAddr = True Then
        Me.textTM79.Tag = Me.textTM79 '記錄每次修改,改回仍需存基本檔
        frm020102_23.Hide
        Set frm020102_23.UpForm = Me
        frm020102_23.m_CP09 = m_CP09
        frm020102_23.stModApply = "3;" & Me.textTM79
        frm020102_23.QueryData
        frm020102_23.Show vbModal
   End If
End Sub

'add by nickc 2007/01/02
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
   Dim bolShowAddr As Boolean 'Add by Amy 2018/10/30 是否彈修改地址畫面
   
   Cancel = False
   textTM80_2 = Empty
   If IsEmptyText(textTM80) = False Then
     '申請人代號補滿9碼
      Me.textTM80.Text = ChangeCustomerL(Me.textTM80.Text)
      Dim oState As Boolean
      oState = True
      textTM80_2 = GetCustomerNameAndState(textTM80, 0, oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textTM80_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM80 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM80_GotFocus
      End If
   End If
   If Cancel = False Then
      'Modify by Amy 018/10/29
      'If Me.textTM80.Text <> m_strCust4 Then
      If Me.textTM80.Text <> Me.textTM80.Tag Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then
            Cancel = True
         Else
            bolShowAddr = True
         End If
      End If
      'end 201810/29
   End If
   If Cancel = True Then textTM80_GotFocus
   'Add by Amy 2018/10/30
   If bolShowAddr = True Then
        Me.textTM80.Tag = Me.textTM80 '記錄每次修改,改回仍需存基本檔
        frm020102_23.Hide
        Set frm020102_23.UpForm = Me
        frm020102_23.m_CP09 = m_CP09
        frm020102_23.stModApply = "4;" & Me.textTM80
        frm020102_23.QueryData
        frm020102_23.Show vbModal
   End If
End Sub

'add by nickc 2007/01/02
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
   Dim bolShowAddr As Boolean 'Add by Amy 2018/10/30 是否彈修改地址畫面
   
   Cancel = False
   textTM81_2 = Empty
   If IsEmptyText(textTM81) = False Then
     '申請人代號補滿9碼
      Me.textTM81.Text = ChangeCustomerL(Me.textTM81.Text)
      Dim oState As Boolean
      oState = True
      textTM81_2 = GetCustomerNameAndState(textTM81, 0, oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textTM81_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textTM81 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM81_GotFocus
      End If
   End If
   If Cancel = False Then
      'Modify by Amy 2018/10/30
      'If Me.textTM81.Text <> m_strCust5 Then
      If Me.textTM81.Text <> Me.textTM81.Tag Then
         If Not PUB_EditCustOk(m_CP09, m_TM01, m_TM02, m_TM03, m_TM04) Then
            Cancel = True
         Else
            bolShowAddr = True
         End If
      End If
   End If
   If Cancel = True Then textTM81_GotFocus
   'Add by Amy 2018/10/30
   If bolShowAddr = True Then
        Me.textTM81.Tag = Me.textTM81 '記錄每次修改,改回仍需存基本檔
        frm020102_23.Hide
        Set frm020102_23.UpForm = Me
        frm020102_23.m_CP09 = m_CP09
        frm020102_23.stModApply = "5;" & Me.textTM81
        frm020102_23.QueryData
        frm020102_23.Show vbModal
   End If
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
   Else
        'Add by Amy 2015/05/26 +T大陸案續展催審期限控管
        'Modify by Amy 2016/03/10 +TF馬德里商標續展同 T 大陸案
        If m_CP10 = "102" And (m_TM10 = "020" Or m_TM10 = "238") Then
            textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
            Cancel = Not (SetUargeDate_102)
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
   'Add by Amy 2021/12/23檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
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
   
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF", "TS"
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
    '若代理人非台灣而申請國家為台灣時
'2010/9/16 MODIFY BY SONIA 改為國外部收文之查名可不輸申請人,有FC代理人即可
'    If m_FA10 > "000" And m_TM10 = "000" Then
'        '可只輸入申請人或代理人
'        If Me.textCP44.Text = "" And Me.textTM23.Text = "" Then
'            MsgBox "請輸入代理人或申請人 !!!", vbExclamation + vbOKOnly
'            Me.textCP44.SetFocus
'           GoTo EXITSUB
'        End If
   If m_FA10 > "010" And Mid(m_CP12, 1, 1) = "F" And m_CP10 = "001" Then
   '2010/9/16 END
   Else
'         ' 申請人
'         If IsEmptyText(textTM23) = True Then
'            strTit = "檢核資料"
'            strMsg = "請輸入申請人"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTM23.SetFocus
'            GoTo EXITSUB
'         End If
      'Modify By Sindy 2011/01/06
      '內商(TS)申請人1或FC代理人至少要輸入一個
      '其他的一定要輸入申請人1
      If m_TM01 = "TS" Then
           If textTM23 = "" And m_TM44 = "" Then
               MsgBox "申請人1或FC代理人至少要輸入一個!!!", vbExclamation + vbOKOnly
               Me.textTM23.SetFocus
               textTM23_GotFocus
               GoTo EXITSUB
           End If
      Else
           If textTM23 = "" Then
               MsgBox "申請人1不可空白!!!", vbExclamation + vbOKOnly
               Me.textTM23.SetFocus
               textTM23_GotFocus
               GoTo EXITSUB
           End If
      End If
   End If
    
   ' 代理人(申請國家非台灣時不可空白)
   If m_TM10 >= "010" Then
      If IsEmptyText(textCP44) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入代理人"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 專用期限
   If m_CP10 = "102" Then
      If IsEmptyText(textTM21) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入專用期限起日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM21.SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(textTM22) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入專用期限止日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM22.SetFocus
         GoTo EXITSUB
      End If
      'Add By Sindy 2012/5/3
      '延展後專用期止日不可小於等於基本檔專用期止日
      If Val(DBDATE(textTM22)) <= Val(DBDATE(m_TM22)) Then
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
      '2012/5/3 End
   End If
   
   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         ' 商品類別
         'Modify By Cheng 2002/07/16
         '若案件性質為"申請"(101), 且商標種類<'7'時, 商品種類不可空白
         If (m_CP10 = "101" Or m_CP10 = "308") And Me.textTM08.Text < "7" Then
            If IsEmptyText(textTM09) = True Then
               strTit = "檢核資料"
               strMsg = "請輸入商品類別"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM09.SetFocus
               GoTo EXITSUB
            End If
         End If
         ' 申請國家為大陸時可以不輸入商品組群
         'If m_TM10 <> "020" Then
         '   ' 商品組群
         '   If IsEmptyText(textTM32) = True Then
         '      strTit = "檢核資料"
         '      strMsg = "請輸入商品組群"
         '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '      GoTo ExitSub
         '   End If
         'End If
        'Modify By Cheng 2002/12/02
        '若商標種類為團體標章及證明標章可不輸入商品組群
        If Me.textTM08.Text <> "8" And Me.textTM08.Text <> "7" Then
            ' 90.06.21
            ' 只有台灣且案件性質為101的才一定要輸入商品組群
            If m_TM10 < "010" And (m_CP10 = "101" Or m_CP10 = "308") Then
               If IsEmptyText(textTM32) = True Then
                  strTit = "檢核資料"
                  strMsg = "請輸入商品組群"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  'add by nickc 2006/10/18
                  Me.SSTab1.Tab = 0
                  
                  textTM32.SetFocus
                  GoTo EXITSUB
               End If
            End If
        End If
         ' 商標種類不可空白
         If IsEmptyText(textTM08) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入商標種類"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            'Modified by Lydia 2023/11/16
            'textTM08.SetFocus
            cboTM08.SetFocus
            GoTo EXITSUB
         End If
         ' 商標種類為聯合商標, 防護商標, 聯合服務標章, 防護服務標章時正商標號數不可空白
         If IsEmptyText(textTM27) = True Then
            Select Case textTM08
               Case "2", "3", "5", "6":
                  strTit = "檢核資料"
                  strMsg = "請輸入正商標號數"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  'Modified by Lydia 2023/11/16
                  'textTM08.SetFocus
                  cboTM08.SetFocus
                  GoTo EXITSUB
            End Select
         End If
      Case Else:
   End Select
   
   ' 案件性質為申請, 申請國家為台灣, 是否郵寄申請為空白時, 申請日及申請案號一定要輸入
   If (m_CP10 = "101" Or m_CP10 = "308") And m_TM10 < "010" And IsEmptyText(textMail) = True Then
      ' 申請日
      If IsEmptyText(textTM11) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入申請日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM11.SetFocus
         GoTo EXITSUB
      End If
      ' 申請案號
      If IsEmptyText(textTM12) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入申請案號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM12.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 案件性質為刊登廣告時
   If m_CP10 = "702" And m_TM10 = "020" Then
      If IsEmptyText(textMediaName) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入雜誌社, 報社"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textMediaName.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 91.09.02 modyfy by louis
   ' 檢查本所案號輸入是否完整
   If textTM01 <> Empty Then
      ' 第二欄為可為空白
      If textTM02 = Empty Then
         strTit = "檢核資料"
         strMsg = "本所案號輸入不完整"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM01.SetFocus
         GoTo EXITSUB
      End If
      ' 如果系統類別為TF時第三碼一定要輸入
      If textTM01 = "TF" Then
         If textTM02_2 = Empty Then
            strTit = "檢核資料"
            strMsg = "本所案號輸入不完整"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM01.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
    'Add By Cheng 2003/06/16
    '若有輸入查名本所案號
    If Me.textTM01.Text <> "" And Me.textTM02.Text <> "" Then
        StrSQLa = "Select * From ServicePractice Where " & ChgService(Me.textTM01.Text & Me.textTM02.Text & Me.textTM02_2.Text & Left(Me.textTM03.Text & "0", 1) & Left(Me.textTM04.Text & "00", 2))
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount <= 0 Then
            MsgBox "您輸入的查名本所案號錯誤，請重新輸入!!!", vbExclamation + vbOKOnly
            Me.textTM01.SetFocus
            textTM01_GotFocus
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            GoTo EXITSUB
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
   
   'Add By Sindy 2012/7/30
   '請加入 申請101案件性質 發文時,若該案號有收文主張優先權或有輸入優先權資料時,若基本資料2的補優先權文件期限沒有輸入時,顯示訊息提醒 "是否要管制優先權文件期限?",但仍可選擇不管制
   If IsEmptyText(textPriorityDate) = True And m_CP10 = "101" Then '申請案且無輸入補優先權文件期限
      CheckOC3
      strSql = "Select cp09 From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10='108'"
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      '有收文主張優先權或有輸入優先權資料
      If AdoRecordSet3.RecordCount > 0 Or m_Priority(1) <> "" Then
         '顯示訊息提醒
         If MsgBox("是否要管制優先權文件期限？", vbInformation + vbYesNo) = vbYes Then
            '2013/10/11 add by sonia 以發文日起算三個月為法定期限（T-189043）
            textPriorityDate = DBDATE(DateAdd("m", 3, ChangeWStringToWDateString(DBDATE(Me.textCP27.Text))))
            SSTab1.Tab = 1
            '2013/10/11 end
            textPriorityDate.SetFocus
            textPriorityDate_GotFocus
            CheckOC3
            GoTo EXITSUB
         End If
      End If
      CheckOC3
   End If
   '2012/7/30 End
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textUargeDate_GotFocus()
   InverseTextBox textUargeDate
End Sub

Private Sub textPetition_GotFocus()
   InverseTextBox textPetition
End Sub

Private Sub textPriorityDate_GotFocus()
   InverseTextBox textPriorityDate
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textMail_GotFocus()
   InverseTextBox textMail
End Sub

Private Sub textMediaType_GotFocus()
   InverseTextBox textMediaType
End Sub

Private Sub textMediaDate_GotFocus()
   InverseTextBox textMediaDate
End Sub

Private Sub textMediaName_GotFocus()
   InverseTextBox textMediaName
End Sub

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
End Sub

' 91.09.02 marked by louis
'Private Sub textCP09S_GotFocus()
'   InverseTextBox textCP09S
'End Sub

Private Sub textCP22_GotFocus()
   InverseTextBox textCP22
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

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
End Sub

Private Sub textTM05_GotFocus()
   InverseTextBox textTM05
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM05.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM06_GotFocus()
   InverseTextBox textTM06
End Sub

Private Sub textTM07_GotFocus()
   InverseTextBox textTM07
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTM07.IMEMode = 1
   OpenIme
End Sub

Private Sub textTM08_GotFocus()
   InverseTextBox textTM08
End Sub

Private Sub textTM09_GotFocus()
   InverseTextBox textTM09
End Sub

Private Sub textTM11_GotFocus()
   InverseTextBox textTM11
End Sub

Private Sub textTM12_GotFocus()
   InverseTextBox textTM12
End Sub

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

Private Sub textTM58_GotFocus()
   InverseTextBox textTM58
End Sub

Private Sub textTM67_GotFocus()
   InverseTextBox textTM67
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
Private Sub InsExpField()
Dim strTM23Nation As String
Dim strSql As String
'Add By Sindy 2010/11/12
Dim IsMaCase As Boolean
Dim arrTM09 As Variant, strGoodsKind As String
'2010/11/12 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   Select Case m_CP10
      ' 查名
      Case "001":
         ' 申請國家
         Select Case m_TM10
            ' 大陸
            Case "020":
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "00", strUserNum
                    ' 案件性質分類
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & _
                          "','案件性質分類','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
                    cnnConnection.Execute strSql
                    ' 回音
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & _
                          "','回音','" & IIf(textCF09 <> "", "大約" & textCF09 & "後可接獲回音。", "") & "')"
                    '      "','回音','" & textCF09 & "')"
                    cnnConnection.Execute strSql
                End If
         End Select
      ' 申請
      Case "101":
         ' 申請國家
         Select Case m_TM10
            ' 大陸
            Case "020":
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "01", m_CP09, "00", strUserNum
                   ' 案件性質分類
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & _
                         "','案件性質分類','" & "註冊申請" & "')"
                   cnnConnection.Execute strSql
                   ' 回音
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & _
                         "','回音','" & IIf(textCF09 <> "", "大約" & textCF09 & "後可接獲回音。", "") & "')"
    '                     "','回音','" & textCF09 & "')"
                   cnnConnection.Execute strSql
                End If
            ' 馬德里
            Case "238":
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "01", strUserNum
                    ' 案件性質分類
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                          "','案件性質分類','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
                    cnnConnection.Execute strSql
                    ' 回音
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                          "','回音','" & textCF09 & "')"
                    cnnConnection.Execute strSql
                 End If
         End Select
         'Add By Sindy 2012/1/12 原本在 InsExpField1 函數裡移出來
         If Me.textMail.Text = "" And m_TM10 < "010" Then
            Select Case m_TM01
               Case "TC":
                 'add by nickc 2006/06/30
                 If textPrint = "1" Then
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "02", m_CP09, "04", strUserNum
                 End If
               Case "T", "TF":
                  If m_TM10 < "010" Then
                     ' 申請人國籍為台灣
                     'edit by nickc 2006/06/30
                     'If strTM23Nation < "010" Then
                     If textPrint = "1" Then
                        ' 清除定稿例外欄位檔原有資料
                        EndLetter "02", m_CP09, "01", strUserNum
                     ' 申請人國籍非台灣
                     'edit by nickc 2006/06/30
                     'Else
                     ElseIf textPrint = "2" Then
                        ' 清除定稿例外欄位檔原有資料
                        EndLetter "02", m_CP09, "02", strUserNum
                        cnnConnection.Execute strSql
                        ' 回音
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "02" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
                                 "'" & "回音" & "','" & textCF09 & "')"
                        cnnConnection.Execute strSql
                     End If
                  End If
            End Select
         End If
      ' 領土延申
      Case "104":
         ' 申請國家
         Select Case m_TM10
            ' 馬德里
            Case "238":
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "01", strUserNum
                    ' 案件性質分類
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                          "','案件性質分類','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
                    cnnConnection.Execute strSql
                    ' 回音
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                          "','回音','" & textCF09 & "')"
                    cnnConnection.Execute strSql
                End If
         End Select
      ' 延展
      'Modify By Sindy 2009/06/16 增加109.被異議續展
      'Case "102":
      Case "102", "109":
         'add by nickc 2006/06/02
         If m_TM01 = "TF" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                EndLetter "01", m_CP09, "01", strUserNum
                Dim otmpCountry As String
                'edit by nickc 2007/02/15 改格式，有請作單
'                strSQL = "select * from nation where na01 in (" & strLicenceCountry & ")  "
'                CheckOC3
'                AdoRecordSet3.CursorLocation = adUseClient
'                AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'
'                If AdoRecordSet3.RecordCount <> 0 Then
'                    AdoRecordSet3.MoveFirst
'                    Do While Not AdoRecordSet3.EOF
'                        If otmpCountry <> "" Then
'                            otmpCountry = otmpCountry & "、"
'                        End If
'                        otmpCountry = otmpCountry & CheckStr(AdoRecordSet3.Fields("na03"))
'                        AdoRecordSet3.MoveNext
'                    Loop
'                End If
                Dim otmpTm09 As Variant
                Dim oII As Integer
                    otmpCountry = ""
                    otmpTm09 = Split(textTM09, ",")
                    For oII = 0 To UBound(otmpTm09)
                        '2009/4/20 MODIFY BY SONIA 領土延伸案的子案也要抓TF-000420
                        'strSQL = "select distinct tm03 ,na03 from nation,trademark,caseprogress where tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and tm10=na01(+) and tm04<>'00' and tm03='" & Trim(oII + 1) & "' and tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm16='1' and tm29 is null   order by tm03 "
                        strSql = "select distinct tm03 ,na03 from nation,trademark,caseprogress where tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and tm10=na01(+) and tm04<>'00' and tm03='" & Trim(oII + 1) & "' and tm01='" & m_TM01 & "' and substr(tm02,1,5)='" & Mid(m_TM02, 1, 5) & "' and tm16='1' and tm29 is null   order by tm03 "
                        CheckOC3
                        AdoRecordSet3.CursorLocation = adUseClient
                        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                        If AdoRecordSet3.RecordCount <> 0 Then
                            otmpCountry = otmpCountry & "　　第 " & otmpTm09(oII) & "類："
                            AdoRecordSet3.MoveFirst
                            Do While Not AdoRecordSet3.EOF
                                If AdoRecordSet3.AbsolutePosition > 1 Then
                                    otmpCountry = otmpCountry & "、"
                                End If
                                otmpCountry = otmpCountry & CheckStr(AdoRecordSet3.Fields("na03"))
                                AdoRecordSet3.MoveNext
                            Loop
                            otmpCountry = otmpCountry & "。" & vbCrLf
                        End If
                    Next oII
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                      "','例外國家欄位','" & otmpCountry & "')"
                cnnConnection.Execute strSql
            End If
         Else
            ' 申請國家為台灣
            If m_TM10 < "010" Then
               ' 申請人國籍為台灣
               'edit by nickc 2006/06/30
               'If strTM23Nation < "010" Then
               If textPrint = "1" Then
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "01", m_CP09, "02", strUserNum
                  ' 回音
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                        "','回音','" & IIf(textCF09 <> "", "大約" & textCF09 & "後可接獲回音。", "") & "')"
                  '      "','回音','" & textCF09 & "')"
                  cnnConnection.Execute strSql
               ' 申請人國籍非台灣
               'edit by nickc 2006/06/30
               'Else
               ElseIf textPrint = "2" Then
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "01", m_CP09, "03", strUserNum
               End If
            ' 申請國家為大陸
            ElseIf m_TM10 = "020" Then
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "01", m_CP09, "00", strUserNum
                  ' 案件性質分類
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & _
                        "','案件性質分類','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
                  cnnConnection.Execute strSql
                  ' 回音
                  'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  '      "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & _
                  '      "','回音','" & textCF09 & "')"
                  'cnnConnection.Execute strSQL
                End If
            End If
         End If
      ' 補換發證書
      Case "103":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "02", strUserNum
               ' 回音
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                     "','回音','" & IIf(textCF09 <> "", "大約" & textCF09 & "後可接獲回音。", "") & "')"
               '      "','回音','" & textCF09 & "')"
               cnnConnection.Execute strSql
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            ElseIf textPrint = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "03", strUserNum
            End If
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "00", strUserNum
               ' 案件性質分類
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & _
                     "','案件性質分類','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
               cnnConnection.Execute strSql
               ' 回音
               'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '      "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & _
               '      "','回音','" & textCF09 & "')"
               'cnnConnection.Execute strSQL
            End If
         End If
        'Modify By Cheng 2003/02/21
      ' 申請英文證明, 申請中文證明
'      Case "304":
      Case "304", "309":
         ' 申請國家非大陸
         If m_TM10 <> "020" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "02", strUserNum
               ' 回音
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                     "','回音','" & IIf(textCF09 <> "", "大約" & textCF09 & "後可接獲回音。", "") & "')"
               '      "','回音','" & textCF09 & "')"
               cnnConnection.Execute strSql
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            ElseIf textPrint = "2" Then
               EndLetter "01", m_CP09, "03", strUserNum
            End If
         End If
      ' 刊登廣告
      Case "702":
         ' 申請國家為大陸
         If m_TM10 = "020" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                Select Case textMediaType:
                   ' 雜誌
                   Case "1":
                      ' 清除定稿例外欄位檔原有資料
                      EndLetter "01", m_CP09, "04", strUserNum
                      ' 廣告媒體
                      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               "VALUES ('" & "01" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & _
                               "','廣告媒體','" & textMediaName & "')"
                      cnnConnection.Execute strSql
                   ' 報紙
                   Case "2":
                      ' 清除定稿例外欄位檔原有資料
                      EndLetter "01", m_CP09, "05", strUserNum
                      ' 廣告媒體
                      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               "VALUES ('" & "01" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & _
                               "','廣告媒體','" & textMediaName & "')"
                      cnnConnection.Execute strSql
                End Select
            End If
         End If
      'add by nickc 2006/06/30 補對應
      Case "308":
         'Add By Sindy 2010/11/12
         '是否為母案
         IsMaCase = False
         strSql = "SELECT * FROM DivisionCase WHERE DC05='" & m_TM01 & "' AND DC06='" & m_TM02 & "' AND DC07='" & m_TM03 & "' AND DC08='" & m_TM04 & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            IsMaCase = True
         End If
         '1-34商品 35-45服務
         strGoodsKind = "本案指定商品"
         If Trim(textTM09.Text) > "" Then
            arrTM09 = Split(textTM09.Text, ",")
            If Val(arrTM09(0)) >= 35 And Val(arrTM09(0)) <= 45 Then
               strGoodsKind = "本案指定服務"
            End If
         End If
         '2010/11/12 End
         
         ' 申請國家
         Select Case m_TM10
            ' 台灣
            Case "000":
               If textPrint = "1" Then
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "01", m_CP09, "02", strUserNum
                  'Add By Sindy 2010/11/12
                  If IsMaCase = False Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "01" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                                  "','本案指定商品','" & strGoodsKind & "：|?TMGoods:" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "-?|')"
                     cnnConnection.Execute strSql
                  End If
                  'Add By Sindy 2010/11/12
                  If IsMaCase = True Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "01" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                                  "','收據','２．收據。')"
                     cnnConnection.Execute strSql
                  End If
                  
               'Add By Sindy 2010/10/28
               ElseIf textPrint = "2" Then '大->台
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "01", m_CP09, "03", strUserNum
                  'Add By Sindy 2010/11/12
                  If IsMaCase = False Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "01" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                                  "','本案指定商品','" & strGoodsKind & "：|?TMGoods:" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "-?|')"
                     cnnConnection.Execute strSql
                  End If
                  'Add By Sindy 2010/11/12
                  If IsMaCase = True Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "01" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                                  "','收據','２．收據。')"
                     cnnConnection.Execute strSql
                  End If
               End If
            'Add by Amy 2014/10/16 T大陸分割案控制
            Case "020"
                If m_TM01 = "T" Then
                    '母案
                    EndLetter "01", m_CP09, "01", strUserNum
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                                  "','母案申請案號','" & textTM12 & "')"
                    cnnConnection.Execute strSql
                    '子案
                    EndLetter "01", m_CP09s, "01", strUserNum
                    'Modify by Amy 2015/04/15 申請案號改為母案申請號+A
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                  "VALUES ('" & "01" & "','" & m_CP09s & "','" & "01" & "','" & strUserNum & _
                                  "','母案申請案號','" & textTM12 & "A')"
                    cnnConnection.Execute strSql
                End If
         End Select
   End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strTM23Nation As String
'92.1.25 ADD BY SONIA
Dim bolEdit As Boolean
'92.1.25 END
'Add By Sindy 2012/1/12
Dim ET01 As String, ET02 As String, ET03 As String, ET01_1 As String, ET03_1 As String
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/12 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/12
   ET01 = "01"
   ET01_1 = "02"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/12 End
   
   Select Case m_CP10
      ' 查名
      Case "001":
         ' 申請國家
         Select Case m_TM10
            ' 大陸
            Case "020":
               ' 列印定稿
               'add by nickc 2006/06/30
               If textPrint = "1" Then
'                    NowPrint m_CP09, "01", "00", False, strUserNum, 0
                  ET03 = "00" 'Modify By Sindy 2012/1/12
               End If
         End Select
      ' 申請
      Case "101":
         ' 申請國家
         Select Case m_TM10
            ' 大陸
            Case "020":
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 列印定稿
'                    NowPrint m_CP09, "01", "00", False, strUserNum, 0
                  ET03 = "00" 'Modify By Sindy 2012/1/12
                End If
            ' 馬德里
            Case "238":
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 列印定稿
'                    NowPrint m_CP09, "01", "01", False, strUserNum, 0
                  ET03 = "01" 'Modify By Sindy 2012/1/12
                End If
         End Select
         'Add By Cheng 2002/06/14
         '當案件質為"申請"時, 申請國家台灣, 且"是否郵寄申請"欄為NULL時, 列印定稿
'         If Me.textMail.Text = "" And m_TM10 < "010" Then PrintLetter1
         'Modify By Sindy 2012/1/12 原本在 PrintLetter1 函數裡移出來
         If Me.textMail.Text = "" And m_TM10 < "010" Then
            Select Case m_TM01
               Case "TC":
                 'add by nickc 2006/06/30
                 If textPrint = "1" Then
                     ' 列印定稿
'                     NowPrint m_CP09, "02", "04", False, strUserNum, 0
                     ET03_1 = "04" 'Modify By Sindy 2012/1/12
                 End If
               Case "T", "TF":
                   'Modify By Cheng 2002/06/12
                  If m_TM10 < "010" Then
                     ' 申請人國籍為台灣
                     'edit by nickc 2006/06/30
                     'If strTM23Nation < "010" Then
                     If textPrint = "1" Then
                        'Modify By Cheng 2002/06/12
         '                  ' 列印定稿
         '                  NowPrint m_CP09, "02", "01", False, strUserNum, 0
                        ' 列印定稿
                        Select Case m_TM08
                        Case "7", "8" '證明標章, 團體標章
'                           NowPrint m_CP09, "02", "05", False, strUserNum, 0
                           ET03_1 = "05" 'Modify By Sindy 2012/1/12
                        Case Else
'                           NowPrint m_CP09, "02", "01", False, strUserNum, 0
                           ET03_1 = "01" 'Modify By Sindy 2012/1/12
                        End Select
                     ' 申請人國籍非台灣
                     'edit by nickc 2006/06/30
                     'Else
                     ElseIf textPrint = "2" Then
                        ' 列印定稿
'                        NowPrint m_CP09, "02", "02", False, strUserNum, 0
                        ET03_1 = "02" 'Modify By Sindy 2012/1/12
                     End If
                  End If
            End Select
         End If
      ' 領土延申
      Case "104":
         ' 申請國家
         Select Case m_TM10
            ' 馬德里
            Case "238":
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 列印定稿
'                    NowPrint m_CP09, "01", "01", False, strUserNum, 0
                  ET03 = "01" 'Modify By Sindy 2012/1/12
                End If
         End Select
      ' 延展
      'Modify By Sindy 2009/06/16 增加109.被異議續展
      'Case "102":
      Case "102", "109":
         ' 申請國家為台灣
         'add by nickc 2006/06/02
         If m_TM01 = "TF" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
'                NowPrint m_CP09, "01", "01", False, strUserNum, 0
               ET03 = "01" 'Modify By Sindy 2012/1/12
            End If
         Else
            If m_TM10 < "010" Then
               ' 申請人國籍為台灣
               'edit by nickc 2006/06/30
               'If strTM23Nation < "010" Then
               If textPrint = "1" Then
                  ' 列印定稿
'                  NowPrint m_CP09, "01", "02", False, strUserNum, 0
                  ET03 = "02" 'Modify By Sindy 2012/1/12
               ' 申請人國籍非台灣
               'edit by nickc 2006/06/30
               'Else
               ElseIf textPrint = "2" Then
                  ' 列印定稿
'                  NowPrint m_CP09, "01", "03", False, strUserNum, 0
                  ET03 = "03" 'Modify By Sindy 2012/1/12
               'add by sonia 2014/6/4
               ElseIf textPrint = "3" Then
                  ET03 = "04"
               'end 2014/6/4
               End If
            ' 申請國家為大陸
            ElseIf m_TM10 = "020" Then
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 列印定稿
'                    NowPrint m_CP09, "01", "00", False, strUserNum, 0
                  ET03 = "00" 'Modify By Sindy 2012/1/12
                End If
            End If
        End If
      ' 補換發證書
      Case "103":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 列印定稿
'               NowPrint m_CP09, "01", "02", False, strUserNum, 0
               ET03 = "02" 'Modify By Sindy 2012/1/12
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            ElseIf textPrint = "2" Then
               ' 列印定稿
'               NowPrint m_CP09, "01", "03", False, strUserNum, 0
               ET03 = "03" 'Modify By Sindy 2012/1/12
            End If
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                ' 列印定稿
'                NowPrint m_CP09, "01", "00", False, strUserNum, 0
               ET03 = "00" 'Modify By Sindy 2012/1/12
            End If
         End If
      ' 申請英文證明, 申請中文證明
'      Case "304":
      Case "304", "309":
         ' 申請國家非大陸
         If m_TM10 <> "020" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 列印定稿
'               NowPrint m_CP09, "01", "02", False, strUserNum, 0
               ET03 = "02" 'Modify By Sindy 2012/1/12
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            ElseIf textPrint = "2" Then
               ' 列印定稿
'               NowPrint m_CP09, "01", "03", False, strUserNum, 0
               ET03 = "03" 'Modify By Sindy 2012/1/12
            End If
         End If
      ' 刊登廣告
      Case "702":
         ' 申請國家為大陸
         If m_TM10 = "020" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
                '92.1.25 ADD BY SONIA
                If grdList.row > 1 Then
                   bolEdit = True
                Else
                   bolEdit = False
                End If
                '92.1.25 END
                Select Case textMediaType:
                   ' 雜誌
                   Case "1":
                      ' 列印定稿
'                      NowPrint m_CP09, "01", "04", bolEdit, strUserNum, 0
                     ET03 = "04" 'Modify By Sindy 2012/1/12
                   ' 報紙
                   Case "2":
                      ' 列印定稿
'                      NowPrint m_CP09, "01", "05", bolEdit, strUserNum, 0
                     ET03 = "05" 'Modify By Sindy 2012/1/12
                End Select
            End If
         End If
      '93.9.30 ADD BY SONIA
      ' 分割
      Case "308":
         ' 申請國家
         Select Case m_TM10
            ' 台灣
            Case "000":
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 列印定稿
                    'edit by nick 2004/10/12
                    'NowPrint m_CP09, "01", "01", False, strUserNum, 0
'                    NowPrint m_CP09, "01", "02", False, strUserNum, 0
                  ET03 = "02" 'Modify By Sindy 2012/1/12
                'Add By Sindy 2010/10/28
                ElseIf textPrint = "2" Then '大->台
'                    NowPrint m_CP09, "01", "03", False, strUserNum, 0
                  ET03 = "03" 'Modify By Sindy 2012/1/12
                End If
            'Add by Amy 2014/10/16 T大陸分割案
            Case "020"
                If textPrint = "1" Then
                    ET03 = "01"
                End If
         End Select
   End Select
   
   'Add By Sindy 2012/1/12
   If ET03 <> "" Or ET03_1 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, m_CP10 = "102", , bolPlusPaper)
      '判斷是否EMail同時寄紙本
      If Not bolPlusPaper Then
         iCopy = 1
      End If
      If bolEmail Then
         'Add By Sindy 2020/1/7 + 信函總收文號
         If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
            If ET03 <> "" Then
               NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
            End If
            If ET03_1 <> "" Then
               NowPrint ET02, ET01_1, ET03_1, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
            End If
         Else
         '2020/1/7 END
            If ET03 <> "" Then
               NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            End If
            If ET03_1 <> "" Then
               NowPrint ET02, ET01_1, ET03_1, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            End If
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
         End If
      Else
         If ET03 <> "" Then
            'Modify by Amy 2014/10/16 +大陸分割案定稿
            If m_TM01 = "T" And m_TM10 = 大陸國家代號 And m_CP10 = "308" Then
               'Add By Sindy 2019/12/20 + strLD18.信函總收文號
               NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18 '母案
               'Add By Sindy 2019/12/20 + strLD18.信函總收文號
               'NowPrint m_CP09s, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18 '子案
               'Modify By Sindy 2021/1/11
               PUB_AddLetterProgress m_CP09s, 0, True, "", False, m_TM23, m_CP10, m_TM44
               '2021/1/11 END
               NowPrint m_CP09s, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , m_CP09s '子案
            Else
               'Add By Sindy 2019/12/20 + strLD18.信函總收文號
               NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
            End If
            'end 2014/10/16
         End If
         If ET03_1 <> "" Then
            'Add By Sindy 2019/12/20 + strLD18.信函總收文號
            NowPrint ET02, ET01_1, ET03_1, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
         End If
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

''Add By Cheng 2002/06/14
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 列印定稿
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Sub PrintLetter1()
'   Dim strTM23Nation As String
'   strTM23Nation = Empty
'   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
'
'   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
'   InsExpField1
'
'   ' 系統類別TD
'   Select Case m_TM01
'      Case "TC":
'        'add by nickc 2006/06/30
'        If textPrint = "1" Then
'            ' 列印定稿
'            NowPrint m_CP09, "02", "04", False, strUserNum, 0
'        End If
'      Case "T", "TF":
'          'Modify By Cheng 2002/06/12
'         If m_TM10 < "010" Then
'            ' 申請人國籍為台灣
'            'edit by nickc 2006/06/30
'            'If strTM23Nation < "010" Then
'            If textPrint = "1" Then
'               'Modify By Cheng 2002/06/12
''                  ' 列印定稿
''                  NowPrint m_CP09, "02", "01", False, strUserNum, 0
'               ' 列印定稿
'               Select Case m_TM08
'               Case "7", "8" '證明標章, 團體標章
'                  NowPrint m_CP09, "02", "05", False, strUserNum, 0
'               Case Else
'                  NowPrint m_CP09, "02", "01", False, strUserNum, 0
'               End Select
'            ' 申請人國籍非台灣
'            'edit by nickc 2006/06/30
'            'Else
'            ElseIf textPrint = "2" Then
'               ' 列印定稿
'               NowPrint m_CP09, "02", "02", False, strUserNum, 0
'            End If
'         End If
'   End Select
'End Sub
'
'' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
'Private Sub InsExpField1()
'   Dim strTM23Nation As String
'   Dim strSql As String
'   strTM23Nation = Empty
'   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
'
'   ' 系統類別TD
'   Select Case m_TM01
'      Case "TC":
'        'add by nickc 2006/06/30
'        If textPrint = "1" Then
'            ' 清除定稿例外欄位檔原有資料
'            EndLetter "02", m_CP09, "04", strUserNum
'        End If
'      Case "T", "TF":
'         If m_TM10 < "010" Then
'            ' 申請人國籍為台灣
'            'edit by nickc 2006/06/30
'            'If strTM23Nation < "010" Then
'            If textPrint = "1" Then
'               ' 清除定稿例外欄位檔原有資料
'               EndLetter "02", m_CP09, "01", strUserNum
'            ' 申請人國籍非台灣
'            'edit by nickc 2006/06/30
'            'Else
'            ElseIf textPrint = "2" Then
'               ' 清除定稿例外欄位檔原有資料
'               EndLetter "02", m_CP09, "02", strUserNum
'               cnnConnection.Execute strSql
'               ' 回音
'               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                        "VALUES ('" & "02" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
'                        "'" & "回音" & "','" & textCF09 & "')"
'               cnnConnection.Execute strSql
'            End If
'         End If
'   End Select
'End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim dblCountFee As Double
Dim intMoney As Long  '倍數
   
   TxtValidate = False
   ' 91.09.02 marked by louis
   'If Me.textCP09S.Enabled = True Then
   '   Cancel = False
   '   textCP09S_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
   'Add By Sindy 2010/12/24
   If Me.textTM12.Enabled = True Then
      Cancel = False
      textTM12_Validate Cancel
      If Cancel = True Then
         textTM12.SetFocus
         Exit Function
      End If
   End If
   
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
      'Add By Sindy 2014/2/20 檢查延展跨類規費
      If m_CP10 = "102" Then
         If Trim(textTM09) <> "" Then 'Add By Sindy 2014/4/18 +if 證明標章沒有商品類別
            textTM09 = IIf(Right(textTM09, 1) = ",", Mid(textTM09, 1, Len(textTM09) - 1), textTM09)
            textTM09 = IIf(Left(textTM09, 1) = ",", Mid(textTM09, 2, Len(textTM09)), textTM09)
            tmpArr = Split(textTM09, ",")
            '有跨類才需要檢查
            If (Val(UBound(tmpArr)) + 1) > 1 Then
               intMoney = 1
'               If m_CP07 <> "" Then
'                  '若系統日的昨天為非工作天, 則以系統日的前一個工作天做比較
'                  If ChkWorkDay(DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1))))) = False Then
'                     If (Val(DBDATE(m_CP07)) - 19110000) < (Val(CompWorkDay(1, DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1)))), 1)) - 19110000) And (Val(DBDATE(m_CP07)) - 19110000) <> 0 Then
'                        intMoney = 2
'                     End If
'                  Else
'                     If (Val(DBDATE(m_CP07)) - 19110000) < Val(GetTaiwanTodayDate) And (Val(DBDATE(m_CP07)) - 19110000) <> 0 Then
'                        intMoney = 2
'                     End If
'                  End If
'               End If
               'Modify By Sindy 2015/4/24 ex.FCT-23444 原用CP07檢查,改用TM22檢查
               '若系統日的昨天為非工作天, 則以系統日的前一個工作天做比較
               If ChkWorkDay(DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1))))) = False Then
                  If (Val(DBDATE(m_TM22)) - 19110000) < (Val(CompWorkDay(1, DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1)))), 1)) - 19110000) And (Val(DBDATE(m_TM22)) - 19110000) <> 0 Then
                     intMoney = 2
                  End If
               Else
                  If (Val(DBDATE(m_TM22)) - 19110000) < Val(GetTaiwanTodayDate) And (Val(DBDATE(m_TM22)) - 19110000) <> 0 Then
                     intMoney = 2
                  End If
               End If
               dblCountFee = Val(4000 * (Val(UBound(tmpArr)) + 1) * intMoney)
               If Val(textCP84) <> dblCountFee Then
                  MsgBox "規費不符，本案類別數共" & (Val(UBound(tmpArr)) + 1) & "類，每類4,000元，應為" & Format(dblCountFee, "#,##0") & "元" & IIf(intMoney = 2, "（含逾期加倍）", "") & "。" & vbCrLf & _
                         "若非所有類別都要延展，必須將商品類別欄之類別數改正確!!"
                  SSTab1.Tab = 0
                  textTM09.SetFocus
                  Exit Function
               End If
            End If
         End If
      End If
      '2014/2/20 END
   End If
   'edit by nickc 2006/01/27
   'If Me.textCP22.Enabled = True Then
   '   Cancel = False
   '   textCP22_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
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
   
   If Me.textMail.Enabled = True Then
      Cancel = False
      textMail_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textMediaDate.Enabled = True Then
      Cancel = False
      textMediaDate_Validate Cancel
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
   
   If Me.textTM11.Enabled = True Then
      Cancel = False
      textTM11_Validate Cancel
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
   
   'Add by Amy 2018/10/30
   If Me.textTM78.Enabled = True Then
      Cancel = False
      textTM78_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM79.Enabled = True Then
      Cancel = False
      textTM79_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM80.Enabled = True Then
      Cancel = False
      textTM80_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textTM81.Enabled = True Then
      Cancel = False
      textTM81_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2018/10/30
   
   If Me.textTM32.Enabled = True Then
      Cancel = False
      textTM32_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textTM58.Enabled = True Then
      Cancel = False
      textTM58_Validate Cancel
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
   
   'add by nickc 2007/01/02 檢查商品類別及組群，
   '商標種類為 7 or  8 不會有
   '系統別  國家   申請人國籍 案件性質  必要
   'T       000                         類
   'T       000               101       類 + 組
   'T       020                         類
   'TF                                  類
   'TS                                  類
   'TS             <>000                類 + 組
   If textTM08 <> "7" And textTM08 <> "8" Then
       '2007/7/31 MODIFY BY SONIA T,TS分開判斷
       'If (m_TM01 = "T" And m_TM10 = "000" And m_CP10 = "101") Or (m_TM01 = "TS" And m_CU10 <> "000") Then
       If (m_TM01 = "T" And m_TM10 = "000" And m_CP10 = "101") Then
           If Trim(textTM09) = "" Then
               MsgBox "商品類別及組群不可空白！", , "資料錯誤！"
               SSTab1.Tab = 0
               textTM09.SetFocus
               Exit Function
           ElseIf Trim(textTM32) = "" Then
               MsgBox "商品類別及組群不可空白！", , "資料錯誤！"
               SSTab1.Tab = 0
               textTM32.SetFocus
               Exit Function
           End If
       '2007/7/31 ADD BY SONIA 組群可空白,TS-000812查整個類別
       ElseIf (m_TM01 = "TS" And m_CU10 <> "000") Then
           If Trim(textTM09) = "" Then
               MsgBox "商品類別不可空白！", , "資料錯誤！"
               SSTab1.Tab = 0
               textTM09.SetFocus
               Exit Function
           End If
       '2007/7/31 END
       ElseIf (m_TM01 = "T" And m_TM10 = "000" And m_CP10 <> "101") Or (m_TM01 = "T" And m_TM10 = "020") Or (m_TM01 = "TF") Or (m_TM01 = "TS") Then
           If Trim(textTM09) = "" Then
               MsgBox "商品類別不可空白！", , "資料錯誤！"
               SSTab1.Tab = 0
               textTM09.SetFocus
               Exit Function
           End If
       End If
   End If
   

   'Added by Lydia 2024/06/14 對申請人1~5的重複輸入檢查
   If Pub_ChkAppList(strExc(0), textTM23 & "," & textTM78 & "," & textTM79 & "," & textTM80 & "," & textTM81) = False Then
      Me.SSTab1.Tab = 0
      Select Case Val(strExc(0))
         Case 1
            textTM23.SetFocus
            textTM23_GotFocus
         Case 2
            textTM78.SetFocus
            textTM78_GotFocus
         Case 3
            textTM79.SetFocus
            textTM79_GotFocus
         Case 4
            textTM80.SetFocus
            textTM80_GotFocus
         Case 5
            textTM81.SetFocus
            textTM81_GotFocus
      End Select
      Exit Function
   End If
   'end 2024/06/14
   
   'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
   For ii = 1 To 5
      strExc(1) = ""
      Select Case ii
         Case 1 '申請人1
            strExc(1) = ChangeCustomerL(textTM23)
            strExc(2) = ChangeCustomerL(m_TM23)
         Case 2 '申請人2
            strExc(1) = ChangeCustomerL(textTM78)
            strExc(2) = ChangeCustomerL(m_TM78)
         Case 3 '申請人3
            strExc(1) = ChangeCustomerL(textTM79)
            strExc(2) = ChangeCustomerL(m_TM79)
         Case 4 '申請人4
            strExc(1) = ChangeCustomerL(textTM80)
            strExc(2) = ChangeCustomerL(m_TM80)
         Case 5 '申請人5
            strExc(1) = ChangeCustomerL(textTM81)
            strExc(2) = ChangeCustomerL(m_TM81)
      End Select
      If strExc(1) <> "" And strExc(1) <> strExc(2) Then
         If GetCustomerAndState(strExc(1), strExc(3), , , , m_TM01, strExc(8), False, Me.Name, m_TM02, m_TM03, m_TM04) = False Then
            Me.SSTab1.Tab = 0
            If ii = 1 Then
               textTM23.SetFocus
               textTM23_GotFocus
               Exit Function
            ElseIf ii = 2 Then
               textTM78.SetFocus
               textTM78_GotFocus
               Exit Function
            ElseIf ii = 3 Then
               textTM79.SetFocus
               textTM79_GotFocus
               Exit Function
            ElseIf ii = 4 Then
               textTM80.SetFocus
               textTM80_GotFocus
               Exit Function
            ElseIf ii = 5 Then
               textTM81.SetFocus
               textTM81_GotFocus
               Exit Function
            End If
         End If
      End If
   Next
   'end 2024/06/13
   
   'Add By Sindy 2016/12/16
   '檢查有設定副本收受人需提醒並新增信函副本B類收文
   m_990CP09 = ""
   If textPrint = "N" Then '不印定稿
      If PUB_ChkCC(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_990CP09) = False Then
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   '2016/12/16 END
   
    'Added by Lydia 2021/06/04 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
    If Pub_ChkACS112isNull(m_TM01, m_TM02, m_TM03, m_TM04, txtCP113) = True Then
        SSTab1.Tab = 0
        txtCP113.SetFocus
        txtCP113_GotFocus
        Exit Function
    End If
    'end 2021/06/04
    
   TxtValidate = True
End Function

' 91.09.02 modify by louis
' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 8
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "審定號"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "是否存在"
   grdList.ColWidth(2) = 0
   grdList.col = 3
   grdList.Text = "本所案號"
   grdList.ColWidth(3) = 0
   grdList.col = 4
   grdList.Text = "本所案號"
   grdList.ColWidth(4) = 0
   grdList.col = 5
   grdList.Text = "本所案號"
   grdList.ColWidth(5) = 0
   grdList.col = 6
   grdList.Text = "本所案號"
   grdList.ColWidth(6) = 0
   '911015 nick 新增
   grdList.col = 7
   grdList.Text = "案件名稱"
   grdList.ColWidth(7) = 7000
   bolFixed = False 'Added by Lydia 2023/10/13
End Sub

Private Function ExistTM15(ByVal strTM15 As String, ByRef strCP01 As String, ByRef strCP02 As String, ByRef strCP03 As String, ByRef strCP04 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bExist As Boolean
   
   bExist = False
   strSql = "SELECT * FROM TRADEMARK " & _
            "WHERE TM15 = '" & strTM15 & "' AND " & _
                  "TM10 = '" & m_TM10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      bExist = True
      If Not IsNull(rsTmp.Fields("TM01")) Then
         strCP01 = rsTmp.Fields("TM01")
      End If
      If Not IsNull(rsTmp.Fields("TM02")) Then
         strCP02 = rsTmp.Fields("TM02")
      End If
      If Not IsNull(rsTmp.Fields("TM03")) Then
         strCP03 = rsTmp.Fields("TM03")
      End If
      If Not IsNull(rsTmp.Fields("TM04")) Then
         strCP04 = rsTmp.Fields("TM04")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
      
   ExistTM15 = bExist
End Function

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nRow As Integer
   Dim nCol As Integer
   Dim nCurrSel As Integer
   nCurrSel = grdList.row
   For nRow = 1 To grdList.Rows - 1
      grdList.row = nRow
      If nRow = nCurrSel Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            grdList.CellBackColor = &H8000000D
            grdList.CellForeColor = &H80000005
         Next nCol
      Else
         grdList.col = 1
         If grdList.CellBackColor <> &H80000005 Then
            For nCol = 1 To grdList.Cols - 1
               grdList.col = nCol
               If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
               If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
            Next nCol
         End If
      End If
   Next nRow
   grdList.row = nCurrSel
   grdList.col = 0
End Sub

Private Sub CheckGrdList()
   Dim nIndex As Integer
   If grdList.Rows > 2 Then
      For nIndex = 0 To grdList.Rows - 1
         If IsEmptyText(grdList.TextMatrix(nIndex, 1)) Then
            grdList.RemoveItem nIndex
            Exit For
         End If
      Next nIndex
   End If
End Sub

' 刪除項目
Private Sub cmdDelItem_Click()
   If grdList.row > 0 And grdList.row < grdList.Rows Then
      If grdList.Rows = 2 Then
         grdList.TextMatrix(grdList.row, 1) = Empty
      Else
         grdList.RemoveItem grdList.row
      End If
   End If
   CheckGrdList
End Sub

' 更改項目
Private Sub cmdModItem_Click()
   If IsEmptyText(textTM15S) Then
      Exit Sub
   End If
   
   If grdList.row > 0 And grdList.row < grdList.Rows Then
      grdList.TextMatrix(grdList.row, 1) = textTM15S.Text
   End If
   CheckGrdList
End Sub

' 新增項目
Private Sub cmdAddItem_Click()
   Dim bFind As Boolean
   Dim nIndex As Integer
   
   If IsEmptyText(textTM15S) Then
      Exit Sub
   End If
   bFind = False
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 1) = textTM15S.Text Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If Not bFind Then
      grdList.Rows = grdList.Rows + 1
      nIndex = grdList.Rows - 1
      grdList.TextMatrix(nIndex, 1) = textTM15S.Text
      '911015 nick 新增
      grdList.TextMatrix(nIndex, 7) = nick911015(0).Caption & "   " & IIf(Len(nick911015(1).Caption) = 0, IIf(Len(nick911015(2).Caption) = 0, IIf(Len(nick911015(3).Caption) = 0, "", nick911015(3).Caption), nick911015(2).Caption), nick911015(1).Caption)
      textTM15S.Text = ""
      nick911015(0).Caption = ""
      nick911015(1).Caption = ""
      nick911015(2).Caption = ""
      nick911015(3).Caption = ""
      textTM15S.SetFocus
      ' 顯示Focus的項目
      grdList.row = grdList.Rows - 1
      grdList_ShowSelection
      'Added by Lydia 2023/10/13
      If bolFixed = False Then
         bolFixed = True
         grdList.FixedRows = 1
      End If
      'end 2023/10/13
   End If
   CheckGrdList
End Sub

'Modify By Cheng 2002/11/06
'Private Sub OnCopyCPData(ByVal strCP09)
Private Function OnCopyCPData(ByVal strCP09) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim strFieldName As String
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnCopyCPData = True

   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & strCP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      For nIndex = 1 To 80
         If (nIndex >= 10) Then
            strFieldName = "CP" & CStr(nIndex)
         Else
            strFieldName = "CP0" & CStr(nIndex)
         End If
         If Not IsNull(rsTmp.Fields(strFieldName)) Then
            If rsTmp.Fields(strFieldName).Type = adNumeric Then
               SetNewCPFieldData strFieldName, rsTmp.Fields(strFieldName), 1
            Else
               SetNewCPFieldData strFieldName, rsTmp.Fields(strFieldName), 0
            End If
         End If
      Next nIndex
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   ' 若為本所案件則產生B類案件進度資料
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 2) = 1 Then
         SetNewCPFieldData "CP01", grdList.TextMatrix(nIndex, 3), 0
         SetNewCPFieldData "CP02", grdList.TextMatrix(nIndex, 4), 0
         SetNewCPFieldData "CP03", grdList.TextMatrix(nIndex, 5), 0
         SetNewCPFieldData "CP04", grdList.TextMatrix(nIndex, 6), 0
         '911023 nick  cp16 cp17 cp18 cp19 cp60 要為 null    cp26='N'
         '***** start
         SetNewCPFieldData "CP16", "null", 1
         SetNewCPFieldData "CP17", "null", 1
         SetNewCPFieldData "CP18", "null", 1
         SetNewCPFieldData "CP19", "null", 1
         SetNewCPFieldData "CP60", "null", 1
         SetNewCPFieldData "CP26", "N", 0
         '***** end
         SetNewCPFieldData "CP09", AutoNo("B", 6), 0
         '911015 nick 新增   邱小姐說把新增到案件進度檔的該筆資料的收文日更新成跟發文日相同
         '***** start
         SetNewCPFieldData "CP05", DBDATE(textCP27), 0
         '***** end
         '911104 nick 邱小姐說新增B 類的備註，要和本來的相同
         
         strSql = GetNewCPSQL()
         cnnConnection.Execute strSql
      End If
   Next nIndex
   
   ClearNewCPFieldList
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnCopyCPData = False
End Function

'''''''''''''''''''''''
Private Function GetNewCPSQL() As String
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   
   strSql = "INSERT INTO CaseProgress ("
   For nIndex = 0 To m_NewCPListCount - 1
      If nIndex <> 0 Then strSql = strSql & ","
      strSql = strSql & m_NewCPList(nIndex).fiName
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   For nIndex = 0 To m_NewCPListCount - 1
      If nIndex <> 0 Then strSql = strSql & ","
      If m_NewCPList(nIndex).fiType = 0 Then
         strSql = strSql & "'" & m_NewCPList(nIndex).fiNewData & "'"
      Else
         strSql = strSql & m_NewCPList(nIndex).fiNewData
      End If
   Next nIndex
   strSql = strSql & ") "
   GetNewCPSQL = strSql
End Function

'Add By Cheng 2002/11/18
Private Function GetFagentNation(strTM44 As String) As String
    Dim rsA As New ADODB.Recordset
    Dim StrSQLa As String
    
    GetFagentNation = ""
    StrSQLa = "Select FA10 From Fagent Where FA01='" & Left(strTM44, 8) & "' And FA02='" & Right(strTM44, 1) & "'"
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        GetFagentNation = "" & rsA("FA10").Value
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Function

'2007/8/7 ADD BY SONIA
Private Function GetDelayTime(strTM10 As String) As Integer
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

StrSQLa = "Select NA15 From Nation Where NA01='" & strTM10 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   GetDelayTime = Val("0" & rsA.Fields(0).Value)
Else
   GetDelayTime = 0
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'取得部分核駁商品服務
Public Function GetGoods1205(stCP01 As String, stCP02 As String, stCP03 As String, stCP04 As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQuery As String
    Dim intQ As Integer
    
    GetGoods1205 = ""
    
    strQuery = "Select CP64 From CaseProgress Where CP01='" & stCP01 & "' And CP02='" & stCP02 & "' " & _
                    "And CP03='" & stCP03 & "' And CP04='" & stCP04 & "' And CP10='1205' "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQuery)
    If intQ = 1 Then
        GetGoods1205 = "" & RsQ.Fields("CP64")
    End If
End Function

'Add by Amy 2015/05/26 T大陸商標續展計算催審期限
Private Function SetUargeDate_102() As Boolean
    'Modify by Amy 2015/09/10
    Dim strTM22 As String
    
    SetUargeDate_102 = False
    
    strExc(0) = "Select  TM22 From TradeMark Where TM01='" & m_TM01 & "' And TM02='" & m_TM02 & "' And TM03='" & m_TM03 & "' And TM04='" & m_TM04 & "' "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        If IsNull(RsTemp.Fields("TM22")) Then
            MsgBox "請至基本檔補輸專用期間", , MsgText(5)
            textCP27.SetFocus
            Exit Function
        End If
        strTM22 = RsTemp.Fields("TM22")
        If Val(textCP27) <= Val(strTM22) Then
            '判斷發文日<=基本檔專用期間止日
            If Val(textUargeDate) > Val(TAIWANDATE(DateAdd("M", 3, ChangeTStringToTDateString(strTM22)))) Then
                '比較催審期限與基本檔專用期間止日+3個月 ,改掛較小者
                textUargeDate = TAIWANDATE(DateAdd("M", 3, ChangeTStringToTDateString(strTM22)))
            End If
        Else
            '判斷發文日>基本檔專用期間止日,催審期限=基本檔專用期間止日+5個月
            textUargeDate = TAIWANDATE(DateAdd("M", 5, ChangeTStringToTDateString(strTM22)))
        End If
        SetUargeDate_102 = True
    Else
        MsgBox "基本檔無資料請確認", , MsgText(5)
        textCP27.SetFocus
        Exit Function
    End If
    'end 2015/09/10
End Function

'add by sonia 2018/9/26
Private Function Get102TM21TM22(strTMF As String) As String

   If Val(m_TM22) = 0 Then Exit Function
   
   If strTMF = "TM21" Then
      If m_TM01 = "TF" Then
         Get102TM21TM22 = DBDATE(m_TM22)
      Else
         Get102TM21TM22 = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22))))
      End If
   Else  'TM22
      'Modified by Lydia 2019/11/13 改用共用模組, 第1次專用期間=公告日+10年-1天，之後延展102沒有減１天；與專利不一樣
      'If Mid(ChangeWDateStringToWString(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22)))), 5) = "0228" Then
      '     If Mid(ChangeWDateStringToWString(DateAdd("d", 1, DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))), 5) = "0229" Then
      '         Get102TM21TM22 = DBDATE(DateAdd("d", 1, DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22)))))
      '     Else
      '         Get102TM21TM22 = DBDATE(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
      '     End If
      'Else
      '     Get102TM21TM22 = DBDATE(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
      'End If
      'Modify By Sindy 2022/3/7 + m_TM10 : 延展後之專用期限年度倘有2月29日時，專用期限止日應為2月29日，而非以加10年之方式計算為2月28日
      Get102TM21TM22 = PUB_GetEndDate(DBDATE(m_TM22), Val(m_NA14), "N", m_TM10)
      'end 2019/11/13
   End If

End Function
'end 2018/9/26

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
