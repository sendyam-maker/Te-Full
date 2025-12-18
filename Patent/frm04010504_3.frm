VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010504_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "一般來函輸入"
   ClientHeight    =   6060
   ClientLeft      =   108
   ClientTop       =   816
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9000
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   4512
      TabIndex        =   79
      Top             =   48
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   78
      Top             =   60
      Width           =   495
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   77
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   76
      Top             =   60
      Width           =   255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   75
      Top             =   60
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   7464
      TabIndex        =   34
      Top             =   0
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   6840
      TabIndex        =   33
      Top             =   0
      Width           =   600
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   8388
      TabIndex        =   35
      Top             =   0
      Width           =   600
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4995
      Left            =   90
      TabIndex        =   55
      Top             =   1080
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   8805
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "來函輸入1"
      TabPicture(0)   =   "frm04010504_3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label28"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label26"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label25"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label23"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label24"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label22"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label13"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label14"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label37"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label43"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label41(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label41(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label39(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label36(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label36(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label9"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label15(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label17"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblDispDate"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label27"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label29"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label48"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label49"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label3(0)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Image1"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lblsNP23"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text29"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text30"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "frm307"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text7"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text6"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Frame4"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Frame3"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Frame2"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "MSHFlexGrid1"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text14(1)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Frame1"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text18"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text17"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text16"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text14(0)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text9"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text21(0)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Text21(1)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Text22"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Text15(1)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Text15(0)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "text8"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Combo2"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Combo3"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtDispDate"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Text33"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "cmdDeadLine"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "Text13"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "Text37"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtSNP23"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).ControlCount=   61
      TabCaption(1)   =   "來函輸入2"
      TabPicture(1)   =   "frm04010504_3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtIDSFee(1)"
      Tab(1).Control(1)=   "txtIDSPt(1)"
      Tab(1).Control(2)=   "txtIDSFee(2)"
      Tab(1).Control(3)=   "txtIDSPt(2)"
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(5)=   "Text23"
      Tab(1).Control(6)=   "Text19"
      Tab(1).Control(7)=   "Text26"
      Tab(1).Control(8)=   "Text27(0)"
      Tab(1).Control(9)=   "Text20(5)"
      Tab(1).Control(10)=   "Text20(4)"
      Tab(1).Control(11)=   "Text20(3)"
      Tab(1).Control(12)=   "Text20(2)"
      Tab(1).Control(13)=   "Text20(1)"
      Tab(1).Control(14)=   "Text20(0)"
      Tab(1).Control(15)=   "Text31"
      Tab(1).Control(16)=   "Label53"
      Tab(1).Control(17)=   "Label52"
      Tab(1).Control(18)=   "Label10"
      Tab(1).Control(19)=   "Label36(0)"
      Tab(1).Control(20)=   "Label35"
      Tab(1).Control(21)=   "Label34"
      Tab(1).Control(22)=   "Label33"
      Tab(1).Control(23)=   "Label32"
      Tab(1).Control(24)=   "Label31"
      Tab(1).Control(25)=   "Label30"
      Tab(1).Control(26)=   "Label38"
      Tab(1).Control(27)=   "Label39(0)"
      Tab(1).Control(28)=   "Label40"
      Tab(1).Control(29)=   "Label46"
      Tab(1).ControlCount=   30
      TabCaption(2)   =   "被舉發"
      TabPicture(2)   =   "frm04010504_3.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkItem(7)"
      Tab(2).Control(1)=   "chkItem(8)"
      Tab(2).Control(2)=   "txtDay(0)"
      Tab(2).Control(3)=   "txtDay(1)"
      Tab(2).Control(4)=   "txtMonth(1)"
      Tab(2).Control(5)=   "txtYear(1)"
      Tab(2).Control(6)=   "txtYear(0)"
      Tab(2).Control(7)=   "txtMonth(0)"
      Tab(2).Control(8)=   "txtItemCount"
      Tab(2).Control(9)=   "chkItem(0)"
      Tab(2).Control(10)=   "chkItem(1)"
      Tab(2).Control(11)=   "chkItem(4)"
      Tab(2).Control(12)=   "chkItem(3)"
      Tab(2).Control(13)=   "chkItem(5)"
      Tab(2).Control(14)=   "chkItem(2)"
      Tab(2).Control(15)=   "txtItemList"
      Tab(2).Control(16)=   "chkItem(6)"
      Tab(2).Control(17)=   "Label51"
      Tab(2).Control(18)=   "Label15(0)"
      Tab(2).Control(19)=   "Label50"
      Tab(2).Control(20)=   "Label47"
      Tab(2).Control(21)=   "Label16(5)"
      Tab(2).ControlCount=   22
      Begin VB.TextBox txtIDSFee 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   1
         Left            =   -67890
         MaxLength       =   6
         TabIndex        =   40
         Top             =   420
         Width           =   765
      End
      Begin VB.TextBox txtIDSPt 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   1
         Left            =   -66960
         MaxLength       =   3
         TabIndex        =   41
         Top             =   420
         Width           =   375
      End
      Begin VB.TextBox txtIDSFee 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   2
         Left            =   -67890
         MaxLength       =   6
         TabIndex        =   42
         Top             =   720
         Width           =   765
      End
      Begin VB.TextBox txtIDSPt 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   2
         Left            =   -66960
         MaxLength       =   3
         TabIndex        =   43
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtSNP23 
         Height          =   270
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   160
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "同一人於同日就相同創作分別申請發明及新型專利，已於申請時分別聲明，而其發明及新型專利權同時並存者"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   -71310
         TabIndex        =   156
         Top             =   720
         Width           =   4605
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "同一人就相同創作，於同日分別申請發明專利及新型專利，其發明專利審定前，新型專利權已當然消滅或撤銷確定者"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   -71310
         TabIndex        =   155
         Top             =   1080
         Width           =   4605
      End
      Begin VB.TextBox Text37 
         Height          =   270
         Left            =   5895
         MaxLength       =   6
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   372
         Width           =   735
      End
      Begin VB.TextBox Text13 
         Height          =   270
         Left            =   7020
         MaxLength       =   2
         TabIndex        =   6
         Top             =   975
         Width           =   375
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72030
         MaxLength       =   2
         TabIndex        =   145
         Top             =   2400
         Width           =   285
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -69645
         MaxLength       =   2
         TabIndex        =   148
         Top             =   2400
         Width           =   285
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70230
         MaxLength       =   2
         TabIndex        =   147
         Top             =   2400
         Width           =   285
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -70905
         MaxLength       =   3
         TabIndex        =   146
         Top             =   2400
         Width           =   420
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -73290
         MaxLength       =   3
         TabIndex        =   143
         Top             =   2400
         Width           =   420
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -72570
         MaxLength       =   2
         TabIndex        =   144
         Top             =   2400
         Width           =   285
      End
      Begin VB.TextBox txtItemCount 
         Enabled         =   0   'False
         Height          =   270
         Left            =   -72075
         TabIndex        =   135
         Top             =   1020
         Width           =   375
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "被請求撤銷全部請求項：共計"
         Height          =   210
         Index           =   0
         Left            =   -74730
         TabIndex        =   134
         Top             =   1050
         Width           =   2670
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "被請求撤銷部分之請求項："
         Height          =   210
         Index           =   1
         Left            =   -74730
         TabIndex        =   136
         Top             =   1260
         Width           =   2625
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "共有專利申請權非由全體共有人提出申請者"
         Height          =   210
         Index           =   4
         Left            =   -71310
         TabIndex        =   140
         Top             =   1680
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "專利權人為非專利申請權人者"
         Height          =   210
         Index           =   3
         Left            =   -71310
         TabIndex        =   139
         Top             =   1470
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "專利權人所屬國家對中華民國申請專利不予受理者"
         Height          =   210
         Index           =   5
         Left            =   -71310
         TabIndex        =   141
         Top             =   1890
         Width           =   4335
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "請求撤銷設計專利權"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   -74730
         TabIndex        =   138
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox txtItemList 
         Enabled         =   0   'False
         Height          =   450
         Left            =   -74460
         TabIndex        =   137
         Text            =   "第項"
         Top             =   1470
         Width           =   2580
      End
      Begin VB.CommandButton cmdDeadLine 
         Caption         =   "補件資料"
         Height          =   285
         Left            =   7380
         TabIndex        =   133
         Top             =   365
         Width           =   1300
      End
      Begin VB.TextBox Text33 
         Height          =   270
         Left            =   6645
         MaxLength       =   15
         TabIndex        =   28
         Top             =   2880
         Width           =   1875
      End
      Begin VB.TextBox txtDispDate 
         Height          =   270
         Left            =   3465
         MaxLength       =   8
         TabIndex        =   1
         Top             =   372
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame Frame5 
         Caption         =   "檢索報告種類(可複選)"
         Height          =   990
         Left            =   -74820
         TabIndex        =   116
         Top             =   3885
         Visible         =   0   'False
         Width           =   8265
         Begin VB.TextBox Text28 
            Height          =   270
            Left            =   1305
            MaxLength       =   8
            TabIndex        =   53
            Top             =   570
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "國際檢索報告"
            Height          =   315
            Index           =   2
            Left            =   5520
            TabIndex        =   52
            Top             =   240
            Width           =   1485
         End
         Begin VB.CheckBox Check1 
            Caption         =   "傳送國際檢索報告和國際檢索單位書面意見或宣布通知書"
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   51
            Top             =   240
            Width           =   5025
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "官方發文日:"
            Height          =   180
            Left            =   180
            TabIndex        =   118
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label20 
            Height          =   200
            Left            =   2750
            TabIndex        =   117
            Top             =   490
            Width           =   1215
         End
      End
      Begin VB.ComboBox Combo3 
         Height          =   276
         Left            =   5940
         TabIndex        =   26
         Top             =   2580
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Left            =   1320
         TabIndex        =   25
         Top             =   2580
         Width           =   3045
      End
      Begin VB.ComboBox text8 
         Height          =   276
         Left            =   4680
         TabIndex        =   4
         Text            =   "text8"
         Top             =   660
         Width           =   1515
      End
      Begin VB.TextBox Text23 
         Height          =   270
         Left            =   -70560
         MaxLength       =   3
         TabIndex        =   39
         Top             =   732
         Width           =   684
      End
      Begin VB.TextBox Text15 
         Height          =   270
         Index           =   0
         Left            =   1845
         MaxLength       =   1
         TabIndex        =   30
         Top             =   3825
         Width           =   255
      End
      Begin VB.TextBox Text15 
         Height          =   270
         Index           =   1
         Left            =   4905
         MaxLength       =   1
         TabIndex        =   31
         Top             =   3825
         Width           =   255
      End
      Begin VB.TextBox Text22 
         Enabled         =   0   'False
         Height          =   270
         Left            =   7440
         MaxLength       =   7
         TabIndex        =   24
         Top             =   2304
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text19 
         Height          =   270
         Left            =   -73320
         TabIndex        =   38
         Top             =   732
         Width           =   1215
      End
      Begin VB.TextBox Text21 
         Height          =   270
         Index           =   1
         Left            =   2800
         TabIndex        =   23
         Top             =   2304
         Width           =   975
      End
      Begin VB.TextBox Text21 
         Height          =   270
         Index           =   0
         Left            =   1320
         TabIndex        =   22
         Top             =   2304
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   5
         Top             =   960
         Width           =   3990
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   15
         Top             =   1740
         Width           =   975
      End
      Begin VB.TextBox Text16 
         Height          =   270
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   19
         Top             =   2028
         Width           =   975
      End
      Begin VB.TextBox Text17 
         Height          =   270
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   20
         Top             =   2028
         Width           =   975
      End
      Begin VB.TextBox Text18 
         Height          =   270
         Left            =   7440
         MaxLength       =   1
         TabIndex        =   21
         Top             =   2028
         Width           =   375
      End
      Begin VB.TextBox Text26 
         Height          =   270
         Left            =   -73320
         MaxLength       =   8
         TabIndex        =   36
         Top             =   432
         Width           =   1215
      End
      Begin VB.TextBox Text27 
         Height          =   270
         Index           =   0
         Left            =   -71130
         MaxLength       =   1
         TabIndex        =   37
         Top             =   432
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   1350
         TabIndex        =   56
         Top             =   1200
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "文到次日"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   8
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "文到當日"
            Height          =   180
            Index           =   0
            Left            =   144
            TabIndex        =   7
            Top             =   180
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox Text14 
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   4680
         MaxLength       =   7
         TabIndex        =   16
         Top             =   1740
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   810
         Left            =   945
         TabIndex        =   32
         Top             =   4140
         Width           =   7755
         _ExtentX        =   13674
         _ExtentY        =   1439
         _Version        =   393216
         Cols            =   12
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
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
         _Band(0).Cols   =   12
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   4080
         TabIndex        =   57
         Top             =   1200
         Width           =   4215
         Begin VB.TextBox Text12 
            Height          =   252
            Left            =   2760
            MaxLength       =   7
            TabIndex        =   14
            Top             =   150
            Width           =   975
         End
         Begin VB.TextBox Text11 
            Height          =   270
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   12
            Top             =   150
            Width           =   375
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Left            =   840
            MaxLength       =   2
            TabIndex        =   10
            Top             =   150
            Width           =   375
         End
         Begin VB.OptionButton Option4 
            Caption         =   "                      日"
            Height          =   225
            Index           =   2
            Left            =   2520
            TabIndex        =   13
            Top             =   180
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "        月"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   11
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "文到           天"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   180
            Value           =   -1  'True
            Width           =   1400
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  '沒有框線
         Height          =   420
         Left            =   6048
         TabIndex        =   108
         Top             =   1608
         Visible         =   0   'False
         Width           =   2676
         Begin VB.TextBox Text14 
            Enabled         =   0   'False
            Height          =   270
            Index           =   2
            Left            =   1392
            MaxLength       =   7
            TabIndex        =   17
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "延緩公告日:"
            Height          =   252
            Left            =   72
            TabIndex        =   109
            Top             =   153
            Width           =   972
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  '沒有框線
         Height          =   420
         Left            =   6048
         TabIndex        =   110
         Top             =   1608
         Visible         =   0   'False
         Width           =   2676
         Begin VB.TextBox Text32 
            Enabled         =   0   'False
            Height          =   270
            Left            =   555
            MaxLength       =   10
            TabIndex        =   18
            Top             =   144
            Width           =   1875
         End
         Begin VB.Label Label8 
            Caption         =   "時間:"
            Height          =   252
            Left            =   72
            TabIndex        =   111
            Top             =   144
            Width           =   972
         End
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   0
         Top             =   372
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   2
         Top             =   660
         Width           =   975
      End
      Begin VB.Frame frm307 
         Height          =   1035
         Left            =   5985
         TabIndex        =   122
         Top             =   3090
         Width           =   2715
         Begin VB.TextBox Text36 
            Height          =   270
            Left            =   2160
            TabIndex        =   127
            Top             =   705
            Width           =   510
         End
         Begin VB.TextBox Text35 
            Height          =   270
            Left            =   840
            TabIndex        =   126
            Top             =   705
            Width           =   870
         End
         Begin VB.TextBox Text34 
            Height          =   270
            Left            =   2160
            TabIndex        =   125
            Top             =   420
            Width           =   510
         End
         Begin VB.TextBox Text24 
            Height          =   270
            Left            =   840
            MaxLength       =   1
            TabIndex        =   123
            Top             =   120
            Width           =   375
         End
         Begin VB.TextBox Text25 
            Height          =   270
            Left            =   840
            TabIndex        =   124
            Top             =   420
            Width           =   870
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "點數:"
            Height          =   180
            Left            =   1755
            TabIndex        =   132
            Top             =   750
            Width           =   405
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "實審費用:"
            Height          =   180
            Left            =   45
            TabIndex        =   131
            Top             =   750
            Width           =   765
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "點數:"
            Height          =   180
            Left            =   1755
            TabIndex        =   130
            Top             =   465
            Width           =   405
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "是否分割:           (Y:分割)"
            Height          =   180
            Left            =   45
            TabIndex        =   129
            Top             =   150
            Width           =   1950
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "分割費用:"
            Height          =   180
            Left            =   45
            TabIndex        =   128
            Top             =   465
            Width           =   765
         End
      End
      Begin VB.CheckBox chkItem 
         Caption         =   "請求撤銷自「        年      月      日」至「        年      月      日」之專利權期間延長"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   -74730
         TabIndex        =   142
         Top             =   2430
         Width           =   7440
      End
      Begin MSForms.TextBox Text30 
         Height          =   300
         Left            =   1320
         TabIndex        =   27
         Top             =   2880
         Width           =   3810
         VariousPropertyBits=   671107099
         MaxLength       =   32
         Size            =   "6720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text29 
         Height          =   420
         Left            =   105
         TabIndex        =   29
         Top             =   3390
         Width           =   5865
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "10345;741"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   300
         Index           =   5
         Left            =   -73320
         TabIndex        =   49
         Top             =   2535
         Width           =   6735
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "11880;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   300
         Index           =   4
         Left            =   -73320
         TabIndex        =   48
         Top             =   2220
         Width           =   6735
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "11880;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   300
         Index           =   3
         Left            =   -73320
         TabIndex        =   47
         Top             =   1935
         Width           =   6735
         VariousPropertyBits=   671107099
         MaxLength       =   600
         Size            =   "11880;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   300
         Index           =   2
         Left            =   -73320
         TabIndex        =   46
         Top             =   1635
         Width           =   6735
         VariousPropertyBits=   671107099
         MaxLength       =   100
         Size            =   "11880;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   300
         Index           =   1
         Left            =   -73320
         TabIndex        =   45
         Top             =   1335
         Width           =   6735
         VariousPropertyBits=   671107099
         MaxLength       =   100
         Size            =   "11880;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text20 
         Height          =   300
         Index           =   0
         Left            =   -73320
         TabIndex        =   44
         Top             =   1035
         Width           =   6735
         VariousPropertyBits=   671107099
         MaxLength       =   100
         Size            =   "11880;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text31 
         Height          =   930
         Left            =   -73290
         TabIndex        =   50
         Top             =   2880
         Width           =   6735
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "11880;1640"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         Caption         =   "ＩＤＳ報價:  1. 第一階段                    (           P)"
         Height          =   180
         Left            =   -69900
         TabIndex        =   162
         Top             =   465
         Width           =   3540
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "2. 第二階段                    (           P)"
         Height          =   180
         Left            =   -68865
         TabIndex        =   161
         Top             =   765
         Width           =   2505
      End
      Begin VB.Label lblsNP23 
         Caption         =   "約定期限:"
         Height          =   255
         Left            =   3840
         TabIndex        =   159
         Top             =   2312
         Width           =   975
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "( 例如：第 1,3,5-12 項 )"
         Height          =   180
         Left            =   -74460
         TabIndex        =   158
         Top             =   1920
         Width           =   1800
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "請求撤銷全部專利權"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   -71310
         TabIndex        =   157
         Top             =   480
         Width           =   1620
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "請求撤銷全部或部分請求項"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74730
         TabIndex        =   154
         Top             =   780
         Width           =   2160
      End
      Begin VB.Image Image1 
         Height          =   15
         Left            =   1920
         Top             =   2160
         Width           =   15
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   0
         Left            =   6705
         TabIndex        =   153
         Top             =   420
         Width           =   630
         VariousPropertyBits=   27
         Caption         =   "判發人"
         Size            =   "1111;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "客戶函判發人:"
         Height          =   180
         Left            =   4725
         TabIndex        =   152
         Top             =   417
         Width           =   1125
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "引證前案檔案數量:"
         Height          =   180
         Left            =   5490
         TabIndex        =   151
         Top             =   1020
         Width           =   1485
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "被請求撤銷發明(新型)專利權"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74730
         TabIndex        =   150
         Top             =   510
         Width           =   2490
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "項"
         Height          =   180
         Index           =   5
         Left            =   -71670
         TabIndex        =   149
         Top             =   1065
         Width           =   180
      End
      Begin VB.Label Label29 
         Caption         =   "審查委員編號:"
         Height          =   180
         Left            =   5445
         TabIndex        =   121
         Top             =   2925
         Width           =   1125
      End
      Begin VB.Label Label27 
         Caption         =   "審查委員名稱:"
         Height          =   180
         Left            =   120
         TabIndex        =   120
         Top             =   2925
         Width           =   1125
      End
      Begin VB.Label lblDispDate 
         AutoSize        =   -1  'True
         Caption         =   "機關發文日:"
         Height          =   180
         Left            =   2430
         TabIndex        =   119
         Top             =   417
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "各式修正補件書:"
         Height          =   180
         Left            =   4605
         TabIndex        =   115
         Top             =   2640
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "主管機關:"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   114
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label Label10 
         Caption         =   "對造案件數代號:"
         Height          =   255
         Left            =   -71970
         TabIndex        =   113
         Top             =   735
         Width           =   1500
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "PS:1.其他來函印定稿時,請輸入進度備註! 2.開庭時間,第X法庭"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   1020
         TabIndex        =   112
         Top             =   3195
         Width           =   4770
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "是否列印客戶通知函:       (N:不印)"
         Height          =   180
         Index           =   1
         Left            =   135
         TabIndex        =   104
         Top             =   3870
         Width           =   2625
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "是否修改通知函內容:       (Y:是)"
         Height          =   180
         Index           =   2
         Left            =   3195
         TabIndex        =   103
         Top             =   3870
         Width           =   2445
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "代理人通知日:"
         Enabled         =   0   'False
         Height          =   180
         Index           =   1
         Left            =   6120
         TabIndex        =   102
         Top             =   2349
         Visible         =   0   'False
         Width           =   1128
      End
      Begin VB.Label Label36 
         Caption         =   "對造名稱(日):"
         Height          =   252
         Index           =   0
         Left            =   -74760
         TabIndex        =   101
         Top             =   2532
         Width           =   1212
      End
      Begin VB.Label Label35 
         Caption         =   "對造名稱(英):"
         Height          =   252
         Left            =   -74760
         TabIndex        =   100
         Top             =   2232
         Width           =   1212
      End
      Begin VB.Label Label34 
         Caption         =   "對造名稱(中):"
         Height          =   252
         Left            =   -74760
         TabIndex        =   99
         Top             =   1932
         Width           =   1212
      End
      Begin VB.Label Label33 
         Caption         =   "對造案件名稱(日):"
         Height          =   255
         Left            =   -74760
         TabIndex        =   98
         Top             =   1635
         Width           =   1455
      End
      Begin VB.Label Label32 
         Caption         =   "對造案件名稱(英):"
         Height          =   255
         Left            =   -74760
         TabIndex        =   97
         Top             =   1335
         Width           =   1455
      End
      Begin VB.Label Label31 
         Caption         =   "對造案件名稱(中):"
         Height          =   255
         Left            =   -74760
         TabIndex        =   96
         Top             =   1035
         Width           =   1455
      End
      Begin VB.Label Label30 
         Caption         =   "對造號數:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   95
         Top             =   735
         Width           =   855
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "點數:"
         Height          =   180
         Index           =   2
         Left            =   2400
         TabIndex        =   94
         Top             =   2355
         Width           =   405
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "大陸費用:"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   93
         Top             =   2355
         Width           =   765
      End
      Begin VB.Label Label43 
         Caption         =   "進度備註:"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   3180
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "本案期限:"
         Height          =   255
         Left            =   120
         TabIndex        =   91
         Top             =   4140
         Width           =   855
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "機關文號:"
         Height          =   180
         Left            =   120
         TabIndex        =   74
         Top             =   1005
         Width           =   765
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "來函性質:"
         Height          =   180
         Left            =   120
         TabIndex        =   73
         Top             =   705
         Width           =   765
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "下一程序:"
         Height          =   180
         Left            =   3840
         TabIndex        =   72
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "一般來函日期:"
         Height          =   180
         Left            =   120
         TabIndex        =   71
         Top             =   417
         Width           =   1125
      End
      Begin VB.Label Label22 
         Caption         =   "來函期限:"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "法定期限:"
         Height          =   255
         Left            =   3840
         TabIndex        =   69
         Top             =   1748
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "本所期限:"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   1755
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "承辦期限:"
         Height          =   255
         Left            =   3840
         TabIndex        =   67
         Top             =   2036
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "承辦人:"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label38 
         Caption         =   "專利權消滅日:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   64
         Top             =   432
         Width           =   1212
      End
      Begin VB.Label Label39 
         Caption         =   "是否閉卷:"
         Height          =   255
         Index           =   0
         Left            =   -71970
         TabIndex        =   63
         Top             =   435
         Width           =   855
      End
      Begin VB.Label Label40 
         Caption         =   "(Y:閉卷)"
         Height          =   255
         Left            =   -70680
         TabIndex        =   62
         Top             =   435
         Width           =   855
      End
      Begin VB.Label Label46 
         Caption         =   "案件備註:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   61
         Top             =   2832
         Width           =   852
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   3
         Left            =   6300
         TabIndex        =   60
         Top             =   696
         Width           =   1956
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "3440;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   4
         Left            =   2430
         TabIndex        =   59
         Top             =   705
         Width           =   1260
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "1879;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   7
         Left            =   2400
         TabIndex        =   58
         Top             =   2064
         Width           =   1344
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "2371;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "是否算案件數:               (N:不算)"
         Height          =   180
         Left            =   6120
         TabIndex        =   65
         Top             =   2073
         Width           =   2568
      End
   End
   Begin VB.Label lblNote 
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   6780
      TabIndex        =   163
      Top             =   420
      Width           =   1920
      WordWrap        =   -1  'True
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   54
      Top             =   330
      Width           =   5535
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "9763;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "是否閉卷"
      Height          =   180
      Index           =   3
      Left            =   6300
      TabIndex        =   107
      Top             =   900
      Width           =   720
   End
   Begin VB.Label lblPA57 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Left            =   7260
      TabIndex        =   106
      Top             =   900
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(Y:閉卷)"
      Height          =   180
      Index           =   4
      Left            =   7980
      TabIndex        =   105
      Top             =   900
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請日:"
      Height          =   180
      Left            =   3960
      TabIndex        =   90
      Top             =   675
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   89
      Top             =   360
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   3525
      TabIndex        =   88
      Top             =   90
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   87
      Top             =   60
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   86
      Top             =   645
      Width           =   765
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   1
      Left            =   1155
      TabIndex        =   85
      Top             =   675
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3334;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   2
      Left            =   5040
      TabIndex        =   84
      Top             =   645
      Width           =   1140
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2011;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "收文號"
      Height          =   180
      Index           =   1
      Left            =   150
      TabIndex        =   83
      Top             =   885
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日"
      Height          =   180
      Index           =   2
      Left            =   3960
      TabIndex        =   82
      Top             =   915
      Width           =   900
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   5
      Left            =   1155
      TabIndex        =   81
      Top             =   915
      Width           =   1890
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3334;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   6
      Left            =   5040
      TabIndex        =   80
      Top             =   915
      Width           =   1140
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2011;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm04010504_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Morgan 2021/12/20 改成Form2.0 (Text29,Text30,Text31,Text20,Combo1,Label3)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Modified by Morgan 2021/8/12 智財法院-->智商法院
'2005/7/5整理
Option Explicit

Dim strReceiveNo As String, strTemp As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String, cp(1 To T_CP) As String
Dim pa() As String, cp() As String
Dim intWhere As Integer, intLastRow As Integer
Dim m_NewCP09 As String
Public MPa9 As String
'Add By Cheng 2001/12/12
Dim m_bln_FieldValid As Boolean 'False:欄位值無效, True:欄位值有效
'Add  By Cheng 2002/12/12
Dim m_CP05 As String '原收文日
Dim m_CP27 As String '原發文日
Dim m_CP09 As String '原收文號
'92.6.20 ADD BY SONIA
Dim m_NP08 As String '通知公開之下一程序實審期限或年費已登記之下一程序年費期限

'Add by Morgan 2004/2/18
'若承辦人是王協理且未發文則要發EMail通知
Dim m_stCP09 As String, m_stCP14 As String
'Add by Morgan 2005/1/19 一案兩請之新型相關資料
Dim m_bolIsDualApp As Boolean, m_stCaseNo As String, m_stCertNo As String, m_stAppNo As String, m_stCaseName As String, m_stUPA(1 To 4) As String
Dim m_bolGiveUpUtility As Boolean '一案兩請放棄新型 Added by Morgan 2014/7/21
Dim m_DualAppNP22 As String 'Add by Morgan 2005/3/16 一案兩請新型接洽單
Dim m_strRetSheet2NP07 As String '第二張回覆單案件性質
'Add by Morgan 2006/6/26
Dim m_901CP09 As String '901內部收文之總收文號
Dim m_901CP12 As String '901內部收文之業務區
Dim m_901CP13 As String '901內部收文之智權人員
'Add by Morgan 2006/8/16
Dim m_bolSaveCheck As Boolean '是否為存檔前檢查
'Add by Morgan 2004/2/6
Dim stCP12 As String, stCP13 As String
Dim bolCancelClose As Boolean 'Add by Morgan 2007/5/4 是否取消閉卷
'Add by Morgan 2009/12/1
Dim m_bolFMP As Boolean, stNP23 As String, stCP48Desc As String
Dim m_si880017 As Single '補件期限按鈕回傳狀態
Dim m_strUnSaveData As String '待新增補文件期限
'end 2009/12/1
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/05/17 是否為寰華案
Dim m_blnClosed As Boolean '是否閉卷'Add By Sindy 2012/3/5
Dim oChk As CheckBox 'Added by Morgan 2012/10/5
Dim m_CustX07166 As Boolean   '2012/11/26 add by sonia 是否順德(含關係企業)專利案件
Dim str941ReceiveNo As String '2012/11/26 ADD BY SONIA 內部收文941收文號(P非台灣案若原承辦人非工程師則改抓國內案承辦人) P-093775
Public str941CP14 As String   '內部收文941收文號及承辦人
Dim strNew404CP43 As String '先延期後收文的收文號 Added by Morgan 2013/10/2
Dim strCP43toNP06 As String, strCP43toNPpty As String, strCP43toCP09 As String, strCP43toNp08 As String, strCP43toNp09 As String, strCP43toPty As String 'Added by Lydia 2025/03/05 點選收文的相關收文號之下一程序：是否已收文、法限、所限、案件性質
'Added by Morgan 2014/1/14
Public m_DocNo As String
Public m_AppNo As String
'end 2014/1/14
'Added by Morgan 2014/4/17
Public m_DocWord As String
Public m_DeadLine As String
Public m_NewCP10 As String
'end 2014/4/17
Dim m_PropertyCode As String 'Added by Morgan 2014/8/27
Dim mPty1004() As String, mPtyNo As String  'Add by Lydia 2014/11/26 台灣案主管機關來函,針對1004(延期受理)
Dim m_bolCCC As Boolean 'Added by Lydia 2015/04/30 通知函是否為副本
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END
Dim strChoseBase As String 'Added by Lydia 2017/05/09 被視為未主張的基礎案
Dim strBasePD06 As String  'Added by Lydia 2017/05/09 被視為未主張的基礎案(只有優先權號)
Dim m_USCaseNo As String 'Added by Morgan 2019/5/28 相關美國案本所案號(提IDS用)
Dim m_bolNoCP27 As Boolean '不上發文 Added by Morgan 2020/1/17
Dim m_bolReKeyInOK As Boolean 'Added by Morgan 2020/8/12 是否與2次確認期限一致
Dim m_bolW2001XCase As Boolean 'Added by Morgan 2021/9/22 是否顧服組W2001的4家客戶案件
Dim m_CustX69365 As Boolean 'Added by Morgan 2021/10/6 是否長庚醫院案件
Dim m_str1998CP09 As String 'Added by Morgan 2021/10/6 轉公文收文號
Dim m_bolFMPNoPrint As Boolean 'Added by Morgan 2023/4/10 FMP案是否列印中文定稿
Dim m_bolBPFCase As Boolean '是否寶齡富錦 Added by Morgan 2023/6/27
Dim m_bolAutoBCP As Boolean, m_strEV02 As String 'Added by Morgan 2023/9/13 是否自動內部收文,特殊計件值
Dim m_bolEngCase As Boolean 'Added by Morgan 2024/4/29 是否工程師承辦
Dim bolPCTReport As Boolean 'Added by Morgan 2024/7/22 改全域變數

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 30) As String
Dim Jjj As Integer
'Add By Cheng 2003/04/18
Dim strExceptField As String
Dim ii As Integer
Dim jj As Integer
Dim m_119NP09 As String   '2006/7/3 ADD BY SONIA PCT案進入國家階段119法定期限
Dim strExp As String      '2010/11/15 add by sonia
Dim strDayOrMon As String '2010/11/15 add by sonia

   ' 90.06.28 modify by louis 印定稿以新收文號
   'EndLetter ET01, strReceiveNo, ET03, strUserNum
   EndLetter ET01, m_NewCP09, ET03, strUserNum
   
   Jjj = 1
   
   If Text8 <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','下一程序','" & Text8 & "')"
      Jjj = Jjj + 1
   End If
   
   If m_strRetSheet2NP07 <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','下一程序2','" & m_strRetSheet2NP07 & "')"
      Jjj = Jjj + 1
   End If
   
   If CheckStr(Label3(3).Caption) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','下一程序名稱','" & Label3(3).Caption & "')"
      Jjj = Jjj + 1
   End If
   
   If CheckStr(Text14(0).Text) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','本所期限','" & Text14(0).Text & "')"
      Jjj = Jjj + 1
   End If
   If CheckStr(Text14(1).Text) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','法定期限','" & Text14(1).Text & "')"
      Jjj = Jjj + 1
   End If
   If CheckStr(Text26.Text) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','專利權消滅日','" & Text26.Text & "')"
      Jjj = Jjj + 1
   End If
   
   '20080915 add by Toni 大-->台 1202
    '審查意見通知函 大陸領証費
    'Modified by Morgan 2012/12/27 +最後通知1227
    If (Text7 = "1202" Or Text7 = "1227") And ET03 = "18" Then
         If pa(8) = "1" Then
            Text21(0) = "15000"   '2010/12/29 原為12000
         ElseIf pa(8) = "3" Then
            Text21(0) = "10500"    '2010/12/29 原為7500
         End If
   End If
   'end Toni 20080915
   
   If CheckStr(Text21(0).Text) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','大陸領證費','" & Text21(0).Text & "')"
      Jjj = Jjj + 1
      
   'Added by Morgan 2024/4/29
   ElseIf m_bolEngCase Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','大陸領證費','|#(紅字)【請程序報價】#|')"
      Jjj = Jjj + 1
   'end 2024/4/29
   End If
   
   'Add by Morgan 2006/4/12
   If CheckStr(Text21(1).Text) <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','點數','" & Text21(1).Text & "')"
      Jjj = Jjj + 1
      
      'Added by Morgan 2016/6/13 非台灣信函進度要存報價
      If Val(Text21(0)) <> 0 Or Val(Text21(1)) <> 0 Then
         PUB_UpdateLP2930 m_NewCP09, Text21(0), Text21(1).Text
      End If
      'end 2016/6/13
   End If
   'Add by Morgan 2005/1/20
   If m_bolIsDualApp = True Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','新型專利號數','" & m_stCertNo & "')"
      Jjj = Jjj + 1
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','新型本所號','" & m_stCaseNo & "')"
      Jjj = Jjj + 1
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','新型申請案號','" & m_stAppNo & "')"
      Jjj = Jjj + 1
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
         " VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','新型案件名稱','" & m_stCaseName & "')"
      Jjj = Jjj + 1
   End If
   
   'Add By Cheng 2002/05/30
   If Me.Text14(2).Enabled Then
      If CheckStr(Text14(2).Text) <> "" Then
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
            "','其他公告日','" & Text14(2).Text & "')"
         Jjj = Jjj + 1
      End If
   End If
   If Me.Text32.Enabled Then
      If CheckStr(Text32.Text) <> "" Then
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
            "','已發文時間','" & Text32.Text & "')"
         Jjj = Jjj + 1
      End If
   End If
    'Modify By Cheng 2002/12/23
    '若為台灣案時
    '92.1.28 MODIFY BY SONIA
    'If pa(9) = 台灣國家代號 And (Me.Text7.Text = 通知補文件 Or Me.Text7.Text = 通知修正 Or Me.Text7.Text = 其他來函) Then
    If pa(9) = 台灣國家代號 Then
    '92.1.28 END
        If CheckStr(Me.Text29.Text) <> "" Then
           strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
              "','列印備註','" & ChgSQL(Me.Text29.Text) & "')"
           Jjj = Jjj + 1
        End If
    End If
    '2010/11/15 add by sonia
    If pa(9) = 台灣國家代號 And (cp(10) = "501" Or cp(10) = "503") Then
      strExp = Empty
      strDayOrMon = Empty
      
      strSql = "SELECT * FROM CASEFEE " & _
               "WHERE CF01 = '" & pa(1) & "' AND " & _
                     "CF02 = '" & pa(9) & "' AND " & _
                     "CF03 = '" & cp(10) & "' "
      Set RsTemp = New ADODB.Recordset
      RsTemp.CursorLocation = adUseClient
      RsTemp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If RsTemp.RecordCount > 0 Then
         ' 取得下一救濟程序名稱
         If IsNull(RsTemp.Fields("CF15")) = False Then
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
               "','下一程序名稱2','" & GetCaseTypeName(pa(1), RsTemp.Fields("CF15"), 0) & "')"
            Jjj = Jjj + 1
            ' 取得下一救濟程序主管機關
            strExc(0) = "SELECT CF10 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & RsTemp.Fields("CF15") & "'"
            intI = 1
            Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
                  "','下一程序3','" & "" & AdoRecordSet3.Fields(0) & "')"
               Jjj = Jjj + 1
            End If
         End If
      End If
      RsTemp.Close
      strSql = "SELECT * FROM CASEPROPERTYMAP " & _
               "WHERE CPM01 = '" & pa(1) & "' AND " & _
                     "CPM02 = '" & cp(10) & "' "
      Set RsTemp = New ADODB.Recordset
      RsTemp.CursorLocation = adUseClient
      RsTemp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If RsTemp.RecordCount > 0 Then
         ' 來函期限
         If IsNull(RsTemp.Fields("CPM07")) = False Then
            Select Case RsTemp.Fields("CPM07")
               Case "1": strExp = "文到當日"
               Case "2": strExp = "文到次日"
            End Select
         End If
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
            "','文到當日次日','" & "" & strExp & "')"
         Jjj = Jjj + 1
         ' 期限天數
         If IsNull(RsTemp.Fields("CPM08")) = False Then
            If IsEmptyText(RsTemp.Fields("CPM08")) = False Then
               strDayOrMon = RsTemp.Fields("CPM08") & "日"
            End If
         End If
         ' 期限月數
         If IsEmptyText(strDayOrMon) = True Then
            If IsNull(RsTemp.Fields("CPM09")) = False Then
               If IsEmptyText(RsTemp.Fields("CPM09")) = False Then
                  strDayOrMon = RsTemp.Fields("CPM09") & "個月"
               End If
            End If
         End If
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
            "','期限天數月數','" & "" & strDayOrMon & "')"
         Jjj = Jjj + 1
      End If
      RsTemp.Close
      Set RsTemp = Nothing
    End If
    '2010/11/15 end
    
    '91.12.27 add by sonia
    If Me.Combo2.Text <> "" Then
       strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
         "','來函主管機關','" & Me.Combo2.Text & "')"
       Jjj = Jjj + 1
    End If
    '91.12.27 end
    'Add By Cheng 2003/04/16
    '若有輸入分割費用
    If Me.Text25.Text <> "" Then
        strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
            "','分割費用','" & Me.Text25.Text & "')"
        Jjj = Jjj + 1
    End If
    'Add by Morgan 2009/9/3
    If Me.Text34.Text <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & "','分割點數','" & Me.Text34.Text & "')"
      Jjj = Jjj + 1
    End If
    If Me.Text35.Text <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & "','實審費用','" & Me.Text35.Text & "')"
      Jjj = Jjj + 1
    End If
    If Me.Text36.Text <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & "','實審點數','" & Me.Text36.Text & "')"
      Jjj = Jjj + 1
    End If
    'end 2009/9/3
    
    'Add By Cheng 2003/04/18
    '若有勾選檢索報告
    '2006/6/30 MODIFY BY SONIA 國際初審報告1216改於來函性質輸,故此處不判斷
    'If Me.Check1(0).Value = vbChecked Or Me.Check1(1).Value = vbChecked Or Me.Check1(2).Value = vbChecked Then
    If Me.Check1(1).Value = vbChecked Or Me.Check1(2).Value = vbChecked Then
        ii = 0
        For jj = 1 To Me.Check1.Count
            If Me.Check1(jj).Value = vbChecked Then ii = ii + 1
        Next jj
        If ii = 2 Then
            strExceptField = "「" & Me.Check1(1).Caption & "」及「" & Me.Check1(2).Caption & "」"
        Else
            strExceptField = "「" & Me.Check1(IIf(Me.Check1(1).Value = vbChecked, 1, 2)).Caption & "」"
        End If
        strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
            "','檢索報告','" & strExceptField & "')"
        Jjj = Jjj + 1
    End If
    '92.6.20 ADD BY SONIA
    '2008/9/11 add by Toni 大-->台 實體審查期限定稿
    If Text7 = 通知公開 And (ET03 = "01" Or ET03 = "02") Then
        strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
            "','其他日期','" & m_NP08 & "')"
        Jjj = Jjj + 1
    End If
    
    
   '92.6.20 END
   
    '93.3.6 add by sonia
    If Me.Text7.Text = "1218" And ET03 = "00" Then
       strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
          "','本所期限'," & m_NP08 & ")"
       Jjj = Jjj + 1
    End If
    '93.3.6 END
   
    '2006/7/3 ADD BY SONIA PCT國際初審報告1216之進入國家階段119期限
'2008/7/21 cancel by sonia 不印定稿由工程師處理
'    If pa(9) = "056" And Text7 = "1216" And ET03 = "00" Then
'       strExc(0) = "SELECT NP09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP06 IS NULL AND NP07='119'"
'       intI = 1
'       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'       If intI = 1 Then  ' 有119期限
'          m_119NP09 = RsTemp.Fields(0)
'          strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
'            "','其他日期','" & m_119NP09 & "')"
'          Jjj = Jjj + 1
'       End If
'    End If
'2008/7/21 end
   '2006/7/3 END
   
   'Add by Morgan 2006/12/14 官方發文日
   If Text28.Text <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
          "','官方發文日'," & DBDATE(Text28) & ")"
       Jjj = Jjj + 1
   End If
   'end 2006/12/14
   
   'Add by Morgan 2007/9/11
   '台灣、大陸及澳門之發明及設計案件,若有同時辦美國案(主案,未閉卷,未核准),則於通知審查意見通知書定稿中加入一段美國提IDS之提醒字眼
   'Modify by Morgan 2011/7/18 改控制美國案未領證未閉卷且需為發明案(設計不用)--郭
   'Modified by Morgan 2019/5/28 需輸入IDS報價，改存檔前檢查
   'If Text7 = 通知申復 And (pa(9) = "000" Or pa(9) = "020" Or pa(9) = "044") And (pa(8) = "1" Or pa(8) = "3") Then
   '   strExc(1) = PUB_GetUSCaseNo(pa(1), pa(2), pa(3), pa(4))
   '   If strExc(1) <> "" Then
      If m_USCaseNo <> "" Then
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
            "','美國案本所案號','" & m_USCaseNo & "')"
         Jjj = Jjj + 1
         
         'Modified by Morgan 2019/6/3 第１階段報價金額大於０才寫，定稿要控制不出該報價文字
         If Val(txtIDSFee(1)) > 0 Then
            strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
               "','IDS報價1','" & txtIDSFee(1) & "')"
            Jjj = Jjj + 1
         End If
         
         strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
            "','IDS報價2','" & txtIDSFee(2) & "')"
         Jjj = Jjj + 1
      End If
   'End If
   'end 2019/5/28
      
   'Added by Lydia 2017/05/09 視為未主張
   If Text7 = "1918" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
          "','未主張優先權號','" & strBasePD06 & "')"
       Jjj = Jjj + 1
   End If
   'end 2017/05/09
   'Added by Lydia 2025/03/05 台灣案增加延期受理定稿
   'Modified by Lydia 2025/04/29 去掉限制台灣案 pa(9) = 台灣國家代號
   If Text7 = "1004" And strCP43toNPpty <> "" Then
      strTxt(Jjj) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
          "VALUES ('" & ET01 & "','" & m_NewCP09 & "','" & ET03 & "','" & strUserNum & _
          "','下一程序收文性質','" & strCP43toNPpty & "')"
       Jjj = Jjj + 1
   End If
   'end 2025/03/05
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(Jjj - 1, strTxt) Then
   If Not ClsLawExecSQL(Jjj - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Sub

'Added by Morgan 2012/10/5
'Modified by Morgan 2013/1/14 增加舉發事項
Private Sub chkItem_Click(Index As Integer)
   Dim ii As Integer
      
   If Me.ActiveControl <> chkItem(Index) Then Exit Sub
   
   txtItemCount.Enabled = False
   txtItemList.Enabled = False
   txtYear(0).Enabled = False
   txtYear(1).Enabled = False
   txtMonth(0).Enabled = False
   txtMonth(1).Enabled = False
   txtDay(0).Enabled = False
   txtDay(1).Enabled = False
   
   If Index = 0 Or Index = 1 Then
      If chkItem(Index).Value = vbChecked Then
         For Each oChk In chkItem
            If oChk.Index <> Index Then
               oChk.Value = vbUnchecked
            End If
         Next
         
         Select Case Index
         Case 0
            txtItemCount.Enabled = True
            txtItemCount.SetFocus
            
         Case 1
            txtItemList.Enabled = True
            txtItemList.SetFocus
            If Left(txtItemList, 1) = "第" Then
               txtItemList.SelStart = 1
               txtItemList.SelLength = 0
            End If
         End Select
      End If
   ElseIf Index = 6 Then
      If chkItem(Index).Value = vbChecked Then
         For Each oChk In chkItem
            If oChk.Index <> Index Then
               oChk.Value = vbUnchecked
            End If
         Next
         
         txtYear(0).Enabled = True
         txtYear(0).SetFocus
         txtYear(1).Enabled = True
         txtMonth(0).Enabled = True
         txtMonth(1).Enabled = True
         txtDay(0).Enabled = True
         txtDay(1).Enabled = True
      End If
   ElseIf chkItem(Index).Value = vbChecked Then
      chkItem(0).Value = vbUnchecked
      chkItem(1).Value = vbUnchecked
      chkItem(6).Value = vbUnchecked
   End If
End Sub

'Add by Morgan 2009/12/1
Private Sub cmdDeadLine_Click()
   If Text8 <> "202" Then
      MsgBox "下一程序必須是補文件！"
      Text8.SetFocus
   Else
      If Text14(0) <> "" And Text14(1) <> "" Then
         ModifyAddDeadline1 cp(9), Text14(0), Text14(1), m_si880017, True, m_strUnSaveData
      Else
         MsgBox "請輸入補文件期限！"
         Text14(0).SetFocus
      End If
   End If
End Sub
Private Sub Process(Index As Integer)

Dim bolChk As Boolean, strTmp As String
'Add By Cheng 2002/12/18
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim mbolDual As Boolean, dStr01 As String 'Added by Lydia 2015/04/20
Dim m_StrTo As String, m_StrSub As String, m_StrCont As String 'Added by Lydia 2022/05/30 整理frm880005改用寄信模組
'Dim bolEngLetter As Boolean 'Added by Morgan 2023/5/10 是否產生工程師用定稿 'Removed by Morgan 2024/4/29 改用 m_bolEngCase 控制
Dim bolAdd15Days As Boolean 'Added by Morgan 2024/1/25 是否適用15天在途

   strTmp = ""
   Select Case Index
      Case 0 '確定
      
         
         'Add by Morgan 2007/5/4 若來函有期限但已閉卷
         bolCancelClose = False
         If Text14(1) <> "" Then
            'Modified by Morgan 2014/6/20
            'If Text27(0) = "Y" Then
            '   MsgBox "因為要管制期限，本案不可閉卷！", vbExclamation
            '   Exit Sub
            'ElseIf pa(57) = "Y" Then
            '   If MsgBox("本案目前為閉卷狀態，為管制期限將於存檔時取消閉卷，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
            '      Exit Sub
            If pa(57) = "Y" Then
               If Text27(0) = "Y" Then
                  If MsgBox("本案目前為閉卷狀態，為管制期限將於存檔時取消閉卷，是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                     Me.SSTab1.Tab = 1
                     Text27(0).SetFocus
                     Exit Sub
                  End If
                  Text27(0) = ""
               End If
            'end 2014/6/20
               bolCancelClose = True
            End If
         End If
         'end 2007/5/4
         
         'Added by Lydia 2015/12/17 對於已經閉卷的案件,後續若有官方來函是無期限的,全部都詢問user是否要取消閉卷,由user來判斷
         If pa(57) = "Y" And Text27(0) = "Y" And Text14(1) = "" And bolCancelClose = False Then
            If MsgBox("本案目前為閉卷狀態，您輸入的是無期限的來函，是否要取消閉卷？", vbYesNo + vbDefaultButton1) = vbYes Then
               Text27(0) = ""
               bolCancelClose = True
            End If
         End If
         'end 2015/12/17
         
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         m_bolSaveCheck = True
         If TxtValidate = False Then
            m_bolSaveCheck = False
            Exit Sub
         End If
         m_bolSaveCheck = False
      
         'Add by Morgan 2007/6/12
         If txtDispDate.Visible = True Then
            If txtDispDate = "" Then
               MsgBox "機關發文日不可空白！", vbCritical
               txtDispDate.SetFocus
               Exit Sub
            ElseIf ChkDate(txtDispDate) = False Then
               txtDispDate.SetFocus
               Exit Sub
            End If
         End If
         'end 2007/6/12
         
         ' 90.07.31 modify by louis (來函性質不可為空白)
         If IsEmptyText(Text7) = True Then
            MsgBox "來函性質不可為空白 !", vbCritical
            Text7.SetFocus
            Exit Sub
         End If
         'Add By Cheng 2002/03/11
         If Me.Text14(0).Text <> "" Then
            If Len(Me.Text14(0).Text) = 8 Then
               'Modified by Morgan 2016/6/8 所限不可輸入時自動改為系統日否則無法存檔
               If Text14(0).Enabled Then
                  If Me.Text14(0).Text < strSrvDate(1) Then
                     MsgBox "本所期限不可小於系統日!!!", vbExclamation
                     If Me.Text14(0).Enabled = True Then
                        Me.Text14(0).SetFocus
                        Me.Text14(0).SelStart = 0
                        Me.Text14(0).SelLength = Len(Me.Text14(0).Text)
                     End If
                     Exit Sub
                  End If
               Else
                  Text14(0).Text = strSrvDate(1)
               End If
               'end 2016/6/8
               
            Else
               If Val(Me.Text14(0).Text) + 19110000 < strSrvDate(1) Then
                  'Modified by Morgan 2016/6/8 所限不可輸入時自動改為系統日否則無法存檔
                  If Text14(0).Enabled Then
                     MsgBox "本所期限不可小於系統日!!!", vbExclamation
                     If Me.Text14(0).Enabled = True Then
                        Me.Text14(0).SetFocus
                        Me.Text14(0).SelStart = 0
                        Me.Text14(0).SelLength = Len(Me.Text14(0).Text)
                     End If
                     Exit Sub
                  Else
                     Text14(0).Text = strSrvDate(2)
                  End If
                  'end 2016/6/8
               End If
            End If
         End If
         'Add By Cheng 2001/12/12
         '依來函性質判斷抓本所期限及法定期限
         If Len(Trim(Me.Text14(0).Text)) <= 0 Or Len(Trim(Me.Text14(1).Text)) <= 0 Then
            m_bln_FieldValid = False
            Text7_Validate False
            If Not m_bln_FieldValid Then Exit Sub
         End If
         '檢查法定期限欄位有效性
         m_bln_FieldValid = False
         Text14_Validate 1, False
         If Not m_bln_FieldValid Then Exit Sub
         
         'add by sonia 2018/5/3 1901通知退費不必輸法定期限故不檢查
         If Text7 = "1901" Then
            If Text8 <> "" And Text14(0) = "" Then
               MsgBox "下一程序不為空白時，本所期限不可為空白 !", vbCritical
               If IsEmptyText(Text14(0)) = True Then
                   If Me.Text14(0).Enabled = True Then
                       Text14(0).SetFocus
                   End If
               End If
               Exit Sub
            End If
         Else
         'end 2018/5/3
            If Text8 <> "" And (Text14(0) = "" Or Text14(1) = "") Then
               ' 90.07.31 modify by louis (更新顯示訊息及設定Focus)
               MsgBox "下一程序不為空白時，本所期限與法定期限不可為空白 !", vbCritical
               If IsEmptyText(Text14(0)) = True Then
                  'Modify By Cheng 2002/05/30
                   If Me.Text14(0).Enabled = True Then
                       Text14(0).SetFocus
                   Else
                       If Me.Option4(0).Value Then
                           If Me.Text10.Enabled = True Then Me.Text10.SetFocus
                       ElseIf Me.Option4(1).Value Then
                           If Me.Text11.Enabled = True Then Me.Text11.SetFocus
                       Else
                           If Me.Text12.Enabled = True Then Me.Text12.SetFocus
                       End If
                   End If
               Else
                   'Modify By Cheng 2002/05/30
                   If Me.Text14(1).Enabled = True Then
                       Text14(1).SetFocus
                   Else
                       If Me.Option4(0).Value Then
                           Me.Text10.SetFocus
                       ElseIf Me.Option4(1).Value Then
                           Me.Text11.SetFocus
                       Else
                           Me.Text12.SetFocus
                       End If
                   End If
               End If
               Exit Sub
            End If
         End If         'add by sonia 2018/5/3
                 
        'Add By Cheng 2002/10/30
         If Text8 = "" And (Text14(0) <> "" Or Text14(1) <> "") And Me.Text7.Text <> "1004" Then
            MsgBox "若有期限時, 下一程序不可為空白!!!", vbExclamation + vbOKOnly
            Me.Text8.SetFocus
            Exit Sub
        End If
        
         '若來函性質為被異議(1801)被舉發(1802), 則要檢查對造資料
         'modify by Morgan 2005/6/1 第三人提起技術報告(1810)要檢查對造代號及中文名
         '2008/4/23
         'If Me.Text7.Text = "1802" Or Me.Text7.Text = "1801" Then
         '2008/4/23 modify by sonia 第三人提起技術報告(1810)改為無下一程序時才檢查對造代號及中文名
         'If Me.Text7.Text = "1802" Or Me.Text7.Text = "1801" Or Me.Text7.Text = "1810" Then
         If Me.Text7.Text = "1802" Or Me.Text7.Text = "1801" Or (Me.Text7.Text = "1810" And Me.Text8.Text = "") Then
            'Modify By Cheng 2002/12/03
'            If Len(Trim("" & Me.Text19.Text)) <= 0 And Len(Trim("" & Me.Text20(0).Text)) <= 0 And _
'               Len(Trim("" & Me.Text20(1).Text)) <= 0 And Len(Trim("" & Me.Text20(2).Text)) <= 0 And _
'               Len(Trim("" & Me.Text20(3).Text)) <= 0 And Len(Trim("" & Me.Text20(4).Text)) <= 0 And _
'               Len(Trim("" & Me.Text20(5).Text)) <= 0 Then
            'Modify by Morgan 2010/2/8 中英日任一有就好
            'If Len(Trim("" & Me.Text19.Text)) <= 0 Or Len(Trim("" & Me.Text23.Text)) <= 0 Or _
               Len(Trim("" & Me.Text20(3).Text)) <= 0 Then
            'Modify by Morgan 2011/9/20 台灣才要輸對造案件數代號--玲玲
            'If Trim(Text19) = "" Or Trim(Text23) = "" Or Trim(Text20(3) & Text20(4) & Text20(5)) = "" Then
            If Trim(Text19) = "" Or (pa(9) = 台灣國家代號 And Trim(Text23) = "") Or Trim(Text20(3) & Text20(4) & Text20(5)) = "" Then
               MsgBox "請輸入本案件的對造資料 !", vbCritical
               SSTab1.Tab = 1
               If Trim(Text19) = "" Then
                  Text19.SetFocus
               '2010/5/19 MODIFY BY SONIA
               'ElseIf Trim(Text23) = "" Then
               'Modify by Morgan 2011/9/20 台灣才要輸對造案件數代號--玲玲
               'ElseIf Len(Trim("" & Me.Text23.Text)) <= 1 Then
               ElseIf pa(9) = 台灣國家代號 And Len(Trim("" & Me.Text23.Text)) <= 1 Then
                  Text23.SetFocus
               Else
                  Text20(3).SetFocus
               End If
               Exit Sub
            Else
               PUB_ChkCustNameExist Text20(3), Text20(4), Text20(5)
            End If
         End If
         '94.1.5 add by sonia 若來函性質為受理技術報告申請, 則要檢查對造號數及對造案件數代號
         If Me.Text7.Text = "1405" Then
            If Len(Trim("" & Me.Text19.Text)) <= 0 Or Len(Trim("" & Me.Text23.Text)) <= 1 Then
               MsgBox "請輸入本案件的對造號數及對造案件數代號 !", vbCritical
               Me.SSTab1.Tab = 1
               Me.Text23.SetFocus
               Exit Sub
            End If
         End If
         '94.1.5 end
         'Add By Cheng 2002/05/06
         '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
         If Len(Me.Text14(0).Text) > 0 And Len(Me.Text17.Text) > 0 Then
            If Val(Me.Text14(0).Text) < Val(Me.Text17.Text) Then
               MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
               Exit Sub
            End If
         End If
         
         m_bolAutoBCP = False: m_strEV02 = "" 'Added by Morgan 2023/9/13
        '92.10.11 ADD BY SONIA
        '若為大陸案, 案件性質為通知申復, 通知修正, 被舉發理由, 通知口頭審理時, 大陸費用一定要輸入
'        If pa(9) = 大陸國家代號 And (Me.Text7.Text = 通知申復 Or Me.Text7.Text = 通知修正 Or Me.Text7.Text = 被舉發理由) Then
        '2010/4/8 modify by sonia FMP不必輸入
        'If pa(9) = 大陸國家代號 And (Me.Text7.Text = 通知申復 Or Me.Text7.Text = 通知修正 Or Me.Text7.Text = 被舉發理由 Or Me.Text7.Text = "1401") Then
        'Modified by Lydia 2025/03/18 從常數「通知申復」改回1202審查意見通知函
        If Not m_bolFMP And pa(9) = 大陸國家代號 And (Me.Text7.Text = "1202" Or Me.Text7.Text = 通知修正 Or Me.Text7.Text = 被舉發理由 Or Me.Text7.Text = "1401") Then
            If Trim(Me.Text21(0).Text) = "" Then
               'Added by Morgan 2024/4/29 工程師承辦時可不輸費用
               If m_bolEngCase Then
                  Text21(1) = ""
                  
               'Modified by Morgan 2025/10/20 通知口頭審理點選的收文號為未發文的口頭審理時可不必輸入費用
               'Else
               ElseIf Not (Text7 = "1401" And cp(10) = "408" And m_CP27 = "") Then
               'end 2025/10/20
               'end 2024/4/29
                  MsgBox "請輸入大陸費用!!!", vbExclamation + vbOKOnly
                  Me.Text21(0).SetFocus
                  Text21_GotFocus (0)
                  Exit Sub
                  
               End If
            'Add by Morgan 2006/4/12 通知補正, 審查意見通知函 若有大陸費用時一定要輸入點數
            Else
               Select Case Text7
                  Case "1201", "1202"
                     If Text21(1) = "" Then
                        MsgBox "請輸入點數!!!", vbExclamation + vbOKOnly
                        Me.Text21(1).SetFocus
                        Text21_GotFocus (1)
                        Exit Sub
                     End If
                     
                     'Added by Morgan 2023/9/13
                     '(1201)通知補正,(1202)通知審查意見，只要費用與點數鍵入值為零，便自動於列印客戶通知函上N
                     '詢問是否要修改基數
                     '有IDS報價時除外
                     'Modified by Morgan 2023/11/29 排除X3880503資策會(憑帳單請款)--品薇
                     If Val(Text21(0)) = 0 And Val(Text21(1)) = 0 And m_USCaseNo = "" And InStr(pa(26), "X3880503") = 0 Then
                        If Text15(0).Text <> "N" Then Text15(0).Text = "N"
                        m_bolAutoBCP = True
                        If MsgBox("本案將自動內部收文" & Label3(3) & "，請確認是否要修改基數？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                           Do
                              m_strEV02 = InputBox("內部收文" & Label3(3) & "基數：")
                              If m_strEV02 = "" Then
                                 If MsgBox("是否確定不修改基數？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                                    Exit Do
                                 End If
                              ElseIf Not IsNumeric(m_strEV02) Then
                                 MsgBox "基數輸入錯誤！", vbCritical
                              Else
                                 Exit Do
                              End If
                           Loop
                        End If
                     End If
                     'end 2023/9/13
                     
               End Select
            End If
        End If
        'Modify By Cheng 2002/12/23
        '若為國內案, 案件性質為通知補文件, 通知修正, 其他來函時, 進度備註一定要輸入
'        'Add By Cheng 2002/11/19
'        If Me.Text7.Text = 通知補文件 Or Me.Text7.Text = 通知修正 Or Me.Text7.Text = 其他來函 Then
        If pa(9) = 台灣國家代號 And (Me.Text7.Text = 通知補文件 Or Me.Text7.Text = 通知修正 Or Me.Text7.Text = 其他來函) Then
            If Trim(Me.Text29.Text) = "" Then
                MsgBox "請輸入進度備註!!!", vbExclamation + vbOKOnly
                Me.Text29.SetFocus
                Text29_GotFocus
                Exit Sub
            End If
        End If
        'Add By Cheng 2003/01/15
        '若PCT案且來函性質為檢索報告(1209)
'2008/7/15 CANCEL BY SONIA 不印定稿由工程師處理
'        If pa(9) = "056" And Me.Text7.Text = 檢索報告 Then
'            '若未勾選檢索報告
'            'Modify By Cheng 2003/04/18
''            If Me.Option2(0).Value = False And Me.Option2(1).Value = False Then
'            '2006/6/30 MODIFY BY SONIA 國際初審報告1216改於來函性質輸,故此處不判斷
'            'If Me.Check1(0).Value = vbUnchecked And Me.Check1(1).Value = vbUnchecked And Me.Check1(2).Value = vbUnchecked Then
'            If Me.Check1(1).Value = vbUnchecked And Me.Check1(2).Value = vbUnchecked Then
'               MsgBox "請勾選檢索報告!!!", vbExclamation + vbOKOnly
'               Me.SSTab1.Tab = 1
'               Exit Sub
'            End If
'            'Add by Morgan 2006/12/15
'            If Text28.Text = "" Then
'               MsgBox "請輸入官方發文日!!!", vbExclamation + vbOKOnly
'               SSTab1.Tab = 1
'               Text28.SetFocus
'               Exit Sub
'            End If
'        End If
'2008/7/15 END
        'Add By Cheng 2003/03/26
        '檢查機關文號
        If pa(9) = 台灣國家代號 Then
            If Me.Text9.Tag = Me.Text9.Text Then
                Me.SSTab1.Tab = 0
                MsgBox "請輸入機關文號!!!", vbExclamation + vbOKOnly
                Me.Text9.SetFocus
                Text9_GotFocus
                Exit Sub
            End If
        End If
        'Add By Cheng 2003/04/16
        '若申請國家非台灣, 且來函性質為通知補正(1201)
        'Modify by Morgan 2009/7/21 不必限制案件性質(已控制是否顯示)
        'If pa(9) <> 台灣國家代號 And Me.Text7.Text = "1201" Then
         'Modify by Morgan 2009/9/3 加欄位改控制
         If frm307.Visible = True Then
            If Me.Text24.Text = "Y" And Me.Text25.Text = "" Then
                MsgBox "請輸入分割費用!!!", vbExclamation + vbOKOnly
                Me.Text25.SetFocus
                Text25_GotFocus
                Exit Sub
            End If
            'Add by Morgan 2009/9/8 分割一定要輸實審費用--敏惠
            'Modify by Morgan 2009/11/11 發明才要 --玲玲
            If Me.Text24.Text = "Y" And pa(8) = "1" And Me.Text35.Text = "" Then
                MsgBox "請輸入實審費用!!!", vbExclamation + vbOKOnly
                Me.Text35.SetFocus
                Text35_GotFocus
                Exit Sub
            End If
         End If
        'End If
        
        'Added by Lydia 2015/04/30
        If pa(1) = "P" And pa(9) = 台灣國家代號 And m_DocNo <> "" And m_bolCCC = True Then
           '電子公文輸入若為副本收受者應無須管制期限
        Else
            'Add By Cheng 2004/02/10
            '檢查來函期限--日期
            If Me.Option4(2).Value = True Then
                If Me.Text12.Text = "" Then
                    MsgBox "請輸入來函期限!!!", vbExclamation + vbOKOnly
                    Me.Text12.SetFocus
                    Exit Sub
                End If
            End If
            'End
        End If
        '2015/04/30
        
        'Add by Amy 2022/09/30 cp36放寬至200 檢查大小(存對造號數&對造案件數代號)
        If pa(9) = 台灣國家代號 And Text7.Text = "1810" And Text9 = "2" And Text8 <> "" Then
        Else
            If CheckLengthIsOK(Text19.Text & Text23.Text, 200, False) = False Then
                MsgBox "對造號數+對造案件數代號 " & vbCrLf & _
                MsgText(9205) & "200" & MsgText(9206) & "!", vbExclamation + vbOKOnly
                Me.Text19.SetFocus
                Exit Sub
            End If
        End If
        'end 2022/09/30
        'Add by Morgan 2004/3/10
        '來函性質為’1004’延期受理時，若frm04010504_2點選資料相關總收文號為’C’類時，檢查本案期限至少要選取一筆
        '若frm04010504_2點選資料相關總收文號非’C’類時，則檢查本案期限不可選取任何一筆。
        If Text7 = "1004" Then
            strNew404CP43 = cp(43) 'Added by Morgan 2013/10/2
            'Add by Morgan 2006/8/16
            If Text14(0) = "" Or Text14(1) = "" Then
               MsgBox "來函性質為 [延期受理] 時，本所期限與法定期限不可為空白 !", vbCritical
               Exit Sub
            End If
            'end 2006/8/16
            
            Dim bolCheck As Boolean, iRow As Integer
            Dim stNP08 As String, stNP09 As String, stChkDate As String, stPty As String 'Added by Morgan 2023/4/17
            
            bolCheck = False
            With MSHFlexGrid1
               For iRow = 1 To .Rows - 1
                  If .TextMatrix(iRow, 0) = "v" Then
                    bolCheck = True
                    'Added by Morgan 2023/4/17
                    stNP08 = .TextMatrix(iRow, 2) '所限 Added by Morgan 2023/4/19
                    stNP09 = .TextMatrix(iRow, 3) '法限
                    stPty = .TextMatrix(iRow, 1) '案件性質
                    'end 2023/4/17
                    Exit For
                  End If
               Next
            End With
            If Left(cp(43), 1) = "C" Then
               If bolCheck = False Then
                  'Added by Morgan 2013/10/2 先延期後收文再延期受理時
                  'Modified by Morgan 2022/12/29 +cp01條件否則會抓到CFP的IDS
                  'Modified by Lydia 2025/03/05 改成先抓取
                  'strExc(0) = "select cp09,cp07,cpm04,cp06,cp07 from nextprogress,caseprogress,casepropertymap where np01='" & cp(43) & "' and cp43(+)=np01 and cp10(+)=np07" & _
                  '   " and cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp27 is null and cpm01(+)=cp01 and cpm02(+)=cp10"
                  'intI = 1
                  'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  'If intI = 1 Then
                  '   strNew404CP43 = RsTemp(0)
                  '   '法限 Added by Morgan 2023/4/17
                  '   stNP08 = RsTemp("cp06") '所限 Added by Morgan 2023/4/19
                  '   stNP09 = RsTemp("cp07")
                  '   stPty = RsTemp("cpm04")
                  '   'end 2023/4/17
                  'Else
                  'end 2013/10/2
                  If strCP43toCP09 <> "" Then
                     strNew404CP43 = strCP43toCP09
                     stNP08 = strCP43toNp08
                     stNP09 = strCP43toNp09
                     stPty = strCP43toPty
                  Else
                  'end 2025/03/05
                     'Modified by Morgan 2012/2/1 加已收文提醒--玲玲
                     MsgBox "請先選取未收文期限的資料來做延期受理的處理!!!" & vbCrLf & vbCrLf & "( 若欲延期的程序已收文, 請先人工修改【延期】的相關收文號為已收文程序的收文號後再輸受理函!!! )", vbExclamation + vbOKOnly
                     Exit Sub
                     
                  End If 'Added by Morgan 2013/10/2 先延期後收文再延期受理時
               End If
               
            ElseIf Left(cp(43), 1) <> "C" And bolCheck = True Then
                MsgBox "不可選取任何一筆未收文期限的資料!!!", , vbExclamation + vbOKOnly
                Exit Sub
            'Added by Morgan 2024/1/29 補收文後延期
            Else
               strExc(0) = "select cp09,cp07,cpm04,cp06,cp07 from caseprogress,casepropertymap where cp09='" & cp(43) & "'" & _
                  " and cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp27 is null and cpm01(+)=cp01 and cpm02(+)=cp10"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strNew404CP43 = RsTemp(0)
                  stNP08 = RsTemp("cp06")
                  stNP09 = RsTemp("cp07")
                  stPty = RsTemp("cpm04")
               End If
            End If
            
            'Added by Morgan 2023/4/17
            '外專程序操作時只檢查但不更新(寰華案會輸入來函上含在途的期限系統管制則不含)
            If Pub_StrUserSt03 = "F22" Then
               If Text12 = "" Then
                  'Modified by Morgan 2024/1/26
                  'MsgBox "延期受理請輸入含在途的官方期限！", vbCritical
                  MsgBox "延期受理請輸入大陸官方期限！", vbCritical
                  'end 2024/1/26
                  Exit Sub
               End If
               
               'Modified by Morgan 2024/1/25 官方發文日20240120(含)之後的來函，電子送件案件無在途期間
               'stChkDate = CompDate(2, -15, Text12)
               bolAdd15Days = False
               If Left(cp(43), 1) = "C" Then
                  strExc(0) = "select  a.cp133,b.cp118 from caseprogress a,caseprogress b where a.cp09='" & cp(43) & "' and b.cp09(+)=a.cp43 and a.cp133 is not null and b.cp09<'C'"
               Else
                  strExc(0) = "select  b.cp133,c.cp118 from caseprogress a,caseprogress b,caseprogress c where a.cp09='" & cp(43) & "' and b.cp09(+)=a.cp43 and c.cp09(+)=b.cp43 and b.cp09>'C' and b.cp133 is not null and c.cp09<'C'"
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If Not (RsTemp("cp133") >= 20240120 And RsTemp("cp118") = "Y") Then
                     bolAdd15Days = True
                  End If
               End If
               If bolAdd15Days Then
                  stChkDate = CompDate(2, -15, Text12)
               Else
                  stChkDate = DBDATE(Text12)
               End If
               'end 2024/1/26
               
               stNP08 = DBDATE(stNP08)
               stNP09 = DBDATE(stNP09)
               If stNP09 > stChkDate Then
                  'Modified by Morgan 2024/1/26
                  'MsgBox "本所" & stPty & "的法定期限大於大陸官方期限(含在途的法限-15天)，請確認！", vbCritical
                  If bolAdd15Days Then
                     MsgBox "本所" & stPty & "的法定期限大於大陸官方期限(含在途的法限-15天)，請確認！", vbCritical
                  Else
                     MsgBox "本所" & stPty & "的法定期限大於大陸官方期限，請確認！", vbCritical
                  End If
                  'end 2024/1/26
                  Exit Sub
               'Added by Morgan 2023/4/19 寰華案延期受理來函期限帶系統的期限
               Else
                  Text14(0) = TransDate(stNP08, 1)
                  Text14(1) = TransDate(stNP09, 1)
               'end 2023/4/19
               End If
            End If
            'strExc (1)
            
        End If
        'Add end 2004/3/10
         
        '2013/8/16  add by sonia
        If Me.Text7.Text = "1506" And Label3(1) = "行政訴訟" Then
           If Me.Text30.Text = Val(Left(DBDATE(m_CP27), 4) - 1911) & "年度行專訴字第號" Then
              MsgBox "行政訴訟之智慧局答辯函, 請輸入法院案號, 以便將來智商法院來函可查詢!!!", vbExclamation + vbOKOnly
              Me.Text30.SetFocus
              Exit Sub
           End If
        End If
        'End
         
         'Add by Morgan 2004/6/28
         If Text7 = "1003" Then
            Dim stCP09 As String
            If CheckCP(pa, stCP09) = True Then
               If stCP09 <> "" Then
                  If MsgBox("本案尚有補文件【" & stCP09 & "】未發文，確定要繼續？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
               End If
            Else
               Exit Sub
            End If
         End If
         'End
         
         '2008/11/28 add by sonia 1912通知已轉他所詢問是否閉卷
         'modify by sonia 2016/12/26 +1916解除代理人
         If (Text7 = "1912" Or Text7 = "1916") And Text27(0) = "" Then
            If MsgBox("通知已轉他所或解除代理人，是否要閉卷？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
               Text27(0) = "Y"
            End If
         End If
         '2008/11/28 End
         
         'Add by Morgan 2005/1/19 檢查是否發明案核准且為一案兩請,先作大陸
         m_bolIsDualApp = False
         m_bolGiveUpUtility = False 'Added by Morgan 2014/7/21
         'Modified by Morgan 2012/8/15 102新法 +台灣,程序人員視來函內容判斷是否為准前擇一通知
         'Modified by Morgan 2012/12/27 +最後通知1227
         'Modified by Morgan 2015/3/25 台灣新增1232一案兩請通知擇一申復,原控制取消
         'If (Text7 = "1202" Or Text7 = "1227") And (pa(9) = "020" Or pa(9) = "000") And pa(8) = "1" Then
         'Modified by Morgan 2021/1/20 寰華案除外--敏莉
         'If (Text7 = "1202" Or Text7 = "1227") And pa(9) = "020" And pa(8) = "1" Then
         If (Text7 = "1202" Or Text7 = "1227") And pa(9) = "020" And pa(8) = "1" And Left(Pub_StrUserSt03, 1) <> "F" Then
         'end 2021/1/20
         'end 2015/3/25
            'Modified by Morgan 2016/6/8 已閉卷或已收文放棄專利權不必問 --郭
            'm_bolIsDualApp = PUB_IsDualApply(pa, m_stUPA, m_stCaseNo, m_stCertNo, m_stAppNo, m_stCaseName)
            m_bolIsDualApp = PUB_IsDualApply(pa, m_stUPA, m_stCaseNo, m_stCertNo, m_stAppNo, m_stCaseName, True)
            'Add by Morgan 2007/3/28
            If m_bolIsDualApp = True Then
               'Added by Morgan 2016/6/8
               '已閉卷或已收文放棄專利權不必問 --郭
               strExc(0) = "select 1 from patent where pa01='" & m_stUPA(1) & "' and pa02='" & m_stUPA(2) & "' and pa03='" & m_stUPA(3) & "' and pa04='" & m_stUPA(4) & "' and pa57 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  m_bolIsDualApp = False
               Else
               'end 2016/6/8
                  If MsgBox("本案為一案兩請，是否放棄" & IIf(pa(9) = "020", "實用", "") & "新型案的專利權？", vbYesNo + vbDefaultButton1) = vbNo Then
                     m_bolIsDualApp = False
                  End If
               End If 'Added by Morgan 2016/6/8
               
'Removed by Morgan 2016/6/8
'               'Added by Morgan 2014/7/21
'               '台灣一案兩請案申請日 101.1.1~102.6.12 者不必管控放棄專利權,發證時自動閉卷
'               If pa(9) = "000" And m_bolIsDualApp = True Then
'                  'Modified by Morgan 2014/10/27 102/6/13 以後的也不必管控放棄專利權--玲玲
'                  'If DBDATE(pa(10)) >= "20120101" And DBDATE(pa(10)) < "20130613" Then
'                  If DBDATE(pa(10)) >= "20120101" Then
'                  'end 2014/10/27
'                     'm_bolIsDualApp = False 'Removed by Morgan 2014/8/5 定稿還是一案兩請的
'                     m_bolGiveUpUtility = True
'                  End If
'               End If
'               'end 2014/7/21
'end 2016/6/8

            End If
            'end 2007/3/28
         
         'Added by Morgan 2015/3/25
         '新增1232一案兩請通知擇一申復
         ElseIf Text7 = "1232" Then
            m_bolIsDualApp = PUB_IsDualApply(pa, m_stUPA, m_stCaseNo, m_stCertNo, m_stAppNo, m_stCaseName)
            If m_bolIsDualApp = False Then
               MsgBox "一案兩請未建立關聯案！", vbCritical
               Exit Sub
            ElseIf pa(9) = "000" Then
               m_bolGiveUpUtility = True
            End If
         'end 2015/3/25
         
         End If
         
         '2008/11/28 add by sonia 非台灣案1912通知已轉他所詢問是否計算結餘
         'modify by sonia 2016/12/26 +1916解除代理人
         If (Text7 = "1912" Or Text7 = "1916") Then
            'Modified by Lydia 2015/03/03 +pa01,pa02,pa03,pa04
            'modify by sonia 2024/12/20 改為自動上可結餘不必詢問
            'Pub_EndModCashMsg pa(9), pa(1), pa(2), pa(3), pa(4)
            If pa(9) <> "000" Then bolEndModCash = True  '自動上結餘日
         End If
         '2008/11/28 end
         
         'Added by Lydia 2017/05/09 後案官方來函性質「視為未主張」，若有兩個以上的優先權，則 show出該兩個優先權讓user勾選哪一個被視為未主張，若只有一個優先權，就直接自優先權資料處移至案件備註(刪除PriDate,寫案件備註)，該案若有以優先權計算期限的，則請重新計算期限。
         strChoseBase = ""
         strBasePD06 = ""
         If Text7 = "1918" Then
            Set RsTemp = PUB_ReadPDStateNew(pa, cp(10))
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
         
         'Add By Sindy 2020/7/20
         If m_strIR01 <> "" Then
            '下載信件檔
            'Modify By Sindy 2022/11/10 + IIf(pa(9) <> 台灣國家代號, "PAT", "RX")
            If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", IIf(pa(9) <> 台灣國家代號, "PAT", "RX"), , True) = False Then
               Exit Sub
            End If
            'Add By Sindy 2022/7/21
            'Mark by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail:可取消外專系統收件區，key來函承辦人掛程序人員，則按確定，信件會再打開一次的設定。
            'If Left(Pub_StrUserSt03, 2) = "F2" Then
            '   If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
            '      Exit Sub
            '   End If
            'End If
            '2022/7/21 END
            'end 2023/05/17
         End If
         '2020/7/20 END
         
         '2012/11/26 add by sonia
         If m_CustX07166 = False Then m_CustX07166 = PUB_CheckX07166Remind(cp(1), cp(9), Text7.Text, str941CP14)
         '2012/11/26 END
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         PUB_CtrlDateAlert m_NewCP09 'Add by Morgan 2011/7/14
         
'Remove by Morgan 2007/4/18 P 的不用--郭
'         'Add by Morgan 2007/1/18 申請人為"福興"時彈訊息
'         If InStr(pa(26) & pa(27) & pa(28) & pa(29) & pa(30), "X43179") > 0 And InStr("1006,1202,1203,1205,1206,1209", Text7.Text) > 0 Then
'            MsgBox "請影印一份OA交付智權人員【" & GetStaffName(stCP13) & "】！"
'         End If
'         'end 2007/1/18
'end 2007/4/18
         
        'Add by Morgan 2004/2/18
        '若承辦人是王協理且未發文則要發EMail通知
        'Removed by Morgan 2023/9/14 改存檔時寫MailCache
        'If m_stCP14 = "71011" Then
        '    Call PUB_SendMail(strUserNum, m_stCP14, m_stCP09, "分案通知")
        'End If
        'end 2023/9/14
        
        '2012/11/26 add by sonia 順德及其關係企業案件,若承辦人是王協理且未發文則要發EMail通知
        'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
        If m_CustX07166 = True And str941CP14 = "99050" Then
           Call PUB_SendMail(strUserNum, "99050", str941ReceiveNo, "分案通知")
        End If
        '2012/11/26 end
      
         'add by toni 2008/10/24
         'modify by sonia 2020/4/29 加入台灣案才發MAIL(2019/9/5大陸案增加1211性質)
         If (Label3(1) = "準備程序" Or Label3(1) = "言詞辯論") And pa(9) = "000" Then
            If Text7 = "1210" Or Text7 = "1211" Then
               '2008/11/12 ADD BY SONIA
               Dim m_CP14 As String
               m_CP14 = ""
               'strSql = "select * from CASEPROGRESS where CP09='" & strReceiveNo & "'"
               strSql = " select CASEPROGRESS.*,NVL(ST04,' ') as ST04 from CASEPROGRESS,STAFF where CP14 = ST01(+) and CP09='" & strReceiveNo & "'"
               'Add by Lydia 2014/10/02 請同時發給m_CP14(承辦人=>專利工程師), 但先檢查m_CP14 若為離職人員則改發 特殊人員-P開庭通知離職工程師轉發
               CheckOC
               adoRecordset.CursorLocation = adUseClient
               adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset.RecordCount > 0 Then
                 If adoRecordset.Fields("ST04") = "2" Then '離職人員代理
                  m_CP14 = Pub_GetSpecMan("P開庭通知離職工程師轉發")
                 Else
                  m_CP14 = CheckStr(adoRecordset.Fields("cp14"))
                 End If
               End If
               '2008/11/12 END
               
               'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
               'Load frm880005
               'If Len(Trim(m_CP14)) > 0 Then 'Add by Lydia 2014/10/02
               '  frm880005.txtEmail(0).Text = Pub_GetSpecMan("Q") & ";" & stCP13 & ";" & m_CP14
               'Else
               '  frm880005.txtEmail(0).Text = Pub_GetSpecMan("Q") & ";" & stCP13
               'End If
               '''2008/11/12 modify by sonia 再抓時間地點,法院案號,承辦人抓工程師
               'frm880005.txtEmail(1).Text = "開庭通知--來函案件：" & Text2 & "-" & Text3 & "-" & Text4 & "-" & Text5
              ' frm880005.txtEmail(2).Text = "本所案號：" & Text2 & "-" & Text3 & "-" & Text4 & "-" & Text5 & vbCrLf & _
                                             "案件名稱：" & Me.Combo1.Text & vbCrLf & _
                                             "案件性質：" & Label3(4) & vbCrLf & _
                                             "申請人　：" & GetCustomerName(pa(26)) & vbCrLf & _
                                             "承辦人　：" & GetStaffName(m_CP14) & vbCrLf & _
                                             "智權人員　：" & GetStaffName(stCP13) & vbCrLf & _
                                             "法定期限：" & DBYEAR(Text14(1).Text) - 1911 & " 年 " & DBMONTH(Text14(1).Text) & " 月 " & DBDAY(Text14(1).Text) & " 日 " & vbCrLf & _
                                             "時間地點：" & Text29 & vbCrLf & _
                                             "法院案號：" & Text9
               'frm880005.Form_Activate: DoEvents
               'frm880005.cmdOK_Click 0: DoEvents
               'Modify By Sindy 2023/12/8 法律所調整內專行政訴訟開庭通知之系統通知信也請一併轉陳亮之; 商標一併調整
               'Modified by Lydia 2024/10/30 串法律所案源資料，抓出法律所案件有承辦人且收文日最大的進度，抓承辦人及所有出庭律師。
               'm_StrTo = Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & stCP13 & IIf(Trim(m_CP14) <> "", ";" & m_CP14, "")
               m_StrTo = PUB_GetLosCL02list(Text2, Text3, Text4, Text5)
               m_StrTo = IIf(m_StrTo <> "", m_StrTo & ";", "") & Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & stCP13 & IIf(Trim(m_CP14) <> "", ";" & m_CP14, "")
               'end 2024/10/30
               
               m_StrSub = "開庭通知--來函案件：" & Text2 & "-" & Text3 & "-" & Text4 & "-" & Text5
               'Modified by Morgan 2025/2/21 m_CP14->left(m_CP14,5), 73022退休前過渡期會設通知2人
               m_StrCont = "本所案號：" & Text2 & "-" & Text3 & "-" & Text4 & "-" & Text5 & vbCrLf & _
                                             "案件名稱：" & Me.Combo1.Text & vbCrLf & _
                                             "案件性質：" & Label3(4) & vbCrLf & _
                                             "申請人　：" & GetCustomerName(pa(26)) & vbCrLf & _
                                             "承辦人　：" & GetStaffName(Left(m_CP14, 5)) & vbCrLf & _
                                             "智權人員　：" & GetStaffName(stCP13) & vbCrLf & _
                                             "法定期限：" & DBYEAR(Text14(1).Text) - 1911 & " 年 " & DBMONTH(Text14(1).Text) & " 月 " & DBDAY(Text14(1).Text) & " 日 " & vbCrLf & _
                                             "時間地點：" & Text29 & vbCrLf & _
                                             "法院案號：" & Text9
               PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
               'end 2022/05/30
            End If
         End If
'         'end 2008/10/24
         
       
        'Add by Lydia 2014/10/20 台灣案通知補文件自動作內部收文控制，發E-MAIL通知智權同仁及承辦工程師
        'Remove by Lydia 2019/05/03 取消官方來文輸入(1003)通知補文件系統自動發E-MAIL通知
'        If cp(1) = "P" And pa(9) = "000" And Text7.Text = "1003" And Text8 = "202" Then
'            If Len(pa(5)) > 0 Then strExc(3) = Trim(pa(5))
'            If Len(strExc(3)) = 0 Then strExc(3) = Trim(pa(6))
'            If Len(strExc(3)) = 0 Then strExc(3) = Trim(pa(7))
'            'oSubject
'            strExc(0) = Text2 & "-" & Text3 & "-" & Text4 & "-" & Text5 & "「" & strExc(3) & "」-->" & Label3(4)
'            'oContext
'            strExc(1) = "本所案號：" & Text2 & "-" & Text3 & "-" & Text4 & "-" & Text5 & vbCrLf & _
'                       "案件名稱：" & strExc(3) & vbCrLf & _
'                       "案件性質：" & Label3(4) & vbCrLf & _
'                       "申請人　：" & GetCustomerName(pa(26)) & vbCrLf & _
'                       "本所期限：" & ChangeTStringToTDateString(Trim(Text14(0))) & vbCrLf & _
'                       "來函內容：請至卷宗區參看官方來函"
'            '收件者
'            strExc(2) = stCP13 & ";" & strExc(9) '---智權人員和承辦工程師（使用特殊設定檔AB202）
'
'            PUB_SendMail strUserNum, strExc(2), "", strExc(0), strExc(1)
'
'            '2015/1/14 ADD BY SONIA 加訊息顯示智權人員,以便程序後續處理(陳玲玲P-110089)
'            MsgBox "本案之智權人員為 ( " & GetStaffName(stCP13) & " ) !!! "
'        End If
        'end 2019/05/03
        
         'Added by Morgan 2023/5/10 工程師承辦的1201 通知修正,1002  核駁,1202 審查意見通知函 也要產生定稿以便撰寫信函使用
         'Modified by Morgan 2023/6/27 寶齡富錦且工程師承辦的來函除外
         'Removed by Morgan 2024/4/29 改用 m_bolEngCase 控制
         'If Text15(0).Text = "N" And Not m_bolFMP And Text16 <> strUserNum And Not m_bolBPFCase Then
         '   If Text7 = "1201" Or Text7 = "1202" Then
         '      bolEngLetter = True
         '   End If
         'End If
         'end 2024/4/29
         'end 2023/5/10
        
         'Modified by Morgan 2023/5/10 +bolEngLetter
         'Modified by Morgan 2024/4/29
         'If Text15(0).Text <> "N" Or bolEngLetter Then '通知函
         If Text15(0).Text <> "N" Or m_bolEngCase Then '通知函
         'end 2024/4/29
            If Text15(1).Text = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            ' 90.06.28 modify by louis
            'Select Case cp(10)
            Select Case Text7
               'Add by Morgan 2007/2/13
               Case "1221" '通知申復
                  '台灣新型技術報告通知申復
                  If pa(9) = 台灣國家代號 And pa(8) = "2" And cp(10) = "421" Then
                     strTmp = "02"
                  Else
                     strTmp = "01"
                  End If
                  'end 2007/2/13
               'Modified by Morgan 2012/8/27 102新法 +1227最後通知
               'Modified by Morgan 2015/3/25 +1232一案兩請通知擇一申復
               'Modified by Lydia 2025/03/18 從常數「通知申復」改回1202審查意見通知函
               Case "1202", "1227", "1232"
                  If pa(9) = 台灣國家代號 Then '台灣 1
                     '大-->台 審查意見通知函定稿  ADD BY TONI 2008/09/12
                     If PUB_CheckCuNation(pa(26), Text2, Text3, Text4, Text5) = "1" Then
                        strTmp = "18"
                        'Added by Morgan 2012/8/15
                        If m_bolIsDualApp = True Then
                           strTmp = "20"
                        End If
                        'end 2012/8/15
                     Else
                        strTmp = "01"
                        'Added by Morgan 2012/8/15
                        If m_bolIsDualApp = True Then
                           strTmp = "21"
                           'Added by Morgan 2014/8/5
                           'Modified by Morgan 2015/11/24 通知擇一申復不必判斷申請日
                           'Modified by Morgan 2016/9/26 改回通知擇一申復還是要判斷申請日--郭雅娟
                           'If DBDATE(pa(10)) >= "20130613" Or Text7 = "1232" Then
                           If DBDATE(pa(10)) >= "20130613" Then
                           'end 2016/9/26
                              strTmp = "22"
                              'Added by Lydia 2016/02/02 一案兩請之發明通知擇一申復時,倘若該新型案已結案(或閉卷),則請帶此類定稿。
                              If PUB_IsDualApplyCom(pa, m_stUPA, m_stCaseNo, m_stCertNo, m_stAppNo, m_stCaseName) = False Then
                                 strTmp = "24"
                              End If
                           End If
                           'end 2014/8/5
                        End If
                        'end 2012/8/15
                     End If
                  Else
                     strTmp = "16"             '大陸 16
                     'Add by Morgan 2005/1/19 一案兩請定稿
                     If m_bolIsDualApp = True Then
                        strTmp = "17"
                     'Add by Morgan 2009/7/22 加分割定稿
                     ElseIf Me.Text24.Text = "Y" Then
                        strTmp = "05"
                     'Add by Morgan 2009/7/23
                     ElseIf cp(10) = "107" Then
                        strTmp = "19"
                     End If
                  End If
               'modify by sonia 2018/8/14 +1812通知聽證
               Case 通知面詢, "1812"
                  If pa(9) = 台灣國家代號 Then '台灣 1
                     strTmp = "01"
                  Else
                     strTmp = "05"             '大陸 5
                     Select Case cp(10)
                        Case "804"
                           strTmp = "06"  '無效宣告答辯
                        Case "803"
                           strTmp = "07"  '無效宣告
                           
                        'Added by Morgan 2025/10/20 大陸案口頭審理已收文未發文
                        Case "408"
                           If m_CP27 = "" Then
                              strTmp = "08"
                           End If
                        'end 2025/10/20
                     End Select
                  End If
               'Modified by Lydia 2025/07/23 拿掉被異議理由1801；整理所有國內對客戶的通知函定稿：協理確定要刪除的定稿
               Case 被舉發理由
                  If pa(9) = 台灣國家代號 Then '台灣 3
                     strTmp = "03"
                  Else
                     strTmp = "04"             '大陸 4
                  End If
               Case 發回補理由
                  If pa(9) = 台灣國家代號 Then '台灣 1
                     strTmp = "01"
                  Else
                     strTmp = "06"             '大陸 6
                  End If
               Case 通知領證
                  'strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) And pa(2) And pa(3) And pa(4)) & " AND CP10='" & 被異議理由 & "'"
                  'intI = 1
                  'Set rsTemp = clslawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  'If intI = 1 Then '曾被異議 1
                  '   strTmp = "01"
                  'Else '未被異議 0
                  '   strTmp = "00"
                  'End If
                   If pa(9) = 台灣國家代號 Then '台灣 7
                     strTmp = "07"
                  Else
                     strTmp = "02"             '大陸 2
                  End If
               Case 專利權消滅
                  If pa(9) = 台灣國家代號 Then '台灣 8
                     strTmp = "08"
                  Else                         '大陸

                  End If
               Case 其他來函
                  If pa(9) = 台灣國家代號 Then '台灣
                     If Text8 <> "" Then
                        strTmp = "14"           '有期限 14
                     Else
                        strTmp = "04"           '無期限 04
                     End If
                  Else                         '大陸
                     'Add by Morgan 2006/8/18 審查員依職權修改通知
                     'Mark by Lydia 2025/07/23 整理所有國內對客戶的通知函定稿：協理確定要刪除的定稿
                     'If Combo3.ListIndex = 0 Then
                     '   strTmp = "13"
                     'Else
                     'end 2025/07/23
                        '2008/8/28 modify by sonia 加無期限定稿
                        'strTmp = "02"
                        If Text8 <> "" Then
                           strTmp = "02"         '有期限 02
                        Else
                           strTmp = "05"         '無期限 05
                        End If
                        '2008/8/28 end
                     'End If 'Mark by Lydia 2025/07/23
                  End If
               Case 撤銷原處分
                  If pa(9) = 台灣國家代號 Then '台灣 15
                     '2010/11/12 MODIFY BY SONIA
                     'strTmp = "15"
                     Select Case cp(10)
                        Case "501"    '訴願
                           strTmp = "15"
                        Case "503"    '行政訴訟
                           strTmp = "16"
                        Case "507"    '行政訴訟上訴
                           strTmp = "17"
                     End Select
                     '2010/11/12 END
                  Else
                     strTmp = "02"             '大陸 2
                  End If
               'Add By Cheng 2002/06/19
               Case 准予延緩公告
                  If pa(9) = 台灣國家代號 Then '台灣 21
                     strTmp = "21"
                  Else '大陸 1
                     strTmp = "01"
                  End If
               'Add By Cheng 2002/12/28
               Case 通知智慧局答辯函, "1508"   '92.9.18 增加 1508 by sonia
                  If pa(9) = 台灣國家代號 Then '台灣 14
                     strTmp = "14"
                  Else
                     'Memoed by Morgan 2019/4/17 P-07-1506-02(台->大 復審委員會答辯函)更正為不掛號--郭
                     strTmp = "02"             '大陸 2
                  End If
               Case 通知審查中
                  If pa(9) = 台灣國家代號 Then '台灣 20
                     strTmp = "20"
                  Else
                     strTmp = "02"             '大陸 2
                  End If
               'Add By Cheng 2003/01/15
               Case 檢索報告, 1216        '2006/6/30 加1216 BY SONIA
                  '2008/7/15 CANCEL BY SONIA 不印定稿由工程師處理
                  'If pa(9) = "056" Then
                  '   strTmp = "00"             'PCT 00
                  'End If
                  '2008/7/15 END
               '92.6.20 ADD BY SONIA
               Case 通知公開
                 strExc(0) = "SELECT NP08 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP06 IS NULL AND NP07='" & 實體審查 & "'"
                 intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                 If intI = 1 Then  ' 有實體審查期限
                     If PUB_CheckCuNation(pa(26), Text2, Text3, Text4, Text5) = "1" Then   '2008/09/11 add by Toni 大-->台 實體審查期限定稿
                           strTmp = "02"
                           m_NP08 = RsTemp.Fields(0)
                        Else
                       strTmp = "01"
                       m_NP08 = RsTemp.Fields(0)
                     End If
                 Else              ' 無實體審查期限
                     If PUB_CheckCuNation(pa(26), Text2, Text3, Text4, Text5) = "1" Then '2008/09/11 add by Toni 大-->台 無實體審查期限定稿
                        strTmp = "03"
                        m_NP08 = ""
                     Else
                        strTmp = "00"
                        m_NP08 = ""
                     End If
                 End If
                 
               '93.3.6 ADD BY SONA
               Case "1218"
                 strExc(0) = "SELECT NP08 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP06 IS NULL AND NP07='" & 年費 & "'"
                 intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                 If intI = 1 Then  ' 有下一年年費期限
                    strTmp = "00"
                    m_NP08 = RsTemp.Fields(0)
                 Else              ' 無下一年年費期限
                    strTmp = "01"
                    m_NP08 = ""
                 End If
               '93.3.6 END
               'Add by Morgan 2004/12/13 通知現場勘察
               Case "1212"
                  If pa(9) = 台灣國家代號 Then
                     '若無下一程序
                     If Me.Text8.Text = "" Then
                         strTmp = "04" '無期限
                     Else
                        '94.1.24 MODIFY BY SONIA 通知現場勘察都要印回覆單
                        ''若點選進度是申請現場勘察則不印回覆單
                        'If cp(10) = "219" Then
                        '   strTmp = "01"
                        ''若點選進度非申請現場勘察則要印回覆單
                        'Else
                        '   strTmp = "00"
                        'End If
                        strTmp = "00"
                        '94.1.24 END
                     End If
                  Else
                     strTmp = "02"
                  End If
               '2008/4/23 ADD BY SONIA 第三人提起技術報告(1810)
               Case "1810"
                  If pa(9) = 台灣國家代號 And pa(8) = "2" Then
                     If Text8 = "" Then    '無下一程序期限
                        strTmp = "04"
                     Else
                        strTmp = "05"      '有申復期限
                     End If
                  End If
               '2008/4/23 END
               'Added by Lydia 2017/05/09 視為未主張
               Case "1918"
                   strTmp = "00"
               'Added by Lydia 2017/09/29 初審報告英譯文
               Case "1233"
                   strTmp = "00"
                   'Added by Morgan 2021/10/21 有實審發文出不同定稿
                   If PUB_ChkCPExist(cp, "416", 2) = True Then
                     strTmp = "01"
                   End If
                   'end 2021/10/21
               'Added by Lydia 2025/03/05 台灣案增加延期受理定稿
               Case "1004"
                   'Added by Lydia 2025/04/29 debug: 區分FMP案, 大-->台
                   If PUB_CheckCuNation(pa(26), Text2, Text3, Text4, Text5) = "1" Then
                         strTmp = "02"
                   Else
                      If m_bolFMP = True Then
                         strTmp = "01"
                      Else
                   'end 2025/04/29
                         strTmp = "00"
                      End If
                   End If
               'end 2025/03/05
               'end 2017/05/09
               Case Else '一般
                  If pa(9) = 台灣國家代號 Then '台灣 1
                     strTmp = "01"
                     '若申請國家為台灣,且無下一程序
                     If Me.Text8.Text = "" Then
                         strTmp = "04" '無期限
                     End If
                     'Add By Cheng 2002/12/28
                     '若前畫面點選的案件性質為準備程序(211), 參加訴訟(506), 言詞辯論(212)且為台灣案時  92.2.10 再加 面詢(408), 閱卷(410)
                     'modify by sonia 2018/8/14 +808聽證
                     If pa(9) = 台灣國家代號 And (cp(10) = 準備程序 Or cp(10) = 參加訴訟 Or cp(10) = 言詞辯論 Or cp(10) = 面詢 Or cp(10) = 閱卷 Or cp(10) = "808") Then
                         If rsA.State <> adStateClosed Then rsA.Close
                         Set rsA = Nothing
                         StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(cp(1) & cp(2) & cp(3) & cp(4)) & " And CP05 IS NOT NULL AND CP27 IS NULL AND CP57 IS NULL  "
                         rsA.CursorLocation = adUseClient
                         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                         '若有已收文未發文的資料
                         If rsA.RecordCount > 0 Then
                             strTmp = "03"
                         End If
                         If rsA.State <> adStateClosed Then rsA.Close
                         Set rsA = Nothing
                     End If
                  Else
                     strTmp = "02"             '大陸 2
                        'Add By Cheng 2003/04/16
                        '若來函性質為通知補正(1201), 且分割 05
                        'Modify by Morgan 2004/5/26
                        '大陸通知修正無費用 23
                        'If Me.Text7.Text = "1201" And Me.Text24.Text = "Y" Then strTmp = "05"
                        'Modify by Morgan 2009/7/21 不必限制案件性質(改判斷是否顯示)
                        'If Me.Text7.Text = "1201" Then
                     'Modify by Morgan 2009/9/3 加欄位改控制
                     If frm307.Visible = True Then
                        If Text24.Visible = True Then
                           If Me.Text24.Text = "Y" Then
                              strTmp = "05"
                              
                           'Modified by Morgan 2024/4/29 工程師承辦時都跑有費用的定稿但金額帶???
                           'ElseIf Val(Text21(0)) = 0 Then
                           ElseIf Val(Text21(0)) = 0 And m_bolEngCase = False Then
                           'end 2024/4/29
                              strTmp = "23"
                           End If
                        End If
                     End If
                     
                     'Added by Morgan 2021/10/15 寶齡富錦 Y55435 案件
                     If ChangeCustomerS(pa(75)) = "Y55435" And Text7 = "1234" Then
                        strTmp = "98"
                     End If
                     'end 2021/10/15
                  End If
            End Select
            '2008/7/15 MODIFY BY SONIA
            'StartLetter "07", strTmp
            'If strTmp <> "" Then StartLetter "07", strTmp 'Removed by Morgan 2016/7/25 移到下面
            '2008/7/15 END
            
            'Add by Morgan 2009/12/1
            If m_bolFMP Then
               'Modified by Morgan 2017/4/28 FMP已閉卷改出一般P案定稿以識別閉卷後來函--潘韻丞
               If pa(57) = "Y" Then
                  StartLetter "07", strTmp
                  NowPrint m_NewCP09, "07", strTmp, bolChk, strUserNum, 0, , , , 1, , , , , , , , m_NewCP09
               Else
               'end 2017/4/28
                
                '2010/4/8 modify by sonia 改用通函
                'If strTmp <> "" Then NowPrint m_NewCP09, "07", strTmp, bolChk, strUserNum, 0, , , , 1
                'Modified by Morgan 2016/7/25 FMP案都用通函
                'If strTmp <> "" Then NowPrint m_NewCP09, "07", "99", bolChk, strUserNum, 0, , , , 1, , , , , , , , m_NewCP09
                'Modified by Morgan 2023/4/10 FMP案有EMail通知的就不在列印紙本
                NowPrint m_NewCP09, "07", "99", bolChk, strUserNum, 0, , , , 1, , , , , , , , m_NewCP09, , , , , m_bolFMPNoPrint
               End If 'Added by Morgan 2017/4/28
            Else
            'end 2009/12/1
               If strTmp <> "" Then
                  StartLetter "07", strTmp 'Added by Morgan 2016/7/25
                  'Modified by Morgan 2021/10/6 若有轉公文則優先
                  'NowPrint m_NewCP09, "07", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , m_NewCP09
                  'Modified by Morgan 2023/5/10 +bolEngLetter
                  'Modified by Morgan 2024/4/29
                  'NowPrint m_NewCP09, "07", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , IIf(m_str1998CP09 <> "", m_str1998CP09, m_NewCP09), , , , , bolEngLetter
                  NowPrint m_NewCP09, "07", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , IIf(m_str1998CP09 <> "", m_str1998CP09, m_NewCP09), , , , , m_bolEngCase
                  'end 2024/4/29
               End If
               
               'Added by Morgan 2014/10/14
               '台灣新型通知修正改預設要出定稿但不列印(自動上已列印)--玲玲:因大部分都內部收文故不印,等判發確定要通知才退程序列印
               'Modified by Morgan 2015/7/13 +改請新型302
               'Modified by Morgan 2016/1/14 +不必再限制相關號的案件性質 --玲玲
               'If pa(9) = 台灣國家代號 And (cp(10) = "102" Or cp(10) = "302") And Text7.Text = "1201" Then
               If pa(9) = 台灣國家代號 And Text7.Text = "1201" Then
                  cnnConnection.Execute "update letterdemand set ld16='*' where ld18='" & m_NewCP09 & "'"
               End If
               'end 2014/10/14
            End If
            
         End If
        'Add By Cheng 2002/12/18
        '若輸入的來函性質為延期受理(1004), 則顯示相關總收文號的承辦人
        If Me.Text7.Text = "1004" Then
            'Modified by Morgan 2013/10/2
            'StrSQLa = "Select ST02 From CaseProgress,STAFF WHERE CP14=ST01(+) AND ST02 IS NOT NULL AND CP09=(SELECT CP43 FROM CASEPROGRESS WHERE CP09='" & m_CP09 & "' ) "
            StrSQLa = "Select ST02 From CaseProgress,STAFF WHERE CP14=ST01(+) AND ST02 IS NOT NULL AND CP09='" & strNew404CP43 & "'"
            'end 2013/10/2
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                MsgBox "本案相關總收文號的承辦人為 ( " & Trim(rsA.Fields(0).Value) & " ) !!! ", vbExclamation + vbOKOnly
            Else
                MsgBox "本案無相關總收文號的承辦人資料!!!", vbExclamation + vbOKOnly
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
         'Add by Lydia 2014/11/18 台灣案主管機關來函輸入，若此案有工程師未發文的程序，發E-MAIL通知工程師收到來函的內容
         'Modified by Lydia 2022/08/15 開放P大陸案
         'If pa(9) = "000" And pa(1) = "P" Then
         'Modified by Lydia 2022/10/11 經查此設定並不適用於外專及日專，故請協助排除FMP案
         'If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" Then
         If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" And m_bolFMP = False Then
           If Text7 <> "1004" Then
             'Added by Lydia 2015/04/20 一案兩請案件若新型已收到通知修正(1201)來函,請於通知承辦工程師已作內部收文的E-MAIL中提醒
              'PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), Text7, m_NewCP09
              If pa(8) = "2" And Text7.Text = "1201" Then
                 mbolDual = PUB_IsDualApply(pa, m_stUPA, m_stCaseNo, m_stCertNo, m_stAppNo, m_stCaseName, , True, dStr01)
              Else
                 m_stCaseNo = "": dStr01 = ""
              End If
              'Modified by Lydia 2022/08/16 +申請國家
              'PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), Text7, m_NewCP09, m_stCaseNo, dStr01
              'Modified by Morgan 2023/6/27 寶齡富錦且工程師承辦的來函已有通知,要排除本次來函以免重複通知
              PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), Text7, pa(9), m_NewCP09, m_stCaseNo, dStr01, m_bolBPFCase
           Else
               'end 2025/04/02
               'Add by Lydia 2014/11/26 台灣案主管機關來函,針對1004(延期受理)
               If Val(mPty1004(0)) >= 1 And Val(mPty1004(0)) <= 3 Then Check2mail1004 pa(), Label3(5), mPty1004(0) '非零->發mail通知
           End If
         End If
         'end 'Add by Lydia 2014/11/18
         
         'Add By Sindy 2016/10/5
         If Me.m_strIR01 <> "" Then
            Unload frm04010504_1
            Unload frm04010504_2
            Unload Me
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         ElseIf Me.m_DocNo <> "" Then
         'Added by Morgan 2014/1/14
         'If Me.m_DocNo <> "" Then
         '2016/10/5 END
            Unload frm04010504_1
            Unload frm04010504_2
            Unload Me
            frm04010516.GoNext
         Else
         'end 2014/1/14
         
            frm04010504_1.Show
            ' 90.07.17 modify by louis (回到第一個畫面全部清除)
            frm04010504_1.Clear
            Unload frm04010504_2
            Unload Me
         End If 'Added by Morgan 2014/1/14
         
         
      Case 1
         frm04010504_2.Show
         Unload Me
      Case 2
         Unload frm04010504_1
         Unload frm04010504_2
         Unload Me
   End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
   'Modified by Morgan 2018/8/23 +連點檢查 Ex:P-112288
   '不可用表單的enabled屬性控制,因若資料檢查有錯時設駐點命令導致執行階段錯誤
   '不可用按鈕的enabled屬性控制,因若存檔會卸載表單若再設值會導致表單再次被載入
   Static bolIsRunning As Boolean
   If bolIsRunning Then Exit Sub
   
   bolIsRunning = True
   Screen.MousePointer = vbHourglass
   Process Index
   Screen.MousePointer = vbDefault
   bolIsRunning = False
End Sub

Private Function FormSave() As Boolean
Dim intMax As Long, intStep As Integer, strTxt(1 To 10) As String, strTmp As String, i As Integer
'edit by nickc 2007/02/02
'Dim Ncp(1 To T_CP) As String
Dim Ncp() As String
ReDim Ncp(1 To TF_CP) As String

' 90.07.17 modify by louis (暫存列印接洽結案單下一程序序號)
Dim strProgressNo As String
'2008/7/21 ADD BY SONIA (暫存B類901告知代理人列印接洽結案單總收文號)
Dim strCP09_B As String
'Add By Cheng 2002/01/29
Dim BlnCheck As Boolean '判斷是否有勾選本案期限
Dim strDate1 As String '本所期限
Dim StrDate2 As String '法定期限
'Add By Cheng 2002/12/12
Dim blnAddNewNP As Boolean '是否新增下一程序資料
'Add by Morgan 2006/10/5
Dim bolAddBCP As Boolean '是否新增B類收文
'Dim strCP48 As String '承辦期限 'Remove by Lydia 2021/11/05
Dim st307Msg As String '分割案提醒訊息
'Add by Morgan 2009/12/1
Dim arrData() As String
'Dim bolPCTReport As Boolean 'Added by Morgan 2015/2/26 'Removed by Morgan 2024/7/22 改全域變數
Dim mRCno As String, mCCno As String, oSubject As String, oContext As String   'Add by Lydia 2014/10/16 FMP案會列印 C類接洽單, 請同時E-MAIL給畫面上之承辦人, 副本發給該員之工程師組別主管.
Dim strCMemo As String  'Added by Morgan 2020/9/15
Dim bolSavPdf As Boolean 'Added by Morgan 2018/10/2
Dim strCP20 As String 'Added by Morgan 2019/8/8
Dim bolReKeyInCase As Boolean 'Added by Morgan 2023/4/10
Dim strOldCP14 As String, strST04, strST02 'Added by Morgan 2023/9/14

On Error GoTo ErrorHandler
   'FormSave = True
   FormSave = False
   bolAddBCP = False
   bolReKeyInCase = False 'Added by Morgan 2023/4/10
   m_bolFMPNoPrint = False 'Added by Morgan 2023/4/10
   cnnConnection.BeginTrans
   
   'Removed by Morgan 2024/7/22 移到 Text7_Change
   'If (Text7 = 檢索報告 Or Text7 = "1216") Then bolPCTReport = True 'Added by Morgan 2015/2/26 PCT案一般來函輸1209檢索報告(1216國際初步審查報告)改為不上發文日,不產生告代,由工程師以承辦歷程方式發文. Ex:P-108094
   'end 2024/7/22
   
   intStep = 1
   
   '1
      Ncp(1) = cp(1)
      Ncp(2) = cp(2)
      Ncp(3) = cp(3)
      Ncp(4) = cp(4)
      '2006/6/23 MODIFY BY SONIA
      'Ncp(5) = Label3(6)
      'Modify by Morgan 2008/5/21 要區分台灣,非台灣
      'Ncp(5) = Text6
      If pa(9) = 台灣國家代號 Then
         Ncp(5) = Text6
      Else
         Ncp(5) = Label3(6)
      End If
      Ncp(6) = Text14(0)
      Ncp(7) = Text14(1)
      Ncp(8) = Text9
      'Modify by Morgan 2011/2/24 修正百年收文號問題
      'Ncp(9) = "C" & Left(strSrvDate(2), 2)
      Ncp(9) = "C" & CompAutoNumberYear(GetTaiwanThisYear)
      Ncp(10) = Text7
      '2009/12/30 MODIFY BY SONIA
      'Ncp(12) = cp(12)
      'Ncp(13) = cp(13)
      Ncp(13) = stCP13
      Ncp(12) = stCP12
      '2009/12/30 END
            
      'Added by Morgan 2015/2/26
      '掛相關號工程師,離職掛71011
      'Modified by Morgan 2016/11/3 將有下一程序規則放後面 Ex.P115099
      If bolPCTReport = True Then
         Ncp(6) = CompDate(2, 7, strSrvDate(1))
         'Added by Morgan 2016/7/25 FMP案承辦人抓畫面設定
         If m_bolFMP Then
            Ncp(14) = Text16
         Else
         'end 2016/7/25
            Ncp(6) = ChangeWStringToTString(PUB_GetWorkDay1(Ncp(6), False))
            'Modified by Morgan 2024/7/22 改也抓畫面設定
            'If GetStaffName(cp(14)) = "" Then
            '   'Added by Lydia 2023/04/24 修改王副總退休之相關控制
            '   If strSrvDate(1) >= "20230501" Then
            '      Ncp(14) = "99050"  '5/1起原工程師離職掛李柏翰
            '   Else
            '   'end 2023/04/24
            '      Ncp(14) = "71011"
            '   End If 'Added by Lydia 2023/04/24
            'Else
            '   Ncp(14) = cp(14)
            'End If
            Ncp(14) = Text16
            'end 2024/7/22
         End If
      'end 2015/2/26
      'Modify by Sindy 2024/6/18 1508國知局答辯函設定承辦期限（收文日+5工作天），並將承辦人預設為工程師
      ElseIf Text8 <> "" Or Text7 = "1508" Then
         Ncp(14) = Text16 '工程師
         Ncp(48) = Text17 '承辦期限
      Else
         Ncp(14) = strUserNum
      End If
      
      '2010/11/15 modify by sonia 取消撤銷原處分准/駁欄,此處一定准,駁改在核駁輸
      'If Text7 = 撤銷原處分 Then Ncp(24) = Text13
      If Text7 = 撤銷原處分 Then Ncp(24) = "1"
      '2010/11/15 end
      If Text7 = 專利權消滅 Then Ncp(25) = Text26
      Ncp(26) = Text18
      
      'Modify by Morgan 2009/12/1 FMP案有期限時不上發文日
      'Modified by Morgan 2018/11/19 1003通知補文件除外--敏莉
      'Modify by Sindy 2024/6/18 1508國知局答辯函除外--敏莉
      'If m_bolFMP And Text8 <> "" And Text7 <> "1003" Then
      If m_bolFMP And ((Text8 <> "" And Text7 <> "1003") Or Text7 = "1508") Then
      '2024/6/18 END
         Ncp(27) = ""
      'Added by Morgan 2015/2/26
      ElseIf bolPCTReport = True Then
         Ncp(27) = ""
      'end 2015/2/26
      'Added by Morgan 2020/1/17
      ElseIf m_bolNoCP27 = True Then
         Ncp(27) = ""
      'end 2020/1/17
      'Added by Morgan 2021/9/22
      ElseIf m_bolW2001XCase Then
         Ncp(27) = ""
      'end 2021/9/22
      'Added by Morgan 2023/6/27
      ElseIf m_bolBPFCase Then
         Ncp(27) = ""
      'end 2023/6/27
      'Added by Morgan 2024/4/29
      ElseIf m_bolEngCase Then
         Ncp(27) = ""
      'end 2024/4/29
      Else
         Ncp(27) = strSrvDate(2)
      End If
      
      Ncp(32) = "N"
      'Modify by Morgan 2004/11/26 加對造案件數代號
      'Ncp(36) = Text19
      Ncp(36) = Text19.Text & Text23.Text
      '2008/4/23 ADD BY SONIA第三人提起技術報告有掛下一程序申復者不存對造資料
      If pa(9) = "000" And Text7.Text = "1810" And pa(8) = "2" And Text8 <> "" Then
         Ncp(36) = ""
      End If
      '2008/4/23 END
      
      For i = 0 To 5
         Ncp(i + 37) = Text20(i)
      Next
      Ncp(43) = cp(9)
      
      'Added by Morgan 2012/10/5
      'Modified by Morgan 2013/1/14 增加舉發事項
      If SSTab1.TabVisible(2) = True Then
         For Each oChk In chkItem
            If oChk.Value = vbChecked Then
               If oChk.Index = 0 Then
                  Text29.Text = oChk.Caption & txtItemCount & "項;" & Text29
               ElseIf oChk.Index = 1 Then
                  Text29.Text = oChk.Caption & txtItemList & ";" & Text29
               ElseIf oChk.Index = 6 Then
                  Text29.Text = "請求撤銷自「" & txtYear(0) & "年" & txtMonth(0) & "月" & txtDay(0) & "日」至「" & txtYear(1) & "年" & txtMonth(1) & "月" & txtDay(1) & "日」之專利權期間延長;" & Text29.Text
               Else
                  Text29.Text = oChk.Caption & ";" & Text29
               End If
            End If
         Next
      End If
      'end 2012/10/5
   
    'Modify By Cheng 2002/11/29
    '加存延緩公告日
'      Ncp(64) = Text29
      Ncp(64) = Text29 & IIf(Me.Text14(2).Enabled = True And Me.Text14(2).Text <> "", "," & Me.Text14(2).Text, "")
      
      
      'Added by Morgan 2019/5/28 備註＋IDS報價
      If m_USCaseNo <> "" Then
         'Modified by Morgan 2019/6/3 第１階段報價金額大於０才寫
         'Modified by Morgan 2019/9/9 調整報價欄位名及定稿內容--郭
         If Val(txtIDSFee(1)) > 0 Then
            Ncp(64) = "IDS報價:1.第一階段 " & txtIDSFee(1) & "(" & txtIDSPt(1) & "P), 2.第二階段 " & txtIDSFee(2) & "(" & txtIDSPt(2) & "P);" & Ncp(64)
         Else
            Ncp(64) = "IDS報價:第二階段 " & txtIDSFee(2) & "(" & txtIDSPt(2) & "P);" & Ncp(64)
         End If
      End If
      'end 2019/5/27
   
'Modify by Morgan 2009/12/3 改放官方發文日欄位+CP133,CP134
'    '2009/8/6 add by sonia 非台灣案檢索報告於進度備註加註機關發文日(畫面之一般來函日期),於撰寫信函時方能計算期限
'    If pa(9) <> 台灣國家代號 And Text7 = "1209" Then
'      Ncp(64) = "機關發文日：" & Text6 & ";" & Ncp(64)
'    End If
'    '2009/8/6 end
   
   'Modified by Morgan 2012/4/25 +不必限制大陸(台灣案延期要用)
   'If pa(9) <> "000" Then
      Ncp(133) = DBDATE(Text6.Text)
      Ncp(134) = Val(Text11)
   'End If
   'end 2012/4/25
'end 2009/12/3
      
   'Add by Morgan 2007/6/13 加CP115
   If txtDispDate.Visible = True Then
      Ncp(115) = DBDATE(txtDispDate)
   End If
   
   Ncp(119) = DBDATE(Label3(6)) 'Added by Morgan 2012/4/30 +cp119=櫃檯收文日

   'add by sonia 2017/6/2 P-109072代理人請款
   If Text7 = "1908" Then
      Ncp(26) = "N"
      Ncp(32) = ""
      Ncp(20) = ""
      Ncp(16) = Val(Text21(0))
      Ncp(17) = Val(Text21(0)) - (Val(Text21(1)) * 1000)
      Ncp(18) = Val(Text21(1))
   End If
   'end 2017/6/2
   
   'Modified by Morgan 2019/8/8 FMP案的CP20要抓設定
   If m_bolFMP Then
      Ncp(20) = PUB_GetCP20(pa(1), Text7, , pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
   End If
   
   'Added by Morgan 2023/4/19
   If Text7 = 延期受理 And Pub_StrUserSt03 = "F22" Then
      Ncp(142) = DBDATE(Text12)
      'Modified by Morgan 2025/8/27 官方期限已無在途 --敏莉
      'Ncp(64) = "含在途法限:" & CFDate(TransDate(Ncp(142), 1)) & ";" & Ncp(64)
      Ncp(64) = "官方法限:" & CFDate(TransDate(Ncp(142), 1)) & ";" & Ncp(64)
      'end 2025/8/27
   End If
   'end 2023/4/19
         
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.SaveNewCaseProgressDatabase("C", Ncp, intWhere) Then
      If Not ClsPDSaveNewCaseProgressDatabase("C", Ncp, intWhere) Then
'         Exit Function
        GoTo ErrorHandler
      End If
      
   'Added by Lydia 2025/08/19 輸入C類來函時，去檢查上一道承辦人掛工程師，是否為未請款，若是，則發Mail通知工程師；
   If pa(1) = "P" And m_bolFMP = True And Text16 <> "" Then
      If PUB_ChkFCPtoCP14CP60(pa(1), pa(2), pa(3), pa(4), Text7, Ncp(9), Text16) = True Then
      End If
   End If
   'end 2025/08/19
   
   'Added by Lydia 2025/10/29
   If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
       stNP23 = "" 'Move by Lydia 2025/11/05 從m_bolFMP = False上面移下來
       strSql = PUB_GetPOurDeadline(Text14(1), pa(9), stNP23, pa(1), "429")
   End If
   'end 2025/10/29
   
      'Add by Morgan 2005/1/20 發明核准且為一案兩請則新增新型案自請撤回下一程序
      'Modify by Morgan 2005/5/13 改自請撤回(413)為放棄專利權(429)
      m_strRetSheet2NP07 = ""
      'Modified by Morgan 2014/8/5
      'If m_bolIsDualApp = True Then
      If m_bolIsDualApp = True And m_bolGiveUpUtility = False Then
      'end 2014/8/5
         intMax = GetNextProgressNo
         'Modified by Lydia 2025/10/29 +NP23
         strSql = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22,NP23) values (" + _
            CNULL(Ncp(9)) + "," + CNULL(m_stUPA(1)) + "," + CNULL(m_stUPA(2)) + "," + CNULL(m_stUPA(3)) + _
            "," + CNULL(m_stUPA(4)) + ",429," + TransDate(Text14(0), 2) + "," & TransDate(Text14(1), 2) & _
            "," + CNULL(PUB_GetAKindSalesNo(m_stUPA(1), m_stUPA(2), m_stUPA(3), m_stUPA(4))) + "," + intMax + "," + CNULL(stNP23, True) + ")"
         cnnConnection.Execute strSql
         m_DualAppNP22 = intMax
         m_strRetSheet2NP07 = "429"
      End If
      
      'Add by Morgan 2004/11/30 抓最新的AB類發文代理人更新
      Pub_UpdateFromMaxCP27 pa(1), pa(2), pa(3), pa(4)
   
      ' 90.06.28 modify by louis 暫存新的收文號
      m_NewCP09 = Ncp(9)

   '2
   'Add By Sindy 2012/3/5 原基本檔未閉卷時,才要更新 +And m_blnClosed = False
   If Text27(0) = "Y" And m_blnClosed = False Then
      'Add By Sindy 2012/3/5 +PA58,PA59及Update servicepractice
      If pa(1) = "P" Then
         strTxt(intStep) = "UPDATE PATENT SET PA57='Y',PA58=" & strSrvDate(1) & ",PA59='99',PA17='N' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      Else
         strTxt(intStep) = "UPDATE servicepractice SET sp15='Y',sp16=" & strSrvDate(1) & ",sp17='99' WHERE " & ChgService(pa(1) & pa(2) & pa(3) & pa(4))
      End If
      '2012/3/5 End
      'Add By Cheng 2002/11/08
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   'Add By Sindy 2012/3/5
   ElseIf Text27(0) = "" Then
      If pa(1) = "P" Then
         strTxt(intStep) = "UPDATE PATENT SET PA57=null,PA58=null,PA59=null WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      Else
         strTxt(intStep) = "UPDATE servicepractice SET sp15=null,sp16=null,sp17=null WHERE " & ChgService(pa(1) & pa(2) & pa(3) & pa(4))
      End If
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   '2012/3/5 End
   
   '**************  90.11.13 nick
   If Text31 <> "" Then
      ' 91.03.25 modify by louis (單引號)
      strTxt(intStep) = "UPDATE PATENT SET PA91='" & ChgSQL(Text31) & "' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      'Add By Cheng 2002/11/08
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   '*****************************
   '3
   
   'Added by Morgan 2014/7/21
   '台灣一案兩請案申請日 101.1.1~102.6.13 者不必管控放棄專利權,發證時自動閉卷
   If m_bolGiveUpUtility = True Then
      'Modified by Morgan 2015/7/9 一案兩請是否放棄新型改放PA60
      'strSql = "update patent set pa162='Y' where " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      strSql = "update patent set pa60='Y' where " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   'end 2014/7/21
   
   'Added by Lydia 2016/02/04 申請國家為PCT的案件，設定為收到檢索報告後，下一程序的催審自動上N (By 如嬿)
   If pa(9) = "056" And Text7 = 檢索報告 Then
      strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='N' WHERE NP01='" & strReceiveNo & "' AND NP07='" & 催審 & "' and np06 is null"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   'end 2016/02/04
   
   'Added by Lydia 2016/10/13 PCT案輸入國際初步審查報告(1216),將催審-實體審查上Y
   If pa(9) = "056" And Text7 = "1216" Then
      strTxt(intStep) = "update nextprogress set np06='Y' where (np01,np22) = (select np01,np22 from nextprogress,caseprogress " & _
                        "where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' " & _
                        "and np07='411' and np06 is null and np01=cp09(+) and cp10='416')"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   'end 2016/10/13
   
   If Text7 = 所外鑑定報告結果 Then
      strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07='" & 催審 & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1' WHERE CP09='" & strReceiveNo & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   '4
   ElseIf Text7 = 撤銷原處分 Then
      '智權人員存最近收文A類接洽記錄單的智權人員
      strTmp = CompDate(1, 3, TransDate(Label3(6).Caption, 2))
      intMax = GetNextProgressNo
      'Modify by Morgan 2011/10/12 智權人員改程序--郭
      strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
         "NP07,NP08,NP09,NP10,NP22) VALUES ('" & Ncp(9) & "','" & pa(1) & _
         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & 改變原處分 & "," & _
         strTmp & "," & strTmp & ",'" & strUserNum & "'," & intMax & ")"
         
        'Add By Cheng 2002/11/08
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      '2010/11/15 add by sonia
      strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07='" & 催審 & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25=" & DBDATE(Text6.Text) & " WHERE CP09='" & strReceiveNo & "' AND CP24 IS NULL AND CP25 IS NULL"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      '2010/11/15 end
   '2012/3/7 add by sonia 暫不續行審理,催審及下一程序的通知實審日1204都上N
   ElseIf Text7 = "1911" Then
      strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='N' WHERE NP01='" & strReceiveNo & "' AND NP07 IN ('" & 催審 & "','" & 通知實審日 & "')"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
    'Add By Cheng 2002/12/12
    '預設可新增下一程序資料
    blnAddNewNP = True
    '若來函性質為通知準備程序(1210), 通知言詞辯論(1211), 通知參加訴訟(1504), 92.2.10增加 通知面詢(1401), 通知閱卷(1402) 若該收文號已收文未發文, 則期限不新增至下一程序, 改更新該筆進度資料期限
    'modify by sonia 2018/8/14 +1812通知聽證
    If Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or Me.Text7.Text = "1504" Or Me.Text7.Text = "1401" Or Me.Text7.Text = "1402" Or Me.Text7.Text = "1812" Then
        If m_CP05 <> "" And m_CP27 = "" Then
            strSql = "Update CaseProgress Set CP06=" & CNULL(DBDATE(Me.Text14(0).Text)) & ", CP07=" & CNULL(DBDATE(Me.Text14(1).Text)) & " Where CP09='" & m_CP09 & "'"
            cnnConnection.Execute strSql
            '不新增下一程序資料
            blnAddNewNP = False
        End If
    End If
   
   '5
   If Text8 <> "" Then
        'Modify By Cheng 2002/12/12
        '若可新增下一程序資料
        If blnAddNewNP = True Then
            If Text20(3) <> "" Then
               strTmp = Text20(3)
            ElseIf Text20(4) <> "" Then
               strTmp = Text20(4)
            ElseIf Text20(5) <> "" Then
               strTmp = Text20(5)
            End If
            'Add by Morgan 2009/12/1
            '非台灣的補文件
            If Text8 = "202" And cmdDeadLine.Visible = True Then
               strSql = "UPDATE CASEPROGRESS SET CP06=" & CNULL(DBDATE(Text14(0)), True) & ",CP07=" & CNULL(DBDATE(Text14(1)), True) & _
                  " WHERE CP43='" & cp(9) & "' AND CP27 IS NULL AND CP10='202' AND CP57 IS NULL"
               cnnConnection.Execute strSql, intI
               'Added by Lydia 2025/10/29
               If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                  stNP23 = "" 'Move by Lydia 2025/11/05 從m_bolFMP = False上面移下來
                  strSql = PUB_GetPOurDeadline(DBDATE(Text14(1)), pa(9), stNP23, pa(1), "202")
               End If
               'end 2025/10/29
               'Added by Morgan 2019/5/10 更新NP補文件期限--敏莉,玲玲
               'Modified by Lydia 2025/10/29 +NP23
               strSql = "UPDATE NEXTPROGRESS SET NP08=" & CNULL(DBDATE(Text14(0)), True) & ",NP09=" & CNULL(DBDATE(Text14(1)), True) & ",NP23=" & IIf(stNP23 = "", "NP23", stNP23) & _
                  " WHERE NP01='" & cp(9) & "' AND NP06 IS NULL AND NP07='202'"
               cnnConnection.Execute strSql, intI
               'end 2019/5/10
               If m_strUnSaveData <> "" Then
                  arrData = Split(m_strUnSaveData, vbCrLf)
                  For i = LBound(arrData) To UBound(arrData)
                     If arrData(i) <> "" Then
                        'Modified by Lydia 2025/10/29 +NP23
                        strSql = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22,NP23)" & _
                           " SELECT '" & m_NewCP09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','202'" & _
                           "," & CNULL(DBDATE(Text14(0)), True) & "," & CNULL(DBDATE(Text14(1)), True) & ",'" & stCP13 & "'" & _
                           "," & CNULL(ChgSQL(arrData(i))) & _
                           ",NP22," & CNULL(stNP23, True) & " FROM (SELECT MAX(NP22)+1 NP22 FROM NEXTPROGRESS) X"
                        cnnConnection.Execute strSql, intI
                     End If
                  Next
               End If
            Else
               'Added by Lydia 2025/10/29
               If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                  stNP23 = "" 'Move by Lydia 2025/11/05 從m_bolFMP = False上面移下來
                  strSql = PUB_GetPOurDeadline(DBDATE(Text14(1)), pa(9), stNP23, pa(1), Text7)
               End If
               'end 2025/10/29
                  
               'Modify by Morgan 2005/4/25 加準備程序,言詞辯論
               'If pa(1) = "P" And pa(9) = 台灣國家代號 And (Text7 = 通知補文件 Or Text7 = 通知閱卷) Then
               'Modify by Morgan 2005/7/20 更正2005/4/25修改
               '改加判斷當點選性質為 "已發文" 之 "準備程序(211)" OR "言詞辯論(212)" 且來函性質為 "通知準備程序(1210)" OR "通知言詞辯論(1211)" 時才新增NP且上Y
               'If pa(1) = "P" And pa(9) = 台灣國家代號 And (Text7 = 通知補文件 Or Text7 = 通知閱卷 Or Text7 = 準備程序 Or Text7 = 言詞辯論) Then
               'Modify by Morgan 2005/8/16 加台灣新型的通知修正及通知申復
               'If pa(1) = "P" And pa(9) = 台灣國家代號 And (Text7 = 通知補文件 Or Text7 = 通知閱卷 Or (cp(10) = 準備程序 And m_CP27 <> "" And Text7 = "1210") Or (cp(10) = 言詞辯論 And m_CP27 <> "" And Text7 = "1211")) Then
               'Modify by Morgan 2006/5/3 台灣新型的通知修正及通知申要控制點選新型申請那一道才做
               'If pA(1) = "P" And pA(9) = 台灣國家代號 And (Text7 = 通知補文件 Or Text7 = 通知閱卷 Or (cp(10) = 準備程序 And m_CP27 <> "" And Text7 = "1210") Or (cp(10) = 言詞辯論 And m_CP27 <> "" And Text7 = "1211") Or (pA(8) = "2" And (Text7 = "1201" Or Text7 = "1202"))) Then
               'Modify by Morgan 2006/9/4  台灣新型的通知修正及通知申要控制點選改請新型也要做--玲玲
               'If pA(1) = "P" And pA(9) = 台灣國家代號 And (Text7 = 通知補文件 Or Text7 = 通知閱卷 Or (cp(10) = 準備程序 And m_CP27 <> "" And Text7 = "1210") Or (cp(10) = 言詞辯論 And m_CP27 <> "" And Text7 = "1211") Or (cp(10) = "102" And (Text7 = "1201" Or Text7 = "1202"))) Then
               'Modify by Morgan 2006/10/5 P的通知補文件都要
               'If pA(1) = "P" And pA(9) = 台灣國家代號 And (Text7 = 通知補文件 Or Text7 = 通知閱卷 Or (cp(10) = 準備程序 And m_CP27 <> "" And Text7 = "1210") Or (cp(10) = 言詞辯論 And m_CP27 <> "" And Text7 = "1211") Or ((cp(10) = "102" Or cp(10) = "302") And (Text7 = "1201" Or Text7 = "1202"))) Then
               '2009/10/5 modify by sonia 加台灣發明申請,設計申請的通知修正
               intMax = GetNextProgressNo
               'Modify by Morgan 2009/12/1 FMP案的期限都不自動內部收文
               'If pa(1) = "P" And (Text7 = 通知補文件 Or (pa(9) = 台灣國家代號 And (Text7 = 通知閱卷 Or (cp(10) = 準備程序 And m_CP27 <> "" And Text7 = "1210") Or (cp(10) = 言詞辯論 And m_CP27 <> "" And Text7 = "1211") Or ((cp(10) = "102" Or cp(10) = "302") And (Text7 = "1201" Or Text7 = "1202")) Or ((cp(10) = "101" Or cp(10) = "301" Or cp(10) = "103" Or cp(10) = "303") And Text7 = "1201")))) Then
               'Modify by Morgan 2010/7/22 +1225 依職權電話通知修正
               'Modified by Morgan 2012/12/27 +最後通知1227
               'Modified by Morgan 2014/12/25 台灣新型(102,302),通知修正(1201),審查意見通知(1202),最後通知(1227)改不自動收文,改判發時決定 Or ((cp(10) = "102" Or cp(10) = "302") And (Text7 = "1201" Or Text7 = "1202" Or Text7 = "1227"))
               'Modified by Morgan 2016/1/14 台灣通知修正改都不要內部收文由判發人決定
               'Modified by Morgan 2024/4/25 P大陸依職權電話通知修正改工程師承辦，不自動內部收文，若不報告客戶則再由程序內部收文--品薇
               If Not m_bolFMP And pa(1) = "P" And ((Text7 = 1225 And m_bolNoCP27 = False) Or Text7 = 通知補文件 Or (pa(9) = 台灣國家代號 And (Text7 = 通知閱卷 Or (cp(10) = 準備程序 And m_CP27 <> "" And Text7 = "1210") Or (cp(10) = 言詞辯論 And m_CP27 <> "" And Text7 = "1211")))) Then
                  'Modified by Lydia 2025/10/29 +NP23
                  strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP06," & _
                     "NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22,NP23) VALUES ('" & Ncp(9) & "','" & pa(1) & _
                     "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','Y'," & Text8 & "," & _
                     TransDate(Text14(0), 2) & "," & TransDate(Text14(1), 2) & "," & CNULL(stCP13) & _
                     "," & CNULL(Text9) & "," & CNULL(ChgSQL(strTmp)) & "," & CNULL(Text29) & "," & intMax & "," & CNULL(stNP23, True) & " )"
                     
                  bolAddBCP = True 'Add by Morgan 2006/10/5
                  
                  'Added by Morgan 2023/5/18 P台灣案通知補文件相關管控--陳玲玲
                  '有未發文244補中文說明書,更新期限及相關收文號
                  '有未發文232補優先權證明,更新相關收文號
                  '有未收文才要內部收文補文件
                  If pa(1) = "P" And pa(9) = 台灣國家代號 And Text7 = 通知補文件 Then
                     arrData = Split(Ncp(64), "、")
                     strExc(9) = ""
                     strExc(1) = GetCaseTypeName(pa(1), "244", 0)
                     strExc(2) = GetCaseTypeName(pa(1), "232", 0)
                     For i = LBound(arrData) To UBound(arrData)
                        If arrData(i) <> "" Then
                           '244補中文說明書
                           If arrData(i) = strExc(1) Then
                              If PUB_ChkCPExist(pa, "244", 1, strExc(3)) Then
                                 'Modified by Morgan 2023/5/25 +cp08 -- 玲玲
                                 strSql = "update caseprogress set cp06=" & DBDATE(Text14(0)) & ",cp07=" & DBDATE(Text14(1)) & ",cp08='" & Ncp(8) & "',cp43='" & Ncp(9) & "' where cp09='" & strExc(3) & "'"
                                 cnnConnection.Execute strSql, intI
                                 arrData(i) = ""
                              End If
                           '232補優先權證明
                           ElseIf arrData(i) = strExc(2) Then
                              If PUB_ChkCPExist(pa, "232", 1, strExc(3)) Then
                                 'Modified by Morgan 2023/5/25 +cp08 -- 玲玲
                                 strSql = "update caseprogress set cp43='" & Ncp(9) & "',cp08='" & Ncp(8) & "' where cp09='" & strExc(3) & "'"
                                 cnnConnection.Execute strSql, intI
                                 arrData(i) = ""
                              End If
                           End If
                        End If
                        If arrData(i) <> "" Then
                           strExc(9) = strExc(9) & IIf(strExc(9) <> "", "、", "") & arrData(i)
                        End If
                     Next
                     Ncp(64) = strExc(9)
                     If strExc(9) = "" Then
                        bolAddBCP = False
                     End If
                  End If
                  'end 2023/5/18
               Else
                  '智權人員存最近收文A類接洽記錄單的智權人員
                  'Modify by Morgan 2009/12/1 +NP23
                  'modify by sonia 2018/5/3 1901通知退費不必輸法定期限
                  'strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
                     "NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22,NP23) VALUES ('" & Ncp(9) & "','" & pa(1) & _
                     "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & Text8 & "," & _
                     TransDate(Text14(0), 2) & "," & TransDate(Text14(1), 2) & "," & CNULL(stCP13) & _
                     "," & CNULL(Text9) & "," & CNULL(ChgSQL(strTmp)) & "," & CNULL(Text29) & "," & intMax & "," & CNULL(stNP23, True) & ")"
                  strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
                     "NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22,NP23) VALUES ('" & Ncp(9) & "','" & pa(1) & _
                     "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & Text8 & "," & _
                     TransDate(Text14(0), 2) & "," & CNULL(ChgSQL(TransDate(Text14(1), 2))) & "," & CNULL(stCP13) & _
                     "," & CNULL(Text9) & "," & CNULL(ChgSQL(strTmp)) & "," & CNULL(Text29) & "," & intMax & "," & CNULL(stNP23, True) & ")"
                  strProgressNo = intMax
               End If
               cnnConnection.Execute strTxt(intStep)
               intStep = intStep + 1
               
               'Added by Lydia 2016/09/21 P大陸案電話通知修正:若大陸代理人來此類電話通知修正，輸入1225依職權電話通知修正，系統請一併自動收文204補正，並沖掉下一程序。
               'Modified by Morgan 2023/9/13 +m_bolAutoBCP
               'Modified by Morgan 2024/4/25 非FMP大陸案電話通知修正改不自動內部收文
               'If (pa(1) = "P" And pa(9) = "020" And Me.Text7.Text = "1225") Or m_bolAutoBCP Then
               If (m_bolFMP And pa(1) = "P" And pa(9) = "020" And Me.Text7.Text = "1225") Or m_bolAutoBCP Then
               'end 2024/4/25
                   strSql = "Update NextProgress set NP06='Y' where np01='" & Ncp(9) & "' and np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07='" & Text8 & "' and np06 is null"
                   cnnConnection.Execute strSql, intI
                   bolAddBCP = True
               End If
               'End 2016/09/21
            End If
            
'Remove by Morgan 2009/12/1 改來函不自動上發文日(印C類接洽單)
'
'            'Add by Morgan 2006/6/26
'            '國外部收文若有期限則自動內部收文901告知代理人,承辦人固定為78063黃得峻並列印內部收文接洽單
'            If m_bolFMP Then
'               m_901CP09 = AutoNo("B", 6)
'               '2008/12/2 modify by sonia 改FMP控管方式
'               'm_901CP13 = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
'               m_901CP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
'               '2008/12/2 END
'               m_901CP12 = GetSalesArea(m_901CP13)
'               strExc(1) = GetWorkDays(pa(1), pa(9), "901")
'               If strExc(1) = Empty Then strExc(1) = 7
'               'Add by Morgan 2008/5/26 若來函期限超過(含)3個月則告代的承辦期限為14天--阮威立
'               If Val(strExc(1)) < 14 Then
'                  If DBDATE(Text14(1)) >= CompDate(1, 3, strSrvDate(1)) Then
'                     strExc(1) = 14
'                  End If
'               End If
'               'end 2008/5/26
'
'               'Modify by Morgan 2006/8/4 不必抓工作天--郭
'               'strCP48 = CompWorkDay(Val(strExc(1)), strSrvDate(1), 0)
'               strCP48 = CompDate(2, Val(strExc(1)), strSrvDate(1))
'               '2008/12/3 MODIFY BY SONIA 依FC代理人國籍抓預設承辦人
'               'strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
'                  "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
'                  "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & strCP48 & "," & strCP48 & _
'                  ",'" & m_901CP09 & "','" & 告知代理人 & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
'                  ",'85030','N','N','N','" & Ncp(9) & "'," & strCP48 & ") "    '2008/2/5 MODIFY BY SONIA 78063離職改85030阮威立--郭
'               strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
'                  "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
'                  "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & strCP48 & "," & strCP48 & _
'                  ",'" & m_901CP09 & "','" & 告知代理人 & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
'                  "," & CNULL(PUB_GetFMCASECP14(cp(1), cp(2), cp(3), cp(4))) & ",'N','N','N','" & Ncp(9) & "'," & strCP48 & ") "    '2008/2/5 MODIFY BY SONIA 78063離職改85030阮威立--郭
'               '2008/12/3 END
'               cnnConnection.Execute strSQL
'            End If
'
'end 2009/12/1
            
            
            
            '2012/8/20 ADD BY SONIA FMP審查意見通知1202且為尼康客戶案件仍要內部收文901,承辦期限為系統日起14天,不必抓工作天,不請款--陳毓芳
            '2012/9/18 MODIFY BY SONIA 加入Y51508(因FCP案要加,此代理人雖無FMP案但仍先加入)
            'Modified by Morgan 2012/10/19 +Y52003--陳毓芳
'Modified by Morgan 2013/9/18 改呼叫共用函數
'            If m_bolFMP And Text7 = "1202" And (Left(pa(26), 6) = "X56040" Or Left(pa(26), 6) = "X48340" Or Left(pa(26), 6) = "X45149" Or Left(pa(26), 6) = "X60049" Or Left(pa(75), 6) = "Y51508" Or Left(pa(75), 6) = "Y52003") Then
'               m_901CP09 = AutoNo("B", 6)
'               m_901CP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
'               m_901CP12 = GetSalesArea(m_901CP13)
'               strExc(1) = 14
'               strCP48 = CompDate(2, Val(strExc(1)), strSrvDate(1))
'               strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
'                  "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
'                  "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & _
'                  ",'" & m_901CP09 & "','" & 告知代理人 & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
'                  "," & CNULL(Text16) & ",'N','N','N','" & Ncp(9) & "'," & strCP48 & ") "
'               cnnConnection.Execute strSql
'            End If
'            '2012/8/20 END
'
'            '2012/11/9 ADD BY SONIA FMP審查意見通知1202且為Y20065案件要內部收文901,承辦期限為主管機關發文日15天,不必抓工作天,不請款--陳毓芳
'            '2012/11/15 MODIFY BY SONIA 加入 Y27766
'            If m_bolFMP And Text7 = "1202" And (Left(pa(75), 6) = "Y20065" Or Left(pa(75), 6) = "Y27766") Then
'               m_901CP09 = AutoNo("B", 6)
'               m_901CP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
'               m_901CP12 = GetSalesArea(m_901CP13)
'               strExc(1) = 15
'               strCP48 = CompDate(2, Val(strExc(1)), DBDATE(Text6.Text))
'               If strCP48 < strSrvDate(1) Then strCP48 = strSrvDate(1)
'               strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
'                  "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
'                  "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & _
'                  ",'" & m_901CP09 & "','" & 告知代理人 & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
'                  "," & CNULL(Text16) & ",'N','N','N','" & Ncp(9) & "'," & strCP48 & ") "
'               cnnConnection.Execute strSql
'            End If
'            '2012/11/9 END
'
'            '2012/10/19 ADD BY SONIA Y53309審查意見通知1202或核駁要內部收文901,承辦期限為系統日起7天(日曆天)--吳若芬(因FCP案要加,此代理人雖無FMP案但仍先加入)
'            '2013/1/24 modify by sonia 加Y51542
'            'Modified by Morgan 2013/8/28 ,+ Y34210 & X51446 --邱子瑜
'            'Modified by Morgan 2013/8/30 ,+ Y47453 & X55778 --羅惠蓮
'            'Modified by Morgan 2013/9/6 + Y20065 --邱子瑜
'            If m_bolFMP And Text7 = "1202" And (Left(pa(75), 6) = "Y53309" Or Left(pa(75), 6) = "Y51542" Or Left(pa(75), 6) = "Y20065" Or _
'               (Left(pa(75), 6) = "Y34210" And Left(pa(26), 6) = "X51446") Or _
'               (Left(pa(75), 6) = "Y47453" And Left(pa(26), 6) = "X55778")) Then
'
'               m_901CP09 = AutoNo("B", 6)
'               m_901CP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
'               m_901CP12 = GetSalesArea(m_901CP13)
'               strExc(1) = 7
'               strExc(2) = 告知代理人
'
'               'Added by Morgan 2013/8/28
'               'Y34210 + X51446 14天 --邱子瑜
'               If Left(pa(75), 6) = "Y34210" And Left(pa(26), 6) = "X51446" Then
'                  strExc(1) = 14
'               'Added by Morgan 2013/9/6
'               'Y20065 15天 --邱子瑜
'               ElseIf Left(pa(75), 6) = "Y20065" Then
'                  strExc(1) = 15
'               End If
'               'Y51542 改收其他翻譯 --吳彩菱
'               If Left(pa(75), 6) = "Y51542" Then
'                  strExc(2) = "927"
'               End If
'               'end 2013/8/28
'               strCP48 = CompDate(2, Val(strExc(1)), strSrvDate(1))

         'Add by Lydia 2014/12/3 核駁及審查意見通知函備註
         '   If m_bolFMP And Text7 = "1202" And PUB_ChkAutoRec(pa(1), pa(75), pa(26), DBDATE(Text6), strExc(2), strCP48, , , pa(27), pa(28), pa(29), pa(30)) = True Then
              
            'Modified by Morgan 2018/7/25 +Text7 = "1202"(原來有不知為何被拿掉???) Ex.P-120426 通知補文件1003
            'If m_bolFMP And pa(57) = "" Then 'Added by Morgan 2017/4/28 FMP未閉卷才交工程師報告客戶,已閉卷直接交FCP程序--潘韻丞(David 確認)
            'Modified by Lydia 2020/03/06 比照FCP案+最後通知1227,被舉發理由1802
            'If m_bolFMP And Text7 = "1202" And pa(57) = "" Then
            ''end 2018/7/25
            'Modified by Lydia 2021/05/18 +通知補正1201
            'Modified by Lydia 2021/08/25 國外部凡是C類工程師的來函(排除核准1001、核發1008，另外非核准和一般來函性質1204,1217,1913,1603,1604)，有設核駁及審查意見通知函備註皆要帶備註到接洽單
            'If m_bolFMP And (Text7 = "1202" Or Text7 = "1227" Or Text7 = "1802" Or Text7 = "1201") And pa(57) = "" Then
            'Modified by Lydia 2021/08/31 +判斷來函承辦人為工程師 And PUB_GetST03(Text16) = "F21" ; ex.發生了FCP065502通知即將公開1207直接收文告代
            'Modified by Lydia 2021/09/03 因為只需在C類接洽單列印,這裡反而不用改
            'If m_bolFMP And InStr("1001,1008,1204,1217,1913,1603,1604", Text7) = 0 And pa(57) = "" And PUB_GetST03(Text16) = "F21" Then
            If m_bolFMP And (Text7 = "1202" Or Text7 = "1227" Or Text7 = "1802" Or Text7 = "1201") And pa(57) = "" Then
                Dim sMemo As String
                Dim stBCP16 As String 'Added by Lydia 2022/01/05
                  'Remove by Lydia 2021/11/05
                  'strExc(2) = ""
                  'strExc(7) = "": strExc(3) = "": strExc(4) = "": strExc(5) = ""
                  'If Not IsNull(pa(27)) Then strExc(7) = ChangeCustomerL(pa(27))
                  'If Not IsNull(pa(28)) Then strExc(3) = ChangeCustomerL(pa(28))
                  'If Not IsNull(pa(29)) Then strExc(4) = ChangeCustomerL(pa(29))
                  'If Not IsNull(pa(30)) Then strExc(5) = ChangeCustomerL(pa(30))
                  'end 2021/11/05
                  'Modified by Lydia 2021/11/02 +strExc(6),strExc(8) 記錄C類來函的承辦期限和指定送件日期
                  'sMemo = PUB_GetIncomMemoNew(pa(1) & pa(2) & pa(3) & pa(4), pa(1), strExc(2), ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)), , strCP48, DBDATE(Text6) _
                              , strExc(7), strExc(3), strExc(4), strExc(5))
                  'strExc(7) = "": strExc(3) = "": strExc(4) = "": strExc(5) = ""
                  'Modified by Lydia 2021/11/05 分別傳回B類收文(承辦期限、所限)和C類來函(承辦期限和指定送件日期)
                   Dim stBCP10 As String, stBCP48   As String, stBCP06 As String, stCCP48 As String, stCCP142 As String
                   sMemo = PUB_GetIncomMemoNew(pa(1) & pa(2) & pa(3) & pa(4), pa(1), ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)) & "," & ChangeCustomerL(pa(27)) & "," & ChangeCustomerL(pa(28)) & "," & ChangeCustomerL(pa(29)) & "," & ChangeCustomerL(pa(30)), _
                                   "", DBDATE(Text6), Text7.Text, stCCP48, stCCP142, stBCP10, stBCP48, stBCP06)
                                   
                  'Added by Lydia 2021/11/05 更新C類來函的承辦期限和指定送件日期，一併更新指定送件日期之前CP164=2
                  If stCCP48 <> "" Then
                      'Modified by Lydia 2021/11/16 加註cp64
                      strSql = "Update CaseProgress set cp48=" & stCCP48 & ", cp141='3', cp142=" & stCCP142 & ", cp164='2' " & _
                                   ", cp64='客戶指定" & ChangeWStringToTDateString(stCCP142) & "之前送件;'||cp64 where cp09='" & Ncp(9) & "' "
                      cnnConnection.Execute strSql, intI
                  End If
                  'end 2021/11/05
                  
                  'Added by Lydia 2025/02/05 輸入中間程序來函時自動產生行事曆
                  If PUB_AddSCforIncomMemo(pa(1), pa(2), pa(3), pa(4), Ncp(9), Text7, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)) & "," & ChangeCustomerL(pa(27)) & "," & ChangeCustomerL(pa(28)) & "," & ChangeCustomerL(pa(29)) & "," & ChangeCustomerL(pa(30))) = False Then
                      GoTo ErrorHandler
                  End If
                  'end 2025/02/05
                  
                 'Modified by Lydia 2021/11/05 PUB_GetIncomMemoNew已有另外抓B類收文設定
                 'If m_bolFMP And Len(sMemo) > 0 Then
                 '      If strExc(2) = "" Then strExc(2) = "901"
                 If Len(stBCP10) > 0 Then
                  'end 'Add by Lydia 2014/12/3
                   m_901CP09 = AutoNo("B", 6)
                   m_901CP13 = stCP13
                   m_901CP12 = stCP12
    'end 2013/9/18
                   'Modified by Morgan 2019/8/8 FMP案的CP20要抓設定
                   'Modified by Lydia 2022/01/05 +stBCP16、改抓變數 strExc(2)=> stBCP10
                   strCP20 = PUB_GetCP20(pa(1), stBCP10, stBCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
                   'Modified by Lydia 2021/11/05 改變數
                   'strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
                      "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
                      "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & _
                      ",'" & m_901CP09 & "','" & strExc(2) & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
                      "," & CNULL(Text16) & ",'" & strCP20 & "','N','N','" & Ncp(9) & "'," & strCP48 & ") "
                   'Modified by Lydia 2022/01/05 +CP16
                   strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
                      "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48,CP06,CP16) VALUES " & _
                      "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & _
                      ",'" & m_901CP09 & "','" & stBCP10 & "','90'," & CNULL(m_901CP12) & "," & CNULL(m_901CP13) & _
                      "," & CNULL(Text16) & ",'" & strCP20 & "','N','N','" & Ncp(9) & "'," & stBCP48 & "," & stBCP06 & "," & CNULL(stBCP16, True) & ")  "
                   cnnConnection.Execute strSql, intI
                End If
            End If 'Added by Morgan 2017/4/28
            '2012/10/19 END
        End If
      
      With MSHFlexGrid1
         For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "v" Then
               'Modify by Morgan 2006/1/24 加NP01
               strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP22=" & .TextMatrix(i, 7) & " and np01='" & .TextMatrix(i, 8) & "'"
                'Add By Cheng 2002/11/08
                cnnConnection.Execute strTxt(intStep)
               intStep = intStep + 1
            End If
         Next
      End With
   End If
   
   If bolAddBCP = True Then
      Ncp(1) = cp(1)
      Ncp(2) = cp(2)
      Ncp(3) = cp(3)
      Ncp(4) = cp(4)
      Ncp(5) = Label3(6)
      Ncp(6) = Text14(0)
      Ncp(7) = Text14(1)
      Ncp(8) = Text9
      'Modify by Morgan 2011/2/24 修正百年收文號問題
      'Ncp(9) = "B" & Left(strSrvDate(2), 2)
      Ncp(9) = "B" & CompAutoNumberYear(GetTaiwanThisYear)
      Ncp(10) = Text8
      '2008/7/15 MODIFY BY SONIA
      'Ncp(12) = cp(12)
      'Ncp(13) = cp(13)
      Ncp(13) = stCP13
      Ncp(12) = stCP12
      Ncp(14) = cp(14)
      'Added by Morgan 2023/9/13
      If m_bolAutoBCP Then
         If GetStaffDepartment(Ncp(14)) = "P12" Then
            Ncp(14) = PUB_GetInCaseCP14(pa(1), pa(2), pa(3), pa(4))
         End If
      End If
      strOldCP14 = Ncp(14)
      'end 2023/9/13
      '92.7.7 MODIFY BY SONIA 取消通知面詢
      'If Text7 = 通知面詢 Or Text7 = 通知閱卷 Then
      '93.9.16 CANCEL BY SONIA
      'If Text7 = 通知閱卷 Then
      '93.9.16 END
      '92.7.7 END
      '通知閱卷原承辦人為分所人員或已離職,則預設為王協理
      strExc(0) = "select ST06,ST04,ST02 from STAFF where ST01=" + CNULL(Ncp(14))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Added by Morgan 2023/9/14
         strST02 = RsTemp.Fields("ST02")
         strST04 = RsTemp.Fields("ST04")
         'end 2023/9/14
         If RsTemp.Fields("ST06") <> "1" Then
            'Modify by Morgan 2004/10/11 通知補文件只判離職的才預設為王協理
            'Ncp(14) = "71011"
            'Modify by Morgan 2005/7/20
            '改分所的通知閱卷固定預設王協理
            'If Text7 <> 通知補文件 Then
            If Text7 = 通知閱卷 Then
               'Added by Lydia 2023/04/24 修改王副總退休之相關控制
               If strSrvDate(1) >= "20230501" Then
                  Ncp(14) = "99050"
               Else
               'end 2023/04/24
                  Ncp(14) = "71011"
               End If 'Added by Lydia 2023/04/24
            End If
         End If
         
         If RsTemp.Fields("ST04") <> "1" Then
            'Added by Lydia 2023/04/24 修改王副總退休之相關控制
            If strSrvDate(1) >= "20230501" Then
               Ncp(14) = "99050"
            Else
            'end 2023/04/24
               Ncp(14) = "71011"
            End If 'Added by Lydia 2023/04/24
         End If
      Else
         'Added by Lydia 2023/04/24 修改王副總退休之相關控制
         If strSrvDate(1) >= "20230501" Then
            Ncp(14) = "99050"
         Else
         'end 2023/04/24
            Ncp(14) = "71011"
         End If 'Added by Lydia 2023/04/24
      End If
      '2008/5/20 add by sonia 何尤玉轉調部門,先分給林柄佑
      'modify by sonia 2022/4/29 82026林柄佑及74018杜燕文的改71011重新分案,76028已離職
      'If Ncp(14) = "76028" Then
      '   Ncp(14) = "82026"
      'End If
      If Ncp(14) = "82026" Or Ncp(14) = "74018" Then
         'Added by Lydia 2023/04/24 修改王副總退休之相關控制
         If strSrvDate(1) >= "20230501" Then
            Ncp(14) = "99050"
         Else
         'end 2023/04/24
            Ncp(14) = "71011"
         End If 'Added by Lydia 2023/04/24
      End If
      
      'Add by Lydia 2014/09/22 台灣案一般來函通知補文件(來函性質1003 下一程序:補文件202),自動產生B類補文件(內部收文),承辦人目前均掛「承辦工程師」,將來由程序承辦
      strExc(9) = "" ''Add by Lydia 2014/10/20 台灣案通知補文件自動作內部收文控制，發E-MAIL通知智權同仁及承辦工程師
      If Ncp(1) = "P" And pa(9) = "000" And Text7.Text = "1003" And Text8 = "202" Then
         'Added by Morgan 2025/1/24
         If strSrvDate(1) >= P業務區劃分啟用日 Then
            Ncp(14) = PUB_GetPHandler(pa(1) & pa(2) & pa(3) & pa(4))
         Else
         'end 2025/1/24
            Ncp(14) = Pub_GetSpecMan("AB202") '---- 使用特殊設定檔（固定特例）
         End If 'Added by Morgan 2025/1/24
         strExc(9) = Ncp(14)
      End If
      '2008/5/20 end
      '93.9.16 CANCEL BY SONIA
      'End If
      '93.9.16 END
      
      'Added by Lydia 2016/09/21 P大陸案電話通知修正:204補正 FMP案掛最近一道程序的工程師(畫面預設) ; P案維持不變(所點選的那道進度的承辦人)
      If m_bolFMP And Ncp(1) = "P" And pa(9) = "020" And Me.Text7.Text = "1225" Then
         Ncp(14) = Text16
      End If
      'end 2016/09/21
      
      Ncp(16) = 0
      Ncp(17) = 0
      Ncp(18) = 0
      Ncp(20) = "N"
      'Added by Morgan 2022/8/22 FMP依職權電話通知修正1225自動收文的補正204預設要請款--敏莉
      If m_bolFMP And Ncp(1) = "P" And pa(9) = "020" And Me.Text7.Text = "1225" Then
         Ncp(20) = ""
      End If
      'end 2022/8/22
      Ncp(26) = "N"
      Ncp(27) = ""
      Ncp(32) = "N"
      Ncp(36) = Text19
      For i = 0 To 5
         Ncp(i + 37) = Text20(i)
      Next
      Ncp(43) = m_NewCP09
      Ncp(48) = ""
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.SaveNewCaseProgressDatabase("B", Ncp, intWhere) Then
      If Not ClsPDSaveNewCaseProgressDatabase("B", Ncp, intWhere) Then
        GoTo ErrorHandler
      End If
      
      'Add by Morgan 2004/2/18
      '若承辦人是王協理且未發文則要發EMail通知
      'Modified by Morgan 2023/9/14
      'm_stCP09 = Ncp(9)
      'm_stCP14 = Ncp(14)
      If m_bolAutoBCP Or (Ncp(14) = "99050" And Ncp(14) <> strOldCP14) Then
         '大陸OA來函內部收文通知--品薇112.8.10請作
         If m_bolAutoBCP Then
            '特殊基數(計件值)
            If m_strEV02 <> "" Then
               strSql = "insert into ExValue(EV01,EV02) values ('" & Ncp(9) & "'," & Val(m_strEV02) & ") "
               cnnConnection.Execute strSql, intI
            End If
            strExc(0) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)) & "已內部收文" & Label3(3) & "，"
            If Ncp(14) = "99050" And Ncp(14) <> strOldCP14 Then
               strExc(0) = strExc(0) & "惟原工程師" & IIf(strST04 <> "1", "已離職", "為" & strST02) & "，請重新分案並考慮是否更改基數。"
            Else
               strExc(0) = strExc(0) & "請自行至卷宗區參考。"
            End If
         '原內部收文分案通知
         Else
            strExc(0) = "分案通知"
         End If
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc13)" & _
            " values('" & strUserNum & "','" & Ncp(14) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & strExc(0) & "','" & Ncp(9) & "')"
         cnnConnection.Execute strSql, intI
      End If
      'end 2023/9/14
            
      Call PUB_UpdRelationCaseFixEP(cp(1), cp(2), cp(3), cp(4), Ncp(10), Label3(3)) 'Added by Morgan 2019/12/20
      
   End If
   '92.5.10 END
   
   'Add By Cheng 2001/12/31
   '若來函性質屬於爭議程序(18XX)
   If Left(Me.Text7.Text, 2) = "18" Then
      strTxt(intStep) = "UPDATE PATENT SET PA19='Y' WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        'Add By Cheng 2002/11/08
        cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If

      'Add By Cheng 2002/01/29
      '若來函性質為延期受理(1004), 若未勾選本案期限, 則以相關總收文號去更新案件進度檔的本所期限及法定期限
      '若來函性質為延期受理(1004), 若有勾選本案期限, 則更新下一程序檔的本所期限及法定期限
      'Modified by Morgan 2023/4/17 外專程序操作時只檢查但不更新(寰華案會輸入來函上含在途的期限而系統管制則不含)
      'If Text7 = 延期受理 Then
      If Text7 = 延期受理 And Pub_StrUserSt03 <> "F22" Then
      'end 2023/4/17
         BlnCheck = False
         With Me.MSHFlexGrid1
            For i = 1 To .Rows - 1
               If LCase("" & .TextMatrix(i, 0)) = "v" Then
                  BlnCheck = True
                  '92.7.5 modify by sonia
                  '本所期限
                  'strDate1 = IIf(Len(Trim(.TextMatrix(i, 2))) > 0, _
                  '            IIf(Len(Trim(.TextMatrix(i, 2)) <> 8), Val(Replace(.TextMatrix(i, 2), "/", "")) + 19110000, Replace(.TextMatrix(i, 2), "/", "")), _
                  '            "")
                  '法定期限
                  'strDate2 = IIf(Len(Trim(.TextMatrix(i, 3))) > 0, _
                  '            IIf(Len(Trim(.TextMatrix(i, 3)) <> 8), Val(Replace(.TextMatrix(i, 3), "/", "")) + 19110000, Replace(.TextMatrix(i, 3), "/", "")), _
                  '            "")
                  '本所期限
                  strDate1 = IIf(Len(Trim(Me.Text14(0).Text)) > 0, _
                               IIf(Len(Trim(Me.Text14(0).Text) <> 8), Val(Replace(Me.Text14(0).Text, "/", "")) + 19110000, Replace(Me.Text14(0).Text, "/", "")), _
                               "")
                  '法定期限
                  StrDate2 = IIf(Len(Trim(Me.Text14(1).Text)) > 0, _
                               IIf(Len(Trim(Me.Text14(1).Text) <> 8), Val(Replace(Me.Text14(1).Text, "/", "")) + 19110000, Replace(Me.Text14(1).Text, "/", "")), _
                               "")
                  '92.7.5 end
                  
                  'Modify by Morgan 2006/1/24 加NP01
                  strSql = "UPDATE NEXTPROGRESS SET NP08='" & strDate1 & "' " & _
                           " , NP09 = '" & StrDate2 & "' " & _
                           " WHERE NP22=" & .TextMatrix(i, 7) & " and np01='" & .TextMatrix(i, 8) & "'"
                  cnnConnection.Execute strSql
               End If
            Next i
         End With
         If BlnCheck = False Then
            '本所期限
            strDate1 = IIf(Len(Trim(Me.Text14(0).Text)) > 0, _
                        IIf(Len(Trim(Me.Text14(0).Text) <> 8), Val(Replace(Me.Text14(0).Text, "/", "")) + 19110000, Replace(Me.Text14(0).Text, "/", "")), _
                        "")
            '法定期限
            StrDate2 = IIf(Len(Trim(Me.Text14(1).Text)) > 0, _
                        IIf(Len(Trim(Me.Text14(1).Text) <> 8), Val(Replace(Me.Text14(1).Text, "/", "")) + 19110000, Replace(Me.Text14(1).Text, "/", "")), _
                        "")
            
            'Modified by Morgan 2013/10/2 改用更新strNew404CP43(原來用 cp(43) )
            'Modified by Morgan 2025/9/2 若已發文則不更新 + and cp27 is null
            strSql = "UPDATE CaseProgress SET CP06 = '" & strDate1 & "' " & _
                  " ,CP07 = '" & StrDate2 & "' " & _
                  " WHERE CP09 = '" & "" & strNew404CP43 & "' and cp27 is null"
            cnnConnection.Execute strSql
            
            '92.6.30 ADD BY SONIA
            'cancel by sonia 2015/9/7 延期後不需工程師再輸收卷註記 P-104887
            'strSql = "UPDATE EngineerProgress SET EP27=NULL,EP31=NULL " & _
            '         " WHERE EP02 = '" & "" & strNew404CP43 & "'"
            'cnnConnection.Execute strSql
            '92.6.30 END
         End If
      End If
   
'   End If

   'Add by Morgan 2008/5/13
   '2013/8/16 add by sonia 行政訴訟之智慧局答辯函存在智慧局答辯函P-099556
   'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP35='" & ChgSQL(Text30) & "',CP117='" & ChgSQL(Text33) & "' WHERE CP09='" & strReceiveNo & "'"
   If Me.Text7.Text = "1506" And Label3(1) = "行政訴訟" Then
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP35='" & ChgSQL(Text30) & "',CP117='" & ChgSQL(Text33) & "' WHERE CP09='" & Ncp(9) & "'"
   Else
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP35='" & ChgSQL(Text30) & "',CP117='" & ChgSQL(Text33) & "' WHERE CP09='" & strReceiveNo & "'"
   End If
   '2013/8/16 end
   cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1
   'END 2008/5/13

   '2008/7/15 ADD BY SONIA PCT檢索報告不印定稿由工程師處理,自動產生B類告知代理人內部收文,期限7天
   '2008/7/21 modify by sonia 加 PCT 1216國際初步審查報告
   'If pa(9) = "056" And Text7 = 檢索報告 Then
   strCP09_B = ""
'Modified by Morgan 2015/2/26 改不發文,印C類接洽單
'   If pa(9) = "056" And (Text7 = 檢索報告 Or Text7 = "1216") Then
'      '新增B類收文
'      Erase Ncp
'      ReDim Ncp(1 To TF_CP) As String
'      Ncp(1) = cp(1)
'      Ncp(2) = cp(2)
'      Ncp(3) = cp(3)
'      Ncp(4) = cp(4)
'      Ncp(5) = Label3(6)
'      Ncp(6) = ChangeWStringToTString(CompDate(2, 7, strSrvDate(1)))
'      Ncp(7) = ChangeWStringToTString(CompDate(2, 7, strSrvDate(1)))
'      'Modify by Morgan 2011/2/24 修正百年收文號問題
'      'Ncp(9) = "B" & Left(strSrvDate(2), 2)
'      Ncp(9) = "B" & CompAutoNumberYear(GetTaiwanThisYear)
'      Ncp(10) = "901"
'      Ncp(11) = "90"
'      Ncp(13) = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
'      'Modify by Morgan 2004/2/9
'      Ncp(12) = GetSalesArea(Ncp(13))
'
'      'Modify by Morgan 2008/9/30 改不預設承辦人需先由協理分案 --郭
'      ''原承辦人已離職,則預設為王協理
'      'Ncp(14) = cp(14)
'      'strExc(0) = "select ST06,ST04 from STAFF where ST01=" + CNULL(cp(14))
'      'intI = 1
'      'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'      'If intI = 1 Then
'      '   If RsTemp.Fields("ST04") <> "1" Then
'      '      Ncp(14) = "71011"
'      '   End If
'      'Else
'      '   Ncp(14) = "71011"
'      'End If
'      'm_stCP14 = Ncp(14)
'      ''若承辦人是王協理且未發文則要發EMail通知
'      'm_stCP09 = Ncp(9)
'      Ncp(14) = ""
'      'end 2008/9/30
'
'      Ncp(20) = "N"
'      Ncp(26) = "N"
'      Ncp(32) = "N"
'      Ncp(43) = m_NewCP09
'      If Not ClsPDSaveNewCaseProgressDatabase("B", Ncp, intWhere) Then
'         GoTo ErrorHandler
'      End If
'
'      '2008/7/21 ADD BY SONIA
'      strCP09_B = Ncp(9)
'      '2008/7/21 END
'   End If
   If bolPCTReport = True Then
      strCP09_B = Ncp(9)
   End If
   '2008/7/15 END
   
   'Add by Morgan 2007/5/4
   If bolCancelClose = True Then
      strSql = "UPDATE PATENT SET PA57=NULL,PA58=NULL,PA59=NULL" & _
         " WHERE PA01 = '" & pa(1) & "' AND PA02 = '" & pa(2) & "'" & _
         " AND PA03 = '" & pa(3) & "' AND PA04 = '" & pa(4) & "' "
      cnnConnection.Execute strSql
   End If
   'end 2007/5/4
      
   '2008/11/28 add by sonia 非台灣案1912通知已轉他所計算結餘
   'modify by sonia 2016/12/26 +1916解除代理人
   If (Text7 = "1912" Or Text7 = "1916") Then
      Pub_UpdateEndModCash pa(1), pa(2), pa(3), pa(4)
   End If
   '2008/11/28 end
   
   'Add by Morgan 2009/7/16 大陸有領證或陳述意見期限時檢查是否有分割案期限需更新
   'Modified by Morgan 2011/12/6 +台灣案的申復,再審期限或申復,再審的延期受理(更新函數內才檢查)
   'If pa(9) = "020" And (text8 = "205" Or text8 = "601") Then
   If (pa(9) = "020" And (Text8 = "205" Or Text8 = "601")) Or (pa(9) = "000" And (Text8 = "205" Or Text8 = "107" Or Text7 = "1004")) Then
      strSql = "select cp09 from divisioncase,caseprogress" & _
         " where dc05='" & cp(1) & "' and dc06='" & cp(2) & "'" & _
         " and dc07='" & cp(3) & "' and dc08='" & cp(4) & "'" & _
         " and cp01(+)=dc01 and cp02(+)=dc02 and cp03(+)=dc03 and cp04(+)=dc04 and cp10='307' and cp57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         st307Msg = ""
         '會有多個分割案
         Do While Not RsTemp.EOF
            If pa(9) = "020" Then
               strExc(1) = PUB_Update307Ref(RsTemp(0))
               
            'Added by Morgan 2011/12/6
            ElseIf pa(9) = "000" Then
               strExc(1) = PUB_Update307RefTw(RsTemp(0))
            End If
            
            If strExc(1) <> "" Then
               st307Msg = st307Msg & strExc(1) & vbCrLf
            End If
            RsTemp.MoveNext
         Loop
      End If
   End If
   
   'Add by Morgan 2009/9/16
   'P通知修正,申復時以發文日+6個月更新相關總收文號的催審期限(機關來函才要,代理人來函不必--郭)
   'Modify by Morgan 2011/3/22 台灣發明通知修正取消,因為審查前就可能會來函--郭
   'If Text7 = "1201" Or Text7 = "1202" Then
   'Modified by Morgan 2012/12/27 +最後通知1227
   'modify by sonia 2014/10/30 PCT申請109不更新催審期限,P-105752
   'If (Text7 = "1201" And Not (pa(8) = "1" And pa(9) = "000")) Or Text7 = "1202" Or Text7 = "1227" Then
   'Modified by Morgan 2019/7/9 +大陸發明的通知修正也不更新--郭
   'If ((Text7 = "1201" And Not (pa(8) = "1" And pa(9) = "000")) Or Text7 = "1202" Or Text7 = "1227") And cp(10) <> "109" Then
   If ((Text7 = "1201" And Not (pa(8) = "1" And (pa(9) = "000" Or pa(9) = "020"))) Or Text7 = "1202" Or Text7 = "1227") And cp(10) <> "109" Then
      'Modify by Morgan 2009/12/14 改1年
      'PUB_UpdateChkResultDate CompDate(1, 6, strSrvDate(1)), cp, m_NewCP09, Text7, cp(9)
      PUB_UpdateChkResultDate CompDate(0, 1, strSrvDate(1)), cp, m_NewCP09, Text7, cp(9)
   'Add by Morgan 2010/8/9 通知審查中1905更新相關號催審期限為來函日+6個月
   ElseIf Text7 = "1905" Then
      PUB_UpdateChkResultDate CompDate(1, 6, DBDATE(Text6)), cp, m_NewCP09, Text7, cp(9)
   End If
   
   '2012/11/26 ADD BY SONIA  順德及其關係企業加內部收文941分析,原工程師離職掛王副總,自動上齊備日(分所上下一工作日)以計算承辦期限,本所期限=承辦期限,但因ENGINEERPROGRESS_BEFORE5及CASEPROGRESS_AFTER6會造成承辦期限=本所期限-1天
   If m_CustX07166 = True Then
'2013/2/8 cancel by sonia 移至basQuery的PUB_CheckX07166Remind
'      str941CP14 = cp(14)
'      '2012/12/5 ADD BY SONIA 記錄上個畫面所點選收文號的承辦人(P非台灣案若原承辦人非工程師則改抓國內案承辦人) P-093775
'      If pa(9) <> 台灣國家代號 Then
'         strExc(0) = "SELECT ST03 FROM STAFF WHERE ST01='" & str941CP14 & "' AND ST03>='P1' AND ST03<='P11'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 0 Then
'            str941CP14 = PUB_GetInCaseCP14(pa(1), pa(2), pa(3), pa(4))
'         End If
'      End If
'      '2012/12/5 END
'2013/2/8 end
      strExc(0) = "SELECT ST04,DECODE(ST04,'1',ST06,'1') ST06 FROM STAFF WHERE ST01='" & str941CP14 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If "" & RsTemp(0).Value <> "1" Then
            'Added by Lydia 2023/04/24 修改王副總退休之相關控制
            If strSrvDate(1) >= "20230501" Then
                str941CP14 = "99050"  '5/1起原工程師離職掛李柏翰
            Else
            'end 2023/04/24
                str941CP14 = "71011" '原工程師離職掛王副總
            End If
         End If
      End If
      
      str941ReceiveNo = AutoNo("B", 6)
      'Modified by Morgan 2015/7/6 承辦人要加單引號
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
         "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43) VALUES " & _
         "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & _
         ",'" & str941ReceiveNo & "','941','90'," & CNULL(stCP12) & "," & CNULL(stCP13) & _
         ",'" & str941CP14 & "','N','N','N','" & Ncp(9) & "') "
      cnnConnection.Execute strSql
      If "" & RsTemp(1).Value <> "1" Then '分所
         strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & CompWorkDay(2, strSrvDate(1), 0) & " WHERE EP02='" & str941ReceiveNo & "'"
      Else
         strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & str941ReceiveNo & "'"
      End If
      cnnConnection.Execute strSql
      '更新本所期限=承辦期限,但因ENGINEERPROGRESS_BEFORE5及CASEPROGRESS_AFTER6會造成承辦期限=本所期限-1天
      strSql = "UPDATE CASEPROGRESS SET CP06=CP48 WHERE CP09='" & str941ReceiveNo & "' AND CP06 IS NULL"
      cnnConnection.Execute strSql
   End If
   '2012/11/26 END
   
   'Added by Morgan 2021/10/6
   If m_CustX69365 = True Then
      '長庚醫院案件要收[轉公文]先簡單報告
      'Removed by Morgan 2022/3/28 取消轉公文,改同其他3家直接報告,但本所期限改為 +14天-3個工作天 --黃教威
      'm_str1998CP09 = AutoNo("D", 6)
      'strSql = "INSERT INTO CASEPROGRESS(cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp27" & _
         ",cp32,cp43) SELECT cp01,cp02,cp03,cp04,cp05,cp06,cp07,'" & m_str1998CP09 & "','1998',cp12,cp13,'" & strUserNum & "'" & _
         ",'N','N'," & strSrvDate(1) & ",'N',cp09 FROM CASEPROGRESS WHERE CP09='" & m_NewCP09 & "'"
      'cnnConnection.Execute strSql, intI
      
      ''P案轉公文用系統的來函定稿，判發人同來函
      'PUB_AddLetterProgress m_str1998CP09, 0, True, Text37, IIf(Val(Text14(1)) > 0, True, False), pa(26), Text7, pa(75)
      
      'PUB_SetX69365Case1998CP06 m_str1998CP09 '設定長庚醫院案件轉公文管制日(所限)
      'end 2022/3/28
      
      '設定長庚醫院案件OA發文管制日(所限)
      PUB_SetX69365CaseOACP06 m_NewCP09
      
      'Added by Morgan 2025/6/18
      If Text16 = "99050" Then
         Call PUB_SendMail(strUserNum, "99050", m_NewCP09, "分案通知")
      End If
      'end 2025/6/18
   'Added by Morgan 2022/4/15
   ElseIf m_bolW2001XCase Then
      'Added by Morgan 2022/11/4 原工程師離職只需通知王副總分案，程序分案會再通知新工程師--有跟郭確認過
      'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
      If Text16 = "99050" Then
         Call PUB_SendMail(strUserNum, "99050", m_NewCP09, "分案通知")
      Else
      'end 2022/11/4
      
         strExc(0) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
         strExc(1) = strExc(0) & "案已收到「" & Label3(4) & "」，請於「承辦期限」前完成分析通知函，謝謝。"
         strExc(2) = PUB_GetW2001InCC(pa(26), pa(158))
         'Modified by Morgan 2022/8/9 應該要發給承辦人CC給窗口及智權人員
         'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " select  '" & strUserNum & "','" & strExc(2) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",replace('" & ChgSQL(strExc(1)) & "','承辦期限',sqldatet(cp48)),'如旨' from caseprogress where cp09='" & m_NewCP09 & "'"
         'Modified by Morgan 2023/5/10 W2001-->stCP13
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " select  '" & strUserNum & "',cp14,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",replace('" & ChgSQL(strExc(1)) & "','承辦期限',sqldatet(cp48)),'如旨','" & stCP13 & ";" & strExc(2) & "'" & _
            " from caseprogress where cp09='" & m_NewCP09 & "'"
         cnnConnection.Execute strSql, intI
         
      End If 'Added by Morgan 2022/11/4
   'end 2022/4/15
   End If
   'end 2021/10/6
   
   'Added by Morgan 2013/10/9
   '台灣新型收文通知申復時若有申請技術報告的催審期限則後延6個月
   If pa(9) = "000" And pa(8) = "2" And Text7 = "1221" Then
      strExc(0) = "select np01,np08,np09,np22 from nextprogress,caseprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06 is null and np07='411' and cp09(+)=np01 and cp10='421'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(1) = CompDate(1, 6, RsTemp("np09"))
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01='" & RsTemp("np01") & "' and np07='411' and np22=" & RsTemp("np22")
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2013/10/9
   
   'Added by Morgan 2014/1/14
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, m_NewCP09, pa(1), pa(2), pa(3), pa(4), Text7
   End If
   'end 2014/1/14
   
   'Add by Sindy 2016/10/5
   If m_strIR01 <> "" Then
      'Added by Morgan 2019/3/8
      '依職權電話通知修正1225要自動歸卷--品薇
      If Text7 = "1225" Then
         'Modify By Sindy 2022/11/10 + IIf(pa(9) <> 台灣國家代號, "PAT", "RX")
         'Modified by Morgan 2024/5/8 大陸案的副檔名也改為ALTR--品薇
         'If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, m_NewCP09, IIf(Pub_StrUserSt03 = "F22", "ALTR", IIf(pa(9) <> 台灣國家代號, "PAT", "RX"))) = False Then
         If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, m_NewCP09, IIf(pa(9) <> 台灣國家代號, "ALTR", "RX")) = False Then
         'end 2024/5/8
            GoTo ErrorHandler
         End If
      End If
      'end 2019/3/8
      'Add By Sindy 2019/12/13 一般來函輸入，選擇(1201)通知補正，(1202)審查意見通知，輸入後請將整封郵件存入系統
      If Text7 = "1201" Or Text7 = "1202" Then
         'Modify By Sindy 2022/11/9 + IIf(pa(9) <> 台灣國家代號, "PAT", "RX")
         If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, m_NewCP09, IIf(Pub_StrUserSt03 = "F22", "ALTR", IIf(pa(9) <> 台灣國家代號, "PAT", "RX"))) = False Then 'PAT.陸代郵件
            GoTo ErrorHandler
         End If
      End If
      '2019/12/13 END
      
      'Modified by Morgan 2020/8/12 +傳 m_NewCP09, m_bolReKeyInOK
      'Modified by Morgan 2020/8/18 有法限才要傳
      'Modify By Sindy 2022/6/16 F2外專不做2次確認 + And Left(Pub_StrUserSt03, 2) <> "F2"
      If Text14(1) <> "" And Left(Pub_StrUserSt03, 2) <> "F2" Then
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010504_1", m_NewCP09, m_bolReKeyInOK
         bolReKeyInCase = True 'Added by Morgan 2023/4/10
      Else
         'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", m_NewCP09, "")
         'Modified by Lydia 2023/05/18 +不開啟附件 , , , False
          PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010504_1", IIf(Pub_StrUserSt03 = "F22", m_NewCP09, ""), , , False
      End If
   End If
   '2016/10/5 END
   
   'Added by Morgan 2014/4/10 電子化-新增信函進度檔
   If pa(9) = "000" Then
      'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
      'Modified by Morgan 2020/3/11 延期受理有輸入期限但非掛號(一般不通知) Ex:P-122323
      'Modified by Morgan 2021/10/6 +長庚醫院案件會自動收[轉公文]先通知客戶，來函分析信由工程師撰寫
      'Modified by Morgan 2022/3/28 長庚醫院案件取消轉公文 --黃教威
      'PUB_AddLetterProgress m_NewCP09, 1 + Val(Text13), IIf(Text15(0) <> "N" And m_CustX69365 = False, True, False), Text37, IIf(Text14(1) <> "" And Text7 <> "1004", True, False), pa(26), Text7, pa(75)
      PUB_AddLetterProgress m_NewCP09, 1 + Val(Text13), IIf(Text15(0) <> "N", True, False), Text37, IIf(Text14(1) <> "" And Text7 <> "1004", True, False), pa(26), Text7, pa(75)
      'end 2022/3/28
   'Added by Morgan 2016/6/14 非臺灣案電子化
   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      intI = 2
      
      'Modified by Morgan 2019/3/8 依職權電話通知修正1225沒有公文
      If Text7 = "1225" Then
         strSql = "update caseprogress set cp121='Y' where cp09='" & m_NewCP09 & "'"
         cnnConnection.Execute strSql, intI
         intI = 1
      End If
      'end 2019/3/8
      
      'Modified by Morgan 2021/10/6 +長庚醫院案件會自動收[轉公文]先通知客戶，來函分析信由工程師撰寫
      'Modified by Morgan 2022/3/28 長庚醫院案件取消轉公文 --黃教威
      'PUB_AddLetterProgress m_NewCP09, intI + Val(Text13), IIf(Text15(0) <> "N" And m_CustX69365 = False, True, False), Text37, IIf(Text14(1) <> "", True, False), pa(26), Text7, pa(75)
      'Modified by Morgan 2024/4/24 大陸 1815 第三方意見, 1508 國知局答辯函 要掛號直寄 Ex:P-097468
      PUB_AddLetterProgress m_NewCP09, intI + Val(Text13), IIf(Text15(0) <> "N", True, False), Text37, IIf(Text14(1) <> "", True, IIf(Text7 = "1815" Or Text7 = "1508", True, False)), pa(26), Text7, pa(75)
      'end 2022/3/28
   'end 2016/6/14
   End If
   'end 2014/4/10
      
   'Add by Lydia 2014/11/26 台灣案主管機關來函,針對1004(延期受理)
   'Modified by lydia 2022/08/31  (2022/08/15) 開放P大陸案
   'If pa(9) = "000" And Me.Text7.Text = "1004" Then
   'Modified by Lydia 2022/10/11 經查此設定並不適用於外專及日專，故請協助排除FMP案
   'If (pa(9) = "000" Or pa(9) = "020") And Me.Text7.Text = "1004" Then
   If (pa(9) = "000" Or pa(9) = "020") And Me.Text7.Text = "1004" And m_bolFMP = False Then
      ReDim mPty1004(0 To 3) As String
      Check2mail1004 pa(), Label3(5), 0 '判斷
      'Modified by Lydia 2025/04/02 debug: 2022/08/15 開放P大陸案，此案有工程師未發文的程序，發E-MAIL通知工程師收到來函的內容
      'If mPty1004(0) = "3" Then '已發文
      If mPty1004(0) = "3" And pa(9) = "000" Then '限:台灣案已發文
        mPtyNo = AutoNo("B", 6) '產生自動內部收文933
        strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
           "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43) VALUES " & _
           "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & _
           ",'" & mPtyNo & "','933','90'," & CNULL(stCP12) & "," & CNULL(stCP13) & _
           "," & CNULL(mPty1004(1)) & ",'N','N','N','" & mPty1004(2) & "') "
        cnnConnection.Execute strSql
      End If
   End If
   'end 'Add by Lydia 2014/11/26
   
   'Added by Lydia 2017/05/09 後案官方來函性質「視為未主張」1918
   If Me.Text7.Text = "1918" Then
       '目前主張國內優先權發文後，被主張的前案會閉卷，往後若輸入來函性質為視為未主張且閉卷原因為88被主張國內優先權的，請系統自動取消前案之閉卷。
       If cp(10) = "121" Then
          Set RsTemp = PUB_ReadPDStateNew(pa, cp(10), True)
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
                strSql = "DELETE FROM PRIDATE WHERE PD01='" & pa(1) & "' AND PD02='" & pa(2) & "' AND PD03='" & pa(3) & "' AND PD04 ='" & pa(4) & "' "
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
          'Modified by Lydia 2023/06/15 改備註的順序
          'strSql = "UPDATE PATENT SET PA91=PA91||'" & ChangeTStringToTDateString(strSrvDate(2)) & " 視為未主張的優先權資料:" & strExc(4) & " ;' WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "'  "
          'Modified by Lydia 2025/08/06 debug: 發現語法被Mark, 取消Mark
          strSql = "UPDATE PATENT SET PA91='" & ChangeTStringToTDateString(strSrvDate(2)) & " 視為未主張的優先權資料:" & strExc(4) & " ;'||PA91 WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "'  "
          cnnConnection.Execute strSql
       End If
       
       '更新公開和實審期限
       strExc(5) = PUB_GetFirstPriDate(pa)
       strExc(9) = ""
       
         '公開或實審期限的相關總收文號用申請程序的收文號
         strSql = "select cp09 from caseprogress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and instr('" & NewCasePtyList & "',cp10)>0 and cp159=0 order by cp05 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strExc(9) = RsTemp(0)
         End If
       PUB_UpdCfpDate2 pa(1), pa(2), pa(3), pa(4), strExc(5), strExc(9)
   End If
   'end 2017/05/09
      
   If m_USCaseNo <> "" Then PUB_SetUsIDS pa(1), pa(2), pa(3), pa(4), m_NewCP09, Text6.Text, , , , True 'Added by Morgan 2020/12/18 美國IDS期限管制
   
   'Added by Morgan 2018/10/2
   'Modified by Lydia 2023/04/24 修改王副總退休之相關控制
   'If strCP09_B <> "" And m_bolFMP = False And Ncp(14) <> "71011" Then
   If strCP09_B <> "" And m_bolFMP = False And (Ncp(14) <> "71011" Or (strSrvDate(1) >= "20230501" And Ncp(14) <> "99050")) Then
      Pub_COrderInform strCP09_B
      bolSavPdf = True
   End If
   'end 2018/10/2
   
   'Added by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail
   'Modified by Lydia 2023/05/26 已閉卷不通知
   'Move by Lydia 2023/05/26 從commit上方移過來,
   Dim bolFMP2mail As Boolean  'Added by Lydia 2023/05/26
   If m_bolFMP = True And m_bolFMP2 = True And pa(57) = "" Then
      'Modified by Lydia 2023/10/31 傳入C類收文號  m_NewCP09
      bolFMP2mail = Pub_SetFMP2toCMail(pa(1), pa(2), pa(3), pa(4), Text7.Text, cp(14), m_NewCP09) '傳入相關收文的承辦人
   End If
   'end 2023/05/17
   
   'Added by Morgan 2020/4/10
   'FMP有期限之案件EMAIL通知(寰華案不必--敏莉)
   'Modified by Lydia 2023/05/26 排除-寰華案無期限之官方來函，系統自動發Mail => And bolFMP2mail = False
   If m_bolFMP = True And Left(Pub_StrUserSt03, 1) <> "F" And bolFMP2mail = False Then
      'Modified by Morgan 2020/9/15 未閉卷的併入工程師通知信
      If pa(57) = "Y" Then
         'Modified by Morgan 2023/5/25 FMP電子化所有來函應該都要EMail通知
         'PUB_FMPCaseInform m_NewCP09
         PUB_FMPCaseInform m_NewCP09, False, True, Left(Pub_StrUserSt03, 1) = "F", bolReKeyInCase
         'end 2023/5/25
      End If
      'end 2020/9/15
   End If
   'end 2020/9/15
   'end 2020/4/10
   
   'Added by Morgan 2021/1/11
   '大陸案主管機關來函輸入(1506)復審委員會答辯函(1508)國知局答辯函，發E-MAIL通知工程師
   'Modified by Morgan 2024/10/16 發信對象應為來函的承辦人而非原程序承辦人(可能離職),cp09改抓m_NewCP09
   If pa(9) = "020" And (Text7.Text = "1506" Or Text7.Text = "1508") Then
      strExc(0) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
      strExc(1) = "已收到 " & strExc(0) & "，主管機關來函（" & Label3(4) & "）內容請參照卷宗區的電子檔，針對答辯函若有需補充說明的內容請於開庭前通知代理人。"
      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
         " select '" & strUserNum & "',cp14,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         ",'" & ChgSQL(strExc(1)) & "','如旨' from caseprogress where cp09='" & m_NewCP09 & "'"
      cnnConnection.Execute strSql, intI
      
   'Added by Morgan 2024/5/8
   ElseIf m_bolNoCP27 Then
      strExc(0) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
      strExc(1) = "已收到 " & strExc(0) & "，主管機關來函（" & Label3(4) & "）內容請參照卷宗區的電子檔。"
      If Text7 = 1225 Then
         strExc(2) = "本案請自行確認是否通知客戶，若要通知請轉知郭經理報價；若不通知客戶則請聯絡程序內部收文補正。"
      Else
         strExc(2) = "如旨"
      End If
      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
         " select '" & strUserNum & "',cp14,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(2)) & "' from caseprogress where cp09='" & m_NewCP09 & "'"
      cnnConnection.Execute strSql, intI
   End If
   'end 2021/1/11
   
   'Added by Morgan 2023/4/10 從下面移上來
   If IsEmptyText(strProgressNo) = False Then
      If m_bolFMP And pa(57) = "" Then
         
         'Added by Morgan 2025/3/19 寰華案改正本給工程師,副本給主管--敏莉/Wilson
         If Left(Pub_StrUserSt03, 1) = "F" Then
            mRCno = Trim(Text16.Text)
            mCCno = PUB_GetFCPEngSup(mRCno)
         Else
         'end 2025/3/19
            mCCno = Trim(Text16.Text)
            mRCno = PUB_GetFCPEngSup(mCCno)
         End If
         
         If mCCno = mRCno Then mCCno = ""
         strExc(0) = "SELECT NVL(PA05,NVL(PA06,PA07)) pa05,nvl(FA05||' '||FA63,'') as faname1, nvl(FA04,'') as faname2, nvl(FA06,'') as faname3,CP48,NP23 " & _
                     "FROM PATENT,FAGENT,caseprogress,nextprogress WHERE substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) and CP01(+)=PA01 And CP02(+)=PA02 And CP03(+)=PA03 And CP04(+)=PA04 " & _
                     "And CP09 = '" & m_NewCP09 & "' and np01(+)=cp09 and np06(+) is null and PA01='" & pa(1) & "' and PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(0) = "" & RsTemp.Fields("PA05")
            strExc(1) = "" & RsTemp.Fields("faname1")
            strExc(2) = "" & RsTemp.Fields("faname2")
            strExc(3) = "" & RsTemp.Fields("faname3")
            strExc(4) = "" & RsTemp.Fields("CP48") '承辦期限
            strExc(5) = "" & RsTemp.Fields("NP23") '約定期限
         End If
         If Len(strExc(1)) > 0 Then '代理人名稱(英->中->日)
            strExc(1) = "代理人　：" & strExc(1)
         ElseIf Len(strExc(2)) > 0 Then
            strExc(1) = "代理人　：" & strExc(2)
         Else
            strExc(1) = "代理人　：" & strExc(3)
         End If
         oSubject = pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
         '發E-Mail通知承辦人
         oContext = "※若改承辦人,請工程師主管轉寄給新的承辦人,及c.c.原承辦人" & vbCrLf  'Added by Morgan 2020/9/17 從最後面移到最前面--敏莉
         'Added by Lydia 2024/06/05 為避免OA不請款又舜禹翻譯OA產生費用情形，除了在改承辦人時(附件請作單)，在輸入C類來函通知信，判斷若承辦人已是內專工程師時新增內文
         If Left(Pub_StrUserSt03, 1) = "F" And Mid(Text16, 4, 1) = "9" Then
            '1. 請通知信一併cc國外部對接主管(Stellar, Alina, Red, Chunyu)，
            '2. 請新增內文：※請外專主管確認本案OA如有不能請款之情形，則通知Wilson是否轉回外專處理，以及通知Sharon不分舜禹翻譯OA (1日以內)
            oContext = oContext & "※請外專主管確認本案OA如有不能請款之情形，則通知Wilson是否轉回外專處理，以及通知Sharon不分舜禹翻譯OA (1日以內)" & vbCrLf
            strSql = PUB_GetFCPEngSup(Text16, , , True)
            If InStr(mRCno & mCCno, strSql) = 0 Then
               mCCno = mCCno & ";" & strSql
            End If
            'Added by Lydia 2024/10/25 增加sharon為 副本收受者
            strSql = Pub_GetSpecMan("M")
            If strSql <> "" And InStr(mCCno & ";", strSql) = 0 Then
               mCCno = mCCno & IIf(mCCno <> "", ";", "") & strSql
            End If
            'end 2024/10/25
         End If
         oContext = oContext & vbCrLf '將多空一行調到這裡
         'end 2024/06/05
         
         oContext = oContext & _
                    "本所案號：" & oSubject & "　　" & vbTab & vbTab & "來函收文日：" & ChangeTStringToTDateString(Trim(Label3(6))) & vbCrLf & _
                    "專利名稱：" & strExc(0) & vbCrLf & _
                       strExc(1) & vbCrLf & _
                    "承辦人　：" & Trim(Label3(7)) & vbCrLf & _
                    "本所期限：" & ChangeTStringToTDateString(Trim(Text14(0))) & "　　　　" & vbTab & vbTab & "法定期限：" & ChangeTStringToTDateString(Trim(Text14(1))) & vbCrLf & _
                    "承辦期限：" & IIf(Len(strExc(4)) > 0, ChangeWStringToTDateString(strExc(4)), "　　　　") & "　　　　" & vbTab & vbTab & "來函性質：" & Trim(Label3(4)) & vbCrLf & _
                    "約定期限：" & ChangeWStringToTDateString(strExc(5)) & vbCrLf
             
         oSubject = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
          
         If stBCP10 = "901" Then
            oSubject = oSubject & "，同時內部收文【告代】"
            oContext = oContext & vbCrLf & "案件性質：告代" & vbCrLf
            oContext = oContext & "承辦期限：" & ChangeWStringToTDateString(stBCP48) & vbCrLf
            oContext = oContext & "本所期限：" & ChangeWStringToTDateString(stBCP06) & vbCrLf
         End If
         
         If Left(Pub_StrUserSt03, 1) = "F" Then '寰華
            'Modified by Lydia 2024/04/26 機械組要加註
            'Modified by Morgan 2025/3/19 --敏莉/Wilson
            'oSubject = IIf(pa(150) = "4", "【機械設計組】", "") & "FMP(寰華)案" & Trim(Label3(4)) & "通知:" & oSubject & "，主管請分案，工程師請處理後續流程，謝謝！"
            oSubject = IIf(pa(150) = "4", "【機械設計組】", "") & "FMP(寰華案)" & Trim(Label3(4)) & "通知:" & oSubject & "，工程師請處理後續流程，謝謝！"
            'end 2025/3/19
            strExc(0) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
            If InStr(mCCno & ";", strExc(0)) = 0 And strExc(0) <> "" Then
                mCCno = mCCno & ";" & strExc(0)
            End If
            If InStr(mRCno & ";" & mCCno, strUserNum) = 0 Then
               mCCno = mCCno & IIf(mCCno <> "", ";", "") & strUserNum
            End If
            If m_bolReKeyInOK Then oSubject = "(重發，請以此封為準)" & oSubject 'Added by Morgan 2023/4/17
         Else
            'Removed by Morgan 2023/5/25 不必再CC給FMP案外專程序窗口--敏莉
            'mCCno = mCCno & ";" & Pub_GetSpecMan("FMP案外專程序窗口")
            'end 2023/5/25
            oSubject = "FMP案" & Trim(Label3(4)) & "通知:" & oSubject & "，主管請分案，工程師請處理後續流程，謝謝！"
         End If
         strCMemo = PUB_FCPCFormMemo(m_NewCP09)  'Added by Morgan 2023/6/21
         If strCMemo <> "" Then
            oContext = oContext & vbCrLf & "備註:" & vbCrLf & strCMemo
         End If
          
         '需2次確認的來函改等職代輸入並確認後才通知
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc12,mc13)" & _
            " values('" & strUserNum & "','" & mRCno & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & ChgSQL(oSubject) & "','" & ChgSQL(oContext) & "','" & mCCno & "'," & IIf(bolReKeyInCase, "99999999", "0") & ",'" & m_NewCP09 & "')"
         cnnConnection.Execute strSql, intI
         m_bolFMPNoPrint = True
      End If
   End If
   'end 2023/4/10
   
   'Added by Morgan 2023/6/27
   If m_bolBPFCase Then Pub_COrderInform m_NewCP09, , IIf(Text16 = "A0029", "", "A0029")
   'end 2023/6/27
   
   'Added by Morgan 2024/12/3
   'FMP通知審查中1905預設不出定稿，由系統自動發信告知承辦人員，主旨為，◎P-XXXXX本案經催審後，官方回覆通知審查中
   'Modified by Morgan 2025/1/2 補案件性質1905條件
   'Modified by Morgan 2025/10/20 +復審受理通知1234--品薇/Joanne
   If m_bolFMP And Not m_bolFMP2 And (Text7.Text = "1905" Or Text7.Text = "1234") Then
      oSubject = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
      If Text7.Text = "1905" Then
         oSubject = oSubject & "本案經催審後，官方回覆通知審查中。"
      'Added by Morgan 2025/10/20
      Else
         oSubject = oSubject & "已收到" & Label3(4)
      'end 2025/10/20
      End If
      
      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
         " select '" & strUserNum & "',cp13,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         ",'" & oSubject & "'" & _
         ",'如旨',st52 from caseprogress,staff where cp09='" & m_NewCP09 & "' and st01(+)=cp13"
      cnnConnection.Execute strSql, intI
      
      m_bolFMPNoPrint = True
   End If
   'end 2024/12/3
cnnConnection.CommitTrans
FormSave = True

   'Add by Morgan 2009/7/16
   If st307Msg <> "" Then
      MsgBox st307Msg
   End If

   '2013/2/8 add by sonia 順德及其關係企業案件,若承辦人是王協理且未發文則要發EMail通知
   If m_CustX07166 = True Then
      'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
      If str941CP14 = "99050" Then
         Call PUB_SendMail(strUserNum, "99050", str941ReceiveNo, "分案通知")
         
      'Added by Morgan 2018/11/13 分析B類接洽單改承辦人是王副總要印紙本分案,其他存卷宗區並EMail通知承辦工程師
         g_PrtForm001.PrintCForm str941ReceiveNo
      Else
         g_PrtForm001.PrintCForm str941ReceiveNo, , , True
         Pub_COrderInform str941ReceiveNo, True
      'end 2018/11/13
      
      End If
   End If
   '2013/2/8 end
   
'Removed by Morgan 2015/11/17 配合無紙化取消列印--郭
'   '2013/3/18 ADD BY SONIA 順德及其關係企業案件之分析B類收文加印接洽單
'   If m_CustX07166 = True Then
'      g_PrtForm001.PrintCForm str941ReceiveNo
'   End If
'   '2013/3/18 END
'end 2015/11/17
   
   'Added by Morgan 2024/10/4
   If pa(1) = "P" And pa(9) <> "000" And m_bolFMP = False Then
      If Pub_B911NotPay(pa(1), pa(2), pa(3), pa(4)) = True Then
          MsgBox "此案有未收款！", vbExclamation
      End If
   End If
   'end 2024/10/4
   
   If FormSave = True Then
     
      If IsEmptyText(strProgressNo) = False Then
         g_PrtForm001.PrintForm strProgressNo, pa(1), pa(2), pa(3), pa(4)
         'Modified by Morgan 2017/4/28 FMP未閉卷才交工程師報告客戶,已閉卷直接交FCP程序--潘韻丞(David 確認)
         'If m_bolFMP Then
         If m_bolFMP And pa(57) = "" Then
         'end 2017/4/28
         
            'Modify by Morgan 2009/12/3 改來函不自動上發文日(印C類接洽單)
            'bol901 = True
            'g_PrtForm001.PrintForm strProgressNo, pa(1), pa(2), pa(3), pa(4), m_901CP09
            'bol901 = False
            'Modified by Lydia 2020/04/06 因應防疫在家上班作業，請將FMP案key來函產生的C類接洽記錄單回存到卷宗區
            '                                         比照FCP案C類接洽單同時列印並且上傳到卷宗區frm06010603_3: 原本就不傳入特殊備註,等到列印時再抓特殊備註
            'g_PrtForm001.PrintCForm m_NewCP09, , stCP48Desc
            'Modified by Morgan 2020/9/15 +strCMemo
            g_PrtForm001.PrintCFormNew m_NewCP09, , stCP48Desc, True, strCMemo
            'end 2009/12/3
            
'Removed by Morgan 2023/4/10 EMail要2次確認完才要發，改移到上面先寫暫存
'            'Add by Lydia 2014/10/16 FMP案會列印 C類接洽單, 請同時E-MAIL給畫面上之承辦人, 副本發給該員之工程師組別主管.
'            'Modified by Lydia 2020/08/24 改用模組
'            'strExc(0) = "SELECT ST01,ST04,decode(ST16,'1','T','2','R','3','S','4','T1','') mst16 FROM STAFF WHERE ST01='" & Trim(Text16.Text) & "' "
'            'intI = 1
'            'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            'If intI = 1 Then
'            '  strExc(0) = "" & RsTemp.Fields("mst16")
'            '  mCCno = Pub_GetSpecMan(strExc(0))
'            '  mRCno = RsTemp.Fields("ST01")
'            '  If mCCno = mRCno Then mCCno = "" '承辦人已是主管則不必再發副本
'            'End If
'            'Modified by Morgan 2020/9/15 改寄承辦工程師主管,cc承辦工程師; Phoebe(FMP案外專程序窗口)
'            'mRCno = Trim(Text16.Text)
'            'mCCno = PUB_GetFCPEngSup(mRCno)
'            'If mCCno = mRCno Then mCCno = ""
'            mCCno = Trim(Text16.Text)
'            mRCno = PUB_GetFCPEngSup(mCCno)
'            If mCCno = mRCno Then mCCno = ""
'            'end 2020/9/15
'            'end 2020/08/24
'
'            strExc(0) = "SELECT NVL(PA05,NVL(PA06,PA07)) pa05,nvl(FA05||' '||FA63,'') as faname1, nvl(FA04,'') as faname2, nvl(FA06,'') as faname3,CP48,NP23 " & _
'                        "FROM PATENT,FAGENT,caseprogress,nextprogress WHERE substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) and CP01(+)=PA01 And CP02(+)=PA02 And CP03(+)=PA03 And CP04(+)=PA04 " & _
'                        "And CP09 = '" & m_NewCP09 & "' and np01(+)=cp09 and np06(+) is null and PA01='" & pa(1) & "' and PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "' "
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strExc(0) = "" & RsTemp.Fields("PA05")
'               strExc(1) = "" & RsTemp.Fields("faname1")
'               strExc(2) = "" & RsTemp.Fields("faname2")
'               strExc(3) = "" & RsTemp.Fields("faname3")
'               strExc(4) = "" & RsTemp.Fields("CP48") '承辦期限
'               strExc(5) = "" & RsTemp.Fields("NP23") '約定期限
'            End If
'            If Len(strExc(1)) > 0 Then '代理人名稱(英->中->日)
'               strExc(1) = "代理人　：" & strExc(1)
'            ElseIf Len(strExc(2)) > 0 Then
'               strExc(1) = "代理人　：" & strExc(2)
'            Else
'               strExc(1) = "代理人　：" & strExc(3)
'            End If
'            oSubject = pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
'            '發E-Mail通知承辦人
'            oContext = "※若改承辦人,請工程師主管轉寄給新的承辦人,及c.c.原承辦人" & vbCrLf & vbCrLf 'Added by Morgan 2020/9/17 從最後面移到最前面--敏莉
'            oContext = oContext & _
'                       "本所案號：" & oSubject & "　　" & vbTab & vbTab & "來函收文日：" & ChangeTStringToTDateString(Trim(Label3(6))) & vbCrLf & _
'                       "專利名稱：" & strExc(0) & vbCrLf & _
'                          strExc(1) & vbCrLf & _
'                       "承辦人　：" & Trim(Label3(7)) & vbCrLf & _
'                       "本所期限：" & ChangeTStringToTDateString(Trim(Text14(0))) & "　　　　" & vbTab & vbTab & "法定期限：" & ChangeTStringToTDateString(Trim(Text14(1))) & vbCrLf & _
'                       "承辦期限：" & IIf(Len(strExc(4)) > 0, ChangeWStringToTDateString(strExc(4)), "　　　　") & "　　　　" & vbTab & vbTab & "來函性質：" & Trim(Label3(4)) & vbCrLf & _
'                       "約定期限：" & ChangeWStringToTDateString(strExc(5)) & vbCrLf
'
'             'Modified by Morgan 2020/9/15
'             'oSubject = oSubject & "　收文-" & Trim(Label3(4)) & "，請自行去調卷處取卷，謝謝！"
'             oSubject = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
'
'            'Added by Morgan 2021/11/10 有收文告代時，主旨及內文都要加--淑華
'            If stBCP10 = "901" Then
'               oSubject = oSubject & "，同時內部收文【告代】"
'               oContext = oContext & vbCrLf & "案件性質：告代" & vbCrLf
'               oContext = oContext & "承辦期限：" & ChangeWStringToTDateString(stBCP48) & vbCrLf
'               oContext = oContext & "本所期限：" & ChangeWStringToTDateString(stBCP06) & vbCrLf
'            End If
'            'end 2021/11/10
'
'             'Modified by Morgan 2020/9/17 寰華案不用cc給Phoebe--敏莉
'             If Left(Pub_StrUserSt03, 1) = "F" Then '寰華
'               'Modified by Lydia 2022/05/10 寰華案與FMP案主旨一致
'               'oSubject = "FMP(寰華)案" & Trim(Label3(4)) & "通知:" & oSubject & "，主管請分案，卷隨後附上，謝謝！"
'               oSubject = "FMP(寰華)案" & Trim(Label3(4)) & "通知:" & oSubject & "，主管請分案，工程師請處理後續流程，謝謝！"
'               'Added by Lydia 2022/04/22 核駁及一般來函皆CC給程序
'               strExc(0) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
'               If InStr(mCCno & ";", strExc(0)) = 0 And strExc(0) <> "" Then
'                   mCCno = mCCno & ";" & strExc(0)
'               End If
'               'end 2022/04/22
'                'Added by Lydia 2023/01/09 FCP和寰華案 key C類來函，若key來函人員沒有在系統自動發Outlook的收件者中，副本請加上key來函人員;
'                If InStr(mRCno & ";" & mCCno, strUserNum) = 0 Then
'                   mCCno = mCCno & IIf(mCCno <> "", ";", "") & strUserNum
'                End If
'                'end 2023/01/09
'             Else
'               mCCno = mCCno & ";" & Pub_GetSpecMan("FMP案外專程序窗口")
'               'Modified by Morgan 2021/11/10 --淑華
'               'oSubject = "FMP案" & Trim(Label3(4)) & "通知:" & oSubject & "，主管請分案，工程師請自行去調卷處取卷，紙本公文後補，謝謝！"
'               oSubject = "FMP案" & Trim(Label3(4)) & "通知:" & oSubject & "，主管請分案，工程師請處理後續流程，謝謝！"
'               'end 2021/11/10
'             End If
'             If strCMemo <> "" Then
'               oContext = oContext & vbCrLf & "備註:" & vbCrLf & strCMemo
'             End If
'             'end 2020/9/15
'            PUB_SendMail strUserNum, mRCno, "", oSubject, oContext, "", "", , , , mCCno, "", "", ""
'            'end Lydia 2014/10/16
'end 2023/4/10

         End If
      End If
      'Add by Morgan 2005/3/16 若發明核准且為一案兩請則列印新型案自請撤回接洽單
      If m_DualAppNP22 <> "" Then
         g_PrtForm001.PrintForm m_DualAppNP22
      End If
      '2008/7/21 add by sonia
      If IsEmptyText(strCP09_B) = False Then
         'Modify by Morgan 2010/6/1 改用C類格式(分案要用 Ex.P-91914 告代)
         'g_PrtForm001.PrintForm "", cp(1), cp(2), cp(3), cp(4), strCP09_B
         g_PrtForm001.PrintCForm strCP09_B, , , bolSavPdf
      End If
      '2008/7/21 end
    End If
    
Exit Function
ErrorHandler:
'    cnnConnection.RollbackTrans
    If FormSave = False Then cnnConnection.RollbackTrans
'    FormSave = False
End Function

Private Sub Combo3_Click()
    'Add By Cheng 2002/12/18
    '將各式修正補件書填入進度備註
    If Me.Combo3.Text <> "" Then
      'Modified by Morgan 2023/5/18 多文件時以、號區隔
      'Text29.Text = Text29.Text & Me.Combo3.Text
      Text29.Text = Text29.Text & IIf(Text29 <> "", "、", "") & Me.Combo3.Text
      'end 2023/5/18
    End If
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
ReDim cp(1 To TF_CP) As String
End Sub

Private Sub Form_Load()
Dim ret As Long

   MoveFormToCenter Me
   ' 90.07.10 modify by louis (預設頁籤)
   SSTab1.Tab = 0
   ' 90.10.09 modify by louis
   'EnableTextBox Text13, False   '2010/11/15 cancel by sonia
   intWhere = 國內
   With frm04010504_2
      pa(1) = .Text2
      pa(2) = .Text3
      pa(3) = .Text4
      pa(4) = .Text5
      strReceiveNo = .Tag
      ReadPatent
   End With
   Combo1.ListIndex = 0
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010504_2.m_strIR01
   m_strIR02 = frm04010504_2.m_strIR02
   m_strIR03 = frm04010504_2.m_strIR03
   m_strIR04 = frm04010504_2.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
   '若申請國家非台灣時, 鎖住來函期限欄位
    'Remove by Morgan 2008/5/12 不用鎖了
    'If pa(9) <> 台灣國家代號 Then
    '    Me.Frame1.Enabled = False
    '    Me.Frame2.Enabled = False
    'End If
    
   'Text6 = strSrvDate(2)
   ' 90.08.11 modify by sonia
   If pa(9) = 台灣國家代號 Then
      Text6 = Label3(6)
      '游標預設在來函性質欄
      'SendKeys "{Tab}" 'Remove by Morgan 2009/12/25
   Else
      'Modify by Morgan 2008/5/21 非台灣不預設來函日,期限預設 文到次日,月
      'Text6 = strSrvDate(2)
      Option1(1).Value = True
      Option4(1).Value = True
      Text6.MaxLength = 8
      Text12.MaxLength = 8
   End If
       
   'Add by Morgan 2006/3/28 承辦人欄位不可修改
   Text16.Enabled = False
   m_CustX07166 = False '2012/11/26 add by sonia
   
   SSTab1.Tab = 0
   
   'Added by Morgan 2014/9/9 電子化-台灣案定稿要轉pdf故修改只能從定稿維護作業
   If pa(9) = "000" Then
      Text15(1).Enabled = False
   'Added by Morgan 2016/6/14  非臺灣案電子化
   ElseIf (內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F") Then
      Text15(1).Enabled = False
   'end 2016/6/14
   End If
   'end 2014/9/9
End Sub

Private Sub ReadPatent()
 Dim Lbl As Object, i As Integer
 Dim strTmp As String, rsTemp1 As New ADODB.Recordset, bolTmp As Boolean
   ' 90.07.06 modify by louis
   Text8.Clear
    'Add By Cheng 2002/11/29
   Text8.AddItem "其他"
   
   Text8.AddItem "補文件"
   Text8.AddItem "修正"
   Text8.AddItem "補充說明"
   Text8.AddItem "申復"
   Text8.AddItem "領證及繳年費"
   Text8.AddItem "異議答辯"
   Text8.AddItem "舉發答辯"
   Text8.AddItem "變更"
   Text8.AddItem "退費"
   Text8.AddItem "改請發明"
   Text8.AddItem "改請新型"
   Text8.AddItem "改請設計"
   Text8.AddItem "改請追加"
   Text8.AddItem "改請聯合"
   Text8.AddItem "改請獨立"
   Text8.AddItem "分割"
   
   m_bolCCC = False 'Added by Lydia 2015/04/30
   
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   
   'Add By Cheng 2001/12/31
   Me.lblPA57.Caption = ""
   
   Label3(6) = frm04010504_1.Text5
   Label3(5) = strReceiveNo
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   If pa(1) = "P" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         '申請日
         Label3(2) = pa(10)
         '申請案號
         Text1 = pa(11)
         '實體聯絡人(中)
         Text31 = pa(91)
         '是否閉卷
         Text27(0) = pa(57)
         'Add By Sindy 2012/3/5
         If pa(57) = "Y" Then
            m_blnClosed = True
         Else
            m_blnClosed = False
         End If
         '2012/3/5 End
         'Add By Cheng 2001/12/31
         Me.lblPA57.Caption = pa(57)
         '申請國家
         MPa9 = pa(9)
      End If
   Else
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadServicePracticeDatabase(pA(), intWhere) Then
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
         Label3(2) = pa(10)
         Text1 = pa(11)
         Text31 = pa(18)
         Text27(0) = pa(15)
         'Add By Sindy 2012/3/5
         If pa(15) = "Y" Then
            m_blnClosed = True
         Else
            m_blnClosed = False
         End If
         '2012/3/5 End
         'Add By Cheng 2001/12/31
         Me.lblPA57.Caption = pa(15)
         MPa9 = pa(9)
      End If
   End If
   
   If pa(9) = 台灣國家代號 Then
      Text14(0).Enabled = False
      Text14(1).Enabled = False
      strTmp = "CPM03"
   Else
      Text14(0).Enabled = True
      Text14(1).Enabled = True
      strTmp = "CPM04"
   End If
   
   'Add by Morgan 2006/1/24 加NP01
   strExc(0) = "SELECT ''," & strTmp & "," & SQLDate("NP08") & "," & SQLDate("NP09") & ",NP13," & _
      "NP14," & SQLDate("NP11") & ",NP22,NP01 FROM NEXTPROGRESS,CASEPROPERTYMAP " & _
      "WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND (NP06<>'Y' OR NP06 IS NULL) AND NP02=CPM01(+) AND NP07=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   ' 90.10.5 modify by sonia (機關文號設預設內容)
   'Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTemp = Left(strSrvDate(2), 2)
   Else
      strTemp = Left(strSrvDate(2), 3)
   End If
   If pa(9) = 台灣國家代號 Then
      Text9.Text = "（" & strTemp & "）智專一（二）字第號"
        'Add By Cheng 2003/03/26
        '記錄機關文號的預設值
        Me.Text9.Tag = Me.Text9.Text
   End If
   'Add By Cheng 2002/02/19
   '內專若來函性質為"通知退證註銷"(1907), 且申請國家為"台灣"時, 機關文號預設為(年度)智專一一字第
   If pa(9) = 台灣國家代號 And Text7.Text = 通知退證註銷 Then
      Text9.Text = "（" & strTemp & "）智專一（一）字第號"
        'Add By Cheng 2003/03/26
        '記錄機關文號的預設值
        Me.Text9.Tag = Me.Text9.Text
   End If
   cp(9) = strReceiveNo
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.ReadCaseProgressDatabase(cp(), intWhere) Then
   If ClsPDReadCaseProgressDatabase(cp(), intWhere) Then
      If cp(10) <> "" Then
         If pa(9) = 台灣國家代號 Then
            bolTmp = False
         Else
            bolTmp = True
         End If
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(cp(1), cp(10), strExc(0), BolTmp) Then Label3(1) = strExc(0)
         If ClsPDGetCaseProperty(cp(1), cp(10), strExc(0), bolTmp) Then Label3(1) = strExc(0)
      End If
      
      'Add by Morgan 2008/5/13
      Text30 = cp(35)
      Text33 = cp(117)
      'end 2008/5/13
      
      If Left(cp(10), 1) = "1" Then
         strExc(0) = "SELECT CP14 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10=" & 翻譯
         intI = 1
         Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'MODIFY BY SONIA 90.11.27不預設承辦人
         End If
      Else
         'MODIFY BY SONIA 90.11.27不預設承辦人
      End If
      ' 90.06.28 modify by louis 進度備註不須帶出來
      'Text31 = cp(64)
      'Add By Cheng 2002/12/12
      '取得收文日
      m_CP05 = "" & cp(5)
      '取得發文日
      m_CP27 = "" & cp(27)
      '取得收文號
      m_CP09 = "" & cp(9)
      
     'Added by Morgan 2021/1/28 從 Formsave 移來以便共用,順便整合重複的變數
      stCP13 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
      stCP12 = GetSalesArea(stCP13)
      'end 2021/1/28
      
      'Add by Amy 2014/09/17 承辦人期限隱藏
      Label25.Visible = False
      Text17.Enabled = False
      Text17.Visible = False
      'end 2014/09/17
      
      '92.5.8 ADD BY SONIA 承辦人預設輸入人員及不算案件數
      'Modify by Morgan 2009/12/1 FMP 案承辦人另外預設
      'Text16 = strUserNum: ChgType 16
      'Modified by Morgan 2021/1/28
      'If Left(cp(12), 1) = "F" And pa(9) <> "000" Then
      If Left(stCP12, 1) = "F" And pa(9) <> "000" Then
      'end 2021/1/28
         m_bolFMP = True
         'Added by Lydia 2015/06/29
         txtSNP23.Locked = True
         'Text16 = PUB_GetFmpCP14(pa): ChgType 16 'Removed by Morgan 2017/10/11 取消,比照FCP改依來函性質預設
         'Add By Sindy 2024/6/18 請將FMP案（包含寰華案）1508 國知局答辯函 設定承辦期限（收文日+5工作天）
         'If Text7 = "1508" Then
            Text17.Enabled = True
         'End If
         '2024/6/18 EMD
      Else
         m_bolFMP = False
         'Added by Lydia 2015/06/29 外專寰華的案件，在輸入各式審查機關來函的畫面，能帶出約定期限欄位
         lblsNP23.Visible = False: txtSNP23.Visible = False: txtSNP23.Locked = True
         
         Text16 = strUserNum: ChgType 16
         
         'Add by Morgan 2010/10/4
         If Not PUB_IfSetCP48() Then
            Text17.Enabled = False
         End If
         
      End If
      'End 2009/12/1
      'Added by Lydia 2023/05/17 判斷寰華案
      m_bolFMP2 = False
      If m_bolFMP = True Then
         If PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4)) = True Then
            m_bolFMP2 = True
         End If
      End If
      'end 2023/05/17
      Text18 = "N"
      '92.5.8 END
   End If
   
   'Add by Morgan 2007/6/13 檢查65002是否為最後的代理人
   lblDispDate.Visible = False
   txtDispDate.Visible = False
   txtDispDate = ""
   If pa(9) = "000" Then
      If PUB_IsLatestAgent(pa(1), pa(2), pa(3), pa(4)) = True Then
         lblDispDate.Visible = True
         txtDispDate.Visible = True
         txtDispDate.MaxLength = 7
      End If
   End If
   'end 2007/6/13
   
   'Add by Morgan 2009/12/1
   If pa(9) = "000" Then
      Text6.Enabled = False
      cmdDeadLine.Visible = False
   Else
      Text6.Enabled = True
      Label11.Caption = "官方發文日"
      cmdDeadLine.Visible = True
   End If
   
   'Added by Lydia 2025/03/05 點選收文的相關收文號之下一程序：是否已收文、法限、所限、案件性質
   If Left(cp(43), 1) = "C" Then
      'Modified by Morgan 2025/9/2 改用函數以NP24抓收文且不必限制未發文(因台灣案申復期限可能收文申復或修正) Ex:P133079
      'strExc(0) = "select np06, " & IIf(pa(9) = 台灣國家代號, "nvl(cpm03,cpm04)", "nvl(cpm04,cpm03)") & " as nppty,cp43,cp09,cp06,cp07,cppty " & _
                  "from nextprogress,casepropertymap ,(select cp43,cp09,cp06,cp07, " & IIf(pa(9) = 台灣國家代號, "nvl(cpm03,cpm04)", "nvl(cpm04,cpm03)") & " as cppty " & _
                  "from nextprogress,caseprogress,casepropertymap where  np01='" & cp(43) & "' and np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' " & _
                  "and np01=cp43(+) and np07=cp10(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp27 is null) " & _
                  "where  np01='" & cp(43) & "' and np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "' " & _
                  "and np02=cpm01(+) and np07=cpm02(+) and np01=cp43(+) "
      'intI = 1
      'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      'If intI = 1 Then
      '   strCP43toNP06 = "" & RsTemp.Fields("np06")
      '   strCP43toNPpty = "" & RsTemp("nppty")
      '   strCP43toCP09 = "" & RsTemp.Fields("cp09")
      '   strCP43toNp08 = "" & RsTemp.Fields("cp06")
      '   strCP43toNp09 = "" & RsTemp.Fields("cp07")
      '   strCP43toPty = "" & RsTemp("cppty")
      'End If
      Call PUB_GetCPAftExt(pa(9), cp(9), strCP43toCP09, strCP43toNP06, strCP43toNp08, strCP43toNp09, strCP43toNPpty, strCP43toPty)
      'end 2025/9/2
   Else
      strCP43toNP06 = "Y"
   End If
End Sub

Private Function ChgType(i As Integer, Optional SstrKind As Integer) As Boolean
 Dim strTempName As String, bolTmp As Boolean
   ChgType = False
   If pa(9) = 台灣國家代號 Then
      bolTmp = False
   Else
      bolTmp = True
   End If
   Select Case i
      Case 7 '來函性質
         LblNote.Caption = "" 'Added by Morgan 2023/4/19
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(pA(1), Text7.Text, strTempName, BolTmp) Then
         If ClsPDGetCaseProperty(pa(1), Text7.Text, strTempName, bolTmp) Then
            Label3(4) = strTempName
            'Added by Morgan 2023/4/19
            If Text7 = 延期受理 And Pub_StrUserSt03 = "F22" Then
               LblNote.Caption = "寰華案延期受理不更新系統的期限!!!"
            End If
            'end 2023/4/19
            Text14(0) = ""
            Text14(1) = ""
            '2008/4/22 MODIFY BY SONIA
            'Text8 = ""
            'Label3(3) = ""
            If pa(9) = "000" And Text7.Text = "1810" And pa(8) = "2" And Text8 = "申復" Then
            Else
               Text8 = ""
               Label3(3) = ""
            End If
            '2008/4/22 END
            
            If pa(9) = 台灣國家代號 Then
            
               'Modified by Morgan 2014/10/28 從 Text7_LostFocus 移來
               'Added by Morgan 2014/1/14
               '期限
               If m_DeadLine <> "" Then
                  Option1(1).Value = True
                  If Len(m_DeadLine) >= 7 Then
                     Option4(2).Value = True
                     Text10 = ""
                     Text11 = ""
                     Text12 = m_DeadLine
                     Text12_Validate False 'Added by Morgan 2014/10/27
                  'Modified by Morgan 2014/8/18 有日的期限
                  ElseIf Right(m_DeadLine, 1) = "日" Then
                     Option4(0).Value = True
                     Text11 = ""
                     Text12 = ""
                     Text10 = Val(m_DeadLine)
                     Text10_Validate False 'Added by Morgan 2014/10/27
                  ElseIf Right(m_DeadLine, 1) = "月" Then
                     Option4(1).Value = True
                     Text10 = ""
                     Text12 = ""
                     Text11 = Val(m_DeadLine)
                     Text11_Validate False 'Added by Morgan 2014/10/27
                  'end 2014/8/18
                  End If
               Else
      
                  strExc(0) = ""
                  'Modify By Cheng 2003/06/16
                  '若來函性質為延期受理
                  If Me.Text7.Text = "1004" Then
                       '若無相關總收文號
                       If cp(43) = "" Then
                           strExc(0) = "SELECT CF27,CF22,CF25 FROM CASEFEE WHERE CF01='" & pa(1) & "' And CF02='" & pa(9) & "' AND CF03='" & Text7.Text & "'"
                       '若相關總收文號非C類
                       ElseIf Left(cp(43), 1) <> "C" Then
                           strExc(0) = "SELECT CF27,CF22,CF25 FROM CASEFEE WHERE CF01='" & pa(1) & "' And CF02='" & pa(9) & "' AND CF03=(Select CP10 From CaseProgress Where CP09='" & cp(43) & "') "
                       '若相關總收文號為C類
                       Else
                           strExc(0) = "SELECT CF27,CF22,CF25 FROM CASEFEE WHERE CF01='" & pa(1) & "' And CF02='" & pa(9) & "' AND CF03=(Select NP07 From NextProgress Where NP01='" & cp(43) & "' And " & ChgNextProgress(cp(1) & cp(2) & cp(3) & cp(4)) & ") "
                       End If
                  'Add by Morgan 2008/9/17 台灣新型的通知申復期限設1個月(原為30天)
                  ElseIf pa(9) = "000" And Text7.Text = "1221" And pa(8) = "2" Then
                     'Added by Morgan 2014/10/9 改 28 天
                     If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                        Option4(0).Value = True
                        Text10 = 28
                     Else
                     'end 2014/10/9
                        Option4(1).Value = True
                        Text11 = 1
                     End If 'Added by Morgan 2014/10/9
                     'Modify by Morgan 2010/1/15 呼叫共用函數以免漏改
                     'Text14(1) = TransDate(CompDate(1, Text11, TransDate(Label3(6), 2)), 1)
                     ''技術報告的通知申復所限=法限-12天
                     'If cp(10) = "421" Then
                     '   i = -12
                     'Else
                     '   i = -2
                     'End If
                     'Text14(0) = TransDate(CompDate(2, i, TransDate(Text14(1), 2)), 1)
                     'Text14(0).Text = TransDate(PUB_GetWorkDay1(Me.Text14(0).Text, True), 1)
                     GetTime
                     'end 2010/1/15
                  '其他
                  Else
                      'modify by sonia 90/7/19
                      strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & pa(1) & "' AND CPM02='" & Text7.Text & "'"
                  End If
                  If strExc(0) <> "" Then
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     With RsTemp
                        If intI = 1 Then
                           If Not IsNull(.Fields(1)) Then
                              Option4(0).Value = True
                              Text10 = .Fields(1)
                              '2005/7/5 ADD BY SONIA 台灣新型之審查意見通知函為文到次日起30天,發明或設計用原設定
                              'Modified by Morgan 2012/12/27 +最後通知1227
                              If pa(9) = "000" And (Text7.Text = "1202" Or Text7.Text = "1227") And pa(8) = "2" Then
                                 Text10 = 30
                              End If
                              '2005/7/5 END
                              
                              'Removed by Morgan 2014/10/9 2008/9/17移到上面了
                              ''Add by Morgan 2007/2/12 原1202[通知申復]案件性質改為1221 上面程式保留(台灣新型應該不會輸到1202案件性質)
                              'If pa(9) = "000" And Text7.Text = "1221" And pa(8) = "2" Then
                              '   Text10 = 30
                              'End If
                              ''End 2007/2/12
                              'end 2014/10/9
                              
                              'Add by Morgan 2007/3/30 非申請程序也為30天
                              'Modified by Morgan 2012/12/27 +最後通知1227
                              If pa(9) = "000" And (Text7.Text = "1202" Or Text7.Text = "1227") And InStr("101,102,103", cp(10)) = 0 Then
                                 Text10 = 30
                              End If
                              'end 2007/3/30
                              
                              Text14(1) = TransDate(CompDate(2, Text10, TransDate(Label3(6), 2)), 1)
                              '2008/4/22 ADD BY SONIA第三人提起技術報告且下一程序為申復時法限設30日(設在案件性質檔)所限設18日
                              If pa(9) = "000" And Text7.Text = "1810" And pa(8) = "2" And Text8 = "" Then
                                 Text14(0) = "": Text14(1) = ""
                                 Text10 = "": Text11 = "": Text12 = ""
                              End If
                              '2008/4/22 END
                              If Text7 = "1901" Then Text14(1) = ""   'add by sonia 2018/5/3 1901通知退費不必輸法定期限
                              
                           ElseIf Not IsNull(.Fields(2)) Then
                              Option4(1).Value = True
                              Text11 = .Fields(2)
                              Text14(1) = TransDate(CompDate(1, .Fields(2), TransDate(Label3(6), 2)), 1)
                           Else
                              Option4(0).Value = True
                              Text10 = ""
                              Text11 = ""
                           End If
                           If Text14(1) <> "" And Not IsNull(.Fields(0)) Then
                              If .Fields(0) = "1" Then
                                 Option1(0).Value = True
                                 Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
                              Else
                                 Option1(1).Value = True
                              End If
                           'add by sonia 2014/5/1 P-107928 通知補文件
                           ElseIf Not IsNull(.Fields(0)) Then
                              If .Fields(0) = "1" Then
                                 Option1(0).Value = True
                              Else
                                 Option1(1).Value = True
                              End If
                           '2014/5/1 end
                           End If
                           'Modify by Morgan 2005/7/6
                           'If Not IsNull(Text10) Then
                           If Text10 <> "" Then
                              'Modify by Morgan 2008/9/19 改60天以上--玲玲
                              'If Text10 = "60" Or Text10 = "90" Then
                              If Val(Text10) >= 60 Then
                                 i = -4
                              Else
                                 i = -2
                                 
                                 'Add by Morgan 2007/2/12 技術報告的通知申復法限設30日所限設18日
                                 If pa(9) = "000" And Text7.Text = "1221" And pa(8) = "2" And cp(10) = "421" Then
                                    i = -12
                                 End If
                                 'End 2007/2/12
                                 '2008/4/22 ADD BY SONIA 第三人提起技術報告且下一程序為申復時法限設30日(設在案件性質檔)所限設18日
                                 If pa(9) = "000" And Text7.Text = "1810" And pa(8) = "2" Then
                                    i = -12
                                 End If
                                 '2008/4/22 END
                                 
                              End If
                           ElseIf Not IsNull(.Fields(2)) Then
                              'Modify by Morgan 2008/9/19 改2個月以上--玲玲
                              'If .Fields(2) = 2 Then
                              If Val(.Fields(2)) >= 2 Then
                                 i = -4
                              Else
                                 i = -2
                              End If
                           End If
                           If Text14(1) <> "" Then
                              'Added by Lydia 2025/10/29
                              If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                                 Text14(0) = TransDate(PUB_GetPOurDeadline(Text14(1), pa(9)), 1)
                              Else
                              'end 2025/10/29
                                 'Added by Morgan 2014/10/9
                                 If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                                    Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
                                 Else
                                 'end 2014/10/9
                                    Text14(0) = TransDate(CompDate(2, i, TransDate(Text14(1), 2)), 1)
                                 End If 'Added by Morgan 2014/10/9
                              End If 'Added by Lydia 2025/10/29
                           End If
                          'Add By Cheng 2003/12/08
                          '本所期限若非工作天則抓最近工作天
                          Me.Text14(0).Text = TransDate(PUB_GetWorkDay1(Me.Text14(0).Text, True), 1)
                        End If
                     End With
                  End If
               End If 'end 2014/10/28
            End If
            'Add By Cheng 2002/11/21
            If Text14(1) <> "" Then
                '93/4/15取消通知參加訴訟(1505)
                'If pa(9) = 台灣國家代號 And (Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or Me.Text7.Text = "1212" Or Me.Text7.Text = "1401" Or Me.Text7.Text = "1402" Or Me.Text7.Text = "1505" Or Me.Text7.Text = "1506") Then
                'modify by sonia 2018/8/14 +1812通知聽證
                If pa(9) = 台灣國家代號 And (Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or Me.Text7.Text = "1212" Or Me.Text7.Text = "1401" Or Me.Text7.Text = "1402" Or Me.Text7.Text = "1506" Or Me.Text7.Text = "1812") Then
                    Me.Text14(0).Text = Me.Text14(1).Text
'cancel by sonia 2018/8/10 專利處請作單
'                    '2008/10/8 add by sonia P-083487 參加訴訟外之案件性質本所=法定-2天,定稿改通知法定
'                    If Me.Text7.Text <> "1506" Then
'                        'Added by Morgan 2014/10/9
'                        If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
'                           Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
'                        Else
'                        'end 2014/10/9
'                           Text14(0) = TransDate(CompDate(2, -2, TransDate(Text14(1), 2)), 1)
'                        End If 'Added by Morgan 2014/10/9
'                    End If
'                    '2008/10/8 end
'end 2018/8/10
                    'Add By Cheng 2003/12/08
                    '本所期限若非工作天則抓最近工作天
                    Me.Text14(0).Text = TransDate(PUB_GetWorkDay1(Me.Text14(0).Text, True), 1)
                End If
            End If
            '2009/10/5 add by sonia 台灣發明申請,設計申請不預設來函期限
            If Text7.Text = "1201" And pa(9) = 台灣國家代號 And (cp(10) = "101" Or cp(10) = "103" Or cp(10) = "301" Or cp(10) = "303") Then
               Option1(1).Value = True: Option4(2).Value = True
               Text10 = "": Text11 = ""
            End If
            '2009/10/5 end
            
            '承辦期限
            Text17 = ""
            
            'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
'            strExc(0) = "SELECT CF04 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & Text7.Text & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               With RsTemp
'                  If Not IsNull(.Fields(0)) Then
'                     Text17 = TransDate(CompDate(2, Val(.Fields(0)), TransDate(Label3(6).Caption, 2)), 1)
'                     If Text17 > Text14(0) Then Text17 = Text14(0)
'                  End If
'               End With
'            End If
            If Text17.Enabled Then 'Add by Morgan 2010/10/4
               'Modify By Sindy 2024/6/18 pa(1) => IIf(m_bolFMP = True, "FCP", pa(1))
               Text17 = TransDate(Pub_GetHandleDay(IIf(m_bolFMP = True, "FCP", pa(1)), pa(9), Text7.Text, TransDate(Label3(6).Caption, 2), Text14(0)), 1)
            End If
            'end 2007/10/11
            
            '下一程序
            'modify by sonia 90/7/19
               strExc(0) = "SELECT CF15 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & Text7 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With RsTemp
                  If Not IsNull(.Fields(0)) Then Text8 = .Fields(0): ChgType 8, Val(Text8)
               End With
            End If
            
            'Added by Morgan 2014/4/11 電子化-判發人
            If m_PropertyCode <> Text7 Then 'Added by Morgan 2014/8/27
               strExc(1) = Text37 'Added by Morgan 2021/5/27
               If pa(9) = 台灣國家代號 Then
                  'Modified by Morgan 2018/8/1
                  'Text37 = PUB_GetLetterJudge(pa(1), Text7, cp(10), , pa(1), pa(2), pa(3), pa(4))
                  Text37 = PUB_GetLetterJudgeNew("1", pa(1), Text7, , cp(10))
                  'If Text37 <> "" Then Text37_Validate False 'Removed by Morgan 2021/5/27
               'Added by Morgan 2016/6/14  非臺灣案電子化
               ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
                  'Modified by Morgan 2018/8/1
                  'Text37 = PUB_GetLetterJudge(pa(1), Text7, cp(10), pa(9), pa(1), pa(2), pa(3), pa(4))
                  Text37 = PUB_GetLetterJudgeNew("1", pa(1), Text7, pa(9), cp(10), , m_bolFMP)
               'end 2016/6/14
               Else
                  Text37 = ""
                  'Label3(0) = "" 'Removed by Morgan 2021/5/27
               End If
               
               'Added by Morgan 2021/5/27 若有先輸入客戶函判發人此處加提醒 Ex:P-121647
               If Text37 = "" Then
                  Label3(0) = ""
               Else
                  Label3(0) = GetPrjSales(Text37)
               End If
               If strExc(1) <> "" And strExc(1) <> Text37 Then
                  MsgBox "客戶函判發人已重設！", vbInformation
               End If
               'end 2021/5/27
               
               m_PropertyCode = Text7 'Added by Morgan 2014/8/27
            End If 'Added by Morgan 2014/8/27
            'end 2014/4/11
            
            ChgType = True
         Else
            Label3(4) = ""
         End If
      Case 8
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(pA(1), Format(SstrKind), strTempName, BolTmp) Then
         If ClsPDGetCaseProperty(pa(1), Format(SstrKind), strTempName, bolTmp) Then
            Label3(3) = strTempName
            ChgType = True
         Else
            Label3(3) = ""
         End If
         '2008/4/22 ADD BY SONIA第三人提起技術報告且下一程序為申復時法限設30日(設在案件性質檔)所限設18日
         If pa(9) = "000" And Text7.Text = "1810" And pa(8) = "2" And Text8 = "申復" Then
            ChgType (7)
            'text8 = "205"
         End If
         '2008/4/22 END
      Case 16
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetStaff(Text16.Text, strTempName) Then
         If ClsPDGetStaff(Text16.Text, strTempName) Then
            Label3(7) = strTempName
            ChgType = True
         Else
            Label3(7) = ""
         End If
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2018/10/2
  Dim ret As Long
    If prevWndProc <> 0 Then
        ret = SetWindowLong(Text8.hWnd, GWL_WNDPROC, prevWndProc)
        prevWndProc = 0
    End If
   'Set frm04010504_3 = Nothing 'Removed by Morgan 2021/12/20 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub Option1_Click(Index As Integer)
    'Add By Cheng 2002/10/24
    If Me.Option4(0).Value Then
        Text10_Validate False
    ElseIf Me.Option4(1).Value Then
        Text11_Validate False
    ElseIf Me.Option4(2).Value Then
        Text12_Validate False
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   On Error Resume Next
   Select Case Me.SSTab1.Tab
   Case 0 '來函輸入1
      If PreviousTab <> Me.SSTab1.Tab Then
         Text7.SetFocus
      End If
   Case 1 '來函輸入2
      If PreviousTab <> Me.SSTab1.Tab Then
        '94.1.5 modify by sonia
        'If Me.Text7.Text = "1801" Or Me.Text7.Text = "1802" Then
        'Modify by Morgan 2005/6/1 加1810
        If Me.Text7.Text = "1801" Or Me.Text7.Text = "1802" Or Me.Text7.Text = "1405" Or Me.Text7.Text = "1810" Then
        '94.1.5 end
            Me.Text23.SetFocus
        Else
             Text20(3).SetFocus
        End If
      End If
   End Select
End Sub

Private Sub Text10_GotFocus()
    TextInverse Text10
    CloseIme
End Sub

Private Sub Text10_LostFocus()
   'Add by Morgan 2008/5/23 非台灣"天"跳離時到"本所期限"欄位
   If pa(9) <> 台灣國家代號 Then
      If Text14(0).Enabled = True Then Text14(0).SetFocus
   End If
   'Added by Lydia 2015/04/30
   If pa(1) = "P" And pa(9) = 台灣國家代號 And m_DocNo <> "" And m_bolCCC = False Then
      Call msgCCC("" & Text10.Text & Text11.Text & Text12.Text)
   End If
End Sub

Private Sub Text11_GotFocus()
  TextInverse Text11
  CloseIme
End Sub

Private Sub Text11_LostFocus()
   'Add by Morgan 2008/5/23 非台灣"月"跳離時到"本所期限"欄位
   'If pa(9) <> 台灣國家代號 Then
   '   If Text14(0).Enabled = True Then Text14(0).SetFocus
   'End If
   'Added by Lydia 2015/04/30
   If pa(1) = "P" And pa(9) = 台灣國家代號 And m_DocNo <> "" And m_bolCCC = False Then
      Call msgCCC("" & Text10.Text & Text11.Text & Text12.Text)
   End If
End Sub

Private Sub Text12_GotFocus()
 TextInverse Text12
End Sub

Private Sub Text12_LostFocus()
   'Add by Morgan 2008/5/23 非台灣"日"跳離時到"本所期限"欄位
   If pa(9) <> 台灣國家代號 Then
      If Text14(0).Enabled = True Then Text14(0).SetFocus
   End If
   'Added by Lydia 2015/04/30
   If pa(1) = "P" And pa(9) = 台灣國家代號 And m_DocNo <> "" And m_bolCCC = False Then
      Call msgCCC("" & Text10.Text & Text11.Text & Text12.Text)
   End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
   Else
      If ChkDate(Text12) Then
         'Add by Morgan 2008/5/23
         If pa(9) <> 台灣國家代號 Then
            If Val(Text12) < Val(strSrvDate(1)) Then
               MsgBox "來函期限不可小於系統日 !", vbCritical
               Cancel = True
            Else
               '轉民國年
               Text14(1) = TransDate(Text12, 1)
               'Added by Lydia 2025/10/29
               If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                  Text14(0) = TransDate(PUB_GetPOurDeadline(Text14(1), pa(9)), 1)
               Else
               'end 2025/10/29
                  '大陸案的所限=法限-10天
                  '2010/11/17 modify by sonia FMP改7天 P-088376
                  'Text14(0) = TransDate(CompDate(2, -10, TransDate(Text14(1), 2)), 1)
                  If m_bolFMP Then
                     Text14(0) = TransDate(CompDate(2, -7, TransDate(Text14(1), 2)), 1)
                  Else
                     Text14(0) = TransDate(CompDate(2, -10, TransDate(Text14(1), 2)), 1)
                  End If
                  '2010/11/17 END
               End If 'Added by Lydia 2025/10/29
               Text14(0) = TransDate(PUB_GetWorkDay1(Text14(0), True), 1)
            End If
         Else
         'end 2008/5/23
            If Val(Text12) < Val(strSrvDate(2)) Then
               MsgBox "來函期限不可小於系統日 !", vbCritical
               Cancel = True
            Else
               Text14(1) = Text12
               
               'Added by Morgan 2014/10/9
               If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                  Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
               Else
               'end 2014/10/9
            
                  '2008/4/29 MODIFY BY SONIA 第三人提起技術報告且下一程序為申復時法限設30日(設在案件性質檔)所限設18日
                  'Text14(0) = TransDate(CompDate(2, -2, TransDate(Text14(1), 2)), 1)
                  If pa(9) = "000" And Text7.Text = "1810" And pa(8) = "2" Then
                     Text14(0) = TransDate(CompDate(2, -12, TransDate(Text14(1), 2)), 1)
                  Else
                     Text14(0) = TransDate(CompDate(2, -2, TransDate(Text14(1), 2)), 1)
                  End If
                  '2008/4/29 END
                  'Add By Cheng 2003/12/08
                  '本所期限若非工作天則抓最近工作天
                  Me.Text14(0).Text = TransDate(PUB_GetWorkDay1(Me.Text14(0).Text, True), 1)
                  
               End If 'Added by Morgan 2014/10/9
               
               'Add By Cheng 2002/05/30
               If Text14(1) <> "" Then
                  'Modify By Cheng 2003/01/22
                  '取消通知參加訴願(1504)
                  'If pa(9) = 台灣國家代號 And (Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or Me.Text7.Text = "1212" Or Me.Text7.Text = "1401" Or Me.Text7.Text = "1402" Or Me.Text7.Text = "1504" Or Me.Text7.Text = "1505" Or Me.Text7.Text = "1506") Then
                  '2008/10/8取消通知參加訴訟(1505),參考ChgType
                  'If pa(9) = 台灣國家代號 And (Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or Me.Text7.Text = "1212" Or Me.Text7.Text = "1401" Or Me.Text7.Text = "1402" Or Me.Text7.Text = "1505" Or Me.Text7.Text = "1506") Then
                  'modify by sonia 2018/8/14 +1812通知聽證
                  If pa(9) = 台灣國家代號 And (Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or Me.Text7.Text = "1212" Or Me.Text7.Text = "1401" Or Me.Text7.Text = "1402" Or Me.Text7.Text = "1506" Or Me.Text7.Text = "1812") Then
                       Me.Text14(0).Text = Me.Text14(1).Text
'cancel by sonia 2018/8/10 專利處請作單
'                       '2008/10/8 add by sonia P-083487 參加訴訟外之案件性質本所=法定-2天,定稿改通知法定
'                       If Me.Text7.Text <> "1506" Then
'                           'Added by Morgan 2014/10/9
'                           If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
'                              Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
'                           Else
'                           'end 2014/10/9
'                              Text14(0) = TransDate(CompDate(2, -2, TransDate(Text14(1), 2)), 1)
'                           End If 'Added by Morgan 2014/10/9
'                       End If
'                       '2008/10/8 end
'end 2018/8/10
                       'Add By Cheng 2003/12/08
                       '本所期限若非工作天則抓最近工作天
                       Me.Text14(0).Text = TransDate(PUB_GetWorkDay1(Me.Text14(0).Text, True), 1)
                  End If
               End If
            End If
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub GetTime()
   Dim i As Integer
   'Add by Morgan 2007/6/13
   Dim strFromDate As String '期限起算日
   Dim iDays1 As Integer, iDays2 As Integer, iDays3 As Integer, iDays4 As Integer 'Add by Morgan 2009/12/1
   
   If txtDispDate.Visible = True Then
      strFromDate = DBDATE(txtDispDate)
      'Add by Amy 2022/09/30 沒資料在算期限會當掉
      If Trim(txtDispDate) = MsgText(601) Then
        MsgBox "機關發文日不可為空"
        Exit Sub
      End If
   Else
      'Modify by Morgan 2008/5/23 大陸案用核駁日期計算
      If pa(9) <> 台灣國家代號 Then
         strFromDate = DBDATE(Text6)
      Else
         strFromDate = DBDATE(Label3(6))
      End If
   End If
   
   '文到天數
   If Option4(0).Value = True Then
      Text14(1) = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
      If Option1(0).Value = True Then Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
      'Modify by Morgan 2008/9/19 改60天以上--玲玲
      'If Text10 = "60" Or Text10 = "90" Then
      If Val(Text10) >= 60 Then
         i = -4
      Else
         i = -2
         'Add by Morgan 2007/2/12 技術報告的通知申復法限設30日所限設18日
         If pa(9) = "000" And Text7.Text = "1221" And pa(8) = "2" And cp(10) = "421" Then
            i = -12
         End If
         'End 2007/2/12
         '2008/4/29 ADD BY SONIA 第三人提起技術報告且下一程序為申復時法限設30日(設在案件性質檔)所限設18日
         If pa(9) = "000" And Text7.Text = "1810" And pa(8) = "2" Then
            i = -12
         End If
         '2008/4/29 END
      End If
   '文到月數
   ElseIf Option4(1).Value = True Then
      ' 90.12.05 modify by louis (加上月數的方式有變)
      'Text14(1) = TransDate(CompDate(1, Val(Text11), TransDate(Label3(6), 2)), 1)
      Text14(1) = TAIWANDATE(AddMonth(strFromDate, Val(Text11)))
      If Option1(0).Value = True Then Text14(1) = TransDate(CompDate(2, -1, TransDate(Text14(1), 2)), 1)
      'Modify by Morgan 2008/9/19 改2個月天以上--玲玲
      'If Text11 = "2" Then
      
      'Add by Morgan 2010/1/15
      '技術報告的通知申復所限=法限-12天
      '第三人提起技術報告且下一程序為申復時法限設30日(設在案件性質檔)所限設18日
      If (pa(9) = "000" And Text7.Text = "1221" And pa(8) = "2" And cp(10) = "421") Or (pa(9) = "000" And Text7.Text = "1810" And pa(8) = "2") Then
         i = -12
      Else
      'end 2010/1/15
         If Val(Text11) >= 2 Then
            i = -4
         Else
            i = -2
         End If
      End If
   End If
   
   'Modify by Morgan 2008/5/12 大陸案的所限=法限-10天
   'Modify by Morgan 2009/12/1 FMP案所限=法限-7天
   If m_bolFMP Then
      i = -7
   'end 2009/12/1
   ElseIf pa(9) <> 台灣國家代號 Then
      i = -10
   End If
   'end 2008/5/12
   
   If Text14(1) <> "" Then
      'Added by Lydia 2025/10/29
      If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
         Text14(0) = TransDate(PUB_GetPOurDeadline(Text14(1), pa(9)), 1)
      Else
      'end 2025/10/29
         'Added by Morgan 2014/10/9
         If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
            Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
         Else
         'end 2014/10/9
            Text14(0) = TransDate(CompDate(2, i, TransDate(Text14(1), 2)), 1)
         End If 'Added by Morgan 2014/10/9
      End If 'Added by Lydia 2025/10/29
   End If
    'Add By Cheng 2003/12/08
    '本所期限若非工作天則抓最近工作天
    Me.Text14(0).Text = TransDate(PUB_GetWorkDay1(Me.Text14(0).Text, True), 1)
   'Add By Cheng 2002/05/30
   If Text14(1) <> "" Then
        'Modify By Cheng 2003/01/22
        '取消通知參加訴願(1504)
'        If pa(9) = 台灣國家代號 And (Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or Me.Text7.Text = "1212" Or Me.Text7.Text = "1401" Or Me.Text7.Text = "1402" Or Me.Text7.Text = "1504" Or Me.Text7.Text = "1505" Or Me.Text7.Text = "1506") Then
        '2008/10/8 modify by sonia 取消通知參加訴訟(1505), 參考ChgType
        'If pa(9) = 台灣國家代號 And (Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or Me.Text7.Text = "1212" Or Me.Text7.Text = "1401" Or Me.Text7.Text = "1402" Or Me.Text7.Text = "1505" Or Me.Text7.Text = "1506") Then
        'modify by sonia 2018/8/14 +1812通知聽證
        If pa(9) = 台灣國家代號 And (Me.Text7.Text = "1210" Or Me.Text7.Text = "1211" Or Me.Text7.Text = "1212" Or Me.Text7.Text = "1401" Or Me.Text7.Text = "1402" Or Me.Text7.Text = "1506" Or Me.Text7.Text = "1812") Then
            Me.Text14(0).Text = Me.Text14(1).Text
'cancel by sonia 2018/8/10 專利處請作單
'            '2008/10/8 add by sonia P-083487 參加訴訟外之案件性質本所=法定-2天,定稿改通知法定
'            If Me.Text7.Text <> "1506" Then
'               'Added by Morgan 2014/10/9
'               If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
'                  Text14(0) = TransDate(PUB_GetOurDeadline(Text14(1)), 1)
'               Else
'               'end 2014/10/9
'                  Text14(0) = TransDate(CompDate(2, -2, TransDate(Text14(1), 2)), 1)
'               End If 'Added by Morgan 2014/10/9
'            End If
'            '2008/10/8 end
'end 2018/8/10
            'Add By Cheng 2003/12/08
            '本所期限若非工作天則抓最近工作天
            Me.Text14(0).Text = TransDate(PUB_GetWorkDay1(Me.Text14(0).Text, True), 1)
        End If
   End If
   
   'Add by Morgan 2009/12/1
   '承辦期限,承辦期限狀況,約定期限
   If m_bolFMP And Val(Text11) > 0 Then
      strExc(1) = PUB_GetFmpCP48(DBDATE(Label3(6)), DBDATE(Text14(0)), DBDATE(Text14(1)), strFromDate, Text11, stNP23, stCP48Desc)
      Text17 = TransDate(strExc(1), 1)
      'Added by Lydia 2015/06/29 外專寰華的案件，在輸入各式審查機關來函的畫面，能帶出約定期限欄位
      txtSNP23.Text = TransDate(stNP23, 1)
   End If
   
   If Text7 = "1901" Then Text14(1) = ""   'add by sonia 2018/5/3 1901通知退費不必輸法定期限
   SetException 'Added by Morgan 2015/4/20
End Sub

Private Sub Text13_GotFocus()
   CloseIme
   TextInverse Text13
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
   End If
End Sub

'2010/11/15 cancel by sonia
'Private Sub Text13_GotFocus()
'  TextInverse Text13
'End Sub
'
'Private Sub Text13_KeyPress(KeyAscii As Integer)
'   If (KeyAscii > 50 Or KeyAscii < 49) And KeyAscii <> 8 Then
'      KeyAscii = 0
'      Beep
'   End If
'End Sub
'2010/11/15 end
Private Sub Text14_GotFocus(Index As Integer)
    TextInverse Text14(Index)
End Sub

Private Sub Text14_Validate(Index As Integer, Cancel As Boolean)
   If Text14(Index) <> "" Then
      If Not ChkDate(Text14(Index)) Then
         Cancel = True
      Else
         '若非台灣案, 則預設本所期限與法定期限相同
         'Remove by Morgan 2008/5/23 不必在預設相同 --玲玲
         'If pa(9) <> 台灣國家代號 And Text7 <> "1401" Then
         '   If Index = 0 Then
         '      Me.Text14(1).Text = Me.Text14(0).Text
         '   End If
         'End If
         
         If Index = 1 Then
            If Not ChkRange(Text14(0), Text14(1), "本所期限、法定期限") Then
               Cancel = True
            Else
               ' 90.07.10 modify by louis
               If pa(9) < "010" Then
                  'edit by nickc 2007/02/05 不用 dll 了
                  'If objLawDll.ChkMRec(TransDate(Label3(6).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
                  If ClsLawChkMRec(TransDate(Label3(6).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
                     If Text14(0) <> TransDate(strExc(1), 1) Then
                        If MsgBox("與櫃台之來函收文記錄本所期限 ( " & TransDate(strExc(1), 1) & ") 不符，請確認 !", vbCritical + vbDefaultButton2 + vbYesNo) = vbNo Then
                            '停留在目前畫面
                            Cancel = True
                        Else
                            'Modify By Cheng 2002/11/19
                            '按下確定仍可執行
'                           Text14(0) = ""
'                           Text14(1) = ""
                        End If
                     ElseIf Text14(1) <> TransDate(strExc(2), 1) Then
                        If MsgBox("與櫃台之來函收文記錄法定期限 ( " & TransDate(strExc(2), 1) & ") 不符，請確認 !", vbCritical + vbDefaultButton2 + vbYesNo) = vbNo Then
                            '停留在目前畫面
                            Cancel = True
                        Else
                        End If
                     End If
                  'Modified by Morgan 2014/5/5 排除無期限電子公文
                  'Else
                  'modify by sonia 2018/2/8 電子公文都不檢查來函記錄檔
                  'ElseIf m_DocNo = "" Or Text14(1) <> "" Then
                  ElseIf m_DocNo = "" Then
                  'end 2014/5/5
                     If MsgBox("來函記錄檔無此記錄，請確認 !", vbCritical + vbDefaultButton2 + vbYesNo) = vbNo Then Cancel = True
                  End If
               End If
            End If
         End If
      End If
            
      If Cancel = False Then
         '若本所期限非工作天則直接調整至最近的工作天
         If Index = 0 Then
             Me.Text14(0).Text = TransDate(PUB_GetWorkDay1(Me.Text14(0).Text, True), 1)
         End If
         'Modify by Morgan 2008/5/23
         'If Val(Me.Text14(0).Text) + 19110000 < ServerDate Then
         '   MsgBox "本所期限不可小於系統日!!!", vbExclamation
         '   Cancel = True
         'End If
         '2008/11/3 modify by sonia P-83487可等於系統日
         'If Val(TransDate(Text14(0), 2)) <= Val(strSrvDate(1)) Then
         If Val(TransDate(Text14(0), 2)) < Val(strSrvDate(1)) Then
            MsgBox "本所期限不可小於系統日!!!", vbExclamation
            Cancel = True
         End If
      End If
   'Add By Cheng 2002/05/29
   Else
      If Index = 2 And Me.Text14(2).Enabled Then
         If Len(Me.Text14(2).Text) <= 0 Then
            MsgBox "請輸入延緩公告日!!!", vbExclamation + vbOKOnly
            Cancel = True
         End If
      End If
   End If
   If Cancel = True Then TextInverse Text14(Index)
   'Add By cheng 2001/12/12
   If Cancel = False And Index = 1 Then m_bln_FieldValid = True
End Sub

Private Sub Text15_GotFocus(Index As Integer)
  TextInverse Text15(Index)
End Sub

Private Sub Text15_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/05/17
   Select Case Index
   Case 0 '是否列印客戶通知函
      If KeyAscii <> 78 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 1 '是否修改通知函內容
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   End Select
   
End Sub

Private Sub Text15_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If Text15(Index) <> "" And Text15(Index) <> "N" Then
            MsgBox "是否列印客戶通知函，只可為空白或 N !", vbCritical
            Cancel = True
         End If
      Case 1
         If Text15(Index) <> "" And Text15(Index) <> "Y" Then
            MsgBox "是否修改通知函內容，只可為空白或 Y !", vbCritical
            Cancel = True
         End If
   End Select
   If Cancel = True Then TextInverse Text15(Index)
End Sub

Private Sub Text16_GotFocus()
    TextInverse Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii) 'Added by Morgan 2024/4/29
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTmp As String
   
   Cancel = False
   Label3(7) = Empty
   If IsEmptyText(Text16) = False Then
      strTmp = Empty
      strTmp = GetStaffName(Text16)
      Label3(7) = strTmp
      If IsEmptyText(strTmp) Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "請輸入正確的承辦人"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         Text16_GotFocus
      End If
   End If
End Sub

Private Sub Text17_GotFocus()
   TextInverse Text17
End Sub

Private Sub Text17_Validate(Cancel As Boolean)
   If Text17 <> "" Then
      If ChkWorkDay(TransDate(Text17, 2)) Then
         'Modify by Morgan 2007/10/31 有本所期限才比
         'If Text17 > Text14(0) Then
         'Modify by Morgan 2010/8/11 百年蟲
         'If Text14(0) <> "" And Text17 > Text14(0) Then
         If Text14(0) <> "" And Val(Text17) > Val(Text14(0)) Then
            MsgBox "承辦期限不可大於本所期限，請重新輸入 !", vbCritical
            Cancel = True
         End If
      Else
         MsgBox "承辦期限不正確，請重新輸入 !", vbCritical
         Cancel = True
      End If
   Else
   End If
   If Cancel = True Then TextInverse Text17
   'Add By cheng 2001/12/12
   m_bln_FieldValid = True
End Sub

Private Sub Text18_GotFocus()
  TextInverse Text18
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text18_Validate(Cancel As Boolean)
   'Modified by Morgan 2013/10/23 考慮程序新人
   'If Text16.Text = "81002" Or Text16.Text = "73017" Then
   If PUB_GetST05(Text16.Text) = "75" Then
   'end 2013/10/23
      Text18.Text = "N"
   End If
End Sub

Private Sub Text19_GotFocus()
  TextInverse Text19
End Sub

Private Sub Text19_Validate(Cancel As Boolean)
'Add By Cheng 2002/05/30
If pa(9) = 台灣國家代號 And (Me.Text7.Text = "1801" Or Me.Text7.Text = "1802") Then
   If Len(Me.Text19.Text) <= 0 Then
      MsgBox "請輸入對造號數!!!", vbExclamation + vbOKOnly
      Cancel = True
      Me.SSTab1.Tab = 1
      Me.Text19.SetFocus
      Text19_GotFocus
   End If
End If
End Sub

Private Sub Text20_GotFocus(Index As Integer)
  TextInverse Text20(Index)
End Sub

Private Sub Text20_LostFocus(Index As Integer)
   If Text19 <> "" Then
      Select Case Index
         Case 2
            If Text20(0) = "" And Text20(1) = "" And Text20(2) = "" Then
               MsgBox "對造案件名稱不可同時空白 !", vbCritical
               Text20(0).SetFocus
            End If
         Case 5
            If Text20(3) = "" And Text20(4) = "" And Text20(5) = "" Then
               MsgBox "對造名稱不可同時空白 !", vbCritical
               Text20(3).SetFocus
            End If
      End Select
   End If
End Sub

Private Sub Text21_GotFocus(Index As Integer)
  TextInverse Text21(Index)
End Sub

Private Sub Text23_GotFocus()
    'Add By Cheng 2002/11/28
    Dim ii As Integer
    
    'Add By Cheng 2002/11/28
    Select Case Me.Text7.Text
    Case "1801"
        ii = InStr(Me.Text23.Text, "P")
        Me.Text23.SelStart = ii
        Me.Text23.SelLength = 0
    Case "1802"
        ii = InStr(Me.Text23.Text, "N")
        Me.Text23.SelStart = ii
        Me.Text23.SelLength = 0
    'Modify by Morgan 2005/6/1 加1810
    Case "1405", "1810"
        ii = InStr(Me.Text23.Text, "e")
        Me.Text23.SelStart = ii
        Me.Text23.SelLength = 0
    Case Else
        TextInverse Me.Text23
    End Select
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
    '94.1.5 modify by sonia 因為受理技術報告須輸 eXX
    'KeyAscii = UpperCase(KeyAscii)
    Select Case Me.Text7.Text
    Case "1405"
    Case Else
       KeyAscii = UpperCase(KeyAscii)
    End Select
    '94.1.5 end
End Sub

Private Sub Text23_LostFocus()
   'Add By Cheng 2002/11/28
   If Me.Text7.Text = "1801" Or Me.Text7.Text = "1802" Then
      Text20(3).SetFocus
   End If
   'Modify by Morgan 2005/6/1 加1810
   'If Me.Text7.Text = "1405" Then
   If Me.Text7.Text = "1405" Or Me.Text7.Text = "1810" Then
      strExc(0) = "SELECT CP36 FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10 in ('1405','1810') AND CP36='" & Text19.Text & Text23.Text & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         MsgBox "對造案件數代號不可重覆 !", vbCritical
         Text23_GotFocus
         Text23.SetFocus
      'Add by Morgan 2005/6/1
      ElseIf Me.Text7.Text = "1810" Then
         Text20(3).SetFocus
      End If
      
   End If
End Sub

Private Sub Text24_GotFocus()
    'Add By Cheng 2003/04/24
    TextInverse Me.Text24
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/04/16
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 89 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text25_GotFocus()
    'Add By Cheng 2003/04/16
    TextInverse Me.Text25
End Sub

Private Sub Text26_GotFocus()
  TextInverse Text26
End Sub

Private Sub Text26_Validate(Cancel As Boolean)
   If Text26 = "" Then
      If Text7 = 專利權消滅 Then
         MsgBox "來函性質為專利權消滅時，不可空白 !", vbCritical
         Cancel = True
      End If
   Else
      If Not ChkDate(Text26) Then
         MsgBox "日期不正確，請重新輸入 !", vbCritical
         Cancel = True
      End If
   End If
End Sub

Private Sub Text27_GotFocus(Index As Integer)
  TextInverse Text27(Index)
End Sub

Private Sub Text27_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
   If Index = 0 Then
      If Text7 = 專利權消滅 Then
         If KeyAscii <> 89 Then
            MsgBox "來函性質為專利權消滅時，必須為 Y !", vbCritical
            KeyAscii = 89
         End If
      End If
   End If
End Sub

Private Sub Text27_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      '91.12.2 modify by sonia 來函性質為專利權公告作廢時, 直接閉卷不必詢問
      'If Text27(0).Text = "Y" Then
      If Text27(0).Text = "Y" And Text7 <> "1606" Then
      '91.12.2 end
         '2008/11/28 modify by sonia 非1912通知已轉他所才詢問
         'modify by sonia 2016/12/26 +1916解除代理人
         If Text7 <> "1912" And Text7 <> "1916" Then
            If MsgBox("是否確定閉卷 ?", vbYesNo + vbQuestion) = vbNo Then
               Cancel = True
               TextInverse Text27(0)
            End If
         End If
      End If
   End If
End Sub

Private Sub Text28_GotFocus()
   TextInverse Text28
End Sub

Private Sub Text28_Validate(Cancel As Boolean)
   If Text28 = "" Then
      '2008/7/15 CANCEL BY SONIA 不印定稿由工程師處理
      'If pa(9) = "056" And Text7 = 檢索報告 Then
      '   MsgBox "來函性質為檢索報告且申請國為PCT時，不可空白 !", vbCritical
      '   Cancel = True
      'End If
      '2008/7/15 END
   Else
      If Not ChkDate(Text28) Then
         MsgBox "日期不正確，請重新輸入 !", vbCritical
         Cancel = True
      End If
   End If
End Sub

Private Sub Text29_GotFocus()
  TextInverse Text29
End Sub

'2008/11/11 ADD BY SONIA
Private Sub Text29_Validate(Cancel As Boolean)
   If (Text7 = "1210" Or Text7 = "1211") And Text29 = "上下午時分,第法庭" Then
      Cancel = True
      MsgBox "請於進度備註欄輸入開庭時間及法庭 !", vbCritical
      Text29_GotFocus
   End If
End Sub
'2008/11/11 END
Private Sub Text30_GotFocus()
Dim intPos As Integer
   
   '2013/8/16 modify by sonia 行政訴訟之智慧局答辯函, 將游標設定在機關文號欄的"第"的後面
   'TextInverse Text30
   If Me.Text7.Text = "1506" And Label3(1) = "行政訴訟" Then
      With Me.Text30
         If Len("" & .Text) > 0 Then
            intPos = InStr("" & .Text, "第")
            If intPos > 0 Then
               .SelStart = intPos
               .SelLength = 0
            End If
         End If
      End With
   Else
      TextInverse Text30
   End If
   '2013/8/16 end
End Sub

Private Sub Text30_Validate(Cancel As Boolean)
   If Not CheckLengthIsOK(Text30, Text30.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub Text31_GotFocus()
  TextInverse Text31
End Sub

Private Sub Text32_GotFocus()
   TextInverse Me.Text32
End Sub

Private Sub Text33_GotFocus()
   TextInverse Text33
End Sub

Private Sub Text33_Validate(Cancel As Boolean)
   If Not CheckLengthIsOK(Text33, Text33.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub Text34_GotFocus()
   TextInverse Text34
End Sub

Private Sub Text35_GotFocus()
   TextInverse Text35
End Sub

Private Sub Text36_GotFocus()
   TextInverse Text36
End Sub

Private Sub Text37_GotFocus()
   TextInverse Text37
End Sub

Private Sub Text37_Validate(Cancel As Boolean)
   Label3(0) = ""
   If Text37 <> "" Then
      If ClsPDGetStaff(Text37, strExc(1)) = True Then
         Label3(0) = strExc(1)
      Else
         Cancel = True
      End If
   End If
End Sub

Private Sub Text6_GotFocus()
    TextInverse Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = "" Then
      MsgBox Label11 & "不可空白 !", vbCritical
      Cancel = True
   Else
      'Add by Morgan 2008/5/21
      If pa(9) <> 台灣國家代號 Then
         If Len(Text6) <> 8 Then
            MsgBox "非台灣案時" & Label11 & "請輸西元格式！"
            Cancel = True
         ElseIf ChkDate(Text6) Then
            If Val(Text6) > Val(strSrvDate(1)) Then
               MsgBox "核駁函日期不可大於系統日 !", vbCritical
               Cancel = True
            End If
         Else
            Cancel = True
         End If
      'end 2008/5/21
      ElseIf ChkDate(Text6) Then
         If Val(Text6) > Val(strSrvDate(2)) Then
            MsgBox "一般來函日期不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text6
End Sub

Private Sub Text7_Change()
   '若來函性質為(1004)延期受理時, 是否列印客戶通知函之預設為N
   'Modify by Morgan 2008/9/4 +1003通知補文件--玲玲
   '2009/10/5 modify by sonia 加台灣發明申請或設計申請1201通知修正--玲玲
   'Modify by Morgan 2010/7/22 +1225 依職權電話通知修正
   'Modified by Morgan 2014/5/27 +新型的1202也預設N
   'Modified by Morgan 2014/10/14 台灣新型通知修正改預設要出定稿但不列印(自動上已列印), 2015/7/13 改請新型 302 也要
   'Modified by Morgan 2016/1/14 台灣通知修正都改預設出定稿但不列印
   'Modified by Morgan 2016/8/29 +1209,1216 (工程師承辦)
   'Modified by Morgan 2024/12/3 FMP通知審查中1905預設不出定稿(EMail通知承辦)--品薇
   If Text7.Text = "1225" Or Text7.Text = "1004" Or Text7.Text = "1003" Or Text7.Text = "1209" Or Text7.Text = "1216" Or (m_bolFMP And Text7.Text = "1905") Then
      'Added by Lydia 2025/03/05 (台灣案)延期受理函時，串關連的總收文號若尚未收文，增加延期受理定稿
      If Text7.Text = "1004" And pa(9) = 台灣國家代號 And strCP43toNP06 & strCP43toCP09 = "" Then
         Me.Text15(0).Text = ""
      Else
      'end 2025/03/05
         Me.Text15(0).Text = "N"
      End If
   'Add by Morgan 2005/3/10 其他性質要清除
   Else
      Me.Text15(0).Text = ""
   End If
    'Add By Cheng 2003/04/16
    '若申請國家非台灣, 且來函性質為通知補正(1201)
    'Modify by Morgan 2009/7/21 +1202 審查意見通知函
    If pa(9) <> 台灣國家代號 And (Text7.Text = "1201" Or Text7.Text = "1202") Then
'Modify by Morgan 2009/9/3 加欄位改控制
'        Me.Text29.Width = 5720
'        Me.Label18.Visible = True
'        Me.Text24.Visible = True
'        Me.Label19.Visible = True
'        Me.Text25.Visible = True
         frm307.Visible = True
    Else
'        Me.Text29.Width = 8320
'        Me.Label18.Visible = False
        Me.Text24.Text = ""
'        Me.Text24.Visible = False
'        Me.Label19.Visible = False
'        Me.Text25.Visible = False
         frm307.Visible = False
    End If

   'Added by Morgan 2012/10/5
   'Modified by Morgan 2012/12/25
   If pa(9) = 台灣國家代號 And Text7.Text = "1802" Then
      SSTab1.TabVisible(2) = True
      If pa(8) = "3" Then
         chkItem(0).Enabled = False
         chkItem(1).Enabled = False
         chkItem(2).Enabled = True
      Else
         chkItem(2).Enabled = False
      End If
   Else
      SSTab1.TabVisible(2) = False
   End If
   'end 2012/10/5
   
   Text37.Enabled = False
   If pa(9) = 台灣國家代號 Then
      If Text7.Text = "1902" Then Text37.Enabled = True
      
   'Added by Morgan 2016/6/15 非臺灣案電子化
   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      Text37.Enabled = True
   'end 2016/6/15
   
   End If
   
   If Len(Text7) = 4 Then SetException 'Added by Morgan 2015/4/20
   
   'Added by Morgan 2024/7/22
   bolPCTReport = False
   If (Text7 = 檢索報告 Or Text7 = "1216") Then bolPCTReport = True
   'end 2024/7/22
End Sub

Private Sub Text7_GotFocus()
   'Added by Morgan 2014/4/17
   If Text7 = "" And m_NewCP10 <> "" Then
      Text7 = m_NewCP10
   End If
   'end 2014/4/17
   TextInverse Text7
End Sub

Private Sub Text7_LostFocus()
   'Add By Cheng 2002/10/30
   If Me.Text7.Text = "" Then Exit Sub
   'Add By Cheng 2002/05/30
   If ChgType(7) = False Then
      Me.SSTab1.Tab = 0
      Me.Text7.SetFocus
      TextInverse Me.Text7
   End If
   
   Label27 = "審查委員名稱:"  '2013/8/16 add by sonia
   'Add By Cheng 2002/11/28
   '若案件性質為被異議(1801), 被舉發(1802), 對造號數預設為申請案號
   Select Case Me.Text7.Text
   Case "1801"
       If Me.Text19.Text = "" Then Me.Text19.Text = Me.Text1.Text
       If pa(9) = 台灣國家代號 Then 'Add by Morgan 2011/9/8 台灣案才要--玲玲
         If Me.Text23.Text = "" Then Me.Text23.Text = "P"
       End If
   Case "1802"
       If Me.Text19.Text = "" Then Me.Text19.Text = Me.Text1.Text
       If pa(9) = 台灣國家代號 Then 'Add by Morgan 2011/9/8 台灣案才要--玲玲
         If Me.Text23.Text = "" Then Me.Text23.Text = "N"
       End If
   '92.1.28 add by sonia 預設進度備註
   Case "1501"
       If Me.Text19.Text = "" Then Me.Text29.Text = "延展二個月為訴願決定"
   Case "1904"
       If Me.Text19.Text = "" Then Me.Text29.Text = "檢還樣品證據"
   '92.1.28 end
   '94.1.5 ADD BY SONIA
   '若案件性質為受理技術報告(1405), 對造號數預設為申請案號
   'Modify by Morgan 2005/6/1 加1810
   'Case "1405"
   Case "1405", "1810"
       If Me.Text19.Text = "" Then Me.Text19.Text = Me.Text1.Text
       If pa(9) = 台灣國家代號 Then 'Add by Morgan 2011/9/8 台灣案才要--玲玲
         If Me.Text23.Text = "" Then Me.Text23.Text = "e"
       End If
   '94.1.5 END
   Case Else
       '2006/1/3 ADD BY SONIA
       Me.Text23.Text = ""
       '2006/1/3 END
       Me.Text19.Text = "" 'Add by Morgan 2011/6/16
       
       '2013/8/16 add by sonia P-099556行政訴訟之智慧局答辯函輸入(存在cp35以便來函可查詢)
       If Text7 = "1506" And Label3(1) = "行政訴訟" Then
         Label27 = "法院案號:"
         Text30 = Val(Left(DBDATE(m_CP27), 4) - 1911) & "年度行專訴字第號"
       End If
       '2013/8/16 end
       
   End Select
   'Add By Cheng 2002/05/29
   If pa(9) = 台灣國家代號 Then
      Select Case Text7
         Case 專利權消滅
            Text9.Text = "（" & strTemp & "）智專一一權字第號"
         Case 通知領證
            Text9.Text = "（" & strTemp & "）智專一（一）字第號"
           'Modify By Cheng 2003/01/22
   '    'Add By Cheng 2002/10/30
   '      Case 通知智慧局答辯函, 通知行政上訴答辯
         Case 通知參加訴願, 通知參加訴訟, "1508"  '92.9.18 增加 1508 by sonia
            Text9.Text = "經訴字第號"
       'Add By Chen 2002/12/28
         Case 准予延緩公告
            Text9.Text = "（" & strTemp & "）智服字第號"
         '2010/11/12 ADD BY SONIA
         Case "1502"   '撤銷原處分
            Select Case cp(10)
               Case 訴願
                  Text9.Text = "經（" & strTemp & "）字第號"
               Case Else
                  Text9.Text = "（" & strTemp & "）智專一（二）字第號"
            End Select
         '2010/11/12 END
         Case Else
            'Modified by Morgan 2013/7/31--玲玲
            'Text9.Text = "（" & strTemp & "）智專一（二）字第號"
            If pa(8) = "1" Then
               Text9.Text = "（" & strTemp & "）智專一（五）字第號"
            ElseIf pa(8) = "2" Then
               Text9.Text = "（" & strTemp & "）智專一（四）字第號"
            Else
               Text9.Text = "（" & strTemp & "）智專一（三）字第號"
            End If
            'end 2013/7/31
      End Select
      
       'Add By Cheng 2003/03/26
       '記錄機關文號的預設值
       Me.Text9.Tag = Me.Text9.Text
       
      'Added by Morgan 2014/1/14
      'Modified by Morgan 2014/4/17 +發文字
      If m_DocWord <> "" Then
         Text9 = m_DocWord & "字第" & m_DocNo & "號"
      ElseIf m_DocNo <> "" Then
         Text9 = Replace(Text9, "第號", "第" & m_DocNo & "號")
      End If
      'end 2014/1/14
      
   End If
   'Add By Cheng 2002/05/29
   If pa(9) = 大陸國家代號 And Me.Text7.Text = "1401" Then
      Me.Frame4.Visible = True
      Me.Text32.Enabled = True
   Else
      Me.Frame4.Visible = False
      Me.Text32.Enabled = False
   End If
   If Me.Text7.Text = 准予延緩公告 Then
      Me.Frame3.Visible = True
      Me.Text14(2).Enabled = True
   Else
      Me.Frame3.Visible = False
      Me.Text14(2).Enabled = False
   End If
   '2015/1/5 ADD BY SONIA 台灣通知補文件1003之下一程序補文件202鎖住--玲玲
   If pa(9) = 台灣國家代號 And Me.Text7.Text = "1003" Then
      Me.Text8.Enabled = False
   Else
      Me.Text8.Enabled = True
   End If
   '2015/1/5 END
   
   'PCT案來函性質為檢索報告(1209)
   '2008/7/15 CANCEL BY SONIA 不印定稿由工程師處理
   'If pa(9) = "056" And Me.Text7.Text = 檢索報告 Then
   '    '顯示檢索報告種類
   '    Me.Frame5.Visible = True
   'End If
   '2008/7/15 END
   
   'Add by Morgan 2006/8/16
   If m_bolSaveCheck = True Then Exit Sub
   
   'Add By Cheng 2002/12/12
   '設定主管機關選項
   Me.Combo2.Clear
   strExc(0) = "SELECT DISTINCT CF10 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF10 IS NOT NULL"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   Do While Not RsTemp.EOF
      Combo2.AddItem RsTemp.Fields("CF10")
      RsTemp.MoveNext
   Loop
   strExc(0) = "SELECT DISTINCT CF10 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & Me.Text7.Text & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.EOF = False Then Me.Combo2.Text = "" & RsTemp.Fields(0).Value
   Else
      If pa(9) = 台灣國家代號 Then 'Added by Morgan 2022/6/20 應該要限制台灣案
         Me.Combo2.Text = "經濟部智慧財產局"
      End If
   End If
   
   '2010/11/15 add by sonia 台灣撤銷原處分依點選案件性質帶主管機關
   If pa(9) = 台灣國家代號 And Me.Text7.Text = "1502" Then
      strExc(0) = "SELECT DISTINCT CF10 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & cp(10) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.EOF = False Then Me.Combo2.Text = "" & RsTemp.Fields(0).Value
      Else
         Me.Combo2.Text = "經濟部智慧財產局"
      End If
   End If
   '2010/11/15 end
   
   'Add By Cheng 2002/12/23
   '若為台灣案時
   If pa(9) = 台灣國家代號 Then
       'Add By Cheng 2002/12/18
       Me.Combo3.Clear
       Select Case Me.Text7.Text
       Case 通知修正
           Me.Combo3.AddItem "說明書"
           Me.Combo3.AddItem "申請專利範圍"
           Me.Combo3.AddItem "說明書及申請專利範圍"
           Me.Combo3.AddItem "圖式"
           Me.Label17.Visible = True
           Me.Combo3.Visible = True
       Case 通知補文件
           Me.Combo3.AddItem "補申請書"
           Me.Combo3.AddItem "補委任書"
           Me.Combo3.AddItem "補讓與申請書, 契約書"
           Me.Combo3.AddItem "補原文說明書"
           'Modified by Morgan 2023/5/18 改抓案件性質名稱
           'Me.Combo3.AddItem "補中譯本"
           Me.Combo3.AddItem GetCaseTypeName(pa(1), "244", 0)
           'end 2023/5/18
           'Add By Cheng 2002/12/19
           Me.Combo3.AddItem "補宣誓書"
           'Me.Combo3.AddItem "補申請權證明書" 'Removed by Morgan 2018/8/8 不用了--潘韻丞
           'Modified by Morgan 2023/5/18 改抓案件性質名稱
           'Me.Combo3.AddItem "補優先權證明文件" 'Added by Morgan 2018/8/8 --潘韻丞
           Me.Combo3.AddItem GetCaseTypeName(pa(1), "232", 0)
           'end 2023/5/18
           Me.Label17.Visible = True
           Me.Combo3.Visible = True
       Case Else
           Me.Label17.Visible = False
           Me.Combo3.Visible = False
       End Select
   'Add by Morgan 2006/8/18 大陸
   ElseIf pa(9) = "020" Then
      If Me.Text7.Text = "1902" Then
         Me.Label17.Caption = "來函文書:"
         Me.Label17.Visible = True
         Me.Combo3.Visible = True
         Me.Combo3.Clear
         Me.Combo3.AddItem "審查員依職權修改通知", 0
      Else
         Me.Label17.Visible = False
         Me.Combo3.Visible = False
         Me.Combo3.Clear
      End If
   End If

   'Add by Morgan 2008/5/23 非台灣"案件性質"跳離時到"月"欄位
   If pa(9) <> 台灣國家代號 Then
      If Option4(0).Value = True Then
         If Text10.Enabled = True Then Text10.SetFocus
      ElseIf Option4(1).Value = True Then
         If Text11.Enabled = True Then Text11.SetFocus
      ElseIf Option4(2).Value = True Then
         If Text12.Enabled = True Then Text12.SetFocus
      End If
   End If

   '2008/11/12 ADD BY SONIA
   If (Text7 = "1210" Or Text7 = "1211") And Text29 = "" Then
      Text29 = "上下午時分,第法庭"
   ElseIf Text7 <> "1210" And Text7 <> "1211" And Text29 = "上下午時分,第法庭" Then
      Text29 = ""
   End If
   '2008/11/12 END
   
   Text16.Enabled = False 'Added by Morgan 2024/4/25
   m_bolNoCP27 = False
   'Added by Lydia 2016/09/21  P大陸案電話通知修正 : P案→1225依職權電話通知修正的承辦人掛品薇 ; FMP案→掛最近一道程序的工程師(畫面預設)
   If Not m_bolFMP And pa(1) = "P" And pa(9) = "020" And Text7 = "1225" Then
      'Modified by Morgan 2018/10/24 承辦人改輸入人員--玲玲
      'Text16 = "98012"
      'Modified by Morgan 2024/4/25 改比照代理人通知修正規則(相關程序為工程師時直接帶入，否則抓台灣案工程師) --品薇
      'Text16 = strUserNum
      If GetStaffDepartment(cp(14)) <> "P12" Then
         Text16 = cp(14)
         'add by sonia 2024/7/15 A7010柯昱安調離也要改為李柏翰99050
         If GetStaffDepartment(Text16) >= "P10" And GetStaffDepartment(Text16) <= "P11" Then
         Else
            Text16 = "99050"
         End If
         'end 2024/7/15
      Else
         Text16 = PUB_GetInCaseCP14(cp(1), cp(2), cp(3), cp(4))
         'add by sonia 2024/7/15 A7010柯昱安調離也要改為李柏翰99050
         If GetStaffDepartment(Text16) >= "P10" And GetStaffDepartment(Text16) <= "P11" Then
         Else
            Text16 = "99050"
         End If
         'end 2024/7/15
         If Text16 = "" Then
            Text16 = strUserNum
         End If
      End If
      Text16.Enabled = True
      m_bolNoCP27 = True
      'end 2024/4/25
      'end 2018/10/24
      ChgType 16
      
   'Added by Morgan 2024/7/22 從 FormSave 移來，改統一抓畫面上承辦人以避免漏改程式 Ex:A7010改部門
   ElseIf Not m_bolFMP And bolPCTReport Then
      'Modified by Morgan 2025/3/17
      '改判斷若承辦人為程序時抓國內案工程師 Ex:P-132531--品薇
      'If GetStaffName(cp(14)) = "" Then
      '   Text16 = "99050"
      'Else
      '   strExc(1) = GetStaffDepartment(cp(14))
      '   If strExc(1) >= "P10" And strExc(1) <= "P11" Then
      '      Text16 = cp(14)
      '   Else
      '      Text16 = "99050"
      '   End If
      'End If
      strExc(1) = cp(14)
      If PUB_GetST03(cp(14)) = "P12" Then
         strExc(1) = PUB_GetInCaseCP14(cp(1), cp(2), cp(3), cp(4))
      End If
      If GetStaffName(strExc(1)) = "" Then
         Text16 = "99050"
      Else
         Text16 = strExc(1)
      End If
      'end 2025/3/17
      ChgType 16
   'end 2024/7/22
   End If
   'end 2016/09/21
   
   'Added by Morgan 2024/4/29
   'P案1201,1202,1002(核駁) 預設承辦人可以讓程序人員改成工程師，當承辦人為工程師時，就不上發文日，不出定稿，不控制金額 --郭
   If Not m_bolFMP And pa(1) = "P" And InStr("1201,1202", Text7) > 0 Then
      Text16.Enabled = True
   End If
   'end 2024/4/29
   
   'Added by Morgan 2017/10/11 FMP預設承辦人比照FCP
   If m_bolFMP Then
      'Modified by Morgan 2022/12/13
      '已閉卷且非寰華案時承辦人預設為輸入人員 Ex:P-125348 --品薇,淑華,敏莉
      If pa(57) = "Y" And Left(Pub_StrUserSt03, 1) <> "F" Then
         Text16 = strUserNum
      Else
      'end 2022/12/13
         Text16 = PUB_GetFCPPromoterNo(cp(9), Me.Text7.Text, cp(14)) '函數不會回傳F編號,不必再剔除
      End If
      ChgType 16
   End If
   'end 2017/10/11
   
   'Added by Morgan 2021/3/12
   '寶齡富錦 Y55435 案件下列來函承辦人預設韻如
   '1202審查意見來函、1002核駁、1006最終核駁、1201通知修正、1209檢索報告、1205通知提供前案、1206通知要求選取、1203通知補充說明
   m_bolBPFCase = False 'Added by Morgan 2023/6/27
   If pa(75) = "Y55435" And Text7 <> "" Then
      If InStr("1202,1002,1006,1201,1209,1205,1206,1203", Text7) > 0 Then
         Text15(0) = "N"
         'Modified by Morgan 2023/6/27 預設最新收文的工程師--郭
         'Text16 = "A0029"
         If PUB_GetLastEng(cp(1), cp(2), cp(3), cp(4), strExc(1)) Then
            Text16 = strExc(1)
         Else
            Text16 = "A0029"
         End If
         m_bolBPFCase = True
         'end 2023/6/27
         ChgType 16
      End If
   End If
   'end 2021/3/12
   
   'Added by Morgan 2021/9/22
   '檢查案件是否為顧服組W2001的4家客戶X69365、X82504、X82708、X83239及其關係企業也掛在顧服組W2001者(只考慮第一申請人即可)
   '來函性質為1202審查意見通知函、1227最後通知、1221通知申復、1810第三人提起技術報告、1802被舉發（理由）、1807對方補充說明時：
   '1. 來函承辦人預設該案最後之工程師(不限收文種類)；
   '2. 來函不上發文日
   m_bolW2001XCase = False
   m_CustX69365 = False 'Added by Morgan 2021/10/6
   'Modified by Morgan 2021/10/6 來函性質改用常數判斷
   'If InStr("1202,1227,1221,1810,1802,1807", Text7) > 0 Then
   If InStr(PatentOAPtyList, Text7) > 0 Then
   'end 2021/10/6
      'Modified by Morgan 2021/10/6
      If PUB_ChkIsW2001XCase(pa(1), pa(2), pa(3), pa(4), strExc(1), m_CustX69365) = True Then
         m_bolW2001XCase = True
         '長庚醫院要收[轉公文]簡單報告
         'Modified by Morgan 2022/3/28 取消轉公文,改同其他3家直接報告,但本所期限改為 +14天-3個工作天 --黃教威
         'If m_CustX69365 = True Then
         '   Text15(0) = ""
         'Else
         '   Text15(0) = "N"
         'End If
         Text15(0) = "N"
         'end 2022/3/28
         Text16 = strExc(1)
         ChgType 16
      End If
   End If
   'end 2021/9/22
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
Dim strTempName As String
Dim bolTmp As Boolean

   If Text7 = "" Then
      MsgBox "來函性質不可空白 !", vbCritical
      Cancel = True
   Else
      If Len(Text7) <> 4 Then
         MsgBox "來函性質錯誤，請重新輸入 !", vbCritical
         Cancel = True
      Else
      'Modified by Morgan 2020/1/17 +1803爭議受理,1804對方延期 Ex:P-123476 出錯定稿
      'If Text7 = 核准 Or Text7 = 核駁 Then
      '   MsgBox "來函性質不可為核准或核駁或改變原處分!", vbCritical
      If Text7 = 核准 Or Text7 = 核駁 Or Text7 = 改變原處分 Or Text7 = 爭議受理 Or Text7 = 對方延期 Then
         If ClsPDGetCaseProperty(pa(1), Text7.Text, strTempName, IIf(pa(9) = "000", False, True)) = True Then
            Label3(4) = strTempName
            MsgBox "來函性質不可為" & strTempName & "!", vbCritical
         End If
      'end 2020/11/7
         Cancel = True
      'Added by Morgan 2022/8/24
      'Modified by Morgan 2025/3/25 +804舉發答辯且不再限制台灣--韻丞
      'ElseIf (Text7 = "1009") And (pa(9) = "000" And cp(10) = "803") Then
      ElseIf Text7 = "1009" And (cp(10) = "803" Or cp(10) = "804") Then
         If ClsPDGetCaseProperty(pa(1), Text7.Text, strTempName, IIf(pa(9) = "000", False, True)) = True Then
            Label3(4) = strTempName
         End If
         'MsgBox "台灣舉發的" & Label3(4) & "請至核駁函輸入！", vbCritical
         MsgBox "[" & Label3(1) & "] 的 [" & Label3(4) & "] 請至 [核駁函輸入]！", vbCritical
         Cancel = True
      'Add by Moragn 2007/8/31
      ElseIf Text7 = "1221" And cp(10) = "807" Then
         MsgBox "第三人申請技術報告不可輸通知申復的來函！", vbCritical
         Cancel = True
      '2010/10/11 add by sonia 審查意見通知函鎖定點選案件性質
      'Modified by Morgan 2012/12/19 +衍生設計125,改請衍生設計308
      'Modified by Morgan 2012/12/27 +最後通知1227
      'Modified by Morgan 2015/9/7 +專利權延長415 P-111633 --玲玲
      'Modified by Morgan 2018/1/22 +更正402 P-117005 --玲玲
      ElseIf (Text7 = "1202" Or Text7 = "1227") And InStr("101,102,103,104,105,107,125,301,302,303,304,305,306,307,308,402,415", cp(10)) = 0 Then
         MsgBox "點選的案件性質不可輸入" & Label3(4) & "！", vbCritical
         Cancel = True
      '2010/10/11 end
      'add by sonia 2014/6/25
      ElseIf Text7 = "1221" And cp(10) <> "421" Then
         MsgBox "案件性質為 申請技術報告, 才可輸入通知申復函 ！", vbCritical
         Cancel = True
      '2010/10/11 end
      'Added by Lydia 2017/05/09 新增C類官方來函性質「視為未主張」(代號：1918)，可用在主張國內(121)、國際優先權(106)及優惠期(123)
      ElseIf Text7 = "1918" Then
            If InStr("106,121,123", cp(10)) = 0 And cp(10) <> "" Then
               MsgBox "視為未主張，只可用在主張國內優先權、國際優先權及優惠期"
               Cancel = True
               Exit Sub
            End If
      'end 2107/04/05
      'Added by Lydia 2017/09/29 增加來函性質「1233初審報告英譯文」,僅限於P之申請國家為056ＰＣＴ的案件
      ElseIf Text7 = "1233" Then
           If pa(1) <> "P" Or pa(9) <> "056" Then
               MsgBox "初審報告英譯文，僅限於P之申請國家為056ＰＣＴ的案件!"
               Cancel = True
               Exit Sub
           End If
      'end 2017/09/29
      Else
         'Add By Cheng 2002/06/03
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(pA(1), Text7.Text, strTempName, BolTmp) = False Then
         If ClsPDGetCaseProperty(pa(1), Text7.Text, strTempName, bolTmp) = False Then
            Label3(4).Caption = ""
            Cancel = True
         End If
      End If
      End If
   End If
   If Cancel = False Then
'2010/11/15 cancel by sonia
'      If Text7 = "1502" Then
'         EnableTextBox Text13, True
'      Else
'         EnableTextBox Text13, False
'      End If
'2010/11/15 end
      '91.12.2 add by sonia
      If Text7 = "1606" Then
         Text27(0) = "Y"
      End If
      '91.12.2 end
   End If
   If Cancel = True Then TextInverse Text7
   
   'Add By Cheng 2001/12/12
   If Cancel = False Then m_bln_FieldValid = True
   
End Sub

Private Sub Text8_Change()
   If Len(Text8) = 3 Then SetException 'Added by Morgan 2015/4/20
End Sub

Private Sub Text8_Click()
   Dim nValue As Integer
   If Text8.ListIndex >= 0 Then
      Select Case Text8.List(Text8.ListIndex)
        'Add By Cheng 2002/11/29
         Case "其他"
            nValue = 其他
         
         Case "補文件"
            nValue = 補文件
         Case "修正"
            nValue = 修正
         Case "補充說明"
            nValue = 補充說明
         Case "申復"
            nValue = 申復
         Case "領證及繳年費"
            nValue = 領證及繳年費
         Case "異議答辯"
            nValue = 異議答辯
         Case "舉發答辯"
            nValue = 舉發答辯
         Case "變更"
            nValue = 變更
         Case "退費"
            nValue = 退費
         Case "改請發明"
            nValue = 改請發明
         Case "改請新型"
            nValue = 改請新型
         Case "改請設計"
            nValue = 改請設計
         Case "改請追加"
            nValue = 改請追加
         Case "改請聯合"
            nValue = 改請聯合
         Case "改請獨立"
            nValue = 改請獨立
         Case "分割"
            nValue = 分割
      End Select
      ChgType 8, nValue
   End If
End Sub

Private Sub Text8_Validate(Cancel As Boolean)
   If Text8 <> "" Then
            
      Select Case Text8.Text
        'Add By Cheng 2002/11/29
         Case "其他"
            Text8.Text = 其他
         
         Case "補文件"
            Text8.Text = 補文件
         Case "修正"
            Text8.Text = 修正
         Case "補充說明"
            Text8.Text = 補充說明
         Case "申復"
            Text8.Text = 申復
         Case "領證及繳年費"
            Text8.Text = 領證及繳年費
         Case "異議答辯"
            Text8.Text = 異議答辯
         Case "舉發答辯"
            Text8.Text = 舉發答辯
         Case "變更"
            Text8.Text = 變更
         Case "退費"
            Text8.Text = 退費
         Case "改請發明"
            Text8.Text = 改請發明
         Case "改請新型"
            Text8.Text = 改請新型
         Case "改請設計"
            Text8.Text = 改請設計
         Case "改請追加"
            Text8.Text = 改請追加
         Case "改請聯合"
            Text8.Text = 改請聯合
         Case "改請獨立"
            Text8.Text = 改請獨立
         Case "分割"
            Text8.Text = 分割
      End Select
      
      '2008/4/23 ADD BY SONIA第三人提起技術報告且有下一程序時必須為申復
      If pa(9) = "000" And Text7.Text = "1810" And pa(8) = "2" And Text8 <> "205" Then
         MsgBox "第三人提起技術報告之下一程序時必須為申復!!!", vbExclamation + vbOKOnly
         Cancel = True
         Text8.SelStart = 0
         Text8.SelLength = Len(Text8.Text)
      End If
      '2008/4/23 END
      
      'Add By Cheng 2002/01/04
      If Len(Me.Text8.Text) <> 3 Then
         'Add By Cheng 2002/11/29
         MsgBox "下一程序代碼錯誤!!!", vbExclamation + vbOKOnly
         Cancel = True
         Text8.SelStart = 0
         Text8.SelLength = Len(Text8.Text)
         Exit Sub
      End If
      
      If Not ChgType(8, Val(Text8)) Then
         Cancel = True
         'TextInverse Text8
         Text8.SelStart = 0
         Text8.SelLength = Len(Text8.Text)
      End If
   '2008/4/28 ADD BY SONIA
   
   'Modified by Morgan 2014/10/28 非電子公文才要清除
   'Else
   ElseIf m_DeadLine = "" Then
   'end 2014/10/28
   
      Text14(0) = "": Text14(1) = ""
      Text10 = "": Text11 = "": Text12 = ""
   '2008/4/28 END
      Option4(0).Value = True 'Added by Morgan 2014/9/9
   End If
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1500: .Text = "下一程序"
      .CellAlignment = flexAlignCenterCenter
        'Modify By Cheng 2002/11/20
'      .Col = 2: .ColWidth(2) = 1500: .Text = "本所期限"
      .col = 2: .ColWidth(2) = 1000: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
        'Modify By Cheng 2002/11/20
'      .Col = 3: .ColWidth(3) = 1500: .Text = "法定期限"
      .col = 3: .ColWidth(3) = 1000: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1500: .Text = "機關文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1500: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1500: .Text = "解除期限日期"
      .col = 7: .ColWidth(7) = 0
      'Add by Morgan 2006/1/24
      .col = 8: .ColWidth(8) = 0 '總收文號
      .Visible = True
   End With
End Sub

Private Sub MSHFlexGrid1_Click()
   'Removed by Morgan 2024/5/13 取消點選功能避免誤沖期限(原FMP通知補文件的期限已改為更新不新增)--玲玲、敏莉
   If Text7 = "1004" Then 'Added by Morgan 2024/5/15 延期受理除外
      GridClick MSHFlexGrid1, intLastRow, 0, 1
   End If
End Sub

Private Sub Text9_GotFocus()
'  TextInverse Text9
Dim intPos As Integer
   'Modify By Cheng 2002/04/22
   '當來函性質為"1601"或"1604"時, 將游標設定在機關文號欄的"第"的後面, 其餘則放在"專"的後面
   With Me.Text9
      If Len("" & .Text) > 0 Then
         '92.9.18 modify by sonia
         'intPos = InStr("" & .Text, IIf(Me.Text7.Text = "1601" Or Me.Text7.Text = "1604" Or Me.Text7.Text = "1506" Or Me.Text7.Text = "1507", "第", IIf(Me.Text7.Text = 准予延緩公告, "服", "專")))
         intPos = InStr("" & .Text, IIf(Me.Text7.Text = "1601" Or Me.Text7.Text = "1604" Or Me.Text7.Text = "1506" Or Me.Text7.Text = "1507" Or Me.Text7.Text = "1508", "第", IIf(Me.Text7.Text = 准予延緩公告, "服", "專")))
         '92.9.18 end
         If intPos > 0 Then
            .SelStart = intPos
            .SelLength = 0
         End If
      End If
   End With
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   'Modify by Morgan 2011/1/3 機關文號欄位改長度(百年問題)改抓MaxLength屬性控制
   If CheckLengthIsOK(Text9, Text9.MaxLength) = False Then
      Cancel = True
   End If
   If pa(9) = 台灣國家代號 Then
      If Text9.Text = "" Then
         MsgBox "申請國家為台灣時不得空白，請重新輸入 !", vbCritical
         Cancel = True
      End If
   End If
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim bPaper As Boolean
Dim arrCaseNo() As String 'Added by Morgan 2021/2/25

TxtValidate = False

   'Added by Morgan 2021/12/20 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/20
   
'Added by Morgan 2016/12/22 P105805 核駁報價費用少輸1個0
If 1000 * Val(Text21(1)) > Val(Text21(0)) Then
   MsgBox "費用輸入錯誤(不可少於點數)！", vbExclamation
   Text21(0).SetFocus
   Exit Function
End If
If 1000 * Val(Text34) > Val(Text25) Then
   MsgBox "費用輸入錯誤(不可少於點數)！", vbExclamation
   Text25.SetFocus
   Exit Function
End If
If 1000 * Val(Text36) > Val(Text35) Then
   MsgBox "費用輸入錯誤(不可少於點數)！", vbExclamation
   Text35.SetFocus
   Exit Function
End If
'end 2016/12/22

'2008/4/23 ADD BY SONIA
If pa(9) = "000" And Text7.Text = "1810" And pa(8) = "2" And Text8 = "" Then
   If MsgBox("第三人提起技術報告是否掛下一程序期限？", vbYesNo + vbDefaultButton1) = vbYes Then
      'Added by Morgan 2014/10/9 若確定自動帶申復--陳玲玲
      Text8 = "申復"
      ChgType 8, 申復
      Text8_Validate False
      'end 2014/10/9
      Exit Function
   End If
End If
'2008/4/23 END

'Remove by Morgan 2008/5/23 存檔時不必再檢查否則期限會重算
'If Me.Text10.Enabled = True Then
'   Cancel = False
'   Text10_Validate Cancel
'   If Cancel = True Then
'      Me.Text10.SetFocus
'      Text10_GotFocus
'      Exit Function
'   End If
'End If

'If Me.Text11.Enabled = True Then
'   Cancel = False
'   Text11_Validate Cancel
'   If Cancel = True Then
'      Me.Text11.SetFocus
'      Text11_GotFocus
'      Exit Function
'   End If
'End If

'If Me.Text12.Enabled = True Then
'   Cancel = False
'   Text12_Validate Cancel
'   If Cancel = True Then
'      Me.Text12.SetFocus
'      Text12_GotFocus
'      Exit Function
'   End If
'End If
'end 2008/5/23

For Each objTxt In Text14
   If objTxt.Enabled = True Then
      Cancel = False
      Text14_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.Text14(objTxt.Index).SetFocus
         Text14_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next

For Each objTxt In Text15
   If objTxt.Enabled = True Then
      Cancel = False
      Text15_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.Text15(objTxt.Index).SetFocus
         Text15_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next

If Me.Text16.Enabled = True Then
   Cancel = False
   Text16_Validate Cancel
   If Cancel = True Then
      Me.Text16.SetFocus
      Text16_GotFocus
      Exit Function
   End If
End If

If Me.Text17.Enabled = True Then
   Cancel = False
   Text17_Validate Cancel
   If Cancel = True Then
      Me.Text17.SetFocus
      Text17_GotFocus
      Exit Function
   End If
End If

If Me.Text18.Enabled = True Then
   Cancel = False
   Text18_Validate Cancel
   If Cancel = True Then
      Me.Text18.SetFocus
      Text18_GotFocus
      Exit Function
   End If
End If

If Me.Text19.Enabled = True Then
   Cancel = False
   Text19_Validate Cancel
   If Cancel = True Then
      Me.Text19.SetFocus
      Text19_GotFocus
      Exit Function
   End If
End If
'91.10.27 cancel by sonia
'If Me.Text22.Enabled = True Then
'   Cancel = False
'   Text22_Validate Cancel
'   If Cancel = True Then
'      Me.Text22.SetFocus
'      Text22_GotFocus
'      Exit Function
'   End If
'End If
'91.10.27 end
If Me.Text26.Enabled = True Then
   Cancel = False
   Text26_Validate Cancel
   If Cancel = True Then
      Me.Text26.SetFocus
      Text26_GotFocus
      Exit Function
   End If
End If

For Each objTxt In Text27
   If objTxt.Enabled = True Then
      Cancel = False
      Text27_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.Text27(objTxt.Index).SetFocus
         Text27_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next

If Me.Text6.Enabled = True Then
   Cancel = False
   Text6_Validate Cancel
   If Cancel = True Then
      Me.Text6.SetFocus
      Text6_GotFocus
      Exit Function
   End If
End If

If Me.Text7.Enabled = True Then
   Cancel = False
   Text7_Validate Cancel
   If Cancel = True Then
      Me.Text7.SetFocus
      Text7_GotFocus
      Exit Function
   End If
End If

If Me.Text9.Enabled = True Then
   Cancel = False
   Text9_Validate Cancel
   If Cancel = True Then
      Me.Text9.SetFocus
      Text9_GotFocus
      Exit Function
   End If
End If
'2008/11/12 ADD BY SONIA
If Me.Text29.Enabled = True Then
   Cancel = False
   Text29_Validate Cancel
   If Cancel = True Then
      Me.Text29.SetFocus
      Text29_GotFocus
      Exit Function
   End If
End If
'2008/11/12 END
   
   'Add by Morgan 2009/12/1
   If cmdDeadLine.Visible And Text8 = "202" And m_si880017 = 0 Then
      MsgBox "有補件期限時必須點選【補件資料】輸入補件資料並按確定！"
      Exit Function
   End If
   
'Add by Morgan 2011/6/21
'大陸審查意見通知收文第二次以上時控制期限非 2 個月時要提醒並做確認
If pa(9) = "020" And Text7 = "1202" Then
   If CompDate(1, 2, Text6) <> DBDATE(Text14(1)) Then
      If PUB_ChkCPExist(pa, "1202") = True Then
          If MsgBox("本案為第二次以上輸入審查意見通知函，法定期限應為來函日期加 2 個月。目前期限資料不符，是否仍要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
            Text14(1).SetFocus
            Exit Function
          End If
      End If
   End If
End If
'end 2011/6/21
   
'Added by Morgan 2012/10/5
'102新法被舉發要勾撤銷被請求項目
'Modified by Morgan 2013/1/14 增加舉發事項
If SSTab1.TabVisible(2) = True And Val(strSrvDate(1)) > 20130000 Then
   Cancel = True
   For Each oChk In chkItem
      If oChk.Value = vbChecked Then
         If oChk.Index = 0 Then
            If txtItemCount = "" Then
               SSTab1.Tab = 2
               MsgBox "請輸入項數", vbExclamation, "舉發聲明"
               If txtItemCount.Enabled Then txtItemCount.SetFocus
               Exit Function
            End If
         ElseIf oChk.Index = 1 Then
            If txtItemList = "第項" Then
               SSTab1.Tab = 2
               MsgBox "請輸入項次", vbExclamation, "舉發聲明"
               If txtItemList.Enabled Then txtItemList.SetFocus
               Exit Function
            ElseIf PUB_ChkItemList(txtItemList) = False Then
               SSTab1.Tab = 2
               MsgBox "撤銷部分請求項格式錯誤！", vbExclamation, "舉發聲明"
               If txtItemList.Enabled Then txtItemList.SetFocus
               Exit Function
            End If
         ElseIf oChk.Index = 6 Then
            For intI = 0 To 1
               If txtYear(intI) = "" Then
                  SSTab1.Tab = 2
                  MsgBox "請輸入年度!", vbExclamation, "舉發聲明"
                  txtYear(intI).SetFocus
                  Exit Function
               End If
               If txtMonth(intI) = "" Then
                  SSTab1.Tab = 2
                  MsgBox "請輸入月份!", vbExclamation, "舉發聲明"
                  txtMonth(intI).SetFocus
                  Exit Function
               End If
               If txtDay(intI) = "" Then
                  SSTab1.Tab = 2
                  MsgBox "請輸入日期!", vbExclamation, "舉發聲明"
                  txtDay(intI).SetFocus
                  Exit Function
               End If
               If Not IsDate((Val(txtYear(intI)) + 1911) & "/" & txtMonth(intI) & "/" & txtDay(intI)) Then
                  SSTab1.Tab = 2
                  MsgBox "日期錯誤，請重新輸入！", vbExclamation, "舉發聲明"
                  txtYear(intI).SetFocus
                  Exit Function
               End If
            Next
            If CDate((Val(txtYear(0)) + 1911) & "/" & txtMonth(0) & "/" & txtDay(0)) > CDate((Val(txtYear(1)) + 1911) & "/" & txtMonth(1) & "/" & txtDay(1)) Then
               SSTab1.Tab = 2
               MsgBox "起日不可晚於迄日，請重新輸入！", vbExclamation, "舉發聲明"
               txtYear(0).SetFocus
               Exit Function
            End If
         End If
         Cancel = False
         Exit For
      End If
   Next
   If Cancel = True Then
      SSTab1.Tab = 2
      MsgBox "請選擇撤銷被請求項目！", vbExclamation, "舉發聲明"
      Exit Function
   End If
End If
'end 2012/10/5

   'Added by Morgan 2019/5/27 從寫定稿例外欄位移來此處先作檢查
   'IDS報價檢查
   m_USCaseNo = ""
   'Modified by Morgan 2019/6/3 通知擇一(1232)不用--郭
   'Modified by Morgan 2021/5/14 通知擇一(1232)又改要,Ex:P-124558--郭 (109.11.12 玲玲請作)
   'Modified by Morgan 2022/7/14 +第三方意見1815--郭
   'Modified by Morgan 2022/9/26 +被舉發1802--郭
   'Modified by Lydia 2025/03/18 從常數「通知申復」改回1202審查意見通知函
   If (Text7 = "1202" Or Text7 = "1227" Or Text7 = "1232" Or Text7 = "1815" Or Text7 = "1802") And (pa(9) = "000" Or pa(9) = "020" Or pa(9) = "044") And (pa(8) = "1" Or pa(8) = "3") Then
      m_USCaseNo = PUB_GetUSCaseNo(pa(1), pa(2), pa(3), pa(4))
      
      If m_USCaseNo <> "" Then
         If txtIDSFee(1) = "" Or txtIDSFee(2) = "" Or txtIDSPt(1) = "" Or txtIDSPt(2) = "" Then
            'Added by Morgan 2021/5/14
            If Text7 = "1232" Then
               If MsgBox("是否通知ＩＤＳ報價？", vbYesNo + vbDefaultButton1 + vbExclamation, "ＩＤＳ報價") = vbNo Then
                  m_USCaseNo = ""
               End If
            End If
            'end 2021/5/14
            
            'Modified by Morgan 2020/7/16 若有輸報價表示已確認要通知，不必再問--玲玲,品薇
            'Added by Morgan 2020/3/5
            If m_USCaseNo <> "" Then
               '有輸過審查意見的都要提醒
               If PUB_ChkCPExist(cp, "1202") = True Then
                  strExc(0) = "1.請確認引證前案是否與前一次審查意見通知書相同。" & vbCrLf & _
                     "2.請確認美國案 " & m_USCaseNo & " 是否已提出相同引證前案的IDS。" & vbCrLf & _
                     "若二者均相同可不必通知IDS報價"
                  strExc(0) = strExc(0) & vbCrLf & vbCrLf & "【是】:要通知    【否】:不通知    【取消】:回畫面" 'Added by Morgan 2020/12/18
                  intI = MsgBox(strExc(0), vbYesNoCancel + vbInformation + vbDefaultButton3, "是否通知IDS報價？")
                  If intI = vbCancel Then
                     Exit Function
                  ElseIf intI = vbNo Then
                     m_USCaseNo = ""
                  End If
               End If
            End If
            'end 2020/3/5
            
            If m_USCaseNo <> "" Then
               SSTab1.Tab = 1
               If MsgBox("尚未輸入ＩＤＳ報價，是否 EMail 通知 CFP 程序人員報價？", vbYesNo + vbDefaultButton2 + vbExclamation, "ＩＤＳ報價") = vbYes Then
                  'Modified by Morgan 2023/5/24
                  'strExc(0) = "核駁函"
                  strExc(0) = Label3(4)
                  strExc(5) = ""
                  Do
                     strExc(5) = InputBox("請輸入引證前案檔案數量：")
                     If Val(strExc(5)) > 0 Then
                        Exit Do
                     ElseIf strExc(5) = "" Then
                        MsgBox "未輸入引證前案檔案數量，取消 EMail 通知！", vbExclamation
                        Exit Function
                     Else
                        MsgBox "引證前案檔案數量必須大於 0，請重新輸入！", vbExclamation
                     End If
                  Loop
                  'end 2023/5/24
                  
                  strExc(2) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
                  'Modified by Morgan 2021/2/25 考慮會有多個美國案
                  'strExc(1) = PUB_GetCFPHandler(m_USCaseNo)
                  'strExc(4) = strExc(2) & " 案已收到核駁函，請提供相關美國案( " & m_USCaseNo & " )的IDS報價！"
                  arrCaseNo = Split(m_USCaseNo, "、")
                  For ii = LBound(arrCaseNo) To UBound(arrCaseNo)
                     strExc(4) = strExc(2) & " 案已收到" & strExc(0) & "，請提供相關美國案( " & arrCaseNo(ii) & " )的IDS報價！"
                     strExc(1) = PUB_GetCFPHandler(arrCaseNo(ii))
                  'end 2021/1/25
                     If strExc(1) <> "" Then
                        'Modified by Morgan 2019/9/9 調整報價欄位名及定稿內容--郭
                        'Modified by Morgan 2023/5/24 +引證前案檔案數量
                        strExc(3) = "引證前案共: " & strExc(5) & " 件" & vbCrLf & _
                                    "IDS報價:" & vbCrLf & _
                                    "　1.第一階段　　　(　P)" & vbCrLf & _
                                    "　2.第二階段　　　(　P)" & vbCrLf & vbCrLf & _
                                    "**　若該案已是第二階段，則第一階段請輸　0　**"
      
                        PUB_SendMail strUserNum, strExc(1), "", strExc(4), strExc(3)
                     End If
                  Next 'Added by Morgan 2021/2/25
                  
               ElseIf txtIDSFee(1) = "" Then
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
   End If
   'end 2019/5/24
   
'Added by Morgan 2014/4/9 電子化-檢查pdf檔
'Modified by Morgan 2014/4/25 +1221,1810+205下一程序
'Modified by Morgan 2016/7/5 非臺灣案電子化
'If pa(9) = "000" And Text13 = "" And (Text7 = "1202" Or Text7 = "1221" Or (Text7 = "1810" And text8 = "205")) Then
'Modified by Morgan 2016/8/29 +1209 --玲玲
If Left(Pub_StrUserSt03, 1) <> "F" And Text13 = "" And (Text7 = "1202" Or Text7 = "1209" Or Text7 = "1221" Or (Text7 = "1810" And Text8 = "205")) Then
'end 2016/7/5
   MsgBox "請輸入引證前案檔案數量!!", vbExclamation
   SSTab1.Tab = 0
   Text13.SetFocus
   Exit Function
End If

If pa(9) = "000" Then
   If PUB_CheckPDF(pa(1), pa(2), pa(3), pa(4), 1 + Val(Text13), m_DocNo) = False Then
      Exit Function
   End If
End If
'end 2014/4/9

'Added by Morgan 2015/6/22
'1001,1002,1202,1209,1802,1807,1809,1810 E化提醒
If InStr("1001,1002,1202,1221,1209,1802,1807,1809,1810", Text7) > 0 Then
   If PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4), , , bPaper) = True And bPaper = False Then
      MsgBox "E化案件，不印前案!!", vbExclamation
   End If
End If
'end 2015/6/22

   'Added by Morgan 2020/1/17
   '大陸案,有通知函,程序承辦,非掛號(無期限)
   'm_bolNoCP27 = False 'Removed by Morgan 2024/4/25
   'Removed by Morgan 2024/1/30 取消--郭
   'If pa(9) = "020" And Text15(0) <> "N" And PUB_GetST03(Text16) = "P12" And Text14(1) = "" Then
   '   If PUB_GetCustomerValue(pa(26), "CU182") = "Y" Then
   '      If MsgBox("請確認是否已收到公文正本？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
   '         m_bolNoCP27 = True
   '      End If
   '   End If
   'End If
   'end 2020/1/17
   
   'Added by Morgan 2020/8/12
   '若為來函期限2次確認退回時需檢查法限是否一致
   If m_strIR01 <> "" Then
      If PUB_ChkReKeyInOk(m_strIR01, m_strIR02, m_strIR03, m_strIR04, Text14(1).Text, m_bolReKeyInOK) = False Then
         Text14(1).SetFocus
         Exit Function
      End If
   End If
   'end 2020/8/12
   
   'Added by Morgan 2024/4/29
   m_bolEngCase = False
   'Modified by Morgan 2024/5/7
   'If GetStaffDepartment(Text16) <> "P12" Then
   If Not m_bolFMP And Text16.Enabled And GetStaffDepartment(Text16) <> "P12" Then
   'end 2024/5/7
      m_bolEngCase = True
      Text15(0).Text = "N"
   End If
   'end 2024/4/29
   
TxtValidate = True
End Function

'Add by Morgan 2004/6/28 檢查是否有補文件未發文
Private Function CheckCP(ByRef stPA() As String, ByRef p_stCP09 As String, Optional ByVal stCP10 As String = "202") As Boolean

On Error GoTo ErrHnd
   
   strSql = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & stPA(1) & "' AND CP02='" & stPA(2) & "' AND CP03='" & stPA(3) & "' AND CP04='" & stPA(4) & "'" & _
      " AND CP10='" & stCP10 & "' AND CP27 IS NULL AND CP57 IS NULL"
      
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      p_stCP09 = "" & adoRecordset.Fields("CP09")
   End If
   CheckCP = True
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
End Function

Private Sub txtDispDate_GotFocus()
   TextInverse txtDispDate
End Sub

'Add by Lydia 2014/11/26
'針對1004(延期受理)->判斷所掛的延期(現.cp43)那道相關總收文號(讀前一筆.cp43) 是否已收文，是否已發文
Private Function Check2mail1004(ByRef r_PA() As String, ByRef rPreNo As String, ByVal rKind As String)
Dim rsB As ADODB.Recordset
Dim rsB2 As ADODB.Recordset
Dim bStr01 As String, cStr01 As String, cStr02 As String
On Error GoTo ErrHnd
   'rPreNo 延期掛的單據

If rKind = 0 Then '判斷
         'Modified by Lydia 2014/12/22 修改判斷: 掛的延期的相關總收文號與延期受理之間有申復
'   bStr01 = "SELECT a.CP09 as a01,a.CP10 as a02,a.CP13 as a03,a.CP14 as a04,NVL(a.CP27,0) as a05," & _
            "b.CP09 as b01,b.CP10 as b02,b.CP13 as b03,b.CP14 as b04,NVL(b.CP27,0) as b05,CPM03 as b06 " & _
            "FROM CASEPROGRESS a,CASEPROGRESS b,casepropertymap WHERE a.CP01=b.CP01(+) and a.CP02=b.CP02(+) and a.CP03=b.CP03(+) and a.CP04=b.CP04(+) " & _
            "AND b.cp01=cpm01 and b.cp10=cpm02 AND a.CP01='" & r_PA(1) & "' AND a.CP02='" & r_PA(2) & "' AND a.CP03='" & r_PA(3) & "' AND a.CP04='" & r_PA(4) & "' " & _
            "AND a.CP09='" & rPreNo & "' and a.CP43=b.CP09(+) "
   bStr01 = "SELECT a.CP09 as a01,a.CP10 as a02,a.CP13 as a03,a.CP14 as a04,NVL(a.CP27,0) as a05," & _
            "b.CP09 as b01,b.CP10 as b02,b.CP13 as b03,b.CP14 as b04,NVL(b.CP27,0) as b05,CPM03 as b06,np01,np06 " & _
            "FROM CASEPROGRESS a,CASEPROGRESS b,casepropertymap,nextprogress WHERE a.CP01=b.CP01(+) and a.CP02=b.CP02(+) and a.CP03=b.CP03(+) and a.CP04=b.CP04(+) " & _
            "AND b.cp01=cpm01 and b.cp10=cpm02 AND a.CP01='" & r_PA(1) & "' AND a.CP02='" & r_PA(2) & "' AND a.CP03='" & r_PA(3) & "' AND a.CP04='" & r_PA(4) & "' " & _
            "AND a.CP09='" & rPreNo & "' and a.CP43=b.CP09(+) and b.cp09=np01(+)"

    intI = 1
    Set rsB = ClsLawReadRstMsg(intI, bStr01)
    If intI = 1 Then
       If rsB.RecordCount > 0 Then
         ' cstr01 是否已收文
         cStr01 = "Y"
         ' cstr02 是否已發文
         cStr02 = "Y"
         'Modified by Lydia 2014/12/22 修改判斷: 掛的延期的相關總收文號與延期受理之間有申復
         'If Left(rsB!b01, 1) = "C" Then cStr01 = "N"
'         ' cstr02 是否已發文
'         cStr02 = "Y"
'         If rsB!b05 = "0" Then cStr02 = "N"
         If Left(rsB!b01, 1) = "C" Then
           If rsB!np06 = "Y" Then
              'Modified by Lydia 2014/12/30 修正BUG,排除掛的延期的相關總收文號
'              bStr01 = "SELECT CP09,NVL(CP27,0) CP27 FROM CASEPROGRESS WHERE CP01='" & r_PA(1) & "' AND CP02='" & r_PA(2) & "' AND CP03='" & r_PA(3) & "' AND CP04='" & r_PA(4) & "' " & _
'                       "AND cp43='" & rsB!b01 & "' and substr(cp09,1,1)='A'"
              'Modified by Lydia 2016/07/14 因為有可能申復或再審,所以案件性質只剔除分析941
              'Modified by Lydia 2016/08/02 +抓承辦人和智權人員
              bStr01 = "SELECT CP09,NVL(CP27,0) CP27,NVL(CP13,'') CP13,NVL(CP14,'') CP14 FROM CASEPROGRESS WHERE CP01='" & r_PA(1) & "' AND CP02='" & r_PA(2) & "' AND CP03='" & r_PA(3) & "' AND CP04='" & r_PA(4) & "' " & _
                       "AND cp43='" & rsB!b01 & "' AND instr(cp09,'" & rsB!a01 & "') = 0 and substr(cp09,1,1)='A' and cp10 not in ('941') "
              intI = 1
              Set rsB2 = ClsLawReadRstMsg(intI, bStr01)
              If rsB2.RecordCount > 0 Then
                '已收文, 其次判斷發文狀況
                 If rsB2!Cp27 = "0" Then cStr02 = "N"
              End If
           Else
              cStr01 = "N"
           End If
         Else
           '已收文, 其次判斷發文狀況
           If rsB!b05 = "0" Then cStr02 = "N"
         End If
         'end 'Modified by Lydia 2014/12/22
                  
         If cStr01 = "Y" Then
            'Memo by Lydia 2016/08/02 延期受理,與所掛的那道進度的相關總收文號期間,有收文申復並且已發文,產生b類收文933和發mail
            If cStr02 = "Y" Then
               mPty1004(0) = "3" '訊息類別
               '自動產生內部收文933(覆函),承辦人預設同補文件(202)相同特定人員
               'Added by Morgan 2025/1/24
               If strSrvDate(1) >= P業務區劃分啟用日 Then
                  mPty1004(1) = PUB_GetPHandler(pa(1) & pa(2) & pa(3) & pa(4))
               Else
               'end 2025/1/24
                  mPty1004(1) = Pub_GetSpecMan("AB202")
               End If 'Added by Morgan 2025/1/24
               mPty1004(2) = rsB!b01 '掛的延期(現.cp43)那道相關總收文號(讀前一筆.cp43)
               mPty1004(3) = rsB!b06
            'Memo by Lydia 2016/08/02 若期間內所做的申復尚未發文,請發e-mail通知未發文那道的智權同仁及承辦工程師
            Else
               mPty1004(0) = "1"
               'Modified by Lydia 2016/08/08 判斷是所掛之申復未發文,還是中間程序未發文
               If Left(rsB!b01, 1) <> "C" Then
                    If rsB!b03 = rsB!b04 Then
                      mPty1004(1) = rsB!b03 '智權人員和承辦人員=同一人
                    Else
                      mPty1004(1) = rsB!b03 & ";" & rsB!b04 'mail給智權人員和承辦人員
                    End If
               ElseIf rsB2!cp13 <> "" Then   '期間內所做的申復尚未發文
                    If rsB2!cp13 = rsB2!cp14 Then
                      mPty1004(1) = rsB2!cp13 '智權人員和承辦人員=同一人
                    Else
                      mPty1004(1) = rsB2!cp13 & ";" & rsB2!cp14 'mail給智權人員和承辦人員
                    End If
               End If
               'end 2016/08/08
            End If
         Else
               mPty1004(0) = "2"
               mPty1004(1) = rsB!b03  'mail給智權人員
               mPty1004(2) = rsB!b01 '所掛的延期(現.cp43)那道相關總收文號(讀前一筆.cp43)
         End If
       End If
    End If
Else
'發mail
    strExc(4) = Label3(4) '(現)案件性質
    strExc(5) = ChangeTStringToTDateString(Trim(Text14(0)))  '(現)本所期限

     '1.)
        strExc(0) = "已收到 " & Text2 & "-" & Text3 & "-" & Text4 & "-" & Text5 & " 主管機關來函(" & Label3(4) & ")，本所期限:" & ChangeTStringToTDateString(Trim(Text14(0)))
     '2.)
        If rKind = "2" Then
           strExc(4) = "": strExc(5) = ""
           bStr01 = "SELECT NP01,NP07,CPM03,NP08,NP09,NP10,CP06 FROM nextprogress,casepropertymap,caseprogress " & _
                    "WHERE NP02=CPM01(+) and NP07=CPM02(+) and NP01='" & mPty1004(2) & "' " & _
                    " and np01=cp09(+) and np02=cp01(+) and np03=cp02(+) and np04=cp03(+) and np05=cp04(+) "
            intI = 1
            Set rsB = ClsLawReadRstMsg(intI, bStr01)
            If intI = 1 Then
               If rsB.RecordCount > 0 Then
                 strExc(4) = rsB!cpm03 '下一程序的案件性質
                 strExc(5) = ChangeWStringToTDateString(rsB!np08) '下一程序的本所期限
                 strExc(0) = strExc(0) & "，本案(" & rsB!cpm03 & ")尚未收文請儘速收文"
               End If
            End If
        End If
     '3.)
        If rKind = "3" Then  'ex:100208711

           strExc(4) = mPty1004(3)
           strExc(0) = Text2 & "-" & Text3 & "-" & Text4 & "-" & Text5 & " 已作內部收文(覆函)" & strExc(4) & "，已於" & ChangeTStringToTDateString(strSrvDate(2)) & "發文"
        End If
    strExc(0) = strExc(0) & "，內容請參照卷宗區的電子檔"
    '內文
         If Len(pa(5)) > 0 Then strExc(3) = Trim(pa(5))
         If Len(strExc(3)) = 0 Then strExc(3) = Trim(pa(6))
         If Len(strExc(3)) = 0 Then strExc(3) = Trim(pa(7))

         strExc(1) = "本所案號：" & Text2 & "-" & Text3 & "-" & Text4 & "-" & Text5 & vbCrLf & _
                    "案件名稱：" & strExc(3) & vbCrLf & _
                    "案件性質：" & strExc(4) & vbCrLf & _
                    "申請人　：" & GetCustomerName(pa(26)) & vbCrLf
         If rKind <> "3" Then strExc(1) = strExc(1) & "本所期限：" & strExc(5) & vbCrLf
         strExc(1) = strExc(1) & "來函內容：內容請參照卷宗區的電子檔"
         '收件者
         strExc(2) = mPty1004(1) '判斷時已設定mail收件者

         PUB_SendMail strUserNum, strExc(2), "", strExc(0), strExc(1)
         
End If
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
End Function

'Added by Morgan 2015/4/20
'例外承辦期限控制
Private Sub SetException()
   If m_bolFMP Then
'Modified by Morgan 2016/9/21 改呼叫共用函數
'      '先正達OA承辦期限設7個工作天,若下一程序為 804,501-509時設2個工作天(24Hr)
'      If InStr("Y4830900,Y4830901,Y4830902,Y4830903,Y4830904,Y4830905,Y4830908,Y5132600", Left(pa(75) & "000", 8)) > 0 Then
'         If text8 = "804" Or text8 >= "501" And text8 <= "509" Then
'            Text17 = TransDate(CompWorkDay(2, TransDate(Label3(6).Caption, 2), 0), 1)
'         ElseIf (Text7 = "1202" Or Text7 = "1227") Then
'            Text17 = TransDate(CompWorkDay(7, TransDate(Label3(6).Caption, 2), 0), 1)
'         End If
'      'Added by Morgan 2015/7/3 --吳彩菱
'      'Y51753+X45149010 承辦天數:14 起算日期:系統日
'      ElseIf Left(pa(75) & "000", 8) = "Y5175300" And Left(pa(26) & "000", 8) = "X4514901" Then
'         If (Text7 = "1202" Or Text7 = "1227") Then
'            Text17 = TransDate(CompDate(2, 14, strSrvDate(1)), 1)
'         End If
'      End If
      'Modified by Morgan 2018/12/19 + DBDATE(Text6)
      Call Pub_SetExceptCP48(pa(75), pa(26), Text7.Text, TransDate(Label3(6).Caption, 2), Text17, Text8, , , DBDATE(Text6))
'end 2016/9/21
   End If
End Sub
'Added by Lydia 2015/04/30 通知函是否為副本
Private Sub msgCCC(ByVal chkStr As String)
  If m_DocNo <> "" And chkStr = "" Then
     If MsgBox("請確認此通知函是否為副本", vbYesNo + vbDefaultButton2) = vbYes Then
        Text8 = "": Text14(0).Text = "": Text14(1).Text = ""
        m_bolCCC = True
     End If
  End If
End Sub

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
