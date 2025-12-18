VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090201_b 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標部工作進度資料維護"
   ClientHeight    =   6220
   ClientLeft      =   3370
   ClientTop       =   2940
   ClientWidth     =   10020
   ControlBox      =   0   'False
   FillColor       =   &H0080FFFF&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6220
   ScaleWidth      =   10020
   Begin TabDlg.SSTab SSTab1 
      Height          =   5685
      Left            =   30
      TabIndex        =   31
      Top             =   480
      Width           =   9930
      _ExtentX        =   17498
      _ExtentY        =   10037
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "瀏覽資料"
      TabPicture(0)   =   "frm090201_b.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(3)=   "Combo1"
      Tab(0).Control(4)=   "cmdok2(0)"
      Tab(0).Control(5)=   "cmdok2(1)"
      Tab(0).Control(6)=   "grd1"
      Tab(0).Control(7)=   "Combo3"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "詳細資料"
      TabPicture(1)   =   "frm090201_b.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(7)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(31)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(29)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(6)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(27)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label1(24)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(23)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lbl1(23)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lbl1(19)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lbl1(17)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lbl1(11)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbl1(9)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lbl1(7)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lbl1(5)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lbl1(3)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Label1(32)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lbl1(30)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "lbl1(28)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "lbl1(8)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "lbl1(6)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label1(30)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "lbl1(29)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label1(11)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label1(13)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label1(14)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label1(15)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label1(16)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label1(17)"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label1(18)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Label1(19)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Label1(20)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Label1(21)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label1(8)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "lbl1(0)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "lbl1(1)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Label1(1)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "lbl1(21)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Label1(3)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Label1(22)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Label1(12)"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "lblClose"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Label1(46)"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Label1(39)"
      Tab(1).Control(43).Enabled=   0   'False
      Tab(1).Control(44)=   "lbl1(13)"
      Tab(1).Control(44).Enabled=   0   'False
      Tab(1).Control(45)=   "lbl1(15)"
      Tab(1).Control(45).Enabled=   0   'False
      Tab(1).Control(46)=   "Label1(26)"
      Tab(1).Control(46).Enabled=   0   'False
      Tab(1).Control(47)=   "lbl1(10)"
      Tab(1).Control(47).Enabled=   0   'False
      Tab(1).Control(48)=   "Label6"
      Tab(1).Control(48).Enabled=   0   'False
      Tab(1).Control(49)=   "Label1(47)"
      Tab(1).Control(49).Enabled=   0   'False
      Tab(1).Control(50)=   "Label1(4)"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Label1(5)"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "lblEApp"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "Label18"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Label1(9)"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "lblFee"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "txtEP12"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "txtCP64"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "Combo2"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "Combo6"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "lblCertType"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "Label1(2)"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "txt1(8)"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "txt1(4)"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "txt1(1)"
      Tab(1).Control(64).Enabled=   0   'False
      Tab(1).Control(65)=   "txt1(9)"
      Tab(1).Control(65).Enabled=   0   'False
      Tab(1).Control(66)=   "txt1(12)"
      Tab(1).Control(66).Enabled=   0   'False
      Tab(1).Control(67)=   "txt1(2)"
      Tab(1).Control(67).Enabled=   0   'False
      Tab(1).Control(68)=   "cmd(2)"
      Tab(1).Control(68).Enabled=   0   'False
      Tab(1).Control(69)=   "chk1"
      Tab(1).Control(69).Enabled=   0   'False
      Tab(1).Control(70)=   "cmd(3)"
      Tab(1).Control(70).Enabled=   0   'False
      Tab(1).Control(71)=   "cmdPic"
      Tab(1).Control(71).Enabled=   0   'False
      Tab(1).Control(72)=   "txt1(18)"
      Tab(1).Control(72).Enabled=   0   'False
      Tab(1).Control(73)=   "txt1(3)"
      Tab(1).Control(73).Enabled=   0   'False
      Tab(1).Control(74)=   "cmd(4)"
      Tab(1).Control(74).Enabled=   0   'False
      Tab(1).Control(75)=   "cmd(5)"
      Tab(1).Control(75).Enabled=   0   'False
      Tab(1).Control(76)=   "txt1(7)"
      Tab(1).Control(76).Enabled=   0   'False
      Tab(1).Control(77)=   "textCP143"
      Tab(1).Control(77).Enabled=   0   'False
      Tab(1).Control(78)=   "cmd(6)"
      Tab(1).Control(78).Enabled=   0   'False
      Tab(1).Control(79)=   "cmdDataMail"
      Tab(1).Control(79).Enabled=   0   'False
      Tab(1).Control(80)=   "cmd(7)"
      Tab(1).Control(80).Enabled=   0   'False
      Tab(1).Control(81)=   "txtNote"
      Tab(1).Control(81).Enabled=   0   'False
      Tab(1).Control(82)=   "cmdOK(5)"
      Tab(1).Control(82).Enabled=   0   'False
      Tab(1).Control(83)=   "SSTab2"
      Tab(1).Control(83).Enabled=   0   'False
      Tab(1).Control(84)=   "Frame1"
      Tab(1).Control(84).Enabled=   0   'False
      Tab(1).Control(85)=   "cmdOK(3)"
      Tab(1).Control(85).Enabled=   0   'False
      Tab(1).Control(86)=   "cmdOK(4)"
      Tab(1).Control(86).Enabled=   0   'False
      Tab(1).Control(87)=   "Frame2"
      Tab(1).Control(87).Enabled=   0   'False
      Tab(1).ControlCount=   88
      TabCaption(2)   =   "待辦歷程"
      TabPicture(2)   =   "frm090201_b.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdQuery"
      Tab(2).Control(1)=   "cmdDetail"
      Tab(2).Control(2)=   "Combo5"
      Tab(2).Control(3)=   "grd2"
      Tab(2).Control(4)=   "Label16"
      Tab(2).Control(5)=   "Label1(48)"
      Tab(2).ControlCount=   6
      Begin VB.Frame Frame2 
         Height          =   370
         Left            =   3870
         TabIndex        =   115
         Top             =   2760
         Visible         =   0   'False
         Width           =   5740
         Begin VB.TextBox txt1 
            Height          =   300
            Index           =   19
            Left            =   4330
            MaxLength       =   7
            TabIndex        =   9
            Top             =   0
            Width           =   915
         End
         Begin VB.TextBox txt1 
            Height          =   300
            Index           =   5
            Left            =   1120
            MaxLength       =   6
            TabIndex        =   116
            Top             =   0
            Visible         =   0   'False
            Width           =   900
         End
         Begin MSForms.ComboBox Combo4 
            Height          =   320
            Left            =   1140
            TabIndex        =   8
            Top             =   0
            Width           =   1890
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "3334;564"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "外文核完日："
            Height          =   180
            Index           =   41
            Left            =   3240
            TabIndex        =   118
            Top             =   60
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   1  '靠右對齊
            AutoSize        =   -1  'True
            Caption         =   "外文核稿人："
            Height          =   180
            Index           =   25
            Left            =   -810
            TabIndex        =   117
            Top             =   60
            Width           =   1920
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "完整卷宗"
         Height          =   320
         Index           =   4
         Left            =   7050
         Style           =   1  '圖片外觀
         TabIndex        =   113
         Top             =   2190
         Width           =   960
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "接洽單"
         Height          =   320
         Index           =   3
         Left            =   8940
         Style           =   1  '圖片外觀
         TabIndex        =   102
         Top             =   5100
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Frame Frame1 
         Height          =   1030
         Left            =   7020
         TabIndex        =   109
         Top             =   1500
         Width           =   2830
         Begin VB.CommandButton cmd 
            Caption         =   "智權人員補充資料記錄(&S)"
            Height          =   320
            Index           =   0
            Left            =   0
            TabIndex        =   112
            Top             =   0
            Width           =   2210
         End
         Begin VB.CommandButton cmd 
            Caption         =   "通知補充資料(&M)"
            Height          =   320
            Index           =   1
            Left            =   0
            TabIndex        =   111
            Top             =   330
            Width           =   1590
         End
         Begin VB.CommandButton cmdTSMap 
            Caption         =   "委查結果"
            Height          =   320
            Left            =   1080
            TabIndex        =   110
            Top             =   680
            Width           =   960
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   1420
         Left            =   0
         TabIndex        =   106
         Top             =   4230
         Width           =   3730
         _ExtentX        =   6579
         _ExtentY        =   2505
         _Version        =   393216
         Tab             =   1
         TabHeight       =   360
         TabCaption(0)   =   "條款"
         TabPicture(0)   =   "frm090201_b.frx":0054
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txt1(11)"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "商標描述中文"
         TabPicture(1)   =   "frm090201_b.frx":0070
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "txt1(0)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "商標描述英文"
         TabPicture(2)   =   "frm090201_b.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txt1(6)"
         Tab(2).ControlCount=   1
         Begin VB.TextBox txt1 
            Height          =   1140
            Index           =   11
            Left            =   -74970
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  '垂直捲軸
            TabIndex        =   114
            Top             =   240
            Width           =   3620
         End
         Begin VB.TextBox txt1 
            Height          =   1140
            Index           =   6
            Left            =   -74970
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  '垂直捲軸
            TabIndex        =   108
            Top             =   240
            Width           =   3620
         End
         Begin VB.TextBox txt1 
            Height          =   1140
            Index           =   0
            Left            =   30
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  '垂直捲軸
            TabIndex        =   107
            Top             =   240
            Width           =   3620
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "客戶專區"
         Height          =   320
         Index           =   5
         Left            =   8205
         Style           =   1  '圖片外觀
         TabIndex        =   25
         Top             =   4500
         Width           =   870
      End
      Begin VB.TextBox txtNote 
         Appearance      =   0  '平面
         BackColor       =   &H8000000A&
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   90
         TabIndex        =   100
         Text            =   "※此案屬多案歷程，請參"
         Top             =   330
         Visible         =   0   'False
         Width           =   3645
      End
      Begin VB.CommandButton cmd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "未完稿暫存區"
         Height          =   320
         Index           =   7
         Left            =   7020
         Style           =   1  '圖片外觀
         TabIndex        =   21
         Top             =   1130
         Width           =   1260
      End
      Begin VB.CommandButton cmdDataMail 
         BackColor       =   &H00C0E0FF&
         Caption         =   "寄發指示信"
         Height          =   320
         Left            =   7020
         Style           =   1  '圖片外觀
         TabIndex        =   24
         Top             =   4500
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
         Caption         =   "電子送件"
         Height          =   320
         Index           =   6
         Left            =   6060
         TabIndex        =   17
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox textCP143 
         Enabled         =   0   'False
         Height          =   300
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   775
         Width           =   930
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   7
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   11
         Top             =   3690
         Width           =   915
      End
      Begin VB.CommandButton cmd 
         Caption         =   "承辦歷程(&E)"
         Height          =   320
         Index           =   5
         Left            =   8430
         TabIndex        =   19
         Top             =   600
         Width           =   1200
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "畫面更新(&Q)"
         Height          =   320
         Left            =   -66960
         TabIndex        =   90
         Top             =   480
         Width           =   1125
      End
      Begin VB.CommandButton cmdDetail 
         Caption         =   "明細資料(&D)"
         Height          =   320
         Left            =   -68130
         TabIndex        =   89
         Top             =   480
         Width           =   1125
      End
      Begin VB.ComboBox Combo5 
         Height          =   260
         ItemData        =   "frm090201_b.frx":00A8
         Left            =   -69150
         List            =   "frm090201_b.frx":00B8
         Style           =   2  '單純下拉式
         TabIndex        =   88
         Top             =   510
         Width           =   960
      End
      Begin VB.CommandButton cmd 
         Caption         =   "申請書(&A)"
         Height          =   320
         Index           =   4
         Left            =   7230
         TabIndex        =   18
         Top             =   360
         Width           =   1080
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   3
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   4
         Top             =   1410
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   18
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   6
         Top             =   2030
         Width           =   915
      End
      Begin VB.CommandButton cmdPic 
         BackColor       =   &H00C0C0C0&
         Caption         =   "代表圖(&I)"
         Height          =   320
         Left            =   8205
         Style           =   1  '圖片外觀
         TabIndex        =   23
         Top             =   4080
         Width           =   1440
      End
      Begin VB.CommandButton cmd 
         Caption         =   "撰寫信函(&L)"
         Enabled         =   0   'False
         Height          =   320
         Index           =   3
         Left            =   8430
         TabIndex        =   20
         Top             =   1125
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CheckBox chk1 
         Caption         =   "無圖式"
         Height          =   255
         Left            =   7320
         TabIndex        =   14
         Top             =   4050
         Width           =   915
      End
      Begin VB.CommandButton cmd 
         Caption         =   "承辦單列印(&P)"
         Height          =   320
         Index           =   2
         Left            =   8610
         TabIndex        =   22
         Top             =   5190
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   2
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   1
         Top             =   775
         Width           =   930
      End
      Begin VB.ComboBox Combo3 
         Height          =   260
         ItemData        =   "frm090201_b.frx":00D7
         Left            =   -70260
         List            =   "frm090201_b.frx":00EA
         TabIndex        =   79
         Top             =   390
         Width           =   2430
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
         Height          =   4755
         Left            =   -74940
         TabIndex        =   32
         Top             =   810
         Width           =   9795
         _ExtentX        =   17268
         _ExtentY        =   8378
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         HighLight       =   2
         AllowUserResizing=   3
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
         _Band(0).Cols   =   1
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   12
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   0
         Top             =   465
         Width           =   930
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   9
         Left            =   5010
         MaxLength       =   1
         TabIndex        =   15
         Top             =   4290
         Width           =   360
      End
      Begin VB.CommandButton cmdok2 
         Caption         =   "未發文"
         Height          =   400
         Index           =   1
         Left            =   -66648
         TabIndex        =   30
         Top             =   348
         Width           =   852
      End
      Begin VB.CommandButton cmdok2 
         Caption         =   "當月資料"
         Height          =   400
         Index           =   0
         Left            =   -67656
         TabIndex        =   29
         Top             =   348
         Width           =   972
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   1
         Left            =   5010
         MaxLength       =   1
         TabIndex        =   5
         Top             =   1720
         Width           =   480
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   4
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   7
         Top             =   2340
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   300
         Index           =   8
         Left            =   5010
         MaxLength       =   7
         TabIndex        =   13
         Top             =   3990
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
         Height          =   4755
         Left            =   -74940
         TabIndex        =   91
         Top             =   840
         Width           =   9825
         _ExtentX        =   17339
         _ExtentY        =   8378
         _Version        =   393216
         Cols            =   14
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|目次|流程日期|本所案號|案件名稱|國家|種類|案件性質|本所期限|承辦人|承辦期限|智權人員|目前流程狀態|不顯示"
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
         _Band(0).Cols   =   14
      End
      Begin VB.Label Label1 
         Caption         =   "註：出申請書時，建議勿同時使用Word軟體，因程式執行中會使用到Word。"
         ForeColor       =   &H000000C0&
         Height          =   200
         Index           =   2
         Left            =   3840
         TabIndex        =   86
         Top             =   5430
         Width           =   6030
      End
      Begin VB.Label lblCertType 
         AutoSize        =   -1  'True
         Caption         =   "lblCertType"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   2310
         TabIndex        =   105
         Top             =   2745
         Width           =   1005
      End
      Begin MSForms.ComboBox Combo6 
         Height          =   320
         Left            =   7320
         TabIndex        =   12
         Top             =   3690
         Width           =   1950
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3440;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   315
         Left            =   5010
         TabIndex        =   3
         Top             =   1080
         Width           =   1890
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3334;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCP64 
         Height          =   510
         Left            =   5010
         TabIndex        =   16
         Top             =   4860
         Width           =   3590
         VariousPropertyBits=   -1466941409
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "6332;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEP12 
         Height          =   510
         Left            =   5010
         TabIndex        =   10
         Top             =   3180
         Width           =   3590
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "6332;900"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   315
         Left            =   -74070
         TabIndex        =   103
         Top             =   390
         Width           =   2430
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "4286;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblFee 
         AutoSize        =   -1  'True
         Caption         =   "lblFee"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   2310
         TabIndex        =   101
         Top             =   3420
         Width           =   2100
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "查名齊備日："
         Height          =   180
         Index           =   9
         Left            =   6120
         TabIndex        =   99
         Top             =   835
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "可不跑承辦歷程"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8430
         TabIndex        =   98
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblEApp 
         AutoSize        =   -1  'True
         Caption         =   "電子送件"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3030
         TabIndex        =   97
         Top             =   1050
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "核稿人："
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   4260
         TabIndex        =   96
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "會稿完成日："
         Height          =   260
         Index           =   4
         Left            =   3890
         TabIndex        =   95
         Top             =   3710
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "判發人："
         ForeColor       =   &H00FF0000&
         Height          =   260
         Index           =   47
         Left            =   6560
         TabIndex        =   94
         Top             =   3770
         Width           =   740
      End
      Begin VB.Label Label16 
         Caption         =   "註：雙擊選取時，開啟承辦歷程。　多案歷程：案件名稱欄位顯示紫紅色。   "
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   -74940
         TabIndex        =   93
         Top             =   330
         Width           =   6315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "最近聯絡："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   48
         Left            =   -70050
         TabIndex        =   92
         Top             =   570
         Width           =   900
      End
      Begin VB.Label Label6 
         Caption         =   "智權人員做”爭議案齊備日輸入”之『回覆補充資料』或『齊備日或急件維護』後，系統會自動上齊備日。"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.5
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   450
         Left            =   4410
         TabIndex        =   87
         Top             =   2700
         Visible         =   0   'False
         Width           =   4730
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   10
         Left            =   5040
         TabIndex        =   85
         Top             =   1440
         Visible         =   0   'False
         Width           =   1035
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1826;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "完稿日："
         Height          =   255
         Index           =   26
         Left            =   4260
         TabIndex        =   84
         Top             =   1440
         Width           =   735
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   1080
         TabIndex        =   83
         Top             =   2745
         Width           =   1200
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   13
         Left            =   1080
         TabIndex        =   82
         Top             =   2370
         Width           =   1410
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2487;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "指定會稿日："
         Height          =   255
         Index           =   39
         Left            =   3915
         TabIndex        =   81
         Top             =   2055
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "進度備註："
         Height          =   255
         Index           =   46
         Left            =   4095
         TabIndex        =   80
         Top             =   4890
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "顏色說明："
         Height          =   225
         Left            =   -71160
         TabIndex        =   78
         Top             =   432
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "承辦人： "
         Height          =   180
         Index           =   0
         Left            =   -74904
         TabIndex        =   77
         Top             =   432
         Width           =   792
      End
      Begin VB.Label lblClose 
         AutoSize        =   -1  'True
         Caption         =   "lblClose"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3030
         TabIndex        =   76
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   255
         Index           =   12
         Left            =   45
         TabIndex        =   75
         Top             =   3713
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "(N:  不通知, 不發文)"
         ForeColor       =   &H00FF0000&
         Height          =   260
         Index           =   22
         Left            =   5450
         TabIndex        =   74
         ToolTipText     =   "(N:  不通知, 自動內部收文)"
         Top             =   4350
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "是否通知客戶："
         ForeColor       =   &H00FF0000&
         Height          =   260
         Index           =   3
         Left            =   3650
         TabIndex        =   73
         Top             =   4350
         Width           =   1350
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   21
         Left            =   1080
         TabIndex        =   72
         Top             =   3713
         Width           =   1410
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2487;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "目次："
         Height          =   255
         Index           =   1
         Left            =   45
         TabIndex        =   71
         Top             =   570
         Width           =   540
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   2130
         TabIndex        =   69
         Top             =   570
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   0
         Left            =   645
         TabIndex        =   68
         Top             =   570
         Width           =   630
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1111;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "承辦人："
         Height          =   255
         Index           =   8
         Left            =   1275
         TabIndex        =   67
         Top             =   570
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "總收文號："
         Height          =   255
         Index           =   21
         Left            =   45
         TabIndex        =   66
         Top             =   825
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "收文日："
         Height          =   255
         Index           =   20
         Left            =   45
         TabIndex        =   65
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號："
         Height          =   255
         Index           =   19
         Left            =   45
         TabIndex        =   64
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱："
         Height          =   255
         Index           =   18
         Left            =   45
         TabIndex        =   63
         Top             =   1743
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "是否算案件數："
         Height          =   255
         Index           =   17
         Left            =   45
         TabIndex        =   62
         Top             =   2055
         Width           =   1320
      End
      Begin VB.Label Label1 
         Caption         =   "商標種類："
         Height          =   255
         Index           =   16
         Left            =   45
         TabIndex        =   61
         Top             =   2370
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質："
         Height          =   255
         Index           =   15
         Left            =   45
         TabIndex        =   60
         Top             =   2745
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限："
         Height          =   255
         Index           =   14
         Left            =   45
         TabIndex        =   59
         Top             =   3135
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限："
         Height          =   255
         Index           =   13
         Left            =   45
         TabIndex        =   58
         Top             =   3420
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "點數："
         Height          =   255
         Index           =   11
         Left            =   45
         TabIndex        =   57
         Top             =   4013
         Width           =   540
      End
      Begin MSForms.Label lbl1 
         Height          =   495
         Index           =   29
         Left            =   6615
         TabIndex        =   54
         Top             =   7587
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;873"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   225
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y/N)"
         Height          =   255
         Index           =   30
         Left            =   5550
         TabIndex        =   53
         Top             =   1740
         Width           =   525
      End
      Begin MSForms.Label lbl1 
         Height          =   270
         Index           =   6
         Left            =   5220
         TabIndex        =   52
         Top             =   1770
         Visible         =   0   'False
         Width           =   860
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1508;476"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   180
         Index           =   8
         Left            =   5040
         TabIndex        =   51
         Top             =   870
         Visible         =   0   'False
         Width           =   975
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2408;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   260
         Index           =   28
         Left            =   5010
         TabIndex        =   50
         Top             =   4620
         Width           =   1320
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2328;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   495
         Index           =   30
         Left            =   6270
         TabIndex        =   49
         Top             =   7437
         Visible         =   0   'False
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2408;556"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "(N：不算)"
         Height          =   255
         Index           =   32
         Left            =   2205
         TabIndex        =   48
         Top             =   2055
         Width           =   1065
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   47
         Top             =   825
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   5
         Left            =   1080
         TabIndex        =   46
         Top             =   1140
         Width           =   1590
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2805;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   7
         Left            =   1080
         TabIndex        =   45
         Top             =   1440
         Width           =   1830
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3228;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   1080
         TabIndex        =   44
         Top             =   1743
         Width           =   2865
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5054;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   1470
         TabIndex        =   43
         Top             =   2055
         Width           =   600
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1058;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   17
         Left            =   1080
         TabIndex        =   42
         Top             =   3120
         Width           =   1200
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   19
         Left            =   1080
         TabIndex        =   41
         Top             =   3420
         Width           =   1170
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2064;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   600
         TabIndex        =   40
         Top             =   4013
         Width           =   915
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1614;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "齊備日："
         Height          =   255
         Index           =   23
         Left            =   4260
         TabIndex        =   39
         Top             =   798
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "會稿日："
         Height          =   255
         Index           =   24
         Left            =   4260
         TabIndex        =   38
         Top             =   2370
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "是否會稿："
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   27
         Left            =   4035
         TabIndex        =   37
         Top             =   1743
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "發文日："
         Height          =   260
         Index           =   6
         Left            =   4260
         TabIndex        =   36
         Top             =   4010
         Width           =   740
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "取消收文日："
         Height          =   260
         Index           =   29
         Left            =   3620
         TabIndex        =   35
         Top             =   4620
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "承辦備註："
         Height          =   200
         Index           =   31
         Left            =   4100
         TabIndex        =   34
         Top             =   3200
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "請點選""確定""按鈕存檔!!"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   14.5
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   5970
         TabIndex        =   33
         Top             =   30
         Width           =   3225
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         Caption         =   "承辦期限："
         Height          =   255
         Index           =   7
         Left            =   4035
         TabIndex        =   70
         Top             =   525
         Width           =   960
      End
      Begin MSForms.Label Label3 
         Height          =   285
         Left            =   -73890
         TabIndex        =   104
         Top             =   390
         Width           =   2445
         VariousPropertyBits=   27
         Size            =   "4313;494"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   23
      Left            =   720
      MaxLength       =   1
      TabIndex        =   119
      Top             =   300
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   2070
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   2
      Left            =   8220
      TabIndex        =   28
      Top             =   50
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "本月統計(&T)"
      Height          =   375
      Index           =   0
      Left            =   6270
      TabIndex        =   26
      Top             =   50
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Height          =   375
      Index           =   1
      Left            =   7500
      TabIndex        =   27
      Top             =   50
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "核稿語文：   (1.英2.日)"
      Height          =   180
      Index           =   50
      Left            =   0
      TabIndex        =   120
      Top             =   360
      Visible         =   0   'False
      Width           =   1790
   End
   Begin VB.Label Label4 
      Caption         =   "申請國家："
      Enabled         =   0   'False
      Height          =   180
      Left            =   2715
      TabIndex        =   55
      Top             =   495
      Width           =   900
   End
End
Attribute VB_Name = "frm090201_b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/23 改成Form2.0 ; grd1改字型=新細明體-ExtB、grd2改字型=新細明體-ExtB、Combo1、Combo2、Combo6、lbl1(index)、txt1(10)改為txtEP12、txtCP64
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Modify By Sindy 2012/5/10 從frm090201_2移出來
Option Explicit

Public TextOk As Boolean
Public Combo1_String As String
Dim s As Integer, i As Integer, k As Integer
Dim SWPRow As String, SWPRow2 As String, SWPColor As String, SWPColor2 As String
Dim strTemp(0 To 26) As String
Dim Tmp001 As String, Tmp002 As String, Tmp003 As String, Tmp004 As String
Dim SeekTmpBk As String
Dim ChkNoData As Boolean, ChkData As Boolean
Dim Fobj As FileSystemObject
Dim StrGrp090201 As String
Dim Adorecordset99 As New ADODB.Recordset
Dim m_SqlGrpStr1 As String, m_SqlGrpStr2 As String, m_SqlGrpStr3 As String, m_SqlGrpStr4 As String, m_SqlGrpStr5 As String
Dim m_ST03 As String
Dim m_strCP09 As String '總收文號
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String '本所案號
Dim m_CP10 As String
Dim m_CP13 As String

Dim m_CP14 As String
Dim m_EP05ST03 As String 'Add By Sindy 2024/7/30

Dim m_CP31 As String
Dim m_CP43 As String
Dim m_CP44 As String
Dim m_CP112 As String
Dim m_CP159 As String 'Add By Sindy 2020/1/31
Dim m_CP141 As String, m_CP79 As String 'Add By Sindy 2021/9/13
Dim m_CP142 As String, m_CP164 As String 'Add By Sindy 2023/4/21
Dim m_CP60 As String 'Add By Sindy 2021/12/13
Dim strSQL1 As String
Dim strSQL2 As String
Dim StrSQL6 As String
Dim StrSQL61 As String
Dim StrSQL62 As String
Dim StrSQL63 As String
Dim StrSQL64 As String
Dim StrSTM As String
Dim StrSLC As String
Dim StrSHC As String
Dim StrSSP As String
Dim m_Country As String
Dim m_CaseName As String
Dim m_SaleArea As String
Dim m_CuNo As String
Dim m_FieldList() As FIELDITEM
'紀錄 mail 資料，在 trans 後發
Dim skMail() As SeekMails
Dim m_NA03 As String
Dim bolInsert As Boolean, bolUpdate As Boolean, bolDelete As Boolean, bolSelect As Boolean, bolPrint As Boolean
Dim m_CPM05 As String
Dim m_blnClkSure As Boolean '判斷是否按下確定
Dim m_CP149 As String 'Add By Sindy 2012/10/24
Public cmdState As Integer '紀錄作用按鍵
Dim m_CP140 As String
'Added by Lydia 2015/11/12 新增查名單對應
Public Tmpfrm090130 As Form
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限
'Add By Sindy 2018/4/17
Public intBackTab As Integer
Dim dblPrevRow As Double
Dim m_intRow As Integer, m_intCol As Integer
'2018/4/17 END
'Add By Sindy 2018/4/20
Dim m_PP04 As String '預設核稿人
Dim m_PP05 As String '預設判發人
Dim m_PP01 As String '系統別
Dim m_PP03 As String '案件性質或核判分類
Dim m_EP39 As String '核稿完成日
Dim m_CPM28 As String
Dim m_CPM29 As String
Dim m_CP27 As String '發文日
Public m_chkcmdok1 As Boolean '記錄確定鍵是否存檔成功
Public m_Flow As String '欲新增的下一流程
'2018/4/20 END
Dim m_AttachPath As String 'Add By Sindy 2020/1/31
Dim m_CP16 As String 'Add By Sindy 2020/11/17
Dim m_CP163 As String 'Add By Sindy 2020/12/2
Dim strCRA05 As String, m_strFilePath As String 'Add By Sindy 2023/3/28
'Added by Lydia 2023/11/30
Dim m2_CP10 As String '相關總收文號-案件性質
Dim m2_CP10ex As String '相關總收文號-案件性質=>智慧局的指定名稱
Dim m_TM15 As String '審定號
'end 2023/11/30
Dim m_EMPST16 As String, m_EP41 As String 'Add By Sindy 2024/8/13


'本來區塊 1 文件和申請書都會跑，現在只剩文件檔案才跑，因為申請書改跑定搞
Private Sub cmd_Click(Index As Integer)
Dim strTempName As String
Dim nFrm As Form 'Add By Sindy 2018/4/17
Dim ET01 As String, ET03 As String 'Add By Sindy 2020/10/20

On Error GoTo ErrHand

Select Case Index
Case 0, 1 '智權人員補充資料記錄 或 通知補充資料
   'Me.Hide
   '開啟視窗
   If frm090201_b_1.Process(LBL1(3)) Then
      'Add By Sindy 2012/10/24
      If Index = 0 Then
         frm090201_b_1.Frame1.Visible = True
         frm090201_b_1.Frame2.Visible = False
         frm090201_b_1.Caption = "智權人員補充資料記錄作業"
      ElseIf Index = 1 Then
         frm090201_b_1.Frame1.Visible = False
         frm090201_b_1.Frame2.Visible = True
         frm090201_b_1.Caption = "通知智權人員補充資料記錄作業"
         frm090201_b_1.txtText(1).TabIndex = 0
      End If
      '2012/10/24 End
      frm090201_b_1.Show vbModal
      'Add By Sindy 2012/10/24
      If Index = 1 Then
         Call Process(LBL1(3).Caption)
      End If
      '2012/10/24 End
   End If
   Unload frm090201_b_1
   Set frm090201_b_1 = Nothing
   'Me.Show
   
'Modify By Sindy 2025/7/10 歷程上線此功能已無使用,所以mark起來
'Case 2 '承辦單列印
'   Dim CUID As String
'   Dim oTM12 As String
'   Dim oTM15 As String
'   Dim oTM2122 As String
'   'Added by Lydia 2016/05/31 檢查查名單是否全部完成
'   If cmdTSMap.Visible = True Then
'      intI = 1
'      'Modified by Lydia 2018/03/20 用TMQ20判斷是否已刪除明細(+And nvl(tmq20,'N') = 'N') 'Remove by Lydia 2018/03/21 影響速度
'      strSql = "select tmq01,tmq10,st02 from trademarkquery,staff " & _
'                   "where tmq11 is null and tmq10=st01(+) " & _
'                   "and tmq01 in (select tqc03 from tmqcasemap where tqc02='" & Me.LBL1(3).Caption & "') order by 1"
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         RsTemp.MoveFirst
'         Do While Not RsTemp.EOF
'            MsgBox "委查單號:" & RsTemp.Fields("tmq01") & "，查名人:" & RsTemp.Fields("st02") & "，尚未查覆完畢，不可印承辦單!", vbCritical, "委查結果"
'            RsTemp.MoveNext
'         Loop
'         Exit Sub
'      End If
'   End If
'   'end 2016/05/31
'
'   '抓商品類別
'   oTM12 = ""
'   oTM15 = "'"
'   strSql = "select tm09,tm12,tm15," & SQLDate("tm21") & "||'-'||" & SQLDate("tm22") & " from trademark,customer where tm01='" & SystemNumber(LBL1(7), 1) & "' and tm02='" & SystemNumber(LBL1(7), 2) & "' and tm03='" & SystemNumber(LBL1(7), 3) & "' and tm04='" & SystemNumber(LBL1(7), 4) & "' and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) "
'   strSql = strSql & " union select sp73,'',''," & SQLDate("sp20") & "||'-'||" & SQLDate("sp21") & " from servicepractice,customer where sp01='" & SystemNumber(LBL1(7), 1) & "' and sp02='" & SystemNumber(LBL1(7), 2) & "' and sp03='" & SystemNumber(LBL1(7), 3) & "' and sp04='" & SystemNumber(LBL1(7), 4) & "' and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) "
'   CheckOC3
'   AdoRecordSet3.CursorLocation = adUseClient
'   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   CUID = ""
'   If AdoRecordSet3.RecordCount <> 0 Then
'      AdoRecordSet3.MoveFirst
'      CUID = CheckStr(AdoRecordSet3.Fields(0).Value)
'      oTM12 = CheckStr(AdoRecordSet3.Fields(1).Value)
'      oTM15 = CheckStr(AdoRecordSet3.Fields(2).Value)
'      oTM2122 = CheckStr(AdoRecordSet3.Fields(3).Value)
'   End If
'   CheckOC3
'   '若申請國家非台灣，則空白
'   If m_Country = "000" Then
'       frm090201_2_3.txt1(5).Text = "經濟部智慧財產局"
'       '改抓案件國家收費表,預設值原為智慧局同時改為全名
'       strSql = "select cf10 from casefee where cf01='" & SystemNumber(LBL1(7), 1) & "' and cf02='000' and cf03='" & m_CP10 & "'"
'       AdoRecordSet3.CursorLocation = adUseClient
'       AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'       If AdoRecordSet3.RecordCount <> 0 Then
'          If "" & CheckStr(AdoRecordSet3.Fields(0).Value) <> "" Then
'             frm090201_2_3.txt1(5).Text = CheckStr(AdoRecordSet3.Fields(0).Value)
'          End If
'       End If
'       CheckOC3
'       frm090201_2_3.txt1(5).Enabled = False
'   Else
'       frm090201_2_3.txt1(5).Text = ""
'       frm090201_2_3.txt1(5).Enabled = True
'       '林副理說大陸案加預設代理人名稱
'       If m_CP44 <> "" Then
'          If PUB_GetAgentName(SystemNumber(LBL1(7), 1), m_CP44, strTempName) Then
'             frm090201_2_3.txt1(5).Text = strTempName
'          Else
'             frm090201_2_3.txt1(5).Text = ""
'          End If
'          frm090201_2_3.txt1(5).Enabled = False
'          frm090201_2_3.txt1(8).Text = "掛號"  '林副理需求
'       End If
'   End If
'   frm090201_2_3.txt1(6).Text = "北所、" & IIf(m_SaleArea = "1", "", IIf(m_SaleArea = "2", "中所、", IIf(m_SaleArea = "3", "南所、", IIf(m_SaleArea = "4", "高所、", "")))) & "客戶"
'   frm090201_2_3.txt1(6).Enabled = False
'   '加審定號或申請案號
'   If oTM15 = "" Then
'      frm090201_2_3.txt1(0) = "「" & m_CaseName & "」" & LBL1(15) & "(申請案號 " & oTM12 & ")"
'   Else
'      frm090201_2_3.txt1(0) = "「" & m_CaseName & "」" & LBL1(15) & "(註冊號 " & oTM15 & ")"
'   End If
'   frm090201_2_3.txt1(0).Enabled = False
'   frm090201_2_3.txt1(1).Text = CUID         '商品類別
'   frm090201_2_3.txt1(1).Enabled = False
'   frm090201_2_3.Label2.Caption = "商品(服務)類別："
'
'   If m_CP10 = "301" Then
'       If oTM15 = "" Then '申請中    ◎■□『』□■○●
'         frm090201_2_3.txt1(3).Text = "◎申請第 " & oTM12 & " 號『" & m_CaseName & "』□商標" & vbCrLf & _
'                                                          "◎變更事項：" & vbCrLf & _
'                                                          "　□申請人名稱　□代表人或負責人　　　□代理人印鑑" & vbCrLf & _
'                                                          "　□申請人印鑑　□代表人或負責人印鑑　□代理人地址" & vbCrLf & _
'                                                          "　□申請人地址　□代理人"
'       Else
'         frm090201_2_3.txt1(3).Text = "◎註冊第 " & oTM15 & " 號『" & m_CaseName & "』□商標(前服務標章)" & vbCrLf & _
'                                                          "◎變更□商標(標章)權人" & vbCrLf & _
'                                                          "◎變更事項：" & vbCrLf & _
'                                                          "　□申請人中文名稱　□申請人英文名稱　　□申請人印章　　　□申請人中文地址" & vbCrLf & _
'                                                          "　□申請人英文地址　□代表人印章　　　　□代表人中文名稱　□代表人英文名稱" & vbCrLf & _
'                                                          "　□代理人異動：□變更、□新增、□撤銷　　□變更商標/標章名稱"
'       End If
'   ElseIf m_CP10 = "101" Then
'           frm090201_2_3.txt1(3).Text = "◎中文：" & vbCrLf & _
'                                                          "◎外文：" & vbCrLf & _
'                                                          "　字義：" & vbCrLf & _
'                                                          "◎圖形說明："
'           frm090201_2_3.txt1(1).Enabled = True
'           frm090201_2_3.txt1(0).Enabled = True
'   ElseIf m_CP10 = "102" Then
'         frm090201_2_3.txt1(3).Text = "◎註冊第 " & oTM15 & " 號『" & m_CaseName & "』□商標　□商標(前服務標章)" & vbCrLf & _
'                                                          "◎原註冊期間：" & oTM2122 & vbCrLf & _
'                                                          "◎變更事項：" & vbCrLf & _
'                                                          "　□變更商標/標章名稱：" & vbCrLf & _
'                                                          "　□防護商標/標章變更為商標" & vbCrLf & _
'                                                          "　□代理人異動：□變更、□新增、□撤銷" & vbCrLf & _
'                                                          "◎系統註記變更事項：" & vbCrLf & _
'                                                          "　□申請人中文名稱　□申請人英文名稱　□申請人印鑑　□申請人地址" & vbCrLf & _
'                                                          "　□代表人中文名稱　□代表人英文名稱　□代表人印鑑"
'   Else
'      frm090201_2_3.txt1(3).Text = ""
'   End If
'   frm090201_2_3.oStrA02 = LBL1(7)
'
'   If LBL1(17) <> "" Then
'      frm090201_2_3.oStrA05 = "   " & Mid(Replace(LBL1(17), "/", ""), 1, Len(Replace(LBL1(17), "/", "")) - 4) & "年  " & Left(Right(Replace(LBL1(17), "/", ""), 4), 2) & "月  " & Right(Replace(LBL1(17), "/", ""), 2) & "日"
'   Else
'      frm090201_2_3.oStrA05 = "     年    月    日"
'   End If
'   If LBL1(19) <> "" Then
'      frm090201_2_3.oStrA06 = "   " & Mid(Replace(LBL1(19), "/", ""), 1, Len(Replace(LBL1(19), "/", "")) - 4) & "年  " & Left(Right(Replace(LBL1(19), "/", ""), 4), 2) & "月  " & Right(Replace(LBL1(19), "/", ""), 2) & "日"
'   Else
'      frm090201_2_3.oStrA06 = "     年    月    日"
'   End If
'   frm090201_2_3.oStrA08 = LBL1(3) '收文號
'   frm090201_2_3.oStrA09 = LBL1(23)
'   frm090201_2_3.txt1(4).Text = "附委任狀正本" & vbCrLf & "附委任狀影本" & vbCrLf & "正本參     卷"
'   frm090201_2_3.oStrA10 = ""
'   frm090201_2_3.oStrA11 = ""
'   frm090201_2_3.oStrA12 = ""
'   frm090201_2_3.oStrA13 = ""
'   frm090201_2_3.oStrA14 = PUB_GetST07(m_CP14)
'
'   '商標
'   'If frm090201_b.txt1(1) = "" Then '是否會稿為空白時
'   'Modify By Sindy 2012/6/28 要會稿
'   'If frm090201_b.txt1(1) = "Y" Then '要會稿
'   '2012/6/28 End
'   'Modify By Sindy 2012/9/28
'   If frm090201_b.txt1(1) <> "N" Then '空白和Y均都是要會稿
'   '2012/9/28 End
'      frm090201_2_3.txt1(9) = ""
'      Select Case PUB_GetST06(m_CP13)
'         Case "2"
'            frm090201_2_3.txt1(9) = "中所　"
'         Case "3"
'            frm090201_2_3.txt1(9) = "南所　"
'         Case "4"
'            frm090201_2_3.txt1(9) = "高所　"
'      End Select
'      frm090201_2_3.txt1(9) = frm090201_2_3.txt1(9) & frm090201_b.LBL1(21)
'      frm090201_2_3.oStrA15 = frm090201_2_3.txt1(9)
'   End If
'   frm090201_2_3.Show vbModal
   
'C類來函且未發文才可用撰寫信函按鈕
Case 3  '撰寫信函
   Call Forms(0).SetTmpfrm090401 'Add By Sindy 2015/7/15
   strExc(1) = m_CP10
   'Modify By Sindy 2015/7/15
   'With frm090401
   With Tmpfrm090401
   '2015/7/15 END
      .Hide
      .Text1 = SystemNumber(LBL1(7).Caption, 1)
      .Text1_Validate True 'Add By Sindy 2020/2/11 TS-001774(A5026)
      .Text2 = SystemNumber(LBL1(7).Caption, 2)
      .Text3 = SystemNumber(LBL1(7).Caption, 3)
      .Text4 = SystemNumber(LBL1(7).Caption, 4)
      .Option1(0).Value = True '點選中文
      .Option3.Value = True '讀取案件資料
      .Option3.Value = True '點選申請人
      For intI = 0 To .Combo8.ListCount
         'Modified by Morgan 2017/1/23 strExc(1)會被使用而改變
         'If InStr(.Combo8.List(intI), strExc(1)) > 0 Then
         If InStr(.Combo8.List(intI), m_CP10) > 0 Then
         'end 2017/1/23
            .Combo8.ListIndex = intI
            Exit For
         End If
      Next
      'Add By Sindy 2025/8/15
      If m_CP01 = "T" And m_Country = "000" And m_CP10 = "727" And .Combo8.ListCount > 1 Then
         .Combo8.ListIndex = .Combo8.ListCount - 1
      End If
      '2025/8/15 END
      .Command1.Value = True
      .Command2.Value = True
   End With
   Set Tmpfrm090401 = Nothing 'Add By Sindy 2015/7/15
   
'Add By Sindy 2013/6/17 申請書
Case 4
   'Add By Sindy 2018/8/13
   If InStr(cmd(4).Caption, "指示信") > 0 Then
      frm020107_1.Option1(1).Value = True
      frm020107_1.Text5 = LBL1(3).Caption
      'Modify By Sindy 2021/1/21
      'frm020107_1.Tag = Me.Name '此tag有用在別處
      frm020107_1.cmdOK(2).Tag = Me.Name '離開按鈕
      '2021/1/21 END
      frm020107_1.cmdQuery.Value = 1
   'Add By Sindy 2018/11/30 繳費單
   ElseIf InStr(cmd(4).Caption, "繳費單") > 0 Then
      Call PUB_PrintTFeeForm(m_CP01, m_CP02, m_CP03, m_CP04)
      '2018/11/30 END
   Else
   '2018/8/13 END
      
      'Modified by Lydia 2019/07/05 +申請101
      If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "103" Then  '紙本送件申請書
         'Modified by Lydia 2019/03/28  傳入收文號
         'Call PUB_GetApplBook(lbl1(7), m_CP10)
         'Added by Lydia 2019/07/05 電子送件申請書
         'Mark by Lydia 2019/07/31 因為非申請案要到跑完歷程才能確定是否電子送件,所以分成兩個按鈕
         'If lblEApp.Visible = True Then
         '     Call GetApplBook_T(lbl1(7), lbl1(3).Caption, m_CP10)
         'Else
         'end 2019/07/05
              Call PUB_GetApplBook(LBL1(7), m_CP10, , , , , , LBL1(3).Caption)
         'End If
      ElseIf m_CP10 = "301" Or m_CP10 = "501" Then '變更301,移轉501
         frm090201_b_3.bolCP118 = False 'Added by Lydia 2020/10/07 非電子送件
         frm090201_b_3.m_CP10 = m_CP10
         frm090201_b_3.LBL1(3).Caption = Me.LBL1(3).Caption
         frm090201_b_3.LBL1(7).Caption = Me.LBL1(7).Caption
         frm090201_b_3.LBL1(9).Caption = Me.LBL1(9).Caption
         frm090201_b_3.LBL1(15).Caption = Me.LBL1(15).Caption
         frm090201_b_3.Show vbModal
      'Add By Sindy 2020/10/20 214.陳述聲明
      ElseIf m_CP10 = "214" Then
         ET01 = "90"
         ET03 = "01"
         strLetterDate = strSrvDate(2)
         'If StartLetter2(tm, m_CaseNo, ET01, ET03, pCP09, "2") = False Then Exit Sub
         NowPrint LBL1(3).Caption, ET01, ET03, True, strUserNum, , , , , , , False, , False
      '2020/10/20 END
      End If
   End If
'2013/6/17 End
   
'Added by Lydia 2019/07/31
Case 6 '電子送件-申請書
   'Add By Sindy 2020/9/28 + 725.代辦退費
   'Modified by Lydia 2022/04/19  (比照FCT案): 請於「延展」案，增加「是否要變更申請人選項」 (因延展可同時變更申請人)，俾申請書帶出正確之申請人資料。
   'If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "103" Or m_CP10 = "725" Then
   If m_CP10 = "101" Or m_CP10 = "103" Or m_CP10 = "725" Then
      Call GetApplBook_T(LBL1(7), LBL1(3).Caption, m_CP10)
   'Added by Lydia 2020/10/07 +變更301, 移轉501,授權502
   'Modified by Lydia 2022/04/19 +延展102
   ElseIf m_CP10 = "301" Or m_CP10 = "501" Or m_CP10 = "502" Or m_CP10 = "102" Then
         frm090201_b_3.bolCP118 = True '電子送件
         frm090201_b_3.m_CP10 = m_CP10
         frm090201_b_3.LBL1(3).Caption = Me.LBL1(3).Caption
         frm090201_b_3.LBL1(7).Caption = Me.LBL1(7).Caption
         frm090201_b_3.LBL1(9).Caption = Me.LBL1(9).Caption
         frm090201_b_3.LBL1(15).Caption = Me.LBL1(15).Caption
         frm090201_b_3.Show vbModal
   'end 2020/10/07
   'Added by Lydia 2020/10/07 電子送件-補正申請書(A、B類收文)：其他沒有設定的性質，ex.303延期、201補正、202申請意見書、208補優先權證明、706其他
   Else
        Call GetApplBook_T(LBL1(7), LBL1(3).Caption, m_CP10)
   End If
   
'Add By Sindy 2018/4/17
Case 5 '承辦歷程
      
      '重新檢查欄位有效性
      If TxtValidate = True Then
         
         If SetColTag(False) = False Then
            'Modify By Sindy 2022/4/25 針對有註記「收款後送件」的台灣商標案件，開放承辦人於案件先行作業後，可自行輸入「完稿日」。
            '第一次輸入完稿日
            If Me.LblFee.Tag = "尚待收款" And Val(txt1(3).Tag) = 0 And Val(txt1(3).Text) > 0 Then
               cmdOK(1).Enabled = False
               Call cmdok_Click(1)
               cmdOK(1).Enabled = True
               Exit Sub
            Else
            '2022/4/26 END
               cmdOK(1).Enabled = False
               Call cmdok_Click(1)
               cmdOK(1).Enabled = True
               If m_chkcmdok1 = False Then Exit Sub
            End If
         Else
            Call Process(LBL1(3)) '要重新查詢資料 Add By Sindy 2018/10/4
         End If
         
'         '檢查表單是否已開啟，若是，則關閉
'         For Each nFrm In Forms
'            If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'               Unload frm090202_2
'               Exit For
'            End If
'         Next
         If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
         intBackTab = 1
         frm090202_2.Hide
         frm090202_2.m_EEP01 = LBL1(3) '總收文號
         frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) '案件流程所屬人員
         frm090202_2.intReceiveKind = 0 '0.承辦人工作進度
         frm090202_2.SetParent Me
         'Add By Sindy 2018/8/15 委查結果
         If cmdTSMap.Visible = True Then
            cmdTSMap.Tag = "Y"
         Else
            cmdTSMap.Tag = ""
         End If
         '2018/8/15 END
         If frm090202_2.QueryData = True Then
            frm090202_2.Show
            Me.Hide
         End If
      End If
'2018/4/17 END

'Add By Sindy 2020/3/17
Case 7 '原始檔暫存區
   'Call PUB_ChkFormIsClose("frm100101_M")
   frm100101_M.m_strKey = LBL1(3).Caption '總收文號
   frm100101_M.SetParent Me
   If frm100101_M.QueryData = True Then
      frm100101_M.Show
      Me.Hide
   End If
'2020/3/17 END
Case Else
End Select
Exit Sub

ErrHand:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

Function checkCP49() As Boolean
checkCP49 = False
'條款
If Trim(txt1(11).Text) = "" Then
    checkCP49 = True
    Exit Function
End If
Dim tmpCp49Arr As Variant
Dim intCp49 As Integer
tmpCp49Arr = Split(txt1(11), ",")
For intCp49 = 0 To UBound(tmpCp49Arr)
    '條款不小於3碼
    'C類來函不檢查
    If Len(Trim(tmpCp49Arr(intCp49))) < 3 And Len(Trim(tmpCp49Arr(intCp49))) <> 0 And LBL1(3) < "C" Then
        s = MsgBox(tmpCp49Arr(intCp49) & "，小於 3 碼！", , "條款輸入錯誤！")
        checkCP49 = False
        txt1(11).SetFocus
        Exit Function
    End If
    If Len(Trim(tmpCp49Arr(intCp49))) <> 0 Then
'        '只抓前三碼檢查
'        strSql = "select * from law where lw01='" & Left(tmpCp49Arr(intCp49), 3) & "' "
        'Modify By Sindy 2012/7/12
        '只抓前三碼或前四碼檢查
        strSql = "select * from law where lw01='" & Left(tmpCp49Arr(intCp49), 3) & "' or lw01='" & Left(tmpCp49Arr(intCp49), 4) & "' "
        '2012/7/12 End
        Set Adorecordset99 = New ADODB.Recordset
        Adorecordset99.CursorLocation = adUseClient
        Adorecordset99.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If Adorecordset99.RecordCount = 0 Then
            s = MsgBox("沒有 " & tmpCp49Arr(intCp49) & " 條款！", , "條款輸入錯誤！")
            checkCP49 = False
            txt1(11).SetFocus
            Exit Function
        End If
    End If
Next intCp49
checkCP49 = True
End Function

'Add By Sindy 2018/10/1
'Modified by Lydia 2021/12/23 ComboBox=> Object
Private Sub SetComboData(mCombo As Object)
Dim blnMatch As Boolean
Dim ii As Integer
   
   Call GetPrjSalesNM(Trim(Left("" & Combo1.Text, 6)), , , "st16", m_EMPST16) 'Modify By Sindy 2024/8/13
   
   '代主任以上的人員
   'Modify By Sindy 2024/8/7
   If Left(PUB_GetST03(Trim(Left("" & Combo1.Text, 6))), 1) = "F" Then
      '(ST03>='F10' and ST03<='F11')
      strExc(0) = "SELECT st01||' ==> '||st02 FROM STAFF WHERE (ST93='" & Pub_StrUserSt93 & "' AND ST04='1' AND ST20<='52'"
'      If m_EMPST16 <> "" Then
'         strExc(0) = strExc(0) & " AND ST16='" & m_EMPST16 & "'"
'      End If
      strExc(0) = strExc(0) & ")"
      If m_EMPST16 = "6" Then 'CF
         strExc(0) = strExc(0) & " OR ST01='" & Pub_GetSpecMan("CFT61") & "' OR ST01='" & Pub_GetSpecMan("CFT62") & "'"
      End If
      strExc(0) = strExc(0) & " ORDER BY ST01 asc"
   Else
   '2024/8/7 END
      'Modify By Sindy 2021/10/4 + 林律師和沈佳穎
      'Modify By Sindy 2023/7/11 取消 or ST01='76012'
      strExc(0) = "SELECT st01||' ==> '||st02 FROM STAFF WHERE ((ST03>='P20' and ST03<='P21') AND ST04='1' AND ST20<='52') or ST01='98003' or ST01='96003' or ST01='98020' ORDER BY ST01 asc"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      .MoveFirst
      Do While .EOF = False
         blnMatch = False
         For ii = 0 To mCombo.ListCount - 1
             blnMatch = False
             If Trim(Left(mCombo.List(ii), 6)) = Left(.Fields(0), 5) Then
                 mCombo.ListIndex = ii
                 blnMatch = True
                 Exit For
             End If
         Next ii
         If blnMatch = False Then
            intI = mCombo.ListCount
            mCombo.AddItem "" & .Fields(0), intI
         End If
         .MoveNext
      Loop
      End With
   End If
   blnMatch = False
   For ii = 0 To mCombo.ListCount - 1
      If Trim(Left(mCombo.List(ii), 6)) = mCombo.Tag Then
          mCombo.ListIndex = ii
          blnMatch = True
          Exit For
      End If
   Next ii
   If blnMatch = False Then mCombo.ListIndex = 0
   mCombo.Tag = mCombo.Text
End Sub

Sub Process(strText As String)
Dim stVTB As String
Dim oLbl As Object
Dim oTxt1 As TextBox
Dim strCompDate As String 'Add By Sindy 2012/10/24
'Add By Sindy 2018/4/20
Dim strRefEEP02 As String
'2018/4/20 END
Dim tmpBol As Boolean 'Added by Lydia 2019/05/02
Dim rsTmp As New ADODB.Recordset
Dim objText As Object
Dim IntTemp1 As Long, IntTemp2 As Long
Dim tmpArr As Variant 'Added by Lydia 2023/11/30

   Me.Enabled = False
   Chk1.Value = vbUnchecked
   
   'Modify By Sindy 2015/9/10 + ,CP140
   'Modified by Lydia 2018/12/10 + CP143
   'Modify By Sindy 2015/9/10 + ,CP159
   'Modify By Sindy 2021/9/13 + ,cp141,cp79
   'Modify By Sindy 2021/12/13 + ,cp60
   'Modified by Morgan 2022/12/15 +TM136
   'Modify By Sindy 2023/2/15 + ,cp85
   'Modify By Sindy 2023/4/21 + ,cp142
   'Modified by Lydia 2023/11/30 +TM15
   'Modify By Sindy 2024/1/15 + ,cp164
   'Modify By Sindy 2024/6/12 + ,tm72,tm137,tm138
   stVTB = " SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112" & _
            ",NVL(TM05,NVL(TM06,TM07)) C10,decode(tm10,'000',ptm03,ptm04) C14,'' C26,'' C28,TM29,cp49 as C33,tm10 as m_country,tm23 as cuno,CP43,CP140,CP118,CP143,cp159,cp16,cp163,cp141,cp79,cp60,TM136,cp85,cp142,tm15,cp164,tm72,tm137,tm138" & _
            " FROM CASEPROGRESS,TRADEMARK,PATENTTRADEMARKMAP WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr2 & ") AND TM01(+)=CP01 AND TM02(+)=CP02 AND TM03(+)=CP03 AND TM04(+)=CP04 AND PTM01(+)='2' AND PTM02(+)=TM08"
   stVTB = stVTB + " UNION all SELECT CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP13,CP14,CP15,CP18,CP26,CP27,CP31,cp44,CP48,CP57,cp64,CP97,CP98,CP99,CP106,cp107,cp111,cp112" & _
            ",NVL(SP05,NVL(SP06,SP07)) C10,'' C14,'' C26,'' C28,SP15,'*' C33,sp09 as m_country,sp08 as cuno,CP43,CP140,CP118,CP143,cp159,cp16,cp163,cp141,cp79,cp60,'' TM136,cp85,cp142,'' as tm15,cp164,'' tm72,'' tm137,'' tm138" & _
            " FROM CASEPROGRESS,SERVICEPRACTICE WHERE CP09='" & strText & "' and cp01 in (" & m_SqlGrpStr5 & ") AND SP01(+)=CP01 AND SP02(+)=CP02 AND SP03(+)=CP03 AND SP04(+)=CP04"
   'Modify By Sindy 2015/9/10 +,CP140
   'Modified by Lydia 2018/12/10 +CP143
   'Modify By Sindy 2019/8/16 pp01(+)=cp01 => pp01(+)='T'
   'Modify By Sindy 2015/9/10 +,CP159
   'Modify By Sindy 2021/9/13 + ,cp141,cp79
   'Modify By Sindy 2021/12/13 + ,cp60
   'Modify By Sindy 2023/2/15 + ,cp85
   'Modify By Sindy 2023/4/21 + ,cp142
   'Modified by Lydia 2023/11/30 +TM15
   'Modify By Sindy 2024/1/15 + ,cp164
   'Modify By Sindy 2024/6/12 + ,tm72,tm137,tm138
   strSql = "SELECT EP01,S1.ST02 C2,sqldateT(CP48) C3,CP09,EP13,sqldateT(cp05) C6,EP34,CP01||'-'||CP02||'-'||CP03||'-'||CP04 C8" & _
      ",EP06,C10,EP09,CP26,EP07,C14,EP04,decode(na01,'000',cpm03,cpm04) C16,EP03,sqldateT(CP06) C18,EP08,sqldateT(CP07) C20,CP27" & _
      ",S5.ST02 C22,EP11,CP18,EP12,C26,Nvl(EP35,0) C27,C28,sqldateT(CP57) C29,CP10,CP15,'' PA57,C33,EP27,EP31,cp13,ep05,m_country,cp31,S5.st06 as Area,cp107,cp97,nvl(cp98,0) as cp98,cp99" & _
      ",cp106,cuno,cp111,cp112,ep28,ep32,ep33,na03,cp64,cpm05,cp44,ibf01,S3.ST02 EP04N,pp04,s6.st02 pp04N,s2.st02 EP13N,s4.st02 EP03N" & _
      ",NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) CuName,CP43,CP140,CP118,CP143,EP38,EP39,pp05,EP40,cpm28,cpm29,cpm23,cp159,cp16,cp163,cp141,cp79,cp60,TM136,cp85,cp142,tm15,cp164,tm72,tm137,tm138,EP41" & _
      " from (" & stVTB & ") X,ENGINEERPROGRESS,CASEPROPERTYMAP,nation" & _
      ",STAFF S1,STAFF S2,STAFF S3,STAFF S4,STAFF S5,customer,imgbytefile,promoterproofreader,staff S6" & _
      " where EP02(+)=CP09 and cpm01(+)=CP01 and cpm02(+)=CP10 AND na01(+)=m_country" & _
      " AND S1.ST01(+)=EP05 AND S2.ST01(+)=EP13 AND S3.ST01(+)=EP04 AND S4.ST01(+)=EP03 AND S5.ST01(+)=CP13" & _
      " and cu01(+)=substr(cuno,1,8) and cu02(+)=substr(cuno,9) and pp01(+)='T' and pp02(+)=EP05 and pp03(+)=cp10 and s6.st01(+)=pp04" & _
      " and ibf01(+)=cp01 and ibf02(+)=cp02 and ibf03(+)=cp03 and ibf04(+)=cp04 and ibf05(+)='1'"
   CheckOC
   With adoRecordset
   .CursorLocation = adUseClient
   .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   
   '***** 清除欄位值 *****
   m_CP163 = "" 'Add By Sindy 2020/12/2
   m_CPM28 = ""
   m_CPM29 = ""
   m_EP39 = "" '核稿完成日
   m_CP01 = "": m_CP02 = "": m_CP03 = "": m_CP04 = ""
   m_CP159 = "" 'Add By Sindy 2020/1/31
   'Add By Sindy 2021/9/13
   m_CP141 = "": m_CP79 = "": m_CP60 = "": m_CP142 = "": m_CP164 = ""
   Me.LblFee.Caption = "": Me.LblFee.Tag = "" 'Add By Sindy 2022/4/25
   '2021/9/13 END
   lblCertType.Caption = "" 'Added by Morgan 2022/12/15
   m_TM15 = "" 'Added by Lydia 2023/11/30
   For Each oLbl In LBL1
      oLbl.Caption = ""
   Next
   Me.lblClose.Caption = ""
   For Each oTxt1 In txt1
      oTxt1.Text = ""
   Next
   txtCP64.Text = ""
   Combo2.Clear: Combo2.Tag = ""
   Combo6.Clear: Combo6.Tag = ""
   'Add By Sindy 2024/12/5
   Combo4.Clear: Combo4.Tag = ""
   Me.Combo4.Enabled = True '外文核稿人
   '2024/12/5 END
   '***** 清除欄位值 END *****
   
   If .RecordCount <> 0 And .RecordCount > 0 Then
      .MoveFirst
      m_CP163 = "" & .Fields("CP163") 'Add By Sindy 2020/12/2
      m_CP43 = "" & .Fields("CP43")
      m_CP01 = SystemNumber(Trim(.Fields("C8")), 1)
      m_CP02 = SystemNumber(Trim(.Fields("C8")), 2)
      m_CP03 = SystemNumber(Trim(.Fields("C8")), 3)
      m_CP04 = SystemNumber(Trim(.Fields("C8")), 4)
      
      'Add By Sindy 2018/4/20
      m_EP39 = "" & .Fields("EP39")
      m_CPM28 = "" & .Fields("CPM28")
      m_CPM29 = "" & .Fields("CPM29")
      m_CP159 = "" & .Fields("CP159") 'Add By Sindy 2020/1/31
      'Add By Sindy 2021/9/13
      m_CP141 = "" & .Fields("CP141"): m_CP79 = "" & .Fields("CP79")
      m_CP142 = "" & .Fields("CP142") 'Add By Sindy 2023/4/21
      m_CP164 = "" & .Fields("CP164") 'Add By Sindy 2024/1/15
      m_CP60 = "" & .Fields("CP60") 'Add By Sindy 2021/12/13
      m_TM15 = "" & .Fields("TM15") 'Added by Lydia 2023/11/30
      m_EP41 = "" & .Fields("EP41") 'Add By Sindy 2024/12/5
      
      If m_CP141 = "2" Then '註記收款後送件的案件
         'Add By Sindy 2021/12/13 國內收據才是判斷CP79；
         '國外請款單要抓acc1k0之a1k29，請參考共同查詢frm100101_2之收回
         If Left(m_CP60, 1) = "X" Then
            IntTemp1 = 0: IntTemp2 = 0
            strSql = "select A1k11,0,'','',decode(a1k29,'Y',a1k11,nvl(A1K30,0)),0,A1K25 FROM ACC1K0 WHERE A1K01='" & m_CP60 & "'"
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
               If Not IsNull(adoRecordset1.Fields(0)) Then
                  IntTemp1 = IntTemp1 + adoRecordset1.Fields(0) '台幣金額
               End If
               If Not IsNull(adoRecordset1.Fields(4)) Then
                  IntTemp2 = IntTemp2 + adoRecordset1.Fields(4) 'decode(a1k29,'Y',a1k11,nvl(A1K30,0)) 已收金額
               End If
               If IntTemp1 = IntTemp2 Then
                  Me.LblFee.Caption = "已收款可送件"
               Else
                  Me.LblFee.Caption = "尚待收款"
               End If
            End If
         Else
         '2021/12/13 END
            If Val(m_CP79) = 0 Then '未收金額=0
               Me.LblFee.Caption = "已收款可送件"
            Else
               Me.LblFee.Caption = "尚待收款"
            End If
         End If
      'Add By Sindy 2023/4/21
      ElseIf m_CP141 = "3" Then '有指定日期送件
         Me.LblFee.Caption = "指定" & ChangeWStringToTDateString(m_CP142) & _
                             IIf(m_CP164 = "1", "當天", IIf(m_CP164 = "2", "之前", IIf(m_CP164 = "3", "之後", ""))) & "送件"
      End If
      '2021/9/13 END
      
      '電子送件
      If Not IsNull(.Fields("CP118")) Then
         'Modify By Sindy 2023/2/15 商標要增加判斷有承辦人發文日時,才要顯示電子送件
         'lblEApp.Visible = True
         If Val("" & .Fields("CP85")) > 0 Then
            lblEApp.Visible = True
         Else
            lblEApp.Visible = False
         End If
         '2023/2/15 END
      Else
         lblEApp.Visible = False
      End If
      '2018/4/20 END
      
      'Added by Lydia 2018/12/10 查名齊備日
      textCP143.Tag = ChangeWStringToTString(CheckStr("" & .Fields("CP143")))
      textCP143.Text = textCP143.Tag

      For i = 0 To 29
         'Modify by Morgan 2008/10/13 原來值由lablel改為放tag或text
         'Add By Sindy 2024/12/6
         '外文核稿人(EP03)
         If i = 16 Then
            txt1(5).Text = CheckStr(.Fields(i))
         '2024/12/6 END
         '會稿日
         ElseIf i = 12 Then
            txt1(4).Text = ChangeWStringToTString(CheckStr(.Fields(i)))
         '會稿完成日
         ElseIf i = 18 Then
            txt1(7).Text = ChangeWStringToTString(CheckStr(.Fields(i)))
         '發文日
         ElseIf i = 20 Then
            txt1(8).Text = ChangeWStringToTString(CheckStr(.Fields(i)))
         '是否通知客戶
         ElseIf i = 22 Then
            txt1(9).Text = CheckStr(.Fields(i))
         '承辦備註
         ElseIf i = 24 Then
            txtEP12.Text = CheckStr(.Fields(i))
         '承辦期限
         ElseIf i = 2 Then
            txt1(12).Text = ChangeTDateStringToTString(CheckStr(.Fields(i)))
         '案件性質
         ElseIf i = 15 Then
            If Not IsNull(.Fields("CP43")) Then '有相關總收文號
               LBL1(i) = CheckStr(.Fields(i)) & PUB_GetRelateCasePropertyName(strText, "1")
            Else
               LBL1(i) = CheckStr(.Fields(i))
            End If
         Else
            If i <> 4 And i <> 14 And i <> 16 And i <> 18 And i <> 26 And i <> 25 And i <> 27 Then
               LBL1(i) = CheckStr(.Fields(i))
            End If
         End If
      Next i
      '外文核完日
      txt1(19) = ChangeWStringToTString(CheckStr(.Fields("EP33")))
      
      m_CP13 = "" & .Fields("cp13").Value '智權人員
      m_CP14 = "" & .Fields("ep05").Value
      m_EP05ST03 = PUB_GetST03(m_CP14) 'Add By Sindy 2024/7/30
      
      m_CP10 = "" & .Fields("cp10").Value
      m_CP16 = "" & .Fields("cp16").Value 'Add By Sindy 2020/11/17
      
      m_NA03 = "" & .Fields("NA03").Value
      
      m_CPM05 = "" & .Fields("cpm05")
      m_CP112 = "" & .Fields("cp112")
     
      m_CP44 = "" & .Fields("cp44")
      m_CuNo = "" & .Fields("CuName")
      
      '進度備註
      txtCP64 = CheckStr(.Fields("cp64"))
      
      '指定會稿日
      txt1(18) = ChangeWStringToTString(CheckStr(.Fields("ep28")))
      
      m_Country = "" & .Fields("m_country").Value
      m_CP31 = "" & .Fields("cp31").Value
      '案件名稱
      m_CaseName = "" & .Fields(9).Value
      m_SaleArea = "" & .Fields("Area").Value
      
      '收文號
      m_strCP09 = Me.LBL1(3).Caption
      If Len(Trim(CheckStr(.Fields(20)))) <> 0 Then
         m_CP27 = .Fields(20) 'Add By Sindy 2018/4/25 發文日
      Else
         m_CP27 = "" 'Add By Sindy 2018/4/25 發文日
      End If
      
      m_CP140 = "" & .Fields("CP140").Value '電子表單單號
      'Add By Sindy 2023/3/28
      cmdOK(5).Enabled = False
      strExc(0) = "select CRL01,CRA05 from ConsultRecordList,ConsultRecApp where CRL01='" & m_CP140 & "'" & _
                  " and CRL01=CRA01 and CRA02=1"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If InStr("'" & strT000Sale1CPMList & "'", m_CP10) > 0 And m_CP01 = "T" And m_Country = "000" Then
            strCRA05 = RsTemp.Fields("CRA05") '客戶編號
            m_strFilePath = strTApp1CasePath & "\" & strCRA05
            '檢查資料夾是否存在
            'If Dir(m_strFilePath, vbDirectory) <> "" Then
            If Dir(m_strFilePath & "\*.*") <> "" Then
               cmdOK(5).Enabled = True
            ElseIf Dir(m_strFilePath, vbDirectory) <> "" Then
               If Dir(m_strFilePath & "\*.*") = "" Then '確保無電子檔
                  RmDir m_strFilePath
               End If
            End If
         End If
      End If
      '2023/3/28 END
      
      cmd(6).Enabled = True 'Added by Lydia 2025/05/29
      'Added by Lydia 2023/11/30 自請撤回(306):分申請案和非申請案
      'Modified by Lydia 2024/12/18 非電子送件不檢查
      If m_CP10 = "306" And m_CP43 <> "" Then
         'Added by Lydia 2024/12/18 查據原請作單「只要是FCT案內商承辦，請於內商分案時自動改為紙本送件。」--紙本送件不檢
         If lblEApp.Visible = False Then
            'cmd(6).Enabled = False 'Mark by Lydia 2025/05/29
         Else
         'end 2024/12/18
            strExc(0) = "select cp10 from caseprogress where cp09='" & m_CP43 & "' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m2_CP10 = "" & RsTemp.Fields("cp10")
               If m2_CP10 = "101" Then
                  m2_CP10ex = "註冊申請案自請"
               Else
                  '因為智慧局名稱有部份與CPM03不同，改成指定名稱;延展(102)、補證(103) 、註冊前變更(301)、註冊變更(301)、英證(304)、註冊前分割(308)、註冊後分割(308)、商品減縮(313)、
                                                                 '移轉(501)、授權(502)、再授權(504)、異議(601)、評定(602)、廢止(605)、代辦退費(725)、設定質權(506)
                  strExc(1) = "延展(102)、補證(103)、註冊前變更(3010)、註冊變更(3011)、英證(304)、" & _
                              "註冊前分割(3080)、註冊後分割(3081)、商品減縮(313)、移轉(501)、授權(502)、" & _
                              "再授權(504)、異議案(601)、評定案(602)、廢止案(605)、退費(725)、" & _
                              "質權(506)"
                  tmpArr = Empty
                  tmpArr = Split(strExc(1), "、")
                  For intI = 0 To UBound(tmpArr)
                     If Trim(tmpArr(intI)) <> "" And m2_CP10ex = "" Then
                        If InStr(Trim(tmpArr(intI)), m2_CP10 & IIf(InStr("301,308", m2_CP10) > 0, IIf(m_TM15 <> "", "1", "0"), "")) > 0 Then
                           m2_CP10ex = Mid(Trim(tmpArr(intI)), 1, InStr(Trim(tmpArr(intI)), "(") - 1)
                           Exit For
                        End If
                     End If
                  Next intI
               End If
               If m2_CP10ex = "" Then
                  MsgBox "目前無【" & LBL1(15) & PUB_GetRelateCasePropertyName(LBL1(3), "1") & "】的電子送件申請書！", vbCritical + vbOKOnly, "自請撤回申請書"
                  'Modified by Lydia 2024/12/18 debug
                  'cmdOK(0).Enabled = False
                  cmd(6).Enabled = False
               End If
            End If
         End If 'Added by Lydia 2024/12/18
      End If
      'end 2023/11/30
      
      'Modify By Sindy 2019/9/27 竹平反應顯示出來,反而會覺得資料都在接洽單裡,但其實應該還要進卷宗區看回覆單
'      'Add By Sindy 2015/9/10
'      If m_CP140 = "" Then
'         cmdOK(3).Visible = False
'      Else
'         cmdOK(3).Visible = True
'      End If
'      '2015/9/10 END
      
      'Add By Sindy 2024/7/30
      cmdDataMail.Visible = False
      If Left(m_EP05ST03, 2) = "P2" Then
      '2024/7/30 END
         'Add By Sindy 2020/11/17 寄發指示信:非台灣案201.補正
         '612.補充理由,沒有費用
         If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
            If m_Country <> "000" And _
               (m_CP10 = "201" Or (m_CP10 = "612" And Val(m_CP16) = 0)) Then
               cmdDataMail.Visible = True
            End If
         End If
         '2020/11/17 END
      End If
      
      'Added by Morgan 2022/12/15
      If m_CP01 = "T" And m_Country = "000" And m_CP10 = "717" Then
         If .Fields("TM136").Value = "1" Then
            lblCertType = "電子註冊證"
         End If
      End If
      'end 2022/12/15
      
      If IsNull(.Fields(31).Value) <> 0 Then
          Me.lblClose.Caption = ""
      Else
          Me.lblClose.Caption = "已閉卷"
      End If
      
      '91.08.14 增加若是商標案則加秀條款欄位* 就是商標   nick  start
      'If CheckStr(.Fields(32).Value) <> "*" Then
      If CheckStr(.Fields("C33").Value) <> "*" Then
          txt1(11).Text = CheckStr(.Fields("C33"))
          txt1(11).Visible = True 'Add By Sindy 2024/7/31
          SSTab2.TabVisible(0) = True
      Else
          txt1(11).Text = ""
          txt1(11).Visible = False 'Add By Sindy 2024/7/31
          SSTab2.TabVisible(0) = False
      End If
      
      'Add By Sindy 2024/6/12
      If "" & .Fields("tm72") <> "" Then
         txt1(0).Text = "" & .Fields("tm137")
         txt1(0).Visible = True 'Add By Sindy 2024/7/31
         SSTab2.TabVisible(1) = True
         txt1(6).Text = "" & .Fields("tm138")
         txt1(6).Visible = True 'Add By Sindy 2024/7/31
         SSTab2.TabVisible(2) = True
         If txt1(0).Text <> "" Then
            SSTab2.Tab = 1
         ElseIf txt1(6).Text <> "" Then
            SSTab2.Tab = 2
         Else
            SSTab2.Tab = 1
         End If
      Else
         txt1(0).Text = ""
         txt1(0).Visible = False 'Add By Sindy 2024/7/31
         SSTab2.TabVisible(1) = False
         txt1(6).Text = ""
         txt1(6).Visible = False 'Add By Sindy 2024/7/31
         SSTab2.TabVisible(2) = False
      End If
      '2024/6/12 END
      
      m_ST03 = Pub_StrUserSt03
      
      '合併到最上面的語法
      If Not IsNull(.Fields("ibf01")) Then
         cmdPic.Caption = "已設定代表圖(&I)"
         cmdPic.BackColor = &HC0FFC0
         '無圖式
         Chk1.Enabled = False
      Else
         cmdPic.Caption = "未設定代表圖(&I)"
         cmdPic.BackColor = &HC0C0FF
         '無圖式
         Chk1.Enabled = True
      End If
      CheckOC2
      
      cmd(6).Visible = False 'Added by Lydia 2019/07/31 電子送件-申請書
      'Add By Sindy 2024/7/30
      cmd(4).Visible = False
      Label1(2).Visible = False
      If Left(m_EP05ST03, 2) = "P2" Then
      '2024/7/30 END
         'Add By Sindy 2018/8/13 T大陸案申請書改顯示指示信
         If m_CP01 = "T" And m_Country = "020" And _
            (Left(LBL1(3).Caption, 1) = "A" Or Left(LBL1(3).Caption, 1) = "B") Then
            cmd(4).Caption = "指示信(&A)"
            cmd(4).Visible = True
            Label1(2).Visible = True
         'Add By Sindy 2018/11/30 繳費單
         ElseIf m_CP10 = "717" And m_Country = "000" Then '註冊費
            cmd(4).Caption = "繳費單(&A)"
            cmd(4).Visible = True
            Label1(2).Visible = True
            cmd(6).Visible = True 'Added by Lydia 2020/10/21
            '2018/11/30 END
         ElseIf m_Country = "000" Then
            cmd(4).Caption = "申請書(&A)"
         '2018/8/13 END
            'Add By Sindy 2013/6/17
            'Modified by Lydia 2019/07/05 +申請101
            If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "301" Or m_CP10 = "501" Or m_CP10 = "103" Then
               cmd(4).Visible = True '顯示「申請書」按鈕(紙本)
               Label1(2).Visible = True
               'Added by Lydia 2019/07/31 分別顯示電子送件,紙本送件
               'Modified by Lydia 2020/10/07 + 301變更,501移轉
               If InStr("101,102,103,301,501", m_CP10) > 0 Then
                  cmd(4).Caption = "紙本送件"
                  cmd(6).Visible = True
               End If
               'Added by Lydia 2022/12/19 註冊證形式: 特定性質(補換發註冊證103)之紙本全改人工
               If strSrvDate(1) >= "20230101" And m_CP10 = "103" Then
                   cmd(4).Visible = False
               End If
               'end 2022/12/19
            'Add by Sindy 2020/9/28 + 電子送件725.代辦退費
            'Modified by Lydia 2020/10/21 +電子送件308.分割,313.減縮商品,304.英文證明書,502.授權
            'Mofieid by Lydia 2023/01/05 +電子送件729復權
            'Modified by Lydia 2024/08/06 +309中文證明書
            'Modified by lydia 2025/05/29 +商爭案:加速審查(311),陳述意見(210)
            ElseIf m_CP10 = "725" Or m_CP10 = "308" Or m_CP10 = "313" Or m_CP10 = "304" Or m_CP10 = "309" Or m_CP10 = "502" Or m_CP10 = "729" Or m_CP10 = "313" Or m_CP10 = "210" Then
               cmd(6).Visible = True
               cmd(4).Visible = False
               Label1(2).Visible = True
               '2020/9/28 END
            'Added by Lydia 2025/05/29 商爭承辦的FCT案
            ElseIf m_CP01 = "FCT" And Pub_StrUserSt93 = "T11" Then
               cmd(6).Visible = True
               cmd(4).Visible = False
               Label1(2).Visible = True
             'end 2025/05/29
            'Add by Sindy 2020/10/20
            ElseIf m_CP10 = "214" Then
               cmd(6).Visible = False
               cmd(4).Visible = True
               Label1(2).Visible = True
               '2020/10/20 END
            'Added by Lydia 2020/10/07 電子送件-補正申請書：其他沒有設定的性質，ex.303延期、201補正、202申請意見書、208補優先權證明、706其他
            ElseIf Left(m_strCP09, 1) = "A" Or Left(m_strCP09, 1) = "B" Then
               'Memo by Lydia 2022/12/19 註冊證形式: 特定性質之紙本全改人工; 復權729,註冊證副本314,分割308,註冊費717原本就不顯示紙本按鈕
               cmd(6).Visible = True
               cmd(4).Visible = False
               Label1(2).Visible = True
            'end 2020/10/07
            Else
               cmd(4).Visible = False
               Label1(2).Visible = False
            End If
            '2013/6/17 End
         End If
      End If
      
      'Add By Sindy 2018/4/20
      m_PP04 = "" '核判表設定的核稿人
      m_PP05 = "" '核判表設定的判發人
      Call PUB_ChkIsSetPromoterReader(m_CP14, m_CP01, m_CP10, m_PP04, m_PP05, m_strCP09, m_Country, m_PP01, m_PP03)
      If m_PP04 = Trim(Left("" & Combo1.Text, 6)) Then m_PP04 = "" '為自行核稿,不需再將自己ID放入核稿人欄位
      If m_PP05 = Trim(Left("" & Combo1.Text, 6)) Then m_PP05 = "" '為自行判發,不需再將自己ID放入判發人欄位
      '核稿人:
      Combo2.AddItem "", 0
      '有完稿日時,則不用再預設核稿人
      If Val("" & .Fields("EP09")) <= 0 And Len("" & .Fields("EP04")) = 0 Then
         Combo2.Tag = m_PP04
      Else
         'Add By Sindy 2018/4/25
         If m_CP27 = "" And "" & .Fields("EP04") = "" And m_PP04 <> "" Then
            Combo2.Tag = m_PP04
         Else
         '2018/4/25 END
            Combo2.Tag = "" & .Fields("EP04")
         End If
      End If
      If Combo2.Tag <> "" Then
         Combo2.AddItem Combo2.Tag & " ==> " & GetPrjSalesNM(Combo2.Tag), 1
      End If
      Call SetComboData(Combo2) 'Add By Sindy 2018/10/1
'      blnMatch = False
'      For ii = 0 To Me.Combo2.ListCount - 1
'         If Trim(Left(Me.Combo2.List(ii), 6)) = Combo2.Tag Then
'            Me.Combo2.ListIndex = ii
'            blnMatch = True
'            Exit For
'         End If
'      Next ii
'      If blnMatch = False Then Me.Combo2.ListIndex = 0
'      Combo2.Tag = Combo2.Text
      '判發人:
      Combo6.AddItem "", 0
      '不用完稿日為預設判發的基準點,改用檢查有無送判或判發歷程
      If PUB_ChkEmpFlowExists(LBL1(3), EMP_送判) = False And _
         PUB_ChkEmpFlowExists(LBL1(3), EMP_判發) = False And _
         Len("" & .Fields("EP40")) = 0 Then
         Combo6.Tag = m_PP05
      Else
         Combo6.Tag = "" & .Fields("EP40")
      End If
      If Combo6.Tag <> "" Then
         Combo6.AddItem Combo6.Tag & " ==> " & GetPrjSalesNM(Combo6.Tag), 1
      End If
      Call SetComboData(Combo6) 'Add By Sindy 2018/10/1
'      blnMatch = False
'      For ii = 0 To Me.Combo6.ListCount - 1
'         If Trim(Left(Me.Combo6.List(ii), 6)) = Combo6.Tag Then
'            Me.Combo6.ListIndex = ii
'            blnMatch = True
'            Exit For
'         End If
'      Next ii
'      If blnMatch = False Then Me.Combo6.ListIndex = 0
'      Combo6.Tag = Combo6.Text
      '2018/4/20 END
   End If
   End With
   
   'Add By Sindy 2020/12/1
   txtNote.Visible = False
   If m_CP163 <> "" Then
      If m_CP163 <> LBL1(3) Then
         strSql = "Select CP01,CP02,CP03,CP04,Decode('" & m_Country & "','000',CPM03,CPM04) as 案件性質" & _
                  " from caseprogress,CasePropertyMap" & _
                  " where cp09='" & m_CP163 & "' And CP01=CPM01(+) And CP10=CPM02(+)"
         If rsTmp.State = 1 Then rsTmp.Close
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            txtNote.Text = "※此案屬多案歷程，請參" & rsTmp.Fields("cp01") & "-" & rsTmp.Fields("cp02") & IIf(rsTmp.Fields("cp03") & rsTmp.Fields("cp04") = "000", "", "-" & rsTmp.Fields("cp03") & "-" & rsTmp.Fields("cp04")) & _
                           "(" & rsTmp.Fields("案件性質") & ")"
            txtNote.Width = 5000
            txtNote.Visible = True
         End If
         If rsTmp.State = 1 Then rsTmp.Close
      End If
   End If
   '2020/12/1 END
   
   'Add By Sindy 2024/12/5 顯示核稿語文
   If m_EP41 = "" Or PUB_ChkEmpFlowExists(LBL1(3), EMP_送英核) = False Then '尚未送英核時才預設核稿語文
      txt1(23) = "1" '預設英文
      If m_EMPST16 = "4" Then
         txt1(23) = "2" '日文
      End If
   Else
      txt1(23) = m_EP41
   End If
   SetEngChecker '設定外文核稿人選單
   '2024/12/5 END
   
   CheckOC
   InitialField
   'Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & LBL1(3).Caption & "' "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      UpdateFieldOldData rsTmp
      m_CP149 = "" & rsTmp.Fields("CP149")
   End If
   If rsTmp.State = 1 Then rsTmp.Close
   
   'Add By Sindy 2024/12/5
   Dim tmpInti As Integer
   For tmpInti = 0 To Combo4.ListCount - 1
       If Trim(txt1(5).Text) = Trim(Mid(Combo4.List(tmpInti), 1, InStr(1, Combo4.List(tmpInti), "=") - IIf(InStr(1, Combo4.List(tmpInti), "=") = 0, 0, 1))) Then
           Combo4.Text = Combo4.List(tmpInti)
       End If
   Next tmpInti
   '2024/12/5 END
   
   '是否會稿
   txt1(1).Text = LBL1(6).Caption
   '齊備日
   txt1(2).Text = ChangeWStringToTString(LBL1(8).Caption)
   '完稿日
   txt1(3).Text = ChangeWStringToTString(LBL1(10).Caption)
   If PUB_GetST05(strUserNum) = "91" Or PUB_GetST05(strUserNum) = "92" Then
      '指定會稿日
      txt1(18).Enabled = True
   Else
      txt1(18).Enabled = False
   End If
   
   If ProState = "2" Then
      frm090614.TextOk = True
   End If
   
   If m_blnClkSure = False Then
      'Modified by  Lydia 2019/05/02
      'If Me.Txt1(12).Text = "" And Me.Txt1(2).Text <> "" Then Call txt1_LostFocus(2)
      If Me.txt1(12).Text = "" And Me.txt1(2).Text <> "" Then
         tmpBol = False
         Call txt1_Validate(2, tmpBol)
      End If
   End If
   
   If Left(UCase(strText), 1) = "C" Then
      Label1(3).Visible = True
      '是否通知客戶
      txt1(9).Visible = True
      Label1(22).Visible = True
      '發文日
      txt1(8).Enabled = True
      txt1(8).TabStop = True
   Else
      Label1(3).Visible = False
      '是否通知客戶
      txt1(9).Visible = False
      Label1(22).Visible = False
      '發文日
      txt1(8).Enabled = False
      txt1(8).TabStop = False
   End If
   
   'C類來函未發文才顯示撰寫信函按鈕
   'Modify By Sindy 2025/8/15 T案727分析也要顯示 + Or (m_CP01 = "T" And m_Country = "000" And m_CP10 = "727")
   If (Mid(LBL1(3).Caption, 1, 1) = "C" Or _
       (m_CP01 = "T" And m_Country = "000" And m_CP10 = "727")) And txt1(8) = Empty Then
      cmd(3).Enabled = True
      cmd(3).Visible = True
   Else
      cmd(3).Enabled = False
      cmd(3).Visible = False
   End If
   
   'T及FCT的台灣爭議案,在收文或分案時若勾已齊備,系統會上齊備日
   '不然就要等到智權同仁文件齊備時,再執行”台灣商標爭議案齊備日輸入”
   '若承辦有輸通知補充資料時,系統會拿掉齊備日及承辦期限,等到智權同仁回覆補充資料時再上齊備日
   Label6.Visible = False
   Label1(9).Visible = False: textCP143.Visible = False  'Added by Lydia 2018/12/10
   Label1(9).Tag = "" 'Added by Lydia 2019/11/22
   'Modified by Lydia 2018/12/10 開放T台灣案管控文件齊備
   'If (m_CP01 = "T" Or m_CP01 = "FCT") And _
      m_country = "000" And _
      InStr(TMdebate, m_CP10) > 0 And _
      Val(DBDATE(lbl1(5))) >= Val(TMdebateStarDT) Then
   'Modified by Lydia 2019/04/15 商申案齊備日只管制A類收文
   'If ((m_CP01 = "T" Or m_CP01 = "FCT") And m_country = "000" And InStr(TMdebate, m_CP10) > 0 And Val(DBDATE(lbl1(5))) >= Val(TMdebateStarDT)) Or _
           (m_CP01 = "T" And m_country = "000" And Val(DBDATE(lbl1(5))) >= Val(T案收文齊備啟用日)) Then
   'Modified by Lydia 2022/07/15 + T大陸案之齊備日管控;  TC案之文件齊備日管控
   'If ((m_CP01 = "T" Or m_CP01 = "FCT") And m_Country = "000" And InStr(TMdebate, m_CP10) > 0 And Val(DBDATE(lbl1(5))) >= Val(TMdebateStarDT)) Or _
           (m_CP01 = "T" And m_Country = "000" And Val(DBDATE(lbl1(5))) >= Val(T案收文齊備啟用日) And Left(lbl1(3).Caption, 1) = "A") Then
   'Modify By Sindy 2023/6/5 非上列系統別,均可自行輸入齊備日 + And InStr("'T','FCT','TC'", "'" & m_CP01 & "'") > 0
   'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
   If (((m_CP01 = "T" Or m_CP01 = "FCT") And InStr("000,020", m_Country) > 0 And InStr(TMdebate, m_CP10) > 0 And Not (m_CP01 = "FCT" And InStr(FCT_NotTMdebate, m_CP10) > 0) And Val(DBDATE(LBL1(5))) >= Val(TMdebateStarDT)) Or _
       (m_CP01 = "T" And InStr("000,020", m_Country) > 0 And Val(DBDATE(LBL1(5))) >= Val(T案收文齊備啟用日) And Left(LBL1(3).Caption, 1) = "A") Or _
       (m_CP01 = "TC" And InStr("000,020", m_Country) > 0 And Left(LBL1(3).Caption, 1) = "A") _
      ) And InStr("'T','FCT','TC'", "'" & m_CP01 & "'") > 0 Then
      
      'Modify By Sindy 2019/5/14
      'cmd(0).Visible = True
      cmd(0).Enabled = True
      '2019/5/14 END
      'Add By Sindy 2012/10/24
      '承辦人逾3天不可再輸入通知補充資料
      'Add By Sindy 2012/11/30
      If m_CP149 = "" Then
         'Modify By Sindy 2019/5/14
         'cmd(1).Visible = False
         cmd(1).Enabled = False
         '2019/5/14 END
      Else
      '2012/11/30 End
         strCompDate = CompWorkDay(4, m_CP149)
         If strCompDate > strSrvDate(1) Then
            'Modify By Sindy 2019/5/14
            'cmd(1).Visible = True
            cmd(1).Enabled = True
            '2019/5/14 END
         Else
            'Modify By Sindy 2019/5/14
            'cmd(1).Visible = False
            cmd(1).Enabled = False
            '2019/5/14 END
         End If
      End If
      '2012/10/24 End
      txt1(2).Enabled = False '齊備日鎖住
      If Val(txt1(2)) = 0 Then
         Label6.Visible = True
      End If
      'Added by Lydia 2018/12/10 查名齊備日
      'Modified by Lydia 2022/07/15 + T大陸案之齊備日管控
      'If m_CP01 = "T" And m_Country = "000" And m_CP10 = 申請 Then
      If m_CP01 = "T" And InStr("000,020", m_Country) > 0 And m_CP10 = 申請 Then
           Label1(9).Visible = True: textCP143.Visible = True
           Label1(9).Tag = "Y" 'Added by Lydia 2019/11/22
      End If
      'end 2018/12/10
   Else
      'Modify By Sindy 2019/5/14
      'cmd(0).Visible = False
      cmd(0).Enabled = False
      'cmd(1).Visible = False 'Add By Sindy 2012/10/24
      cmd(1).Enabled = False
      '2019/5/14 END
      txt1(2).Enabled = True '開放齊備日可輸入
   End If
   
   cmd(5).Tag = "" 'Added by Lydia 2018/12/10 加註-基本權限
   'Add By Sindy 2018/4/20
   '個人案件不可用主管權限操作
   If ProState = "2" And m_CP14 = strUserNum Then  '2.主管
      cmd(5).Enabled = False
      cmdDetail.Enabled = False
      cmd(5).Tag = "N" 'Added by Lydia 2018/12/10 加註-基本權限
'   '無齊備日,不可使用歷程 (Sindy 2019/8/1:無齊備日時,進歷程也只能做附加歷程和聯絡)
'   ElseIf Val(txt1(2)) = 0 Then
'      cmd(5).Enabled = False
'      cmdDetail.Enabled = False
   Else
      cmd(5).Enabled = True
      cmdDetail.Enabled = True
      cmd(5).Tag = "Y" 'Added by Lydia 2018/12/10 加註-基本權限
   End If
   
   'Added by Lydia 2019/01/30 判斷T案的文件+查名齊備日
   'Memo by Lydia 2019/01/30 切換詳細資料: 先跑Process, 再跑SStab1_Click
   'Remove by Lydia 2019/11/22 因為T-223944(申請)的文件尚未齊備，但有聯絡歷程，所以改到歷程畫面控制
   'If textCP143.Visible = True And cmd(5).Tag = "Y" Then
   '   If Val(txt1(2)) = 0 Or Val(textCP143) = 0 Then
   '      cmd(5).Enabled = False
   '   Else
   '      cmd(5).Enabled = True
   '   End If
   'End If
   ''end 2019/01/30
   'end 2019/11/22
   
   If m_CPM29 = "N" Then
      Label18.Visible = False '先不顯示
   Else
      Label18.Visible = False
   End If
   '若為個人工作管理及承辦人下拉選單為操作者
   '完稿日
   Me.txt1(3).Enabled = True
   '會稿日
   Me.txt1(4).Enabled = True
   '會稿完成日
   Me.txt1(7).Enabled = True
   '外文核完日
   Me.txt1(19).Enabled = True
   '發文日
   If Left(m_strCP09, 1) = "C" Then Me.txt1(8).Enabled = True
   If ProState = "1" Or Trim(Left("" & Combo1.Text, 6)) = strUserNum Then
      If m_CPM29 = "" Then '要電子簽核的案件性質
         '完稿日
         'Modify By Sindy 2022/4/25 針對有註記「收款後送件」的台灣商標案件，開放承辦人於案件先行作業後，可自行輸入「完稿日」。
         If Me.LblFee.Caption = "尚待收款" And m_Country = "000" And Val(txt1(3)) = 0 And Val(txt1(2)) > 0 Then
            Me.txt1(3).Enabled = True
            Me.LblFee.Tag = "尚待收款"
         Else
         '2022/4/25 END
            Me.txt1(3).Enabled = False
         End If
         '會稿日
         Me.txt1(4).Enabled = False
         '會稿完成日
         Me.txt1(7).Enabled = False
         'Add By Sindy 2024/12/5 外文核完日
         Me.txt1(19).Enabled = False
         '2024/12/5 END
         '發文日
         Me.txt1(8).Enabled = False
         '不自動更新會完日時,則開放可以自行輸入會稿完成日
         If Me.txt1(7).Text = "" Then
            If PUB_ChkEmpFlowExists(LBL1(3), EMP_送會, , strRefEEP02) = True Then
               If PUB_ChkEmpFlowExists(LBL1(3), EMP_不自動更新會完日, strRefEEP02) = True Then
                  Me.txt1(7).Enabled = True
               End If
            End If
         End If
      Else
         '完稿日
         'Modify By Sindy 2022/4/25 針對有註記「收款後送件」的台灣商標案件，開放承辦人於案件先行作業後，可自行輸入「完稿日」。
         If Me.LblFee.Caption = "尚待收款" And m_Country = "000" Then
            If Val(txt1(3)) = 0 And Val(txt1(2)) > 0 Then
               Me.LblFee.Tag = "尚待收款"
            Else
               Me.txt1(3).Enabled = False
            End If
         End If
         '2022/4/25 END
      End If
   End If
   
   '是否會稿
   txt1(1).Enabled = True 'Add Sindy 2018/9/18
   If txt1(1) = "Y" And PUB_ChkEmpFlowExists(LBL1(3), EMP_送會) = True Then
      txt1(1).Enabled = False
   End If
   'Add By Sindy 2019/7/25 是否會稿,空白才要預設
   If Trim(LBL1(6).Caption) = "" Then
   '2019/7/25 END
      'Add By Sindy 2018/4/20 不會稿案件性質,則預設為N
      '商爭案是否會稿由智權人員在填寫接洽單時決定(收文,分案)
      If txt1(1).Text = "" Then
         If m_CPM28 <> "" Then
            txt1(1).Text = m_CPM28
         End If
         '申請案(101)一定要會稿
         'Modify By Sindy 2024/8/19 +剔除外商英日文組
         If m_CP10 = "101" And _
            Not (Left(PUB_GetST03(Trim(Left("" & Combo1.Text, 6))), 1) = "F" And (m_EMPST16 = "2" Or m_EMPST16 = "4")) Then
            'Modify By Sindy 2024/9/13 MCT案預設不會稿
            If m_Country = "000" And PUB_GetST93(m_CP13) = "T21" Then
               txt1(1).Text = "N"
            Else
            '2024/9/13 END
               txt1(1).Text = "Y"
            End If
         'Add By Sindy 2024/8/28 陳蒲璇提核准(1001)-限下一程序為701領證,要會稿
         ElseIf m_CP01 = "CFT" And m_CP10 = "1001" Then
            strExc(0) = "select * from nextprogress where NP01='" & m_strCP09 & "'" & _
                        " and NP07='701'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               txt1(1).Text = "Y"
            End If
         '2024/8/28 END
         End If
      End If
      '2018/4/20 END
      'Add By Sindy 2019/7/12 不電子簽核的案件性質,是否會稿預設為N
      If m_CPM29 = "N" Then
         txt1(1).Text = "N"
      End If
   End If
   
   'Add By Sindy 2021/11/16 主管權限
   If ProState = "2" Then
      '若要增加第二級期限管制人也有修改權限 : Or PUB_GetST52(Trim(Left("" & Combo1.Text, 6)), strUserNum) = True)
      'Modify By Sindy 2023/11/16 + InStr(Pub_GetSpecMan("承辦人工作管理可修改資料人員"), strUserNum) > 0
      If InStr(Pub_GetSpecMan("承辦人工作管理可修改資料人員"), strUserNum) > 0 Or _
         Pub_StrUserSt03 = "M51" Then
         '有修改權限
      Else
         For Each objText In Me.txt1
            objText.Enabled = False
         Next
         Me.Combo2.Enabled = False '核稿人
         Me.Combo6.Enabled = False '判發人
         Me.Combo4.Enabled = False '外文核稿人 Add By Sindy 2024/12/5
      End If
   End If
   '2021/11/16 END
   
   Call SetColTag(True)
   
   'Add By Sindy 2024/7/30
   If Left(m_EP05ST03, 2) <> "P2" Then '外商
      Label1(3).ForeColor = &H80000012
      Label1(22).ForeColor = &H80000012
      cmdOK(5).Visible = False '客戶專區
      LblFee.Visible = False '尚待收款
      Label6.Visible = False '智權人員做”爭議案齊備日輸入”之『回覆補充資料』或『齊備日或急件維護』後，系統會自動上齊備日。
      '查名齊備日
      Label1(9).Visible = False: textCP143.Visible = False
      Label1(9).Tag = ""
   End If
   '2024/7/30 END
   
   'Add By Sindy 2024/12/5 已有送英核歷程,外文核稿人鎖住
   If PUB_ChkEmpFlowExists(LBL1(3), EMP_送英核) = True Then
      Combo4.Enabled = False
   End If
   '2024/12/5 END
   
   Me.Enabled = True
End Sub

'Add By Sindy 2018/4/25
'bolSetTag=true : 將輸入欄位值記錄至.tag裡面
'bolSetTag=false : 比較輸入欄位值.Tag與畫面上資料是否一致
Private Function SetColTag(bolSetTag As Boolean) As Boolean
   If bolSetTag = True Then
      txt1(1).Tag = txt1(1)
      txt1(12).Tag = txt1(12)
      txt1(2).Tag = txt1(2)
      txt1(3).Tag = txt1(3)
      txt1(18).Tag = txt1(18)
      txt1(4).Tag = txt1(4)
      txtEP12.Tag = txtEP12
      txt1(7).Tag = txt1(7)
      txt1(19).Tag = txt1(19) 'Add By Sindy 2024/12/5
      txt1(8).Tag = txt1(8)
      txt1(9).Tag = txt1(9)
      txtCP64.Tag = txtCP64
      Combo2.Tag = Combo2.Text '核稿人
      Combo6.Tag = Combo6.Text '判發人
      'Add By Sindy 2024/12/5
      Combo4.Tag = Combo4.Text '外文核稿人
      '2024/12/5 END
      
      Chk1.Tag = Chk1.Value
      'Add By Sindy 2024/6/13
      txt1(0).Tag = txt1(0)
      txt1(6).Tag = txt1(6)
      '2024/6/13 END
   Else
      SetColTag = True
      'Add By Sindy 2024/9/23 + Or Trim(lbl1(6).Caption) = ""
      If txt1(1) = "" Or Trim(LBL1(6).Caption) = "" Then SetColTag = False: Exit Function '是否會稿欄位空白時,確定鍵會Update成 N
      If txt1(12).Tag <> txt1(12) Then SetColTag = False: Exit Function
      If txt1(2).Tag <> txt1(2) Then SetColTag = False: Exit Function
      If txt1(3).Tag <> txt1(3) Then SetColTag = False: Exit Function
      If txt1(1).Tag <> txt1(1) Then SetColTag = False: Exit Function
      If txt1(18).Tag <> txt1(18) Then SetColTag = False: Exit Function
      If txt1(4).Tag <> txt1(4) Then SetColTag = False: Exit Function
      If txtEP12.Tag <> txtEP12 Then SetColTag = False: Exit Function
      If txt1(7).Tag <> txt1(7) Then SetColTag = False: Exit Function
      If txt1(19).Tag <> txt1(19) Then SetColTag = False: Exit Function 'Add By Sindy 2024/12/5
      If txt1(8).Tag <> txt1(8) Then SetColTag = False: Exit Function
      If txt1(9).Tag <> txt1(9) Then SetColTag = False: Exit Function
      If txtCP64.Tag <> txtCP64 Then SetColTag = False: Exit Function
      If Left(Combo2.Tag, 5) <> Left(Combo2.Text, 5) Then SetColTag = False: Exit Function '核稿人
      If Left(Combo6.Tag, 5) <> Left(Combo6.Text, 5) Then SetColTag = False: Exit Function '判發人
      If Left(Combo4.Tag, 5) <> Left(Combo4.Text, 5) Then SetColTag = False: Exit Function '外文核稿人 Add By Sindy 2024/12/5
      If Chk1.Tag <> Chk1.Value Then SetColTag = False: Exit Function
      'Add By Sindy 2024/6/13
      If txt1(0).Tag <> txt1(0) Then SetColTag = False: Exit Function
      If txt1(6).Tag <> txt1(6) Then SetColTag = False: Exit Function
      '2024/6/13 END
   End If
End Function

'Add By Sindy 2018/4/25
Private Sub ChkEP34ToEP07EP08()
Dim bolChkEmp As Boolean
   
   'add by nickc 2006/09/26 若是輸入不會稿，直接按存檔，他不會自動代
   'If txt1(1) = "N" Then txt1(4) = txt1(3): txt1(7) = txt1(3)
   If txt1(1) = "N" Then
      bolChkEmp = False
      '要電子簽核的案件或有電子歷程的案件
      If m_CPM29 = "" Or _
         m_Flow = EMP_送核 Or _
         m_Flow = EMP_送英核 Or _
         ((PUB_ChkEmpFlowExists(LBL1(3), EMP_送核) = True Or PUB_ChkEmpFlowExists(LBL1(3), EMP_送英核) = True) And PUB_ChkEmpFlowExists(LBL1(3), EMP_核完) = False) Then
         '有核稿人
         If Combo2.Text <> "" And Left(Trim(Combo2.Text), 5) <> m_CP14 Then
            bolChkEmp = True
'                     strExc(0) = "select ep39 From engineerprogress where ep02='" & lbl1(3) & "'"
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        If Val("" & RsTemp.Fields("ep39")) <= 0 Then
'                           bolUpdDate = False
'                        End If
'                     End If
         End If
      End If
      If bolChkEmp = False Then '無電子簽核或無核稿主管
         'Modify By Sindy 2016/5/19 發現核完N不會稿時,系統會上(會稿日)和(會稿完成日),所以要檢查日期是否已有值,以免重新覆蓋掉 ex:P-113197
         'txt1(4) = txt1(3): txt1(7) = txt1(3)
         If Trim(txt1(4).Text) = "" Then
            txt1(4).Text = txt1(3).Text
         End If
         If Trim(txt1(7).Text) = "" Then
            txt1(7).Text = txt1(3).Text
         End If
         '2016/5/19 END
'               Else
'                  If bolUpdDate = True Then
'                     If Trim(txt1(4)) = "" Then
'                        txt1(4) = strSrvDate(2)
'                     End If
'                     If Trim(txt1(7)) = "" Then
'                        txt1(7) = strSrvDate(2)
'                     End If
'                  End If
      End If
   End If
   '2013/10/4 END
End Sub

Public Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
   Exit Sub
End Sub

Public Sub PubShowNextData()
'Dim rsA As New ADODB.Recordset
'Dim stFileName As String
Dim hLocalFile As Long
   
   '***2008/11/21 加註BY SONIA 按確定後很快按結束會因為DoEvents造成錯誤,因使用者未反應故暫不取消DoEvents
   Dim iMouse As Integer
   iMouse = Screen.MousePointer
   
   Select Case cmdState
   Case 0 '本月統計
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      CALCUTE_090201 Trim(Left("" & Combo1.Text, 6)), Text1.Text
      Me.Enabled = True
      
      Screen.MousePointer = iMouse
      
      GRD1.col = 3
      Me.Hide
      frm090201_3_1.m_strEmp = Trim(Left(Me.Combo1.Text, 6)) 'Add By Sindy 2021/9/11
      '若為個人管理
      If ProState = "1" Then
          frm090201_3_1.m_strYear = Left(Me.Text1.Text, 4) - 1911
          frm090201_3_1.m_strMonth = Right(Me.Text1.Text, 2)
      '若為工作維護
      Else
          frm090201_3_1.m_strYear = frm090614.txt1(3).Text
          frm090201_3_1.m_strMonth = frm090614.txt1(4).Text
      End If
      frm090201_3_1.Show
      
   Case 1 '確定
      Select Case ProState
      Case "1", "2"
         m_chkcmdok1 = False 'Add By Sindy 2013/6/7 進入承辦歷程時會先執行一次確定鍵,因有可能已在此畫面先修改資料,且有些日期檢查條件須先執行
         If SSTab1.Tab = 0 Then Exit Sub
         
         'Add By Sindy 2019/7/12 輸入發文日,未輸入完稿日,則完稿日=發文日
         If Len(txt1(3)) = 0 And txt1(8).Enabled = True Then txt1(3) = txt1(8)
         'Add By Sindy 2018/4/25
         If txt1(1) = "" Then txt1(1) = "N"
         Call ChkEP34ToEP07EP08
         '2018/4/25 END
         
         Screen.MousePointer = vbHourglass
         'Modify By Sindy 2018/4/25
         'If SSTab1.Tab = 1 Then
         If SSTab1.Tab = 1 Or Me.m_Flow <> "" Then
         '2018/4/25 END
            If ChkNoData = False Then
               '重新檢查欄位有效性
               If TxtValidate = True Then
                  DoEvents
                  Me.Enabled = False
                  If FormSave = True Then
                     '集中發信
                     'Modify By Sindy 2018/6/20 + If m_Flow = "" Then
                     If m_Flow = "" Then BatctMail
                     
                     '更新mdb暫存資料及第一畫面的Grid內容
                     UpdEngMdb
                     TextOk = False
                     Call SetColTag(True) 'Add By Sindy 2018/4/25
                     m_chkcmdok1 = True 'Add By Sindy 2018/4/25
                  End If
                  Me.Enabled = True
                  'SSTab1.Tab = 0
'                  'Add By Sindy 2018/4/26
'                  If cmdOK(1).Enabled = True Then
'                  '2018/4/26 END
'                     SSTab1.Tab = 0
'                  End If
               'Add By Sindy 2019/12/3
               Else
                  Screen.MousePointer = vbDefault
                  Exit Sub
               '2019/12/3 END
               End If
            End If
         Else
            SSTab1.Tab = 1
         End If
         'Add By Sindy 2018/5/15
         If SSTab1.Tab = 1 Then
            Call Process(LBL1(3).Caption)
         End If
         '2018/5/15 END
         
         'Add By Sindy 2020/10/8
         If Me.m_Flow <> "" Then
            '重查資料,多案單筆歷程時,要更新瀏覽資料日期欄位值
            'Modify By Sindy 2021/2/2 嘉雯在做延展時,會點本所期限做排序再進行歷程,為不影響到畫面資料順序,改寫
            'Call Combo1_Click
            If frm090202_2.m_RetrunRecvSub <> "" Then
               Call StrMenuOneRec_RecvSub(frm090202_2.m_RetrunRecvSub)
            End If
            '2021/2/2 END
         '2020/10/8 END
            Me.m_Flow = "" 'Add By Sindy 2018/4/25
         End If
         Screen.MousePointer = iMouse
      Case Else
      End Select
         
   Case 2 '結束
      Select Case ProState
      Case "1"
          Unload Me
          Exit Sub
      Case "2"
           frm090614.Show
           Unload Me
           Exit Sub
      Case "3"
      Case Else
      End Select
      
   Case 3 '接洽單
      Screen.MousePointer = vbHourglass
      If m_CP140 <> "" Then
         '查詢接洽記錄單
         'Modify By Sindy 2022/12/23 改用共用函數
         Call PUB_Queryfrm090801(m_CP140, DBDATE(LBL1(5).Caption), Me)
'         'Modify By Sindy 2022/9/5
'         If DBDATE(lbl1(5)) >= 接洽單電子收文啟用日 Then
'            frm090801_Q.SetParent Me
'            frm090801_Q.m_blnCallPrint = True
'            frm090801_Q.Text5 = m_CP140
'            Call frm090801_Q.cmdOK_Click(4)
'            'frm090801_Q.ZOrder
'            frm090801_Q.Show vbModal
'         Else
'         '2022/9/5 END
'            frm090801.SetParent Me
'            frm090801.m_blnCallPrint = True 'Add By Sindy 2022/10/19
'            frm090801.Text5 = m_CP140
'            frm090801.m_blnCallPrint_CRL119 = True '是否列印特殊收據頁
'            Call frm090801.cmdOK_Click(4)
'            frm090801.cmdOK(2).Visible = False
'            frm090801.cmdOK(0).Visible = False
'            frm090801.txtPCnt.Visible = False
'            Me.Hide
'         End If
         '2022/12/23 END
         cmdState = 99 '結束
'      Else
'         '檢查是否有接洽單.pdf
'         strExc(0) = "select *" & _
'                     " From casepaperpdf" & _
'                     " where cpp01='" & m_EEP01 & "' and instr(upper(cpp02),upper('" & EMP_接洽單 & ".pdf'))>0 and cpp10<>'D'"
'         rsA.CursorLocation = adUseClient
'         rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount > 0 Then
'            '讀取檔案名稱
'            stFileName = rsA.Fields("cpp02")
'   '         If GetAttachFile_CPP(m_EEP01, stFileName, m_AttachPath & "\" & stFileName) = False Then
'   '            MsgBox "無法儲存欲開啟的檔案[ " & stFileName & " ]！"
'   '         End If
'            If PUB_GetAttachFile_CPP(m_EEP01, stFileName, m_AttachPath) = True Then
'               '開啟檔案
'               ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
'            End If
'         Else
'            MsgBox "無接洽單！"
'         End If
'         rsA.Close
'         Set rsA = Nothing
      End If
      Screen.MousePointer = vbDefault
   
   Case 4 '完整卷宗
      Screen.MousePointer = vbHourglass
      frm100101_L.m_strKey = LBL1(7).Caption
      frm100101_L.SetParent Me
      If frm100101_L.QueryData = True Then
         frm100101_L.Show
         Me.Hide
      Else
         Unload frm100101_L
      End If
      Screen.MousePointer = vbDefault
   
   'Add By Sindy 2023/3/28
   Case 5 '客戶專區
      '直接開啟視窗
      'SHELL "Explorer.exe " & TmpDirNm, vbNormalFocus
      'Lydia:用檔案總管開啟放置1~2分鐘後,檔案總管會出錯(ex. A2037, A4041)
      ShellExecute hLocalFile, "explore", m_strFilePath, vbNullString, vbNullString, 1
   Case Else
   End Select
End Sub

Private Sub cmdok2_Click(Index As Integer)
Dim iMouse As Integer
iMouse = Screen.MousePointer

Screen.MousePointer = vbHourglass
GRD1.Visible = False
Select Case Index
Case 0 '當月資料
      'Modify By Sindy 2015/9/10 增加CP140電子表單單號做排序,自動收文在前面
'      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110010,format(R110026 * r110027 * iif(r110030='N',1,iif(isnull(r110028) or r110028=0,1,r110028)),'0.00') ,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032 FROM R090614 " & _
'                    " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' ORDER BY R110002 desc,R110003,R110004 "
      'Modify By Sindy 2023/2/8 原案件清單排列,以電子收文為優先
      '但現皆為電子收文案件,會導致當日智慧局來文不易被發現,請調整為以"目次"由大至小排列
      '取消 R110033 desc,
      'Modify By Sindy 2024/1/15 +,R110034
      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110034,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032" & _
               " FROM R090614 " & _
               " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
               " ORDER BY R110002 desc,R110003,R110004 "
      CheckOC
      With adoRecordset
          .CursorLocation = adUseClient
          .Open strSql, adoEng, adOpenStatic, adLockReadOnly
          If .RecordCount <> 0 And .RecordCount > 0 Then
              Set GRD1.Recordset = adoRecordset
              ChgGrdColor
          Else
             GRD1.Clear
             GRD1.Rows = 2
          End If
      End With
      CheckOC
      SWPRow2 = 1
      SWPRow = 1 'Add By Sindy 2024/11/7
Case 1 '未發文
      'Modify By Sindy 2015/9/10 增加CP140電子表單單號做排序,自動收文在前面
      'Modify By Sindy 2023/2/8 原案件清單排列,以電子收文為優先
      '但現皆為電子收文案件,會導致當日智慧局來文不易被發現,請調整為以"目次"由大至小排列
      '取消 R110033 desc,
      'Modify By Sindy 2024/1/15 +,R110034
      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110034,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,r110032" & _
               " FROM R090614 " & _
               " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' AND R110018='' and (R110024='' or R110024='0')" & _
               " order by R110002 desc "
      CheckOC
      With adoRecordset
          .CursorLocation = adUseClient
          .Open strSql, adoEng, adOpenStatic, adLockReadOnly
          If .RecordCount <> 0 And .RecordCount > 0 Then
              Set GRD1.Recordset = adoRecordset
              ChgGrdColor
          Else
               GRD1.Clear
               GRD1.Rows = 2
               SetGrd1
          End If
      End With
      CheckOC
      SWPRow2 = 1
      SWPRow = 1 'Add By Sindy 2024/11/7
Case Else
End Select
'Modify By Sindy 2018/8/23
'MouseClick (1)
MouseClick_1 (1)
'2018/8/23 END
GRD1.Visible = True
Screen.MousePointer = iMouse
End Sub

Private Sub CmdPic_Click()
frmPic001.oCP01 = SystemNumber(LBL1(7), 1)
frmPic001.oCP02 = SystemNumber(LBL1(7), 2)
frmPic001.oCP03 = SystemNumber(LBL1(7), 3)
frmPic001.oCP04 = SystemNumber(LBL1(7), 4)
frmPic001.StrMenu
frmPic001.SetSeekCmdok 'Add by Amy 2018/07/19
frmPic001.Show vbModal
'add by nickc 2005/12/14 檢查有無代表圖
'Modify by Amy 2018/07/16  改寫至function
'strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(Lbl1(7), 1) & "' and ibf02='" & SystemNumber(Lbl1(7), 2) & "' and ibf03='" & SystemNumber(Lbl1(7), 3) & "' and ibf04='" & SystemNumber(Lbl1(7), 4) & "' and ibf05='1' "
'CheckOC2
'adoRecordset1.CursorLocation = adUseClient
'adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
If ChkImgByteFile(SystemNumber(LBL1(7), 1), SystemNumber(LBL1(7), 2), SystemNumber(LBL1(7), 3), SystemNumber(LBL1(7), 4)) = True Then
    cmdPic.Caption = "已設定代表圖(&I)"
    cmdPic.BackColor = &HC0FFC0
    'add by nickc 2007/11/29 加入無圖式的格式
    Chk1.Value = vbUnchecked
    Chk1.Enabled = False
Else
    cmdPic.Caption = "未設定代表圖(&I)"
    cmdPic.BackColor = &HC0C0FF
    'add by nickc 2007/11/29 加入無圖式的格式
    Chk1.Value = vbUnchecked
    Chk1.Enabled = True
End If
'CheckOC2
'end 2018/07/19
End Sub

'Add By Sindy 2018/4/17
Private Sub cmdDetail_Click()
   Call grd2_DblClick
End Sub

'Add By Sindy 2018/4/17
Private Sub cmdQuery_Click()
   If QueryData(True) = False Then ShowNoData
End Sub

'Add By Sindy 2018/4/17
Public Function QueryData(bolFirst As Boolean) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer
Dim strQuyDate As String
Dim strVal As String
   
   m_blnColOrderAsc = True
   QueryData = True
   
   If Combo5.ListIndex = 0 Then
      strQuyDate = CompWorkDay(3, strSrvDate(1), 1)  '不含當天,3個工作天
   ElseIf Combo5.ListIndex = 1 Then
      strQuyDate = CompWorkDay(5, strSrvDate(1), 1) '不含當天,5個工作天
   ElseIf Combo5.ListIndex = 2 Then
      strQuyDate = CompWorkDay(7, strSrvDate(1), 1) '不含當天,7個工作天
   Else
      '全部
   End If
   
   grd2.Clear
   SetGrd2
   
   Screen.MousePointer = vbHourglass
   
   'Modify By Sindy 2016/3/3 取消此句,因退件不會上待回覆Y " union select EmpElectronProcess.* from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp27 is null and cp57 is null and eep04='" & EMP_退件 & "' and eep09='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   '                         增加EEP13='Y'
   'Modify By Sindy 2020/11/30 + EEP15,EEP11
   strVal = "select eep01,eep02 from EmpElectronProcess where EEP13='Y' and EEP09 is null and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and EEP04 not in(" & EMP_待辦歷程查詢除外的狀態 & ")" & _
            " union select eep01,eep02 from EmpElectronProcess where EEP13='Y' and EEP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and EEP04='" & EMP_聯絡 & "'" & IIf(strQuyDate <> "", " And EEP06>=" & strQuyDate, "")
'Modify By Sindy 2024/7/30 CP14 改判斷 EP05
   strSql = "Select ' ' as V,EP01 as 目次,SqlDateT(EEP06)||' '||sqltime(EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,TM05||TM06||TM07 as 案件名稱," & _
            "NA03 as 國家,Decode(TM10,'000',PTM03,PTM04) as 種類,Decode(TM10,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "EEP01 as 總收文號,EEP02 as 序號,EP08,EP38,' ' as 不顯示,EEP06 a,EEP07 b,EEP15,EEP11" & _
            " From EmpElectronProcess,CaseProgress,EngineerProgress,Trademark," & _
            "staff s1,staff s2,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
            " Where (EEP01,EEP02) in(" & strVal & ")" & _
            " and EEP01=CP09(+)" & _
            " and cp158=0 and cp159=0" & _
            " And EEP01=EP02(+)" & _
            " And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04" & _
            " And EP05=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And TM10=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '2'=PTM01(+) AND TM08=PTM02(+)" & _
            " And ac01='09' And EEP04=ac02(+)"
   'If ProState = "1" Then '個人
      strSql = strSql & " And EP05='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   'End If
   'Add By Sindy 2018/4/17
   'Modify By Sindy 2020/11/30 + EEP15,EEP11
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(EEP06)||' '||sqltime(EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,SP05||SP06||SP07 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(SP09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "EEP01 as 總收文號,EEP02 as 序號,EP08,EP38,' ' as 不顯示,EEP06 a,EEP07 b,EEP15,EEP11" & _
            " From EmpElectronProcess,CaseProgress,EngineerProgress,servicepractice," & _
            "staff s1,staff s2,nation,CasePropertyMap,allcode" & _
            " Where (EEP01,EEP02) in(" & strVal & ")" & _
            " and EEP01=CP09(+)" & _
            " and cp158=0 and cp159=0" & _
            " And EEP01=EP02(+)" & _
            " And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04" & _
            " And EP05=s1.ST01(+) And CP13=s2.ST01(+)" & _
            " And SP09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And EEP04=ac02(+)" & _
            IIf(Pub_StrUserSt15 = "P22", " And EEP04 not in('" & EMP_判發 & "','" & EMP_退件重送 & "')", "")
   'If ProState = "1" Then '個人
      strSql = strSql & " And EP05='" & Trim(Left("" & Combo1.Text, 6)) & "'"
   'End If
   strSql = strSql & " order by a desc,b desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set grd2.Recordset = rsTmp
      For i = 1 To grd2.Rows - 1
         Call SetColColor(i)
      Next i
      cmdDetail.Enabled = True
   Else
      cmdDetail.Enabled = False
      QueryData = False
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Function
   End If
      
ExitQuery:
   '若有資料時游標停在第一筆
   If bolFirst = True Then
      grd2.Visible = False
      grd2.col = 0
      grd2.row = 1
      If rsTmp.RecordCount > 0 Then
         dblPrevRow = grd2.row
         grd2.Text = "V"
         m_intRow = 1: m_intCol = 0
         For i = 0 To grd2.Cols - 1
            'Modify By Sindy 2020/11/30
            If i <> 4 Then
            '2020/11/30 END
               grd2.col = i
               If grd2.CellBackColor <> &H8080FF Then
                  grd2.CellBackColor = &HFFC0C0
               End If
            End If
         Next i
      End If
      grd2.Visible = True
   End If
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'Add By Sindy 2018/4/17
Private Sub SetColColor(intRow As Integer)
Dim i As Integer
   
   If intRow < 1 Then Exit Sub
   grd2.row = intRow
   '退件要淺紅色表示
   If grd2.TextMatrix(intRow, 12) = "退件" Then
      grd2.col = 12
      grd2.CellBackColor = &HC0C0FF
   End If
   'Add By Sindy 2020/11/30 多案時,案件名稱變桃粉色
   If grd2.TextMatrix(intRow, 20) <> "" And _
      InStr(grd2.TextMatrix(intRow, 21), "多案單筆歷程") > 0 Then
      grd2.col = 4
      grd2.CellBackColor = &HFF00FF 'QBColor(Rnd * 5)
   End If
End Sub

'Add By Sindy 2018/4/17
Private Sub SetGrd2()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2020/11/30 + EEP15,EEP11
   arrGridHeadText = Array("V", "目次", "流程日期", "本所案號", "案件名稱", _
                           "國家", "種類", "案件性質", "本所期限", "承辦人", _
                           "承辦期限", "智權人員", "目前流程狀態", _
                           "總收文號", "序號", "EP08", "EP38", "不顯示", "EEP06 a", "EEP07 b", "EEP15", "EEP11")
   arrGridHeadWidth = Array(200, 400, 800, 1400, 1000, _
                            700, 450, 900, 800, 600, _
                            800, 600, 600, _
                            0, 0, 0, 0, 600, 0, 0, 0, 0)
   grd2.Visible = False
   grd2.Cols = UBound(arrGridHeadText) + 1
   grd2.Rows = 2
   For iRow = 0 To grd2.Cols - 1
      grd2.row = 0
      grd2.col = iRow
      grd2.Text = arrGridHeadText(iRow)
      grd2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      If iRow = 11 Or iRow = 12 Then
         grd2.CellAlignment = flexAlignLeftCenter
      Else
         grd2.CellAlignment = flexAlignCenterCenter
      End If
   Next
   grd2.Visible = True
End Sub

'Add By Sindy 2018/4/20 核稿人
Private Sub Combo2_LostFocus()
   Dim ii As Integer
   Dim blnMatch As Boolean
   
   For ii = 0 To Me.Combo2.ListCount - 1
       blnMatch = False
       If Trim(Left(Me.Combo2.List(ii), 6)) = Trim(Left(Me.Combo2.Text, 6)) Then
           Me.Combo2.ListIndex = ii
           blnMatch = True
           Exit For
       End If
   Next ii
   If blnMatch = False Then
      If Trim(Left(Me.Combo2.Text, 6)) = "" Then
      '   Me.Combo2.ListIndex = 0 'Remove by Lydia 2021/12/23
      Else
         If Len(GetPrjSalesNM(Trim(Left(Me.Combo2.Text, 6)))) = 0 Then
            'Modify By Sindy 2022/12/6
            'Call ShowStaffErr(Trim(Left(Me.Combo2.Text, 6)))
            Call PUB_GetStaffNameDept(Trim(Left(Me.Combo2.Text, 6)), strExc(10), strExc(0), True, False)
            '2022/12/6 END
            Me.Combo2.SetFocus
            Exit Sub
         Else
            Combo2.Text = Trim(Left(Me.Combo2.Text, 6)) & " ==> " & GetPrjSalesNM(Trim(Left(Me.Combo2.Text, 6)))
         End If
      End If
   End If
End Sub
'2018/4/20 END

Private Sub Combo4_LostFocus()
   Dim ii As Integer
   Dim blnMatch As Boolean
   
   For ii = 0 To Me.Combo4.ListCount - 1
       blnMatch = False
       If Trim(Left(Me.Combo4.List(ii), 6)) = Trim(Left(Me.Combo4.Text, 6)) Then
           Me.Combo4.ListIndex = ii
           blnMatch = True
           Exit For
       End If
   Next ii
   If blnMatch = False Then Me.Combo4.ListIndex = 0
End Sub

Private Sub Combo5_Click()
   If Me.Visible = True Then
      If QueryData(True) = False Then ShowNoData 'Add By Sindy 2023/4/12
   End If
End Sub

'Add By Sindy 2018/4/20 判發人
Private Sub Combo6_LostFocus()
   Dim ii As Integer
   Dim blnMatch As Boolean
   
   For ii = 0 To Me.Combo6.ListCount - 1
       blnMatch = False
       If Trim(Left(Me.Combo6.List(ii), 6)) = Trim(Left(Me.Combo6.Text, 6)) Then
           Me.Combo6.ListIndex = ii
           blnMatch = True
           Exit For
       End If
   Next ii
   If blnMatch = False Then
      If Trim(Left(Me.Combo6.Text, 6)) = "" Then
         Me.Combo6.ListIndex = 0
      Else
         If Len(GetPrjSalesNM(Trim(Left(Me.Combo6.Text, 6)))) = 0 Then
            'Modify By Sindy 2022/12/6
            'Call ShowStaffErr(Trim(Left(Me.Combo6.Text, 6)))
            Call PUB_GetStaffNameDept(Trim(Left(Me.Combo6.Text, 6)), strExc(10), strExc(0), True, False)
            '2022/12/6 END
            Me.Combo6.SetFocus
            Exit Sub
         Else
            Combo6.Text = Trim(Left(Me.Combo6.Text, 6)) & " ==> " & GetPrjSalesNM(Trim(Left(Me.Combo6.Text, 6)))
         End If
      End If
   End If
End Sub
'2015/5/21 END

'Add By Sindy 2018/4/17
Private Sub grd2_DblClick()
Dim nFrm As Form
   
   If m_intRow <> 0 Then
      If m_intCol <> 17 Then
         If cmdDetail.Enabled = False Then Exit Sub

         If dblPrevRow = 0 Then
            MsgBox "請點選一筆資料列!", vbExclamation
            Exit Sub
         End If
         
         If grd2.TextMatrix(dblPrevRow, 0) = "V" Then
            Call Process(grd2.TextMatrix(dblPrevRow, 13)) '要重新查詢資料,因核稿人及判發人有預設問題
            If Me.cmd(5).Enabled = True Then
               '重新檢查欄位有效性
               If TxtValidate = True Then
'                  '檢查表單是否已開啟，若是，則關閉
'                  For Each nFrm In Forms
'                     If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'                        Unload frm090202_2
'                        Exit For
'                     End If
'                  Next
                  If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Sub 'Add By Sindy 2020/1/17
                  intBackTab = 2
                  frm090202_2.Hide
                  frm090202_2.m_EEP01 = grd2.TextMatrix(dblPrevRow, 13) '總收文號
                  frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) '案件流程所屬人員
                  frm090202_2.intReceiveKind = 0 '0.承辦人工作進度
                  frm090202_2.SetParent Me
                  'Add By Sindy 2018/8/15
                  If cmdTSMap.Visible = True Then
                     cmdTSMap.Tag = "Y"
                  Else
                     cmdTSMap.Tag = ""
                  End If
                  '2018/8/15 END
                  If frm090202_2.QueryData = True Then
                     frm090202_2.Show
                     Me.Hide
                  End If
               End If
            Else
               Me.SSTab1.Tab = 1
            End If
         End If
      End If
   End If
End Sub

'Add By Sindy 2018/4/17
Private Sub GRD2_SelChange()
Dim j As Integer

grd2.Visible = False
If grd2.MouseRow = 0 Then
   '已選取的資料列清除反白
   For j = 1 To grd2.Rows - 1
      If grd2.TextMatrix(j, 0) = "V" Then
         grd2.col = 0
         grd2.row = j
         grd2.Text = ""
         For i = 0 To grd2.Cols - 1
            'Modify By Sindy 2020/11/30
            If i <> 4 Then
            '2020/11/30 END
               grd2.col = i
               grd2.CellBackColor = QBColor(15)
            End If
         Next i
         Call SetColColor(j)
         Exit For
      End If
   Next j
Else
   '上一筆資料列清除反白
   If dblPrevRow > 0 And dblPrevRow <= (grd2.Rows - 1) Then
      grd2.col = 0
      grd2.row = dblPrevRow
      grd2.Text = ""
      For i = 0 To grd2.Cols - 1
         'Modify By Sindy 2020/11/30
         If i <> 4 Then
         '2020/11/30 END
            grd2.col = i
            If grd2.CellBackColor <> &H8080FF Then
               grd2.CellBackColor = QBColor(15)
            End If
         End If
      Next i
      Call SetColColor(CStr(dblPrevRow))
   End If
   '目前資料列反白
   grd2.col = 0
   grd2.row = grd2.MouseRow
   dblPrevRow = grd2.row
   If grd2.TextMatrix(grd2.row, 1) <> "" Then
      grd2.Text = "V"
      For i = 0 To grd2.Cols - 1
         'Modify By Sindy 2020/11/30
         If i <> 4 Then
         '2020/11/30 END
            grd2.col = i
            If grd2.CellBackColor <> &H8080FF Then
               grd2.CellBackColor = &HFFC0C0
            End If
         End If
      Next i
   End If
End If
grd2.Visible = True
End Sub

'Add By Sindy 2018/4/17
Private Sub grd2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow grd2, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   'grd2.col = nCol
   grd2.row = nRow
   If Me.grd2.row < 1 And Me.grd2.Text <> "V" Then
      If Me.grd2.Text = "目次" Then
         If m_blnColOrderAsc = True Then
            Me.grd2.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.grd2.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.grd2.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.grd2.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub

'Add By Sindy 2018/4/17 不顯示功能
Private Sub grd2_Click()
   m_intRow = grd2.MouseRow
   m_intCol = grd2.MouseCol
   If m_intRow <> 0 Then
      If m_intCol = 17 Then '不顯示
         If grd2.TextMatrix(m_intRow, 13) <> "" And _
            grd2.TextMatrix(m_intRow, 12) <> "核修" And _
            grd2.TextMatrix(m_intRow, 12) <> "核完" And _
            grd2.TextMatrix(m_intRow, 12) <> "會修" And _
            grd2.TextMatrix(m_intRow, 12) <> "會完" And _
            grd2.TextMatrix(m_intRow, 12) <> "繪圖判發" And _
            grd2.TextMatrix(m_intRow, 12) <> "判發" And _
            grd2.TextMatrix(m_intRow, 12) <> "退回" And _
            grd2.TextMatrix(m_intRow, 12) <> "退件" And _
            grd2.TextMatrix(m_intRow, 12) <> "圖修" And _
            grd2.TextMatrix(m_intRow, 12) <> "圖完" Then
            grd2.TextMatrix(m_intRow, 17) = "V"
            If MsgBox("請再次確定不顯示 " & vbCrLf & grd2.TextMatrix(m_intRow, 3) & " " & grd2.TextMatrix(m_intRow, 12) & " 嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
               grd2.TextMatrix(m_intRow, 17) = ""
            Else
               strExc(0) = "update EmpElectronProcess set eep13=null" & _
                           " where eep01='" & grd2.TextMatrix(m_intRow, 13) & "'" & _
                             " and eep02=" & grd2.TextMatrix(m_intRow, 14)
               Pub_SeekTbLog strExc(0) 'Add By Sindy 2020/11/30
               cnnConnection.Execute strExc(0)
               grd2.RowHeight(m_intRow) = 0
            End If
         End If
      End If
   End If
End Sub

'Modified by Morgan 2021/12/24 Form2.0點選同一人不會觸發Click事件，改用DropButtonClick事件但要控制第2次才執行
'Private Sub Combo1_Click()
'Modify By Sindy 2025/7/31
'Private Sub Combo1_DropButtonClick()
Public Sub Combo1_DropButtonClick()
'2025/7/31 END
   Static bClick As Boolean
   If bClick = False Then
      bClick = True
      Exit Sub
   End If
   bClick = False
'end 2021/12/24
   
   Me.Enabled = False 'Add By Sindy 2024/3/14
   
   Dim iMouse As Integer
   iMouse = Screen.MousePointer
   
   Me.GRD1.Visible = False
   Screen.MousePointer = vbHourglass
   Me.MousePointer = vbHourglass
   GRD1.MousePointer = flexArrowHourGlass
   Me.Enabled = False
   Combo1.Enabled = False
   Label3.Caption = GetPrjSalesNM(Trim(Left("" & Combo1.Text, 6)))
   StrMenu1
   StrMenu
'   If ChkNoData = True Then
'      For s = 0 To 10
'         txt1(s).Enabled = False
'      Next s
'   Else
'      For s = 0 To 10
'         txt1(s).Enabled = True
'      Next s
'   End If
   SetGrd1
   DoEvents
   'cmdok2(0).SetFocus
   Combo1.Enabled = True
   Me.Enabled = True
   GRD1.MousePointer = flexDefault
   Me.MousePointer = vbDefault
   Screen.MousePointer = iMouse
   
   Me.GRD1.Visible = True
   
   Me.Enabled = True 'Add By Sindy 2024/3/14
End Sub

Private Sub Form_Activate()
   ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_Load()
Dim iMouse As Integer
Dim nFrm As Form
   
   If strSrvDate(1) < T商標電子化第2階段啟用日 Then
      Label16.Caption = "註：雙擊選取時，開啟承辦歷程。"
   End If
   
   m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
   iMouse = Screen.MousePointer
   ReDim m_FieldList(TF_CP)
   
'   'Add By Sindy 2018/4/26
'   '檢查表單是否已開啟，若是，則關閉
'   For Each nFrm In Forms
'      If StrComp(nFrm.Name, "frm090202_2", vbTextCompare) = 0 Then
'         'Modify By Sindy 2018/10/12 + if
'         '0.承辦人工作進度:又重新登入此作業需要結束上一個已開啟的歷程作業, 因歷程存檔時會使用到此畫面「詳細資料」
'         If frm090202_2.intReceiveKind = 0 Then
'         '2018/10/12 END
'            Unload frm090202_2
'         End If
'         Exit For
'      End If
'   Next
'   '2018/4/26 END

   InitialField
   MoveFormToCenter Me
   '讀取各基本檔可用系統別
   m_SqlGrpStr1 = SQLGrpStr("", 1)
   m_SqlGrpStr2 = SQLGrpStr("", 2)
   m_SqlGrpStr3 = SQLGrpStr("", 3)
   m_SqlGrpStr4 = SQLGrpStr("", 4)
   m_SqlGrpStr5 = SQLGrpStr("", 5)
   
   Call PUB_GetTMQans("1", True) 'Added by Lydia 2016/06/02 求近似本所案
   Frame1.BorderStyle = 0 'Add By Sindy 2024/6/26
   Frame2.BorderStyle = 0 'Add By Sindy 2024/12/5
   
   ReDim skMail(0) As SeekMails
   
   If PUB_GetST05(strUserNum) = "91" Or PUB_GetST05(strUserNum) = "92" Then
      '指定會稿日
      txt1(18).Enabled = True
   Else
      txt1(18).Enabled = False
   End If
   
   Select Case ProState
   Case "1" '個人
      '讀取使用權限
      Me.Caption = "內商工作進度資料維護 (個人)" 'Add By Sindy 2018/8/13
      bolInsert = IsUserHasRightOfFunction("frm090201_4", strAdd, False)
      bolUpdate = IsUserHasRightOfFunction("frm090201_4", strEdit, False)
      bolDelete = IsUserHasRightOfFunction("frm090201_4", strDel, False)
      bolSelect = IsUserHasRightOfFunction("frm090201_4", strFind, False)
      bolPrint = IsUserHasRightOfFunction("frm090201_4", strPrint, False)
      
      TextOk = True
      '統計年月(個人抓系統日的年月)
      Text1.Text = Mid(strSrvDate(1), 1, 6)
   Case "2" '主管 承辦人管理工作進度資料查詢
      Me.Caption = "內商工作進度資料維護 (主管)" 'Add By Sindy 2018/8/13
      bolInsert = IsUserHasRightOfFunction("frm090614", strAdd, False)
      bolUpdate = IsUserHasRightOfFunction("frm090614", strEdit, False)
      bolDelete = IsUserHasRightOfFunction("frm090614", strDel, False)
      bolSelect = IsUserHasRightOfFunction("frm090614", strFind, False)
      bolPrint = IsUserHasRightOfFunction("frm090614", strPrint, False)
      
      frm090614.TextOk = True
      cmdOK(2).Caption = "回前畫面"
      '統計年月(管理抓查詢畫面輸入的發文年月)
      Text1.Text = Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2))
   Case "3" '分所
      bolInsert = IsUserHasRightOfFunction("frm090201_4", strAdd, False)
      bolUpdate = IsUserHasRightOfFunction("frm090201_4", strEdit, False)
      bolDelete = IsUserHasRightOfFunction("frm090201_4", strDel, False)
      bolSelect = IsUserHasRightOfFunction("frm090201_4", strFind, False)
      bolPrint = IsUserHasRightOfFunction("frm090201_4", strPrint, False)
   Case Else
   End Select
   
   Screen.MousePointer = vbHourglass
   Select Case ProState
   Case "1"
         'Add By Sindy 2018/4/18
         Combo1.AddItem strUserNum & " " & "(" & strUserName & ")", 0
         Combo1.Text = Combo1.List(0)
         '2013/9/17 END
         StrMenu1 'Modify By Sindy 2016/9/6 因前句Combo1就會run 到 StrMenu1
         SetEngineer '設定承辦人選單
         '檢查當時是否需要為他人職代
         Call Pub_SetForOthersEmpCombo(strUserNum, Combo1, False)
         '2018/4/18 END
         'StrMenu1
   Case "2" '承辦人管理工作進度資料查詢
         frm090614.Process2
         StrMenu1
   Case "3"
   Case Else
   End Select
   
   Label3.Caption = GetPrjSalesNM(Trim(Left("" & Combo1.Text, 6)))
   'Add By Sindy 2024/7/30
   If Left(PUB_GetST03(Trim(Left("" & Combo1.Text, 6))), 2) <> "P2" Then '外商
      Me.Caption = Replace(Me.Caption, "內商", "外商")
      Me.Frame1.Visible = False
      Me.Frame2.Visible = True '外文核稿人
   End If
   '2024/7/30 END
   
   SetEngChecker '設定外文核稿人選單 Add By Sindy 2024/12/5
   
   StrMenu
   
   Select Case ProState
   Case "1"
      If TextOk = False Then Screen.MousePointer = iMouse: GoTo EXITSUB
      'Add By Sindy 2018/4/18
      'Combo1.Enabled = False
      Combo1.Enabled = True
      '2018/4/18 END
   Case "2"
      If frm090614.TextOk = False Then Screen.MousePointer = iMouse: TextOk = True: GoTo EXITSUB
      Combo1.Enabled = True
   Case "3"
   Case Else
   End Select
   
   SetGrd1
   'Modify By Sindy 2018/8/23
   'MouseClick (1)
   MouseClick_1 (1)
   '2018/8/23 END
   Screen.MousePointer = iMouse
   SSTab1.Tab = 0
   Me.Combo3.ListIndex = 0
   
   'Add By Sindy 2018/4/26
   SSTab1.Tab = 2
   If QueryData(True) = False Then
      SSTab1.Tab = 0
   End If
   '2018/4/26 END
   
   If bolUpdate = False Then
      cmdOK(1).Visible = False
   End If
   
   Me.txt1(12).Enabled = False 'Added by Lydia 2019/05/03 承辦期限不可修改(應該是沒有人強調,所以都沒限制)
   
   Exit Sub

EXITSUB:
   Me.Hide
   Select Case ProState
   Case "1"
        Me.Hide
   Case "2"
        frm090614.Show
        Me.Hide
   Case "3"
   Case Else
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
   
   'Add by Sindy 2022/12/17 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q") = True Then
      Unload frm090801_Q
   End If
   '2022/12/17 END
   
   Set Fobj = New FileSystemObject
   Fobj.DeleteFile DocTempPath & "\*.doc", False
   ClearFieldList
   Set Fobj = Nothing
   
   Set frm090201_b = Nothing
End Sub

Sub StrMenu1()
Dim ManaGrp As String

Me.Enabled = False
DoEvents
On Error GoTo ErrHnd 'Add By Sindy 2024/3/14
adoEng.Execute "drop table R090614 "
'Modify By Sindy 2015/9/10 +,R110033 text
'Modify By Sindy 2024/1/15 +,R110034 text:指定送件日
RunCreateTable: 'Add By Sindy 2024/3/14
adoEng.Execute "create table R090614 (R110001 text,R110002 double,R110003 text,R110004 text,R110005 text" & _
               ",R110006 text,R110007 text,R110008 text,R110009 text,R110010 text" & _
               ",R110011 text,R110012 text,R110013 text,R110014 text,R110015 text" & _
               ",R110016 text,R110017 text,R110018 text,R110019 double,R110020 memo" & _
               ",R110021 text,R110022 text,ID text,R110023 text, R110024 text,R110025 text" & _
               ",R110026 double,R110027 double,R110028 double,R110029 text,R110030 text" & _
               ",R110031 text,R110032 double,R110033 text,R110034 text)"
On Error GoTo 0 'Add by Sindy 2024/3/14 還原錯誤控制

ManaGrp = ""
CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open "select distinct sg02 from staff_group ,staff where st01='" & Trim(Left("" & Combo1.Text, 6)) & "' and st11=sg01(+) ", cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
       .MoveFirst
       Do While .EOF = False
            ManaGrp = ManaGrp & CheckStr(.Fields(0))
            .MoveNext
            If .EOF = False Then
               ManaGrp = ManaGrp & ","
            End If
      Loop
   End If
End With

Select Case ProState
Case "1" '承辦人個人工作進度資料維護
      StrGrp090201 = ""
      StrSQL6 = ""
      strSQL1 = ""
      strSQL2 = ""
      StrSQL61 = ""
      StrSQL62 = ""
      StrSQL63 = ""
      StrSQL64 = ""
      StrSTM = ""
      StrSLC = ""
      StrSHC = ""
      StrSSP = ""
        
      'Modify By Sindy 2024/7/30 CP14 改判斷 EP05
      StrSQL6 = StrSQL6 & " and EP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and cp05>=19980101 "
      StrSQL61 = StrSQL61 & " and CP158=0 and CP159=0 "
      StrSQL62 = StrSQL62 & " and CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP27<=" & Mid(strSrvDate(1), 1, 6) & "31 "
      StrSQL63 = StrSQL63 & " and CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP57<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp158=0 "
      StrSQL64 = StrSQL64 & " and CP05>=" & Mid(strSrvDate(1), 1, 6) & "01 AND CP05<=" & Mid(strSrvDate(1), 1, 6) & "31 and cp159=0 And CP05>CP27 "
      StrSTM = StrSTM & " and ((tm30>=" & Mid(strSrvDate(1), 1, 6) & "01 AND tm30<=" & Mid(strSrvDate(1), 1, 6) & "31) or tm30 is null) "
      StrSLC = StrSLC & " and ((lc09>=" & Mid(strSrvDate(1), 1, 6) & "01 AND lc09<=" & Mid(strSrvDate(1), 1, 6) & "31) or lc09 is null) "
      StrSHC = StrSHC & " and ((hc10>=" & Mid(strSrvDate(1), 1, 6) & "01 AND hc10<=" & Mid(strSrvDate(1), 1, 6) & "31) or hc10 is null) "
      StrSSP = StrSSP & " and ((sp16>=" & Mid(strSrvDate(1), 1, 6) & "01 AND sp16<=" & Mid(strSrvDate(1), 1, 6) & "31) or sp16 is null) "

Case "2" '承辦人管理工作進度資料查詢
      'StrGrp090201 = frm090614.ManaGrp
      StrGrp090201 = ManaGrp
      '改成收文日要小於等於發文年月當月的最後一天
      StrSQL6 = " and cp05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 "
      strSQL1 = ""
      strSQL2 = ""
      StrSQL61 = ""
      StrSQL62 = ""
      StrSQL63 = ""
      StrSQL64 = ""
      StrSTM = " and ((tm30>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and tm30<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or tm30 is null) "
      StrSLC = " and ((lc09>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and lc09<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or lc09 is null) "
      StrSHC = " and ((hc10>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and hc10<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or hc10 is null) "
      StrSSP = " and ((sp16>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and sp16<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31) or sp16 is null) "
      If frm090614.txt1(8) = "N" Then
         'Modify By Sindy 2024/7/30 CP14 改判斷 EP05
         StrSQL6 = StrSQL6 & " and EP05 IN (" & Combo1_String & ")  and cp05>=19980101 "
         StrSQL61 = StrSQL61 & " and CP158=0 and CP159=0 "
         StrSQL62 = StrSQL62 & " and CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 "
         StrSQL63 = StrSQL63 & " and CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP158=0 "
         StrSQL64 = StrSQL64 & " and CP05>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp159=0 And CP05>CP27 "
      Else
         'Modify By Sindy 2024/7/30 CP14 改判斷 EP05
         StrSQL6 = StrSQL6 & " and EP05='" & Trim(Left("" & Combo1.Text, 6)) & "' and cp05>=19980101 "
         StrSQL61 = StrSQL61 & " and CP158=0 and CP159=0 "
         StrSQL62 = StrSQL62 & " and CP27>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 "
         StrSQL63 = StrSQL63 & " and CP57>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP158=0 "
         StrSQL64 = StrSQL64 & " and CP05>=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "01 and CP05<=" & Trim((Val(frm090614.txt1(3)) + 1911)) & Trim(Right(ChgNumByNick(frm090614.txt1(4)), 2)) & "31 and cp159=0 And CP05>CP27 "
      End If
Case Else
End Select

CheckOC

'Modify By Sindy 2015/9/10 增加讀取cp140
'第一次
'Modify By Sindy 2018/8/1 當月資料剔除核准1001(已發文才剔除)及註冊證1701資料(+  and cp10<>'1701' and not(cp10='1001' and cp158>0))
'Modify By Sindy 2024/1/15 +,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142
'Modify By Sindy 2024/7/30 CP14 改判斷 EP05; CP09=EP02(+)=>CP09(+)=EP02: 查詢速度才不會慢
strSql = "SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL61 & StrSTM & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
strSql = strSql + " UNION SELECT EP05,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL61 & StrSLC & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL61 & StrSHC & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL61 & StrSSP & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
'第二次
'Modify By Sindy 2018/8/1 當月資料剔除核准1001(已發文才剔除)及註冊證1701資料(+  and cp10<>'1701' and not(cp10='1001' and cp158>0))
'Modify By Sindy 2024/1/15 +,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142
'Modify By Sindy 2024/7/30 CP14 改判斷 EP05; CP09=EP02(+)=>CP09(+)=EP02: 查詢速度才不會慢
strSql = strSql + " UNION SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL62 & StrSTM & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
strSql = strSql + " UNION SELECT EP05,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL62 & StrSLC & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL62 & StrSHC & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL62 & StrSSP & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
'第三次
'Modify By Sindy 2018/8/1 當月資料剔除核准1001(已發文才剔除)及註冊證1701資料(+  and cp10<>'1701' and not(cp10='1001' and cp158>0))
'Modify By Sindy 2024/1/15 +,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142
'Modify By Sindy 2024/7/30 CP14 改判斷 EP05; CP09=EP02(+)=>CP09(+)=EP02: 查詢速度才不會慢
strSql = strSql + " UNION SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL63 & StrSTM & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
strSql = strSql + " UNION SELECT EP05,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL63 & StrSLC & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL63 & StrSHC & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09=EP02(+) AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL63 & StrSSP & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
'第四次
'Modify By Sindy 2018/8/1 當月資料剔除核准1001(已發文才剔除)及註冊證1701資料(+  and cp10<>'1701' and not(cp10='1001' and cp158>0))
'Modify By Sindy 2024/1/15 +,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142
'Modify By Sindy 2024/7/30 CP14 改判斷 EP05; CP09=EP02(+)=>CP09(+)=EP02: 查詢速度才不會慢
strSql = strSql + " UNION SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),NVL(TM05,NVL(TM06,TM07)),CP26,decode(tm10,'000',ptm03,ptm04),nvl(DECODE(TM10,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,TM10,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,TM30)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,trADEMARK,CASEPROPERTYMAP,PATENTTRADEMARKMAP,NATION " & _
         " WHERE CP09(+)=EP02 AND cp01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '2'=PTM01(+) AND tm08=PTM02(+) AND TM10=NA01(+) " & StrSQL6 & StrSQL64 & StrSTM & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 2) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
strSql = strSql + " UNION SELECT EP05,ep01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),NVL(LC05,NVL(LC06,LC07)),CP26,'',nvl(DECODE(lC15,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,LC15,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,LC09)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,LAWCASE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09(+)=EP02 AND CP01=lc01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND LC15=NA01(+) " & StrSQL6 & StrSQL64 & StrSLC & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 3) & ") "
strSql = strSql + " UNION SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,CP26,'',nvl(CPM03,cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,'000',CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,hc10)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,HIRECASE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09(+)=EP02 AND cp01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '000'=NA01(+) " & StrSQL6 & StrSQL64 & StrSHC & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 4) & ") "
strSql = strSql + " UNION SELECT EP05,EP01,SUBSTR(CP09,1,1),SQLDateT2(CP05),CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),NVL(SP05,NVL(SP06,SP07)),CP26,'',nvl(DECODE(SP09,'000',CPM03,CPM04),cp10),SQLDateT2(CP48),SQLDateT2(CP06),SQLDateT2(CP07),SQLDateT2(EP06),SQLDateT2( EP09 ),SQLDateT2(EP07),nvl(S3.ST02,ep04),SQLDateT2( EP08 ),SQLDateT2(CP27),Nvl(EP35,0),' '||EP12,nvl(S2.ST02,cp13),CP09,S1.ST05,CP01,SP09,CP10,NVL(NA03,NA04),SQLDateT2(NVL(CP57,sp16)), S1.ST02,cp97,cp98,cp111,ep34,cp112,SQLDateT2( EP28) ep28,cp140,SQLDateT2(CP142)||decode(CP164,'1','當天','2','之前','3','之後',cp164) CP142" & _
         " FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,SERVICEPRACTICE,CASEPROPERTYMAP,NATION " & _
         " WHERE CP09(+)=EP02 AND cp01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND EP05=S1.ST01(+) AND CP13=S2.ST01(+) AND EP04=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND SP09=NA01(+) " & StrSQL6 & StrSQL64 & StrSSP & _
         " AND CP01 IN (" & SQLGrpStr(StrGrp090201, 5) & ") and cp10<>'1701' and not(cp10='1001' and cp158>0) "
AddToMdb (strSql)
Me.Enabled = True

'Add By Sindy 2024/3/14
Exit Sub

ErrHnd:
   GoTo RunCreateTable
'2024/3/14 END
End Sub

Sub AddToMdb(oStrSQL As String)
Dim strCP09s As String

CheckOC
With adoRecordset
   .CursorLocation = adUseClient
   .Open oStrSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 And .RecordCount > 0 Then
      .MoveFirst
      k = 0
      strCP09s = "''"
      Do While .EOF = False
         strCP09s = strCP09s & ",'" & .Fields("cp09") & "'"
         For i = 0 To 26
            strTemp(i) = CheckStr(.Fields(i))
            If Len(strTemp(i)) = 8 Then
               If Mid(strTemp(i), 3, 1) = "/" And Mid(strTemp(i), 6, 1) = "/" Then
                  strTemp(i) = " " & strTemp(i)
               End If
            End If
         Next i
         'Modify By Sindy 2015/9/10 +,'" & .Fields("cp140").Value & "'
         'Modify by Sindy 2024/1/15 +CP142
         strSql = "INSERT INTO R090614 VALUES ('" & strTemp(0) & "'," & Val(strTemp(1)) & ",'" & strTemp(2) & "','" & strTemp(3) & "','" & strTemp(4) & "','" & ChgSQL(strTemp(5)) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "','" & strTemp(9) & "','" & strTemp(10) & "','" & strTemp(11) & "','" & strTemp(12) & "','" & strTemp(13) & "','" & strTemp(14) & "','" & strTemp(15) & "','" & strTemp(16) & "','" & strTemp(17) & "'," & Val(strTemp(18)) & ",'" & strTemp(19) & "','" & strTemp(20) & "','" & strTemp(21) & "','" & strUserNum & "','" & "" & .Fields(26).Value & "','" & .Fields(27).Value & "','" & .Fields(28).Value & "'," & Val("" & .Fields("cp97")) & "," & Val("" & .Fields("cp98")) & "," & Val("" & .Fields("cp111")) & ",'" & "" & .Fields("ep34").Value & "','" & "" & .Fields("cp112").Value & "','" & .Fields("ep28").Value & "',0,'" & .Fields("cp140").Value & "','" & .Fields("CP142").Value & "') "
         adoEng.Execute strSql
         .MoveNext
      Loop
   End If
End With
CheckOC
End Sub

Sub ChgGrdColor(Optional iRow As Integer = -1)
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim ColorFlag As String
Dim iStart As Integer, iEnd As Integer
Dim i As Integer 'Add By Sindy 2024/3/14

With GRD1
   If iRow >= 0 Then
      iStart = iRow
      iEnd = iRow
   Else
      iStart = 1
      iEnd = .Rows - 1
   End If
   For i = iStart To iEnd
      DoEvents
      'Add By Sindy 2024/3/14
      If i > GRD1.Rows - 1 Or i < 0 Then
         MsgBox "不正確的資料列值。( .Row=" & i & " )"
         Exit Sub
      ElseIf i = 1 And iEnd = 1 Then
         .row = i
         If .Text = "" Then
            Exit Sub
         End If
      End If
      '2024/3/14 END
      
      .row = i
      .col = 21 '承辦人備註
'      ColorFlag = Mid(.Text, 1, 1)
      '.Text = Mid(.Text, 2)
      .Text = Mid(.Text, 1)
'      If ColorFlag = "1" And Left(Pub_StrUserSt15, 2) <> "P2" Then
'         .col = 4
'         .CellBackColor = QBColor(10) '淡綠色
'      End If
      .col = 24
      Tmp003 = Trim(.Text)
      '若有取消收文日期
      If Tmp003 <> "" Then
         '灰色
         .col = 3
         .CellBackColor = QBColor(8)
         'Modify By Sindy 2024/7/31
         .col = 9
         .CellBackColor = QBColor(8)
         '2024/7/31 END
         .col = 10
         .CellBackColor = QBColor(8)
         .col = 11
         .CellBackColor = QBColor(8)
         .col = 13
         .CellBackColor = QBColor(8)
      Else
         If .TextMatrix(i, 1) & SystemNumber(.TextMatrix(i, 3), 1) <> "CP" And .TextMatrix(i, 25) <> "N" Then
            'Modify By Sindy 2024/7/31
            .col = 11 '10 承辦期限
            Tmp001 = Trim(.Text)
            .col = 16
            Tmp002 = Trim(.Text)
            .col = 24
            Tmp003 = Trim(.Text)
            '若有承辦期限, 無會稿日及取消收文日期
            If Tmp001 <> "" And Tmp002 = "" And Tmp003 = "" Then
               If Val(ChangeTStringToWString(ChangeTDateStringToTString(Tmp001))) < Val(strSrvDate(1)) Then
                  '黃色
                  .col = 3
                  .CellBackColor = &H80FFFF
                  'Modify By Sindy 2024/7/31
                  .col = 9
                  .CellBackColor = &H80FFFF
                  '2024/7/31 END
                  .col = 10
                  .CellBackColor = &H80FFFF
                  .col = 11
                  .CellBackColor = &H80FFFF
                  .col = 13
                  .CellBackColor = &H80FFFF
               End If
            Else
               '若是有會稿日，且過承辦期限，給淡黃色
               If Tmp001 <> "" And Tmp002 <> "" And Tmp003 = "" Then
                  If Val(ChangeTStringToWString(ChangeTDateStringToTString(Tmp001))) < Val(ChangeTStringToWString(ChangeTDateStringToTString(Tmp002))) Then
                     '淡黃色
                     .col = 3
                     .CellBackColor = &HC0FFFF
                     'Modify By Sindy 2024/7/31
                     .col = 9
                     .CellBackColor = &HC0FFFF
                     '2024/7/31 END
                     .col = 10
                     .CellBackColor = &HC0FFFF
                     .col = 11
                     .CellBackColor = &HC0FFFF
                     .col = 13
                     .CellBackColor = &HC0FFFF
                   End If
               End If
            End If
         End If
         .col = 19
         '若無發文日
         If .Text = "" Then
            .col = 9
            '若系統日大於等於本所期限且本所期限有值(逾本所期限未發文)
            If Val(ChangeTStringToWString(ChangeTDateStringToTString(Trim(.Text)))) <= Val(strSrvDate(1)) And Trim(.Text) <> "" Then
               '淺紅色
               .col = 3
               .CellBackColor = &HC0C0FF
               'Modify By Sindy 2024/7/31
               .col = 9
               .CellBackColor = &HC0C0FF
               '2024/7/31 END
               .col = 10
               .CellBackColor = &HC0C0FF
               .col = 11
               .CellBackColor = &HC0C0FF
               .col = 13
               .CellBackColor = &HC0C0FF
            End If
         '若有發文日
         Else
            .col = 13
            Tmp001 = Trim(.Text)
            .col = 14
            Tmp002 = Trim(.Text)
            .col = 16
            Tmp003 = Trim(.Text)
            .col = 18
            Tmp004 = Trim(.Text)
            .col = 24
            If (Tmp001 = "" Or Tmp002 = "" Or Tmp003 = "" Or Tmp004 = "") Then
               .col = 18
               .Text = " ******"
            End If
         End If
      End If
   Next i
   'Modify By Sindy 2024/11/7 mark
'   '預設目前在第一筆的位置
'   With Me.grd1
'      .row = 1
'      .col = 0
'      .CellBackColor = &HFFC0C0
'      .col = 12
'      .CellBackColor = &HFFC0C0
'      SWPColor2 = SWPColor
'      SWPRow2 = .row
'   End With
   SetGrd1
End With
End Sub

Sub StrMenu()
Dim iMouse As Integer
iMouse = Screen.MousePointer
Select Case ProState
Case "1"
      'Modify By Sindy 2015/9/10 增加CP140電子表單單號做排序,自動收文在前面
      'Modify By Sindy 2018/8/1 + AND R110018='' and (R110024='' or R110024='0') : 進入後只出現未發文案件
      'Modify By Sindy 2023/2/8 原案件清單排列,以電子收文為優先
      '但現皆為電子收文案件,會導致當日智慧局來文不易被發現,請調整為以"目次"由大至小排列
      '取消 R110033 desc,
      'Modify By Sindy 2024/1/15 +,R110034
      strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110034,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032" & _
               " FROM R090614 " & _
               " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "' AND R110018='' and (R110024='' or R110024='0')" & _
               " ORDER BY R110002 desc,R110003,R110004 "

Case "2"
      'Modify By Sindy 2024/1/15 +,R110034
      If frm090614.txt1(8) = "N" Then 'N：不區分個人
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110034,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032" & _
                  " FROM R090614 " & _
                  " WHERE ID='" & strUserNum & "' AND R110001 IN (" & Combo1_String & ") ORDER BY R110005,R110002 desc,R110004 "
      Else
         'Modify By Sindy 2015/9/10 增加CP140電子表單單號做排序,自動收文在前面
         'Modify By Sindy 2023/2/8 原案件清單排列,以電子收文為優先
         '但現皆為電子收文案件,會導致當日智慧局來文不易被發現,請調整為以"目次"由大至小排列
         '取消 R110033 desc,
         strSql = "SELECT R110002,R110003,R110004,R110005,R110006,R110023,R110008,R110009,R110007,R110011,R110034,R110010,R110012,R110013,R110014,R110031,R110015,R110016,R110017,R110018,R110019,R110020,R110021,R110022, R110024,r110029,r110025,r110030,R110032" & _
                  " FROM R090614 " & _
                  " WHERE ID='" & strUserNum & "' AND R110001='" & Trim(Left("" & Combo1.Text, 6)) & "'" & _
                  " ORDER BY R110002 desc,R110003,R110004 "
      End If
Case "3"
Case Else
End Select
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        If ProState = "2" Then
            InsertQueryLog (.RecordCount)
        End If
        Set GRD1.Recordset = adoRecordset
        ChkNoData = False
    Else
        If ProState = "2" Then
            InsertQueryLog (0)
        End If
        ChkNoData = True
        GRD1.Clear
        GRD1.Rows = 2
        Screen.MousePointer = iMouse
        Exit Sub
    End If
End With
CheckOC
ChgGrdColor
    SWPRow2 = "1"
    GRD1.row = Val(SWPRow2)
    GRD1.col = 1
End Sub

Private Sub SetGrd1()
With GRD1
    .Visible = False
    .Cols = 29
    .row = 0
    .col = 0:   .Text = "目次"
    .ColWidth(0) = 350
    .CellAlignment = flexAlignCenterCenter
    .col = 1:   .Text = "收文類別"
    .ColWidth(1) = 200
    .CellAlignment = flexAlignCenterCenter
    .col = 2:   .Text = "收文日"
    .ColWidth(2) = 795
    .ColAlignment(2) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 3:   .Text = "本所案號"
    .ColWidth(3) = 1005
    .CellAlignment = flexAlignCenterCenter
    .col = 4:   .Text = "案件名稱"
    .ColWidth(4) = 1155
    .CellAlignment = flexAlignCenterCenter
    .col = 5:   .Text = "國家"
    .ColWidth(5) = 450
    .CellAlignment = flexAlignCenterCenter
    .col = 6:   .Text = "種類"
    .ColWidth(6) = 450
    .CellAlignment = flexAlignCenterCenter
    .col = 7:   .Text = "案件性質"
    .ColWidth(7) = 795
    .CellAlignment = flexAlignCenterCenter
    .col = 8:   .Text = "Y/N"
    .ColWidth(8) = 285
    .CellAlignment = flexAlignCenterCenter
    .col = 9:   .Text = "本所期限"
    .ColWidth(9) = 795
    .ColAlignment(9) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    'Modify By Sindy 2024/1/15
    .col = 10:   .Text = "指定送件日"
    .ColWidth(10) = 1000
    .ColAlignment(10) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 11:  .Text = "承辦期限"
    .ColWidth(11) = 795
    .ColAlignment(11) = flexAlignRightCenter
    .CellAlignment = flexAlignLeftCenter
    '2024/1/15 END
    .col = 12:  .Text = "法定期限"
    .ColWidth(12) = 0
    .ColAlignment(12) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 13:  .Text = "齊備日"
    .ColWidth(13) = 795
    .ColAlignment(13) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 14:  .Text = "完稿日"
    .ColWidth(14) = 795
    .ColAlignment(14) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 15:  .Text = "指會日"
    .ColWidth(15) = 795
    .ColAlignment(15) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 16:  .Text = "會稿日"
    .ColWidth(16) = 795
    .ColAlignment(16) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 17:  .Text = "核稿人"
    .ColWidth(17) = 700
    .CellAlignment = flexAlignCenterCenter
    .col = 18:  .Text = "會稿完成日"
    .ColWidth(18) = 795
    .ColAlignment(18) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 19:  .Text = "發文日"
    .ColWidth(19) = 795
    .ColAlignment(19) = flexAlignRightCenter
    .CellAlignment = flexAlignCenterCenter
    .col = 20:  .Text = "承辦天數"
    .ColWidth(20) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 21:  .Text = "備註"
    .ColWidth(21) = 2000
    .CellAlignment = flexAlignCenterCenter
    .col = 22:  .Text = "智權人員" 'R110021
    .ColWidth(22) = 800
    .CellAlignment = flexAlignCenterCenter
    .col = 23:  .Text = "CP09" 'R110022
    .ColWidth(23) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 24:  .Text = "" 'SQLDateT2(NVL(CP57,TM30)) R110024
    .ColWidth(24) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 25:  .Text = "EP34" 'R110029
    .ColWidth(25) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 26:  .Text = "承辦人" 'R110025
    .ColWidth(26) = 0
    .CellAlignment = flexAlignCenterCenter
    .col = 27:  .Text = "CP112" 'R110030 --cp112
    .ColWidth(27) = 0
    .CellAlignment = flexAlignCenterCenter
    For intI = 28 To .Cols - 1
      .ColWidth(intI) = 0
    Next
    .Visible = True
End With
'Modify By Sindy 2024/11/7 mark
'   '預設目前在第一筆的位置
'   With Me.grd1
'      .row = 1
'      .col = 0
'      .CellBackColor = &HFFC0C0
'      .col = 12
'      .CellBackColor = &HFFC0C0
'      SWPColor2 = SWPColor
'      SWPRow2 = .row
'   End With
End Sub

'Add By Sindy 2024/1/16
Public Function GetValue(pRow As Integer, pFieldName As String) As String
   Dim ii As Integer
   With Me.GRD1
   For ii = 0 To .Cols - 1
      If UCase(.TextMatrix(0, ii)) = UCase(pFieldName) Then
         If pRow = 0 Then
            '回傳第幾欄
            GetValue = ii
         Else
            '回傳欄位內容
            GetValue = .TextMatrix(pRow, ii)
         End If
         Exit For
      End If
   Next
   End With
End Function

Private Sub GRD1_DblClick()
    If Me.GRD1.MouseRow > 0 Then
        '若有資料
        If Me.GRD1.Rows > 1 Then
            SWPRow = str(GRD1.MouseRow)
            '若點選的那筆無資料, 則退出函式
            If Me.GRD1.TextMatrix(SWPRow, 1) = "" Then Exit Sub
            SSTab1.Tab = 1
        End If
    End If
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Strindex As Integer
Dim iMouse As Integer

iMouse = Screen.MousePointer

If Me.GRD1.MouseRow <= 0 Then Exit Sub
If Button = 1 Then
    Screen.MousePointer = vbHourglass
    SWPRow = str(GRD1.MouseRow)
    Strindex = SWPRow
    With GRD1
        DoEvents
        .Visible = False
        If SWPRow2 <> "" Then
           .row = SWPRow2
           .col = 0
           .CellBackColor = QBColor(15)
           .col = 12
           .CellBackColor = QBColor(15)
        End If
        .col = 0
        If Strindex <> 0 Then
            .row = Strindex
        Else
            .row = .MouseRow
        End If
        If .row = 0 Then
            .row = 1
        End If
        .col = 0
        .CellBackColor = &HFFC0C0
        .col = 12
        .CellBackColor = &HFFC0C0
        SWPColor2 = SWPColor
        SWPRow2 = .row
        .Visible = True
    End With
    Screen.MousePointer = iMouse
End If
End Sub

Sub MouseClick(Optional Strindex As Integer = 0)
    Dim iMouse As Integer
    
    iMouse = Screen.MousePointer
    
    Screen.MousePointer = vbHourglass
    With GRD1
        DoEvents
        .Visible = False
        If SWPRow2 <> "" Then
           .row = SWPRow2
           .col = 0
           .CellBackColor = QBColor(15)
           .col = 12
           .CellBackColor = QBColor(15)
        End If
        .col = 0
        If Strindex <> 0 Then
            .row = Strindex
        Else
            .row = .MouseRow
        End If
        If .row = 0 Then
            .row = 1
        End If
        .col = GetValue(0, "CP09") '23
        Process (.Text)
        .col = 0
        .CellBackColor = &HFFC0C0
        .col = 12
        .CellBackColor = &HFFC0C0
        SWPColor2 = SWPColor
        SWPRow2 = .row
        .Visible = True
    End With
    Screen.MousePointer = iMouse
End Sub

'存檔時使用
Sub MouseClick_1(Optional Strindex As Integer = 0)
Dim iMouse As Integer
   
   iMouse = Screen.MousePointer
   
   Screen.MousePointer = vbHourglass
   With GRD1
       DoEvents
       .Visible = False
       If SWPRow2 <> "" Then
          .row = SWPRow2
          .col = 0
          .CellBackColor = QBColor(15)
          .col = 12
          .CellBackColor = QBColor(15)
       End If
       .col = 0
       If Strindex <> 0 Then
           .row = Strindex
       Else
           .row = .MouseRow
       End If
       If .row = 0 Then
           .row = 1
       End If
       .col = 0
       .CellBackColor = &HFFC0C0
       .col = 12
       .CellBackColor = &HFFC0C0
       SWPColor2 = SWPColor
       SWPRow2 = .row
       .Visible = True
   End With
   
   Screen.MousePointer = iMouse
End Sub

Private Sub grd1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.GRD1.MouseRow < 1 Then
        Select Case Me.GRD1.MouseCol
        Case 0
            If m_blnColOrderAsc = True Then
                Me.GRD1.Sort = 3 '昇冪
                m_blnColOrderAsc = False
            Else
                Me.GRD1.Sort = 4 '降冪
                m_blnColOrderAsc = True
            End If
        Case Else
            If m_blnColOrderAsc = True Then
                Me.GRD1.Sort = 5 '昇冪
                m_blnColOrderAsc = False
            Else
                Me.GRD1.Sort = 6 '降冪
                m_blnColOrderAsc = True
            End If
        End Select
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
 Dim aD1 As Integer 'Added by Lydia 2016/05/06
Dim ii As Integer 'Add By Sindy 2018/4/26
   
   'Add By Sindy 2018/4/26
   If SSTab1.Tab = 2 Then
      Call QueryData(True)
   Else
      Call QueryData(False)
   End If
   If PreviousTab = 2 Then
      '若有資料
      If (Me.grd2.Rows - 1) < dblPrevRow Then dblPrevRow = 0 'Add By Sindy 2018/10/2
      If Me.grd2.Rows > 1 And dblPrevRow > 0 Then
         If Me.grd2.TextMatrix(dblPrevRow, 1) <> "" Then
            For i = 1 To Me.GRD1.Rows - 1
               If Me.grd2.TextMatrix(dblPrevRow, 1) = Me.GRD1.TextMatrix(i, 0) Then
                  SWPRow = i
                  Exit For
               End If
            Next i
            MouseClick Val(SWPRow)
            If SSTab1.Tab = 1 Then
               SSTab1.Tab = 1
            End If
         End If
      End If
   End If
   If PreviousTab = 0 Or PreviousTab = 1 Then
      '若有資料
      If (Me.GRD1.Rows - 1) < Val(SWPRow) Then SWPRow = 0 'Add By Sindy 2018/10/2
      If Me.GRD1.Rows > 1 Then
         '若點選的那筆無資料, 則退出函式
         If Me.GRD1.TextMatrix(Val("0" & SWPRow), 1) = "" Then SSTab1.Tab = 0: Exit Sub
         If Val(SWPRow) > 0 Then
            '上一筆資料列清除反白
            If dblPrevRow > 0 And dblPrevRow <= (grd2.Rows - 1) Then
               grd2.col = 0
               grd2.row = dblPrevRow
               grd2.Text = ""
               For ii = 0 To grd2.Cols - 1
                  'Modify By Sindy 2020/11/30
                  If ii <> 4 Then
                  '2020/11/30 END
                     grd2.col = ii
                     If grd2.CellBackColor <> &H8080FF Then
                        grd2.CellBackColor = QBColor(15)
                     End If
                  End If
               Next ii
               dblPrevRow = 0
               Call SetColColor(CStr(dblPrevRow))
            End If
            For i = 1 To Me.grd2.Rows - 1
               If Me.grd2.TextMatrix(i, 1) = Me.GRD1.TextMatrix(Val("0" & SWPRow), 0) Then
                  '目前資料列反白
                  dblPrevRow = i
                  grd2.col = 0
                  grd2.row = dblPrevRow
                  If grd2.TextMatrix(grd2.row, 1) <> "" Then
                     grd2.Text = "V"
                     For ii = 0 To grd2.Cols - 1
                        'Modify By Sindy 2020/11/30
                        If ii <> 4 Then
                        '2020/11/30 END
                           grd2.col = ii
                           If grd2.CellBackColor <> &H8080FF Then
                              grd2.CellBackColor = &HFFC0C0
                           End If
                        End If
                     Next ii
                  End If
                  Exit For
               End If
            Next i
         End If
      End If
   End If
   '2018/4/26 END
   
   'If PreviousTab = 0 Then
   If SSTab1.Tab = 1 Then
      Label6.Visible = False
      cmd(2).Enabled = True 'Added by Lydia 2016/05/13 預設可印承辦單
      '若有資料
      If Me.GRD1.Rows > 1 Then
         '若點選的那筆無資料, 則退出函式
         If Me.GRD1.TextMatrix(Val("0" & SWPRow), 1) = "" Then SSTab1.Tab = 0: Exit Sub
         MouseClick Val(SWPRow)
         SSTab1.Tab = 1
         'Added by Lydia 2015/11/12 新增查名單對應
         cmdTSMap.Visible = False
         'Modified by Lydia 2016/03/28
         'Remove by Lydia 2018/10/11 有可能不是從Promoter->商標處->承辦人作業而來,造成沒有執行檢查(ex.T-217100在10/9覆核結果為近似本所案號,啟動承辦歷程)
         'If strSrvDate(1) >= TMQ電子化啟用日 And TypeName(Tmpfrm090130) <> "Nothing" Then
            'Modified by Lydia 2016/04/25 +TS
            'If m_CP01 = "T" And m_CP10 = "101" Then
            'Modifed by Lydia 2021/11/19 增加737智財協作之T案
            'If (m_CP01 = "T" And m_CP10 = TMQ_T案) Or (m_CP01 = "TS" And m_CP10 = TMQ_TS案) Then
            'Modified by Lydia 2022/07/15 T大陸案之齊備日管控 : 限制臺灣案And m_Country = "000"
            If (m_CP01 = "T" And m_Country = "000" And InStr(TMQ_T案, m_CP10) > 0) Or (m_CP01 = "TS" And InStr(TMQ_TS案, m_CP10) > 0) Then
               'Modified by Lydia 2018/10/11 有可能不是從Promoter->商標處->承辦人作業而來
               'cmdTSMap.Visible = True
               'Modified by Lydia 2021/06/25 改成商標承辦人才彈提醒
               'If TypeName(Tmpfrm090130) <> "Nothing" Then cmdTSMap.Visible = True
               If TypeName(Tmpfrm090130) <> "Nothing" Then
                    cmdTSMap.Visible = True
                    Call ChkTMQmapData(Me.LBL1(3).Caption) 'Added by Lydia 2021/06/25 檢查查名單對照檔和卷宗區TS.menu，有缺就自動補上資料
               'end 2021/06/25
                    '提醒承辦人查名結果
                    intI = 1
                    'Modified by Lydia 2016/05/09 依查名收文對照檔
                    'strExc(0) = "select nvl(min(tmq11),'') mindate from trademarkquery where tmq21='" & Me.LBL1(3).Caption & "' "
                    'Modified by Lydia 2018/03/20 用TMQ20判斷是否已刪除明細(+And nvl(tmq20,'N') = 'N') 'Remove by Lydia 2018/03/21 影響速度
                    strExc(0) = "select nvl(min(tmq11),'') mindate from trademarkquery where tmq01 in (select tqc03 from tmqcasemap where tqc02='" & Me.LBL1(3).Caption & "') "
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    If intI = 1 Then
                       'Modified by Lydia 2016/05/06 只要承辦日期和查覆日期相差20天以上,皆提示
                       'If Not IsNull(RsTemp(0)) And (Val(DBDATE(Me.LBL1(5).Caption)) - Val("" & RsTemp(0))) >= 20 Then
                       'Added by Lydia 2016/05/09
                       If Val("" & RsTemp(0)) > 0 Then
                         'Modified by Lydia 2016/11/28 以系統日和查覆日期來判斷(T-206154);補查本所為承辦人自行查名,不需新增查名單
                         'aD1 = DateDiff("d", AFDate("" & RsTemp(0)), AFDate(DBDATE(Me.LBL1(5).Caption)))
                         aD1 = DateDiff("d", AFDate("" & RsTemp(0)), AFDate(strSrvDate(1)))
                         If Not IsNull(RsTemp(0)) And (aD1 >= 20 Or aD1 <= -20) Then
                            MsgBox "查覆日期距收文日超過20天(平常日),請補查本所!", vbInformation + vbOKOnly
                         End If
                       End If
                    End If
                    '提醒承辦人追蹤覆核結果
                    'Modified by Lydia 2018/12/10 判斷基本權限
                    'If GetCheckTMQ23(Me.lbl1(3).Caption) = False Then
                    If GetCheckTMQ23(Me.LBL1(3).Caption) = False Or cmd(5).Tag = "N" Then
                       cmd(2).Enabled = False
                       cmd(5).Enabled = False 'Add By Sindy 2018/5/23
                    Else
                       cmd(2).Enabled = True
                    End If
               End If 'Added by Lydia 2021/06/15
            'Added by Lydia 2022/07/15
            ElseIf m_CP01 = "T" And m_Country = "020" Then
                'T大陸案之齊備日管控: 大陸案不管控"若文件和查名尚未齊備，則承辦歷程無法使用"
            ElseIf m_CP01 = "TC" And InStr("000,020", m_Country) > 0 And Left(LBL1(3).Caption, 1) = "A" Then
                'TC案之文件齊備日管控: TC案(台灣、大陸)文件尚未齊備，則承辦歷程無法使用
                If cmd(5).Tag = "N" Or Val(txt1(2)) = 0 Then
                    cmd(5).Enabled = False
                End If
            'end 2022/07/15
            End If
         'End If 'Remove by Lydia 2018/10/11
         'end 2015/11/12
         
      End If
   End If
End Sub

'Added by Lydia 2016/03/28 提醒承辦人追蹤覆核結果
'Modified by Lydia 2018/12/10 增加查名是否已齊備bolCP143
Public Function GetCheckTMQ23(ByVal iCp09 As String, Optional ByVal bCP143 As Boolean = False) As Boolean
Dim idx As Integer
Dim rsR As New ADODB.Recordset
Dim strR As String
Dim strA As String 'Added by Lydia 2018/12/10

    GetCheckTMQ23 = False
    
    idx = 1
    
    'Added by Lydia 2019/01/30 增加查名是否已齊備
    If bCP143 = True And DBDATE(LBL1(5)) >= T案收文齊備啟用日 Then
        Call Process(LBL1(3)) '要重新查詢資料
        If Val(textCP143.Text) = 0 Then
             MsgBox "查名尚未齊備 !", vbCritical
             Exit Function
        Else
             MsgBox "查名已齊備 !", vbInformation
        End If
    End If
    'end 2019/01/30
    
    'Modified by Lydia 2016/05/09 依查名收文對照檔
    'strR = "select tmq01,tmq18,tmq22,tmq23,tqd06,tqd09 From trademarkquery,tmqdetail " & _
                "where tmq21='" & iCP09 & "' and tmq18=tqd01(+) and tmq01=tqd02(+) and tqd06 in ('" & TMQ_近似1 & "','" & TMQ_近似2 & "') "
    'Modified by Lydia 2016/05/13 + TQC07 是否已勾選收文
    'Modified by Lydia 2016/07/06 改TQC07
    'strR = "select tmq01,tmq18,tmq22,tmq23,tqd06,tqd09 From trademarkquery,tmqdetail " & _
                "where tmq01 in (select tqc03 from tmqcasemap where tqc02='" & iCP09 & "' and tqc07 is null) and tmq18=tqd01(+) and tmq01=tqd02(+) and tqd06 in ('" & TMQ_近似1 & "','" & TMQ_近似2 & "') "
    'Modified by Lydia 2018/03/20 用TMQ20判斷是否已刪除明細(+And nvl(tmq20,'N') = 'N') 'Remove by Lydia 2018/03/21 影響速度
    If m_CP01 = "T" Then 'Added by Lydia 2024/11/04 因TS案僅就查名結果向客戶端或代理人端提出查名報告; 若查名單歸屬於TS案，無論是否「近似△」或「相同△」皆可進行電子歷程及發文
         '這樣就不用為了發文而先將覆核結果拿掉「近似△」或「相同△」，然後發通知Email。----嘉雯
         strR = "select tmq01,tmq18,tmq22,tmq23,tqd06,tqd09 From trademarkquery,tmqdetail " & _
                     "where tmq01 in (select tqc03 from tmqcasemap where tqc02='" & iCp09 & "' ) " & _
                     "and tmq18=tqd01(+) and tmq01=tqd02(+) and tqd06 in ('" & TMQ_近似1 & "','" & TMQ_近似2 & "') "
         Set rsR = ClsLawReadRstMsg(idx, strR)
         If idx = 1 Then
            strR = TMQ_近似1 & "," & TMQ_近似2
            rsR.MoveFirst
            Do While Not rsR.EOF
               If "" & rsR.Fields("tqd09") = "" Then
                   'Modified by Lydia 2018/10/11 改"列印承辦單"=>"操作歷程"
                   MsgBox "查名結果有與本所案近似或相同,無覆核結果,無法操作歷程!", vbCritical
                   Exit Function
               ElseIf "" & rsR.Fields("tqd09") <> "" And InStr(strR, "" & rsR.Fields("tqd09")) > 0 Then
                   'Modified by Lydia 2018/10/11 改"列印承辦單"=>"操作歷程"
                   MsgBox "查名結果有與本所案近似或相同,覆核結果仍是與本所案近似或相同,無法操作歷程!", vbCritical
                   Exit Function
               End If
               rsR.MoveNext
            Loop
         End If
    End If 'Added by Lydia 2024/11/04
    GetCheckTMQ23 = True
End Function

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Index = 1 Then
      If KeyAscii <> 8 And KeyAscii <> Asc("N") And KeyAscii <> Asc("Y") Then
         KeyAscii = 0
         Beep
      End If
   'Add By Sindy 2018/8/8
   ElseIf Index = 2 Or Index = 3 Or Index = 4 Or Index = 7 Or Index = 8 Or Index = 12 Or Index = 19 Then
      KeyAscii = Pub_NumAscii(KeyAscii)
   '2018/8/8 END
   End If
End Sub

Private Sub txt1_LostFocus(Index As Integer)
'Add by Morgan 2010/9/6 若回第一頁籤時不檢查,否則若有錯誤時會無窮回圈
If Me.SSTab1.Tab = 0 Then Exit Sub

Select Case Index
Case 1 '是否會稿
     Select Case Trim(txt1(1))
     Case "Y", ""
     Case "N"
         'Add By Sindy 2018/4/25
'         Call ChkEP34ToEP07EP08
         txt1_LostFocus (4)
         '2018/4/25 END
     Case Else
         s = MsgBox("是否會稿只能輸入 Y 或 N !!", , "USER 輸入錯誤")
         txt1(1).SetFocus
         txt1(1).SelStart = 0
         txt1(1).SelLength = Len(txt1(1))
         Exit Sub
     End Select
Case 2 '齊備日
'Mark by Lydia 2019/05/02 改到Validate
'     If Len(Txt1(Index)) <> 0 Then
'         If Not ChkWorkDay(ChangeTStringToWString(Txt1(Index))) Then
'            ShowDateErr
'            txt1(Index).SetFocus
'            txt1(Index).SelLength = Len(txt1(Index))
'            Exit Sub
'         End If
''     Else
''        '若未輸入齊備日則清空承辦期限
''        Me.txt1(12).Text = ""
'     End If
Case 3 '完稿日
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
     End If
Case 4 '會稿日
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
     End If
Case 7 '會稿完成日
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
     End If
Case 8 '發文日
     If Len(txt1(Index)) <> 0 Then
        '若發文日為111111則不檢查是否為工作日
        If Me.txt1(Index).Text <> "111111" Then
            If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
               ShowDateErr
               txt1(Index).SetFocus
               txt1(Index).SelLength = Len(txt1(Index))
               Exit Sub
            End If
        End If
        If txt1(1) = "Y" And Len(txt1(4)) = 0 Then txt1(1) = "N" 'Add By Sindy 2019/7/12
     End If
Case 9 '是否通知客戶
     Select Case Trim(txt1(9))
     Case "Y", "N", ""
     Case Else
          s = MsgBox("是否通知客戶只能輸入 Y 或 N !!", , "USER 輸入錯誤")
          txt1(9).SetFocus
          txt1(9).SelStart = 0
          txt1(9).SelLength = Len(txt1(9))
          Exit Sub
     End Select
Case 11 '條款
     checkCP49
Case 12 '承辦期限
   'Add By Sindy 2018/8/8
   'Modify By Sindy 2018/9/25 + 排除(102)延展案
   If Len(txt1(Index)) <> 0 And m_CP10 <> "102" Then
        If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
           ShowDateErr
           txt1(Index).SetFocus
           txt1(Index).SelLength = Len(txt1(Index))
           Exit Sub
        End If
    End If
    '2018/8/8 END
    '若有承辦期限
    If Me.txt1(12).Text <> "" And Me.LBL1(17).Caption <> "" Then
        '若承辦期限大於本所期限
        If Val(txt1(12).Text) > Val(Replace(LBL1(17).Caption, "/", "")) Then
            Me.txt1(12).Text = Replace(LBL1(17).Caption, "/", "")
        End If
    End If
'Add By Sindy 2024/12/5
Case 19 '外文核完日
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
         End If
     End If
Case Else
End Select

'Add By Sindy 2018/4/25
If Index = 1 Or Index = 3 Then
   Select Case Trim(txt1(1))
   Case "N"
'         Call ChkEP34ToEP07EP08
   Case Else
   End Select
End If
'2018/4/25 END
End Sub

Sub ChkTxt(Strindex As String)
    ChkData = False
    '齊備日
    If Strindex = "2" Or Strindex = "" Then
         If Len(txt1(2)) = 0 Then
             If Len(txt1(3)) <> 0 Or Len(txt1(4)) <> 0 Or Len(txt1(8)) <> 0 Then
                 ShowDateRanErr
                 txt1(2).SetFocus
                 txt1_GotFocus (2)
                 Exit Sub
             End If
         End If
    End If
    
    '完稿日
    If Strindex = "3" Or Strindex = "" Then
        If Len(txt1(3)) = 0 Then
            If Len(txt1(4)) <> 0 Or Len(txt1(8)) <> 0 Then
                ShowDateRanErr
                txt1(3).SetFocus
                Exit Sub
            End If
        End If
    End If
    
    '會稿日
    If Strindex = "4" Or Strindex = "" Then
        '無會稿日
        If Len(txt1(4)) = 0 Then
            '有發文日
            If Len(txt1(8)) <> 0 Then
                ShowDateRanErr
                txt1(4).SetFocus
                Exit Sub
            End If
        End If
    End If
    
    '是否通知客戶
    If Strindex = "9" Or Strindex = "" Then
        If Not CheckLengthIsOK(txt1(9), 1) Then
            txt1(9).SetFocus
            txt1_GotFocus (9)
            Exit Sub
        End If
    End If
    
    '承辦備註
    If Strindex = "10" Or Strindex = "" Then
        If Not CheckLengthIsOK(txtEP12, 2000) Then
            txtEP12.SetFocus
            txt1_GotFocus (10)
            Exit Sub
        End If
    End If
    
    '承辦期限
    If Strindex = "12" Or Strindex = "" Then
        If CheckIsTaiwanDate(Me.txt1(12).Text) = False Then
            MsgBox "承辦期限輸入錯誤！", vbExclamation
            Me.txt1(12).SetFocus
            txt1_GotFocus 12
            Exit Sub
        End If
    End If
    
    'Add By Sindy 2024/12/5
    '外文核完日
    If Strindex = "19" Or Strindex = "" Then
        If CheckIsTaiwanDate(Me.txt1(19).Text) = False Then
            MsgBox "外文核完日輸入錯誤！", vbExclamation
            Me.txt1(19).SetFocus
            txt1_GotFocus 19
            Exit Sub
        End If
    End If
    '2024/12/5 END
    
    ChkData = True
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
'Added by Lydia 2019/05/02 從Lostfocus移過來
Case 2
     If Len(txt1(Index)) <> 0 Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            Cancel = True
            Exit Sub
         End If
     End If
'Modified by Lydia 2021/12/23 去掉txt1(10)
Case 3, 4, 8, 9
    '若欄位無資料則不檢查
    If Me.txt1(Index).Text = "" Then Exit Sub
    ChkTxt "" & Index
    If ChkData = False Then
        Cancel = True
        Exit Sub
    End If
    'Add By Sindy 2019/12/3 C類發文日只能小於等於系統日+2個工作天
    'Modify By Sindy 2021/5/24 改判斷不是 系統日 和 19221111 就彈詢問訊息
    If Me.txt1(Index).Enabled = True And Val(Me.txt1(Index).Text) > 0 Then
      If Index = 8 Then '發文日
         If Left(LBL1(3).Caption, 1) = "C" Then
'            If Val(DBDATE(Me.txt1(Index).Text)) > Val(CompWorkDay(3, strSrvDate(1))) Then
'               MsgBox "發文日不可大於系統日+2個工作天！"
'               txt1(8).SetFocus
'               txt1_GotFocus 8
'               Cancel = True
'               Exit Sub
'            End If
            If Not (Val(DBDATE(Me.txt1(Index).Text)) = strSrvDate(1) Or _
                    Val(DBDATE(Me.txt1(Index).Text)) = 19221111) Then
               If MsgBox("確定發文日為 " & ChangeTStringToTDateString(Me.txt1(Index).Text) & " 嗎？" & vbCrLf & vbCrLf & "（注意：不發文應該輸入 11/11/11）", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  txt1(8).SetFocus
                  txt1_GotFocus 8
                  Cancel = True
                  Exit Sub
               End If
            End If
         End If
      End If
    End If
    '2019/12/3 END
    
'指定會稿日
Case 18
     If txt1(Index).Enabled = True And Trim(txt1(Index).Text) <> "" Then
         If ChkWork(ChangeTStringToWString(txt1(Index))) = False Then
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            Cancel = True
            Exit Sub
         End If
         
         If CheckIsTaiwanDate(txt1(Index).Text) = False Then
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            Cancel = True
            Exit Sub
         End If
         
         If txt1(Index) <> txt1(Index).Tag Then
            If Val(DBDATE(txt1(Index))) < Val(strSrvDate(1)) Then
               MsgBox "指定會稿日不可早於系統日！"
               Cancel = True
               Exit Sub
            End If
         End If
     End If
Case 19 '外文核完日
      If txt1(Index).Enabled = True And Trim(txt1(Index).Text) <> "" And txt1(Index).Enabled = True Then
         If Not ChkWorkDay(ChangeTStringToWString(txt1(Index))) Then
            ShowDateErr
            txt1(Index).SetFocus
            txt1(Index).SelLength = Len(txt1(Index))
            Cancel = True
            Exit Sub
         End If
    End If
Case Else
End Select
End Sub

Private Function TxtValidate() As Boolean
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim rsB As New ADODB.Recordset
Dim StrSqlB As String
Dim arrCaseNo '本所案號
Dim ii As Integer
Dim blnMatch As Boolean
Dim tmpBol As Boolean  'Added by Lydia 2019/05/02
TxtValidate = False

''檢查承辦期限
'If Me.txt1(12).Text <> "" And txt1(12).Enabled = True Then
'    If Me.txt1(2).Text = "" Then
'        MsgBox "無齊備日不可輸入承辦期限!!!", vbExclamation + vbOKOnly
'        Exit Function
'    End If
'End If

'add by nickc 2006/10/23 有齊備日的，承辦期限若是空白再抓一次
'Modified by Lydia 2019/05/02 T-217900齊備日輸入5/1(勞動節放假),因為載入資料時直接SetFocus會程式出錯,所以到存檔前才檢查
'If txt1(2).Text <> "" And txt1(12) = "" Then txt1_LostFocus 2
If txt1(2).Text <> "" Then
    tmpBol = False
    Call txt1_Validate(2, tmpBol)
    If tmpBol = True Then
        txt1(2).SetFocus
        txt1_GotFocus 2
        Exit Function
    End If
End If
'end 2019/05/02

'Add By Sindy 2019/12/3
If txt1(8).Text <> "" Then
    tmpBol = False
    Call txt1_Validate(8, tmpBol)
    If tmpBol = True Then
        txt1(8).SetFocus
        txt1_GotFocus 8
        Exit Function
    End If
End If
'2019/12/3 END

'發文日
If Me.txt1(8).Enabled = True And Val(Me.txt1(8).Text) > 0 Then
   'Add By Sindy 2019/7/12
   '若發文日為111111是否通知客戶應為N不通知
   If Me.txt1(8).Text = "111111" Then
       If txt1(9) <> "N" Then
          MsgBox "要通知客戶，發文日不可輸入 111111！", vbInformation
          SSTab1.Tab = 1 'Add By Sindy 2024/3/15
          txt1(8).SetFocus
          txt1_GotFocus 8
          Exit Function
       End If
   'Add By Sindy 2021/1/8
   Else
      If Trim(txt1(9)) = "N" Then
         MsgBox "不通知客戶，發文日只能輸入 111111 ！"
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         txt1(8).SetFocus
         txt1_GotFocus 8
         Exit Function
      End If
      '2021/1/8 END
   End If
   '2019/7/12 END
End If

'若為商標案且申請國家為台灣且案件性質為核駁(1002)或核駁前先行通知(1202), 條款不可空白
'Modify By Sindy 2024/12/17 FCT人員不用在此處輸條款
If Left(PUB_GetST03(Trim(Left("" & Combo1.Text, 6))), 1) <> "F" Then
'2024/12/17 END
   StrSQLa = "Select TM10 From TradeMark Where " & ChgTradeMark(Replace(Me.LBL1(7).Caption, "-", "")) & " "
   StrSQLa = StrSQLa & " Union Select SP09 From ServicePractice Where " & ChgService(Replace(Me.LBL1(7).Caption, "-", "")) & " "
   If rsA.State <> adStateClosed Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
       If "" & rsA.Fields(0).Value = 台灣國家代號 Then
           '判斷案件性質
           StrSqlB = "Select CP10 From CaseProgress Where CP09='" & Me.LBL1(3).Caption & "' "
           If rsB.State <> adStateClosed Then rsB.Close
           rsB.CursorLocation = adUseClient
           rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
           If rsB.RecordCount > 0 Then
               '判斷案件性質, 以決定是否一定要輸條款
               If "" & rsB.Fields(0).Value = "1002" Or "" & rsB.Fields(0).Value = "1202" Or "" & rsB.Fields(0).Value = "1602" Or "" & rsB.Fields(0).Value = "1604" Or "" & rsB.Fields(0).Value = "1606" Then
                   '檢查條款
                   If Me.txt1(11).Text = "" Then
                       MsgBox "請輸入條款!!!", vbExclamation + vbOKOnly
                       SSTab1.Tab = 1 'Add By Sindy 2024/3/15
                       If Me.txt1(11).Enabled = True And txt1(11).Visible = True Then
                           Me.txt1(11).SetFocus
                           If rsB.State <> adStateClosed Then rsB.Close
                           Set rsB = Nothing
                           If rsA.State <> adStateClosed Then rsA.Close
                           Set rsA = Nothing
                           Exit Function
                       End If
                   End If
               End If
           End If
       End If
   End If
End If

'Add By Sindy 2018/4/20
'核稿人不可與承辦人相同
If Combo2.Enabled = True And Val(m_CP27) = 0 Then
   '若核判表有設定核稿人時只可以修改但不可以空白
   If Trim(m_PP04) <> "" And Trim(Left(Combo2.Text, 6)) = "" Then
      MsgBox "核稿人不可空白!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1 'Add By Sindy 2024/3/15
      Combo2.SetFocus
      Exit Function
   End If
   If Trim(Left(Combo2.Text, 6)) <> "" Then
      '增加檢查核稿人是否離職
      If PUB_ChkEmpFlowExists(LBL1(3), EMP_送核) = True And m_EP39 = "" Then
         If ChkStaffST04(Trim(Left(Combo2.Text, 6))) = True Then
            SSTab1.Tab = 1
            Combo2.SetFocus
            Exit Function
         End If
      End If
      If UCase(Trim(Left("" & Combo1.Text, 6))) = UCase(Trim(Left(Combo2.Text, 6))) Then
         MsgBox "核稿人不可與承辦人相同!!!", vbExclamation + vbOKOnly
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         Combo2.SetFocus
         Exit Function
      End If
      'Modify By Sindy 2018/10/1 不鎖權限
'      '只要非系統設定的人員均要檢查權限
'      '承辦人非程序人員時,才需檢查核判權限
'      'Modify By Sindy 2018/9/19 And m_PP03 <> "" ==>有設定核判表的案件性質才要檢查權限
'      If GetStaffDepartment(Trim(Left("" & Combo1.Text, 6))) <> "P22" And m_PP03 <> "" Then
'         If Combo2.Tag <> Combo2.Text And ProState = "1" Then
'            arrCaseNo = Split(Me.Lbl1(7).Caption, "-")
'            If Trim(m_PP04) <> Trim(Left(Combo2.text, 6)) Then
'               If PUB_ChkPromoterReader(IIf(m_PP01 <> "", m_PP01, arrCaseNo(0)), IIf(m_PP03 <> "", m_PP03, m_CP10), "1", Trim(Left(Combo2.text, 6))) = False And _
'                  PUB_ChkPromoterReader(IIf(m_PP01 <> "", m_PP01, arrCaseNo(0)), m_CP10, "1", Trim(Left(Combo2.text, 6))) = False Then
'                  MsgBox "此人無核稿權限，請重新輸入！"
'                  Combo2.SetFocus
'                  Exit Function
'               End If
'            End If
'         End If
'      End If
   End If
End If
If Combo6.Enabled = True And Val(m_CP27) = 0 Then
   '若核判表有設定判發人時只可以修改但不可以空白
   If Trim(m_PP05) <> "" And Trim(Left(Combo6.Text, 6)) = "" Then
      MsgBox "判發人不可空白!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1 'Add By Sindy 2024/3/15
      Combo6.SetFocus
      Exit Function
   End If
   If Trim(Left(Combo6.Text, 6)) <> "" Then
      '增加檢查判發人是否離職
      If PUB_ChkEmpFlowExists(LBL1(3), EMP_送判) = True And _
         PUB_ChkEmpFlowExists(LBL1(3), EMP_判發) = False Then
         If ChkStaffST04(Trim(Left(Combo6.Text, 6))) = True Then
            SSTab1.Tab = 1
            Combo6.SetFocus
            Exit Function
         End If
      End If
      If UCase(Trim(Left(Combo1.Text, 6))) = UCase(Trim(Left(Combo6.Text, 6))) Then
         MsgBox "判發人不可與承辦人相同!!!", vbExclamation + vbOKOnly
         SSTab1.Tab = 1 'Add By Sindy 2024/3/15
         Combo6.SetFocus
         Exit Function
      End If
      'Modify By Sindy 2018/10/1 不鎖權限
'      '當代理狀況時，檢查輸入的判發人是否有判發權限
'      '承辦人非程序人員時,才需檢查核判權限
'      'Modify By Sindy 2018/9/19 And m_PP03 <> "" ==>有設定核判表的案件性質才要檢查權限
'      If GetStaffDepartment(Trim(Left(Combo1.text, 6))) <> "P22" And m_PP03 <> "" Then
'         If Combo6.Tag <> Combo6.Text And ProState = "1" Then
'            arrCaseNo = Split(Me.Lbl1(7).Caption, "-")
'            '只要非系統設定的人員均要檢查權限
'            If Trim(m_PP05) <> Trim(Left(Combo6.text, 6)) Then
'               If PUB_ChkPromoterReader(IIf(m_PP01 <> "", m_PP01, arrCaseNo(0)), IIf(m_PP03 <> "", m_PP03, m_CP10), "2", Trim(Left(Combo6.text, 6))) = False And _
'                  PUB_ChkPromoterReader(IIf(m_PP01 <> "", m_PP01, arrCaseNo(0)), m_CP10, "2", Trim(Left(Combo6.text, 6))) = False Then
'                  For ii = 0 To Me.Combo6.ListCount - 1
'                      blnMatch = False
'                      If Trim(Left(Me.Combo6.List(ii), 6)) = Trim(Left(Me.Combo6.Text, 6)) Then
'                          Me.Combo6.ListIndex = ii
'                          blnMatch = True
'                          Exit For
'                      End If
'                  Next ii
'                  If blnMatch = False Then
'                     MsgBox "此人無判發權限，請重新輸入！"
'                     Combo6.SetFocus
'                     Exit Function
'                  End If
'               End If
'            End If
'         End If
'      End If
   End If
End If
'加入日期檢查
If Trim(txt1(2)) = "" And Trim(txt1(3)) & Trim(txt1(4)) & Trim(txt1(7)) & Trim(txt1(8)) <> "" And _
   txt1(2).Enabled = True Then
    MsgBox "有下列日期，齊備日不能空白！" & vbCrLf & "完稿日、會稿日、會稿完成日、發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(2).SetFocus
    Exit Function
End If
If Trim(txt1(3)) = "" And Trim(txt1(4)) & Trim(txt1(7)) & Trim(txt1(8)) <> "" And _
   txt1(3).Enabled = True Then
    MsgBox "有下列日期，完稿日不能空白！" & vbCrLf & "會稿日、會稿完成日、發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(3).SetFocus
    Exit Function
End If
If Trim(txt1(4)) = "" And Trim(txt1(7)) & Trim(txt1(8)) <> "" And _
   txt1(4).Enabled = True Then
    MsgBox "有下列日期，會稿日不能空白！" & vbCrLf & "會稿完成日、發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(4).SetFocus
    Exit Function
End If
If Trim(txt1(7)) = "" And Trim(txt1(8)) <> "" And txt1(7).Enabled = True Then
    MsgBox "有下列日期，會稿完成日不能空白！" & vbCrLf & "發文日！", vbExclamation + vbOKOnly
    SSTab1.Tab = 1 'Add By Sindy 2024/3/15
    txt1(7).SetFocus
    Exit Function
End If
'2018/4/20 END

'Added by Lydia 2021/12/23 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If txtEP12 <> "" Then
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         txtEP12.SetFocus
         Exit Function
    End If
End If
'end 2021/12/23

Set rsA = Nothing
Set rsB = Nothing
TxtValidate = True
End Function

'Add By Sindy 2021/2/2 確定鍵,多案單筆歷程時,要更新瀏覽資料日期欄位值
Sub StrMenuOneRec_RecvSub(strRecvSub As String)
Dim ii As Integer, intCnt As Integer
Dim PicRs As New ADODB.Recordset
Dim arrID As Variant
   
   arrID = Split(strRecvSub, ",")
   For intCnt = 0 To UBound(arrID)
      For ii = 1 To Me.GRD1.Rows - 1
         '依收文號更新畫面欄位值
         'If Me.grd1.TextMatrix(ii, 23) = arrID(intCnt) Then
         If Me.GRD1.TextMatrix(ii, GetValue(0, "CP09")) = arrID(intCnt) Then
            strSql = "SELECT * from engineerprogress,caseprogress where ep02='" & arrID(intCnt) & "' and ep02=cp09"
            PicRs.CursorLocation = adUseClient
            PicRs.Open strSql, cnnConnection, adOpenStatic, adLockOptimistic
            If PicRs.RecordCount <> 0 Then
               '承辦期限
               If Val("" & PicRs.Fields("cp48")) > 0 Then
                  Me.GRD1.TextMatrix(ii, 11) = ChangeWStringToTDateString("" & PicRs.Fields("cp48"))
               Else
                  Me.GRD1.TextMatrix(ii, 11) = ""
               End If
               '齊備日
               If Val("" & PicRs.Fields("ep06")) > 0 Then
                  Me.GRD1.TextMatrix(ii, 13) = ChangeWStringToTDateString("" & PicRs.Fields("ep06"))
               Else
                  Me.GRD1.TextMatrix(ii, 13) = ""
               End If
               '完稿日
               If Val("" & PicRs.Fields("ep09")) > 0 Then
                  Me.GRD1.TextMatrix(ii, 14) = ChangeWStringToTDateString("" & PicRs.Fields("ep09"))
               Else
                  Me.GRD1.TextMatrix(ii, 14) = ""
               End If
               '指會日
               If Val("" & PicRs.Fields("EP28")) > 0 Then
                  Me.GRD1.TextMatrix(ii, 15) = ChangeWStringToTDateString("" & PicRs.Fields("EP28"))
               Else
                  Me.GRD1.TextMatrix(ii, 15) = ""
               End If
               '會稿日
               If Val("" & PicRs.Fields("EP07")) > 0 Then
                  Me.GRD1.TextMatrix(ii, 16) = ChangeWStringToTDateString("" & PicRs.Fields("EP07"))
               Else
                  Me.GRD1.TextMatrix(ii, 16) = ""
               End If
               '核稿人
               If "" & PicRs.Fields("EP04") <> "" Then
                  Me.GRD1.TextMatrix(ii, 17) = GetPrjSalesNM("" & PicRs.Fields("EP04"))
               Else
                  Me.GRD1.TextMatrix(ii, 17) = ""
               End If
               '會稿完成日
               If Val("" & PicRs.Fields("EP08")) > 0 Then
                  Me.GRD1.TextMatrix(ii, 18) = ChangeWStringToTDateString("" & PicRs.Fields("EP08"))
               Else
                  Me.GRD1.TextMatrix(ii, 18) = ""
               End If
               '發文日
               If Val("" & PicRs.Fields("cp27")) > 0 Then
                  Me.GRD1.TextMatrix(ii, 19) = ChangeWStringToTDateString("" & PicRs.Fields("cp27"))
               Else
                  Me.GRD1.TextMatrix(ii, 19) = ""
               End If
               '承辦備註
               Me.GRD1.TextMatrix(ii, 21) = "" & PicRs.Fields("ep12")
               
               '修正日期欄位排序問題(小於100年的前面補空白)
               For intI = 10 To 21
                  If Len(GRD1.TextMatrix(ii, intI)) = 8 Then
                    If Mid(GRD1.TextMatrix(ii, intI), 3, 1) = "/" And Mid(GRD1.TextMatrix(ii, intI), 6, 1) = "/" Then
                       GRD1.TextMatrix(ii, intI) = " " & GRD1.TextMatrix(ii, intI)
                    End If
                  End If
               Next
               
               ChgGrdColor ii
               PicRs.Close
               Exit For
            End If
            PicRs.Close
         End If
      Next ii
   Next intCnt
   Set PicRs = Nothing
End Sub

Sub StrMenuOneRec(Optional ByVal Strindex As Integer = 1)
Dim ii As Integer
   For ii = 1 To Me.GRD1.Rows - 1
      '若目次相同, 收文號也相同
      'If Me.grd1.TextMatrix(ii, 0) = Me.lbl1(0).Caption And Me.grd1.TextMatrix(ii, 23) = m_strCP09 Then
      If Me.GRD1.TextMatrix(ii, 0) = Me.LBL1(0).Caption And Me.GRD1.TextMatrix(ii, GetValue(0, "CP09")) = m_strCP09 Then
         '承辦期限
         Me.GRD1.TextMatrix(ii, 11) = ChangeTStringToTDateString(Me.txt1(12).Text)
         '齊備日
         Me.GRD1.TextMatrix(ii, 13) = ChangeTStringToTDateString(Me.txt1(2).Text)
         '完稿日
         Me.GRD1.TextMatrix(ii, 14) = ChangeTStringToTDateString(Me.txt1(3).Text)
         '指會日
         Me.GRD1.TextMatrix(ii, 15) = ChangeTStringToTDateString(Me.txt1(18).Text)
         '會稿日
         Me.GRD1.TextMatrix(ii, 16) = ChangeTStringToTDateString(Me.txt1(4).Text)
         '核稿人 Add By Sindy 2018/4/25
         Me.GRD1.TextMatrix(ii, 17) = IIf(Combo2.Text = "", "", IIf(InStr(Combo2, "==>") > 0, Trim(Mid(Combo2, 10)), Trim(Mid(Combo2, 6))))
         '會稿完成日 Add By Sindy 2018/4/25
         Me.GRD1.TextMatrix(ii, 18) = ChangeTStringToTDateString(Me.txt1(7).Text)
         '發文日
         Me.GRD1.TextMatrix(ii, 19) = ChangeTStringToTDateString(Me.txt1(8).Text)
         '承辦備註
         Me.GRD1.TextMatrix(ii, 21) = Me.txtEP12.Text
         
         '修正日期欄位排序問題(小於100年的前面補空白)
         For intI = 10 To 21
            If Len(GRD1.TextMatrix(ii, intI)) = 8 Then
              If Mid(GRD1.TextMatrix(ii, intI), 3, 1) = "/" And Mid(GRD1.TextMatrix(ii, intI), 6, 1) = "/" Then
                 GRD1.TextMatrix(ii, intI) = " " & GRD1.TextMatrix(ii, intI)
              End If
            End If
         Next
         
         ChgGrdColor ii
         Exit For
      End If
   Next ii
   
   SWPRow2 = Strindex
   GRD1.row = Val(SWPRow2)
   GRD1.col = 1
End Sub

' 控制只跟 DB 溝通一次
' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To TF_CP
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "CP" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      '定義數字
      Select Case nIndex
         Case 5, 6, 7, 15, 16, 17, 18, 19, 25, 27, 33, 34, 46, 47, 48, 53, 54, 57, 66, 67, 69, 70, 73, 74, 75, 76, 77, 78, 79, 82, 84, 85, 97, 98, 100, 101, 103, 104, 108, 109, 111:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
End Sub

Private Sub ClearFieldList()
   Erase m_FieldList
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To TF_CP - 1
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   For nIndex = 0 To TF_CP - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

Private Function FormSave() As Boolean
'若勾選無圖式，複製無圖式的圖檔給該案
Dim BytesS() As Byte
Dim BytesVal As String
Dim PicRs As New ADODB.Recordset
Dim iMouse As Integer
Dim p_FileName As String, strFtpPath As String
Dim strTmp As String
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
   
On Error GoTo ErrHand
   
   iMouse = Screen.MousePointer
   If m_Flow = "" Then cnnConnection.BeginTrans
   
   cnnConnection.Execute "begin user_data.user_formname:='" & Me.Name & "';end;"
   
   '若勾選無圖式，複製無圖式的圖檔給該案
   If Chk1.Value = vbChecked Then
      strSql = "select * from ImgByteFile where ibf01='000' and ibf02='000000' and ibf03='0' and ibf04='01'  "
      Set PicRs = New ADODB.Recordset
      PicRs.CursorLocation = adUseClient
      PicRs.Open strSql, cnnConnection, adOpenStatic, adLockOptimistic
      If PicRs.RecordCount <> 0 Then
         BytesVal = PicRs.Fields("ibf13").Value
'         ReDim BytesS(Val(BytesVal))
'         BytesS() = PicRs.Fields("ibf14").GetChunk(Val(BytesVal))
         'Add By Sindy 2017/8/10 下載檔案
         p_FileName = App.path & "\TempFile"
         RidFile p_FileName
         If "" & PicRs.Fields("IBF15") <> "" Then
            If PUB_GetFtpFile(PicRs.Fields("IBF15"), p_FileName, UCase("ImgByteFile")) = False Then
               GoTo ErrHand
            End If
         End If
         '2017/8/10 END
         PicRs.AddNew
         PicRs.Fields("ibf07").Value = strUserNum
         PicRs.Fields("ibf08").Value = Val(strSrvDate(1))
         PicRs.Fields("ibf09").Value = Val(Format(time, "HHMM"))
         PicRs.Fields("ibf01").Value = m_CP01
         PicRs.Fields("ibf02").Value = m_CP02
         PicRs.Fields("ibf03").Value = m_CP03
         PicRs.Fields("ibf04").Value = m_CP04
         PicRs.Fields("ibf05").Value = "1"
         PicRs.Fields("ibf06").Value = "6"
         PicRs.Fields("ibf13").Value = BytesVal
'         PicRs.Fields("ibf14").Value = Null
'         PicRs.Fields("ibf14").AppendChunk BytesS()
         'Modify By Sindy 2017/8/10
         '檔案改放FTP
         If FileExists(p_FileName) Then
            PUB_PutFtpFile p_FileName, m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & "-1", m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & "-1", strFtpPath, UCase("imgbytefile")
            If strFtpPath <> "" Then
               PicRs.Fields("ibf15") = strFtpPath
            End If
         End If
         '2017/8/10 END
         PicRs.UPDATE
      End If
   End If
   
   '目次
   SeekTmpBk = Trim(LBL1(0).Caption)
   'Modify By Sindy 2018/4/20 +EP04,EP40
   'Modify By Sindy 2018/12/25
'   strSql = "Update EngineerProgress Set EP06=" & IIf(ChangeTStringToWString(txt1(2)) = "", "NULL", ChangeTStringToWString(txt1(2))) & ",EP09=" & IIf(ChangeTStringToWString(txt1(3)) = "", "NULL", ChangeTStringToWString(txt1(3))) & ",EP07=" & IIf(ChangeTStringToWString(txt1(4)) = "", "NULL", ChangeTStringToWString(txt1(4))) & _
'      ",EP11='" & txt1(9) & "',EP12='" & txtep12 & "',EP34='" & txt1(1) & "',EP04='" & Trim(Left("" & Combo2.Text, 6)) & "',EP40='" & Trim(Left("" & Combo6.Text, 6)) & "',EP08=" & IIf(ChangeTStringToWString(txt1(7)) = "", "NULL", ChangeTStringToWString(txt1(7))) & _
'      " Where EP02='" & lbl1(3).Caption & "'"
   'Modify By Sindy 2024/5/23 調整檢查欄位有異動再儲存
'****************************************************************
'更新EP
'****************************************************************
   '有預設欄位值
   strSql = "EP04='" & Trim(Left("" & Combo2.Text, 6)) & "',EP40='" & Trim(Left("" & Combo6.Text, 6)) & "'" & _
            ",EP34='" & txt1(1) & "',EP03='" & Trim(Left("" & Combo4.Text, 6)) & "',EP41=" & CNULL(txt1(23))
   If txt1(2).Tag <> txt1(2).Text Then
      Pub_SaveLog strUserNum, "齊備日異動：" & DBDATE(LBL1(8)) & "==>" & DBDATE(txt1(2)) & " ", SystemNumber(LBL1(7).Caption, 1), SystemNumber(LBL1(7).Caption, 2), SystemNumber(LBL1(7).Caption, 3), SystemNumber(LBL1(7).Caption, 4), LBL1(3).Caption
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP06=" & IIf(ChangeTStringToWString(txt1(2)) = "", "NULL", ChangeTStringToWString(txt1(2)))
   End If
   If txtEP12.Tag <> txtEP12 Then
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP12='" & txtEP12 & "'"
   End If
'Modify By Sindy 2019/2/15 改判斷方式
   'If txt1(3).Enabled = True Or Me.m_Flow <> "" Then
   If Val(txt1(3).Tag) <> Val(txt1(3).Text) Then
      Pub_SaveLog strUserNum, "完稿日異動：" & DBDATE(Trim(txt1(3).Tag)) & "==>" & DBDATE(Trim(txt1(3).Text)) & " ", SystemNumber(LBL1(7).Caption, 1), SystemNumber(LBL1(7).Caption, 2), SystemNumber(LBL1(7).Caption, 3), SystemNumber(LBL1(7).Caption, 4), LBL1(3).Caption
      'Modify By Sindy 2022/4/25 針對有註記「收款後送件」的台灣商標案件，開放承辦人於案件先行作業後，可自行輸入「完稿日」。
      '第一次輸入完稿日
      If Me.LblFee.Tag = "尚待收款" And Val(txt1(3).Tag) = 0 And Val(txt1(3).Text) > 0 And Me.m_Flow = "" Then
         Me.LblFee.Tag = "尚待收款-完稿日"
      End If
      '2022/4/25 END
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP09=" & IIf(ChangeTStringToWString(txt1(3)) = "", "NULL", ChangeTStringToWString(txt1(3)))
   End If
   'If txt1(4).Enabled = True Or Me.m_Flow <> "" Then
   If Val(txt1(4).Tag) <> Val(txt1(4).Text) Then
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP07=" & IIf(ChangeTStringToWString(txt1(4)) = "", "NULL", ChangeTStringToWString(txt1(4)))
   End If
   'If txt1(7).Enabled = True Or Me.m_Flow <> "" Then
   If Val(txt1(7).Tag) <> Val(txt1(7).Text) Then
   '2019/2/15 END
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP08=" & IIf(ChangeTStringToWString(txt1(7)) = "", "NULL", ChangeTStringToWString(txt1(7)))
   End If
   '2018/12/25 END
   '指定會稿日異動
   If Val(txt1(18).Tag) <> Val(txt1(18).Text) Then
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP28=" & CNULL(ChangeTStringToWString(txt1(18)))
   End If
   'Add By Sindy 2024/12/5
   '外文核完日
   If Val(txt1(19).Tag) <> Val(txt1(19).Text) Then
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "ep33=" & CNULL(ChangeTStringToWString(txt1(19)))
   End If
   '2024/12/5 END
   'Add By Sindy 2021/1/8
   '加入 是否通知客戶異動時，紀錄
   If txt1(9).Tag <> txt1(9).Text Then
      Pub_SaveLog strUserNum, "是否通知客戶異動：" & txt1(9).Tag & "==>" & txt1(9).Text & " ", SystemNumber(LBL1(7).Caption, 1), SystemNumber(LBL1(7).Caption, 2), SystemNumber(LBL1(7).Caption, 3), SystemNumber(LBL1(7).Caption, 4), LBL1(3).Caption
      strSql = IIf(strSql <> "", strSql & ",", "") & _
               "EP11='" & txt1(9) & "'"
      If Trim(txt1(9)) = "N" Then '不通知客戶
         PUB_NotCusLP LBL1(3).Caption '沒有客戶函,要清空相關欄位值
      End If
   End If
   '2021/1/8 END
   If strSql <> "" Then
      strSql = "Update EngineerProgress Set " & strSql & " Where EP02='" & LBL1(3).Caption & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   '2024/5/23 END
   
'****************************************************************
   '加入 核稿人異動時，紀錄
   If Combo2.Tag <> Combo2.Text Then
      Pub_SaveLog strUserNum, "核稿人異動：" & Combo2.Tag & "==>" & Combo2.Text & " ", SystemNumber(LBL1(7).Caption, 1), SystemNumber(LBL1(7).Caption, 2), SystemNumber(LBL1(7).Caption, 3), SystemNumber(LBL1(7).Caption, 4), LBL1(3).Caption
   End If
   '加入 判發人異動時，紀錄
   If Combo6.Tag <> Combo6.Text Then
      Pub_SaveLog strUserNum, "判發人異動：" & Combo6.Tag & "==>" & Combo6.Text & " ", SystemNumber(LBL1(7).Caption, 1), SystemNumber(LBL1(7).Caption, 2), SystemNumber(LBL1(7).Caption, 3), SystemNumber(LBL1(7).Caption, 4), LBL1(3).Caption
   End If
   'Add By Sindy 2024/12/5
   '加入 外文核稿人異動時，紀錄
   If Combo4.Tag <> Combo4.Text Then
      If txt1(23) = "2" Then
         Pub_SaveLog strUserNum, "日文核稿人異動：" & Combo4.Tag & "==>" & Combo4.Text & " ", m_CP01, m_CP02, m_CP03, m_CP04, LBL1(3).Caption
      Else
         Pub_SaveLog strUserNum, "英文核稿人異動：" & Combo4.Tag & "==>" & Combo4.Text & " ", m_CP01, m_CP02, m_CP03, m_CP04, LBL1(3).Caption
      End If
   End If
   '2024/12/5 END
   
'****************************************************************
'更新商標基本檔
'****************************************************************
   'Add By Sindy 2024/6/12
   strSql = ""
   '商標描述中文
   If Trim(txt1(0).Tag) <> Trim(txt1(0).Text) Then
      strSql = IIf(strSql <> "", strSql & ",", "") & "TM137='" & txt1(0) & "'"
   End If
   '商標描述英文
   If Trim(txt1(6).Tag) <> Trim(txt1(6).Text) Then
      strSql = IIf(strSql <> "", strSql & ",", "") & "TM138='" & txt1(6) & "'"
   End If
   If strSql <> "" Then
      strSql = "Update trademark Set " & strSql & _
               " Where tm01='" & m_CP01 & "' and tm02='" & m_CP02 & "' and tm03='" & m_CP03 & "' and tm04='" & m_CP04 & "'"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   '2024/6/12 END
   
'****************************************************************
'更新進度檔
'****************************************************************
   '若是商標則要回存條款欄位
   If SSTab2.TabVisible(0) = True Then
'      strSql = "Update CaseProgress Set cp49='" & txt1(11).Text & "' where CP09='" & Me.lbl1(3).Caption & "' "
'      cnnConnection.Execute strSql
      SetFieldNewData "CP49", txt1(11).Text
   End If
   If Mid(LBL1(3).Caption, 1, 1) = "C" Then
      '發文日
      If Trim(txt1(8).Tag) <> Trim(txt1(8).Text) Then
         SetFieldNewData "CP27", IIf(Trim(txt1(8)) <> "", ChangeTStringToWString(txt1(8)), "")
      End If
   End If
   '承辦期限
   If Trim(txt1(12).Tag) <> Trim(txt1(12).Text) Then
      SetFieldNewData "CP48", IIf(Trim(txt1(12)) <> "", ChangeTStringToWString(txt1(12)), "")
   End If
   strSql = " UPDATE CASEPROGRESS SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To TF_CP - 1
      strTmp = Empty
      If nIndex < 64 Or nIndex > 69 Then
         If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
            If m_FieldList(nIndex).fiType = 0 Then
               If m_FieldList(nIndex).fiNewData = Empty Then
                  strTmp = m_FieldList(nIndex).fiName & " = NULL "
               Else
                  strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
               End If
            Else
               If m_FieldList(nIndex).fiNewData = Empty Then
                  strTmp = m_FieldList(nIndex).fiName & " = NULL "
               Else
                  strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
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
      End If
   Next nIndex
   strSql = strSql & " " & _
      "WHERE CP09 = '" & Me.LBL1(3).Caption & "' "
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   cnnConnection.Execute "begin user_data.user_formname:=Null;end;"
   
   If m_Flow = "" Then cnnConnection.CommitTrans
   
   FormSave = True
   
   'Modify By Sindy 2022/4/25 為”尚待收款-完稿日”操作時,進入聯絡歷程
   '新增聯絡歷程並且發Mail通知智權人員
   If Me.LblFee.Tag = "尚待收款-完稿日" Then
      Call SetColTag(True)
      If PUB_ChkFormIsClose("frm090202_2") = False Then Exit Function 'Add By Sindy 2020/1/17
      intBackTab = 1
      frm090202_2.Hide
      frm090202_2.m_EEP01 = LBL1(3) '總收文號
      frm090202_2.m_FlowUserNum = Trim(Left("" & Combo1.Text, 6)) '案件流程所屬人員
      frm090202_2.intReceiveKind = 0 '0.承辦人工作進度
      frm090202_2.SetParent Me
      If frm090202_2.QueryData = True Then
         frm090202_2.m_strSpecState = Me.LblFee.Tag
         '只能進入聯絡歷程
         frm090202_2.cmdExit.Enabled = False
         frm090202_2.cmdSend.Enabled = True
         frm090202_2.cmdAdd.Visible = False
         frm090202_2.txtEEP03 = strUserNum
         frm090202_2.txtEEP03_2 = strUserName
         frm090202_2.m_EditMode = 1 '新增
         frm090202_2.CboEEP04.Clear
         frm090202_2.CboEEP04.AddItem EMP_聯絡 & " " & "聯絡"
         frm090202_2.CboEEP04.ListIndex = 0
         frm090202_2.CboEEP04.Enabled = True
         frm090202_2.CboEEP05.Text = frm090202_2.m_SPMan
         frm090202_2.txtEEP08 = "商標案件已完稿，惟案件註記為收款後送件，請儘速收款。"
         '無多案,直接新增聯絡
         Me.LblFee.Tag = "已更新完稿日"
         If frm090202_2.cmdManyCase.Enabled = False And frm090202_2.m_SPMan <> "" Then
            frm090202_2.cmdSend_Click
         Else
            frm090202_2.Show
            Me.Hide
         End If
      End If
'      strSubject = Replace(lbl1(7).Caption, "-0-00", "") & "「" & lbl1(9).Caption & "」-->聯絡"
'      strContent = ""
''      'FC代理人來台
''      If strSubPA75 <> "" And m_Country = "000" Then
''         strContent = "貴方卷號：" & strSubPA77 & vbCrLf
''      End If
'      strContent = strContent & "本所案號：" & Replace(lbl1(7).Caption, "-0-00", "") & vbCrLf
'      strContent = strContent & _
'                   "案件名稱：" & lbl1(9).Caption & vbCrLf & _
'                   "案件性質：" & lbl1(15).Caption & vbCrLf & _
'                   "流程狀態：聯絡" & vbCrLf & _
'                   "內　　容：商標案件已完稿，惟案件註記為收款後送件，請儘速收款" & vbCrLf
''      strContent = MailContentAddEnd(strContent)
'      PUB_SendMail strUserNum, m_CP13, lbl1(3).Caption, strSubject, strContent
'      '取得最大序號
'      intMaxEEP02 = 0
'      strSql = "select eep02 From empelectronprocess where eep01='" & lbl1(3).Caption & "' order by eep02 desc"
'      intI = 1
'      CheckOC3
'      Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         AdoRecordSet3.MoveFirst
'         If AdoRecordSet3.RecordCount > 0 Then
'            intMaxEEP02 = AdoRecordSet3.Fields(0)
'         End If
'      End If
'      strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08,eep11) values(" & _
'               CNULL(lbl1(3).Caption) & "," & intMaxEEP02 + 1 & ",'" & strUserNum & "'," & _
'               CNULL(EMP_聯絡) & "," & _
'               CNULL(m_CP13) & "," & _
'               strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & CNULL(ChgSQL(strContent)) & "," & CNULL(Me.lblFee.Tag) & ")"
'      cnnConnection.Execute strSql
   End If
   '2022/4/25 END
   
   Exit Function
   
ErrHand:
   cnnConnection.Execute "begin user_data.user_formname:=Null;end;"
   If m_Flow = "" Then cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

'批次發Mail
'Modify By Sindy 2018/6/20
'Private Sub BatctMail()
Public Sub BatctMail()
'2018/6/20 END
   Dim i As Integer
   For i = 1 To UBound(skMail)
        PUB_SendMail skMail(i).fiSender, skMail(i).fiReceiver, skMail(i).fiRecriverNo, skMail(i).fiSubject, skMail(i).fiContent
   Next i
   ReDim skMail(0) As SeekMails
   'Trigger 也會產生待發郵件
   PUB_SendMailCache
End Sub

'更新mdb暫存資料及第一畫面的Grid內容
Private Sub UpdEngMdb()

On Error GoTo ErrHand
   
   'R110013.齊備日
   'R110014.完稿日
   'R110015.會稿日
   'R110010.承辦期限
   'R110018.發文日
   'R110020.承辦備註
   strSql = "UPDATE R090614 SET " & _
      "R110013='" & IIf(txt1(2) = "", "", Right(" " & ChangeTStringToTDateString(txt1(2)), 9)) & "'," & _
      "R110014='" & IIf(txt1(3) = "", "", Right(" " & ChangeTStringToTDateString(txt1(3)), 9)) & "'," & _
      "R110015='" & IIf(txt1(4) = "", "", Right(" " & ChangeTStringToTDateString(txt1(4)), 9)) & "'," & _
      "R110017='" & IIf(txt1(7) = "", "", Right(" " & ChangeTStringToTDateString(txt1(7)), 9)) & "'," & _
      "R110016='" & IIf(Combo2.Text = "", "", IIf(InStr(Combo2, "==>") > 0, Trim(Mid(Combo2, 10)), Trim(Mid(Combo2, 6)))) & "'," & _
      "R110010='" & IIf(txt1(12) = "", "", Right(" " & ChangeTStringToTDateString(txt1(12)), 9)) & "'," & _
      "R110018='" & IIf(txt1(8) = "", "", Right(" " & ChangeTStringToTDateString(txt1(8)), 9)) & "'," & _
      "R110020='" & txtEP12 & "' " & _
      " WHERE ID='" & strUserNum & "' AND R110022='" & LBL1(3).Caption & "' "
   adoEng.Execute strSql, intI
   
   m_blnClkSure = True
   For i = 1 To GRD1.Rows - 1
      GRD1.row = i
      GRD1.col = 0
      '若目次相同, 收文號也相同
      'If grd1.Text = SeekTmpBk And Me.grd1.TextMatrix(i, 23) = m_strCP09 Then
      If GRD1.Text = SeekTmpBk And Me.GRD1.TextMatrix(i, GetValue(0, "CP09")) = m_strCP09 Then
         MouseClick_1 (i)
         StrMenuOneRec SWPRow2
         Exit For
      End If
   Next i
   m_blnClkSure = False
      
ErrHand:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'設定承辦人選單
Private Sub SetEngineer()
   strSql = "SELECT Distinct (R110001&' '&'(' & R110025&')') FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Trim(strUserNum) & "' ORDER BY (R110001&' '&'(' & R110025&')') "
   CheckOC
   i = 0
   Combo1.Clear
   Combo1_String = ""
   With adoRecordset
       .CursorLocation = adUseClient
       .Open strSql, adoEng, adOpenStatic, adLockReadOnly
       If .RecordCount <> 0 And .RecordCount > 0 Then
         Do While .EOF = False
           Combo1.AddItem "" & .Fields(0), i
           i = i + 1
           If Combo1_String = "" Then
              Combo1_String = "'" & Trim(Left("" & .Fields(0), 6)) & "'"
           Else
              Combo1_String = Combo1_String + ",'" & Trim(Left("" & .Fields(0), 6)) & "'"
           End If
           .MoveNext
         Loop
         Combo1.Text = Combo1.List(0)
       End If
   End With
End Sub

'設定外文核稿人選單
Private Sub SetEngChecker()
   Combo4.Clear
   Combo4.AddItem "", 0
   If m_EMPST16 = "4" Then '日文
      strExc(0) = "select st01||' ==> '||st02 from staff where st04='1' and st93='J41' and st01<>'99998' order by st01 asc"
   Else
      strExc(0) = "select st01||' ==> '||st02 from staff where st04='1' and st93='F41' and st01<>'99998' order by st01 asc"
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      .MoveFirst
      Do While .EOF = False
          Combo4.AddItem "" & .Fields(0), intI
          intI = intI + 1
          .MoveNext
      Loop
      End With
   End If
   Combo4.Text = Combo4.List(0)
End Sub

'Added by Lydia 2015/11/12 新增查名單對應
Private Sub cmdTSMap_Click()
' iStiu '0:新增收文, 1:修改,  2:查詢
  If TypeName(Tmpfrm090130) <> "Nothing" Then
    'Modified by Lydia 2018/03/20 先傳變數
    'Tmpfrm090130.SetParent Me
    Tmpfrm090130.SetParent Me, m_CP13
    If bolUpdate = False Then
        Tmpfrm090130.iStiu = 2
    Else
        Tmpfrm090130.iStiu = 1
    End If
    Tmpfrm090130.mbolCall = False
    Tmpfrm090130.m_CP09 = LBL1(3).Caption
    'Remove by Lydia 2018/03/20
    'Tmpfrm090130.txtField(0) = m_CP13
    'Tmpfrm090130.lblSname.Caption = LBL1(21).Caption
    Tmpfrm090130.Show
    'Modified by Lydia 2016/05/16 加註案號和收文號
    Tmpfrm090130.Caption = cmdTSMap.Caption & "　(本所案號:" & m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & "　總收文號:" & m_strCP09 & " )"
    Me.Hide
  End If
End Sub
'end 2015/11/12

'Added by Lydia 2019/07/05 產生電子送件申請書
Private Sub GetApplBook_T(ByVal pCaseNo As String, ByVal pCP09 As String, ByVal pCP10 As String)
Dim bolChk As Boolean
Dim tm() As String '商標基本檔
Dim intWhere As Integer
Dim strFolder As String, strFileName As String
Dim ET01 As String, ET03 As String
Dim ET03_1 As String 'Added by Lydia 2020/10/21 基本資料表
Dim ET03type As String
Dim strChkVal As String
Dim m_CaseNo As String
Dim strContent As String
Dim mChkType As String 'Added by Lydia 2023/11/14 確定商標種類：特殊商標TM72>商標種類TM08;
                        '商標註冊種類:要申請的權利主體,商標型態:以立體形狀、聲音等形式呈現，而這些表彰商品或服務來源之標識，為商標法規範之商標「型態」
    ReDim tm(TF_TM)
    Call ChgCaseNo(Replace(pCaseNo, "-", ""), tm)
    intWhere = 國內
    If ClsPDReadTrademarkDatabase(tm(), intWhere) Then
    End If
        
   'Added by Lydia 2023/11/14 確定商標種類：特殊商標TM72>商標種類TM08;
   mChkType = IIf(tm(72) <> "", tm(72), tm(8))
   If mChkType = "" Then mChkType = "1"
   'end 2023/11/14
   'Added by Lydia 2025/01/24 商爭案:相關總收文號為601,603,605都只用主流程的申請書
   If (tm(1) = "T" Or tm(1) = "FCT") And Left(pCP09, 1) = "B" Then
      'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
      strExc(0) = "select c2.cp09,c2.cp10 from caseprogress c1,caseprogress c2 where c1.cp09='" & pCP09 & "' and c1.cp43=c2.cp09(+) and instr('" & TMdebate & "',c2.cp10) > 0 And Not (c2.cp01 = 'FCT' And InStr(" & FCT_NotTMdebate & ", c2.cp10) > 0) "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         pCP09 = "" & RsTemp.Fields("cp09")
         pCP10 = "" & RsTemp.Fields("cp10")
      End If
   End If
   'end 2025/01/24
   
    Screen.MousePointer = vbHourglass
    
    m_CaseNo = PUB_FCPCaseNo2FileName(tm(1), tm(2), tm(3), tm(4))
    '桌面上建立案號資料夾
    'Mark by Lydia 2019/07/10
    'strFolder = PUB_Getdesktop
    'strFolder = strFolder & "\" & m_CaseNo
    'If Dir(strFolder, vbDirectory) = "" Then
    '    MkDir strFolder
    'End If
    
    strLetterDate = strSrvDate(1)
    
    '2.申請書
    ET01 = "90"
    ET03 = ""
    ET03_1 = "00" 'Added by Lydia 2020/10/21 基本資料表:預設共用00
    strExc(0) = ""
    If pCP10 = "101" Then '申請 (處理狀況的編號與FCT案一致)
        'Modified by Lydia 2023/11/14 tm(8)=>mChkType
        Select Case mChkType
            Case "7" '證明標章
                ET03 = "21"
            Case "9" '團體商標
                ET03 = "22"
            Case "A" '立體商標
                ET03 = "23"
            'Modified by Lydia 2023/11/14 "C"=>"B"
            Case "B" '顏色商標
                ET03 = "24"
            'Added by Lydia 2023/11/30 增加各商標種類的電子送件申請書
            Case "8"  '團體標章
                ET03 = "28"
            Case "C"  '聲音商標
                ET03 = "29"
            Case "D", "E", "F" '其他商標
                ET03 = "30"
            Case "G"  '動態商標
                ET03 = "31"
            Case "H"  '全像圖商標
                ET03 = "32"
            Case "I"  '立體團體商標
                ET03 = "33"
            Case "J"  '顏色團體商標
                ET03 = "34"
            Case "K"  '聲音團體商標
                ET03 = "35"
            'end 2023/11/30
            Case Else '商標，含未設定種類
                ET03 = "20"
        End Select
        '商標種類
        'Modified by Lydia 2023/11/14
        'Call ClsPDGetPatentTrademarkKind(商標, IIf(ET03 = "00", "1", tm(8)), strExc(0), False)
        Call ClsPDGetPatentTrademarkKind(商標, mChkType, strExc(0), False)
        ET03type = strExc(0) & "註冊"
         
    ElseIf pCP10 = "102" Or pCP10 = "103" Then '延展 (處理狀況的編號與FCT案一致)
        If pCP10 = "102" Then '延展
             ET03 = "25"
        ElseIf pCP10 = "103" Then '補換發證書
             ET03 = "26"
        End If
        ET03type = LBL1(15).Caption '案件性質
    
    'Add By Sindy 2020/9/28
    ElseIf pCP10 = "725" Then '代辦退費
        ET03 = "01"
    '2020/9/28 END
    'Added by Lydia 2020/10/21
    ElseIf pCP10 = "308" Then '分割
        ET03 = "01"
    ElseIf pCP10 = "313" Then '減縮商品
        ET03 = "01"
    'Added by Lydia 2022/12/19
    ElseIf pCP10 = "314" Then '申請註冊證副本/註冊證副本申請書
        ET03 = "27"
    'end 2022/12/19
    ElseIf pCP10 = "304" Then '英文證明書
        ET03 = "01"
    'Added by Lydia 2024/08/06
    ElseIf pCP10 = "309" Then '中文證明書
        ET03 = "01"
    'end 2024/08/06
    ElseIf pCP10 = "717" Then '註冊費
        ET03 = "01"
    'end 2020/10/21
    'Added by Lydia 2023/01/05
    ElseIf pCP10 = "729" Then '復權
        ET03 = "28"
    'end 2023/01/05
   'Added by Lydia 2023/11/30 增加各商標種類的電子送件申請書
   ElseIf pCP10 = "306" Then  '自請撤回
      If m2_CP10 = "101" Then  '註冊申請案
         ET03 = "01"
      Else
         ET03 = "02"
      End If
   ElseIf pCP10 = "307" Then  '自請拋棄商標權
      ET03 = "01"
   'end 2023/11/30
   'Added by Lydia 2025/01/24 商爭案件:加速審查(311)、陳述意見(210)、異議(601)／評定(603)／廢止(605)
   ElseIf InStr("311,210,601,603,605", pCP10) > 0 Then
      ET03 = "01"
   'end 2025/01/24
    'Added by Lydia 2020/10/07 電子送件-補正申請書(A、B類收文)：其他沒有設定的性質，ex.303延期、201補正、202申請意見書、208補優先權證明、706其他
    Else
         'Memo by Lydia 2025/01/24 包含商爭案件:申請意見書(202)、延期(303,相關總收文非601,603,605)、異議答辯(602)／評定答辯答辯(604)／廢止答辯答辯(606
         ET03 = "10" '補正 (處理狀況的編號與FCT案一致)
    'end 2020/10/07
    End If
    
    'Added by Lydia 2025/01/24 商爭人員承辦FCT案，基本資料表的分機預設" 內商承辦分機"，Email預設為tm@taie.com.tw
    If tm(1) = "FCT" Then ET03_1 = "12"
    
    'Added by Lydia 2025/05/14 基本資料表不同:異議(601)／評定(603)／廢止(605)
    If (tm(1) = "T" Or tm(1) = "FCT") And InStr("601,603,605", pCP10) > 0 Then
       ET03_1 = "02"
    End If
    'end 2025/05/14

    
    If ET03 <> "" Then
         '1.申請書
         If StartLetter2(tm, m_CaseNo, ET01, ET03, pCP09, "2") = False Then Exit Sub
         'Memo by Lydia 2019/07/10 和內專工程師出發明申請一樣，同時產生申請書+基本資料表
         'NowPrint pCP09, ET01, ET03, False, strUserNum, , , True, strExc(9)
         'strFileName = strFolder & "\" & m_CaseNo & "." & ET03type & "申請書"
         'Call PUB_MakeDoc(strExc(9), strFileName)
         NowPrint pCP09, ET01, ET03, False, strUserNum, , strContent, True, strContent
    End If
    
    '1.基本資料
    'Modified by Lydia 2020/10/21 改用ET03_1
    'ET03 = "00"
    If ET03_1 <> "" Then
        If StartLetter2(tm, m_CaseNo, ET01, ET03_1, pCP09, "1") = False Then Exit Sub
            'Memo by Lydia 2019/07/10 和內專工程師出發明申請一樣，同時產生申請書+基本資料表
            'NowPrint pCP09, ET01, et03_1, False, strUserNum, , , True, strExc(9)
            'strFileName = strFolder & "\" & m_CaseNo & ".contact"
            'Call PUB_MakeDoc(strExc(9), strFileName)
            '因為MakeDoc有造字轉檔的引用，所以改用MakeDoc
            'NowPrint pCP09, ET01, et03_1, True, strUserNum, , strContent, , , , , False, , , , , , , , True
            NowPrint pCP09, ET01, ET03_1, False, strUserNum, , strContent, True, strContent
    End If 'Added by Lydia 2020/10/21
    
    'Modified by Lydia 2020/09/25 增加分節處理頁碼
    'Call PUB_MakeDoc(strContent, "")
    strContent = Replace(strContent, vbCrLf & Chr(12), vbCrLf & "|#(分節)#|")    '換頁符號Chr(12)替換為分節符號 "|#(分節)#|"
    Call PUB_MakeDoc(strContent, strFileName, , , , , True)  '分節處理頁碼
    
    MsgBox "資料已產生完畢!!!"
    
ExitSub1:
    Screen.MousePointer = vbDefault
End Sub

'Added by Lydia 2019/07/05 電子送件-申請書
Private Function StartLetter2(ByRef iTM() As String, ByVal iCaseNo As String, ByVal iET01 As String, _
   ByVal iET03 As String, ByVal iCp09 As String, ByVal iKind As String) As Boolean
Dim strTxt(1 To 30) As String
Dim ii As Integer, jj As Integer
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim intA As Integer
Dim iCP10 As String
Dim iCP17 As String
Dim iCP14 As String, iCP14ext As String
Dim iCP110 As String, iCP08 As String
Dim rsAD As New ADODB.Recordset
Dim TempList As ListBox
Dim strTmp As String 'Add By Sindy 2020/9/28
Dim iCP40 As String, iCP41 As String 'Added by Lydia 2025/01/24
   
   'Modify By Sindy 2020/9/28 + cp08
   'Modified by Lydia 2025/01/24 +cp40,cp41
   strSql = "select cp09,cp10,cp14,cp17,cp110,ed01,cp08,cp40,cp41 from caseprogress,ExtensionData where cp09='" & iCp09 & "' and cp14=ed02(+) "
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strSql)
   If intA = 1 Then
      iCP10 = "" & rsAD.Fields("cp10")
      iCP14 = "" & rsAD.Fields("cp14")
      iCP14ext = "" & rsAD.Fields("ed01")
      iCP17 = "" & rsAD.Fields("cp17")
      iCP110 = "" & rsAD.Fields("cp110")
      iCP08 = "" & rsAD.Fields("cp08")
      'Added by Lydia 2025/01/24
      iCP40 = "" & rsAD.Fields("cp40")
      iCP41 = "" & rsAD.Fields("cp41")
      'end 2025/01/24
   End If
   
   EndLetter iET01, iCp09, iET03, strUserNum
   
   ii = 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','本所案號','" & iCaseNo & "')"
   
   '申請人資料
   'Modified by Lydia 2020/10/07 +iCP10
   'Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, iTM(), False, , , iTM(1))
   Call PUB_GetApplFCT_EData(iET01, iET03, iCp09, iCP10, iTM(), False, , , iTM(1))
   
   '出名代理人: 改成共用模組取得資料
   If iCP110 = "" Then
       Call PUB_SetOurAgent(TempList, iTM, iCP110, iCP10)
   End If
   'Memo by Lydia 2020/10/21 分割案：母案和子字會同時收分割，但是申請書只會做母案，子案的卷宗區掛母案的申請書
   strExc(0) = PUB_GetAgentCP110(iCp09, iCP110, "T", "4")
   If strExc(0) <> "" Then
       tmpArr1 = Empty
       tmpArr1 = Split(strExc(0), "|")
       For jj = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(jj)) <> "" Then
               tmpArr2 = Empty
               tmpArr2 = Split(tmpArr1(jj), ",")
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-證書字號','" & tmpArr2(0) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-ID','" & tmpArr2(1) & "')"
                ii = ii + 1
                strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','代理人" & jj + 1 & "-中文姓名','" & PUB_ConvertNameFormat("" & tmpArr2(2)) & "')"
           End If
       Next jj
   End If
   
   If iKind = "1" Then '基本資料表
        ii = ii + 1
        '內商承辦分機
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','內商承辦分機','" & iCP14ext & "')"
   End If
   
   If iKind = "2" Then '電子送件申請書
        ii = ii + 1
        '繳費金額
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','繳費金額','" & Val(iCP17) & "')"
       '收據抬頭(內商才用)
        strExc(1) = ""
        strExc(1) = GetPrjPeople1(ChangeCustomerL(iTM(23)))
        For intI = 78 To 81 '申請人2~4
            If iTM(intI) <> "" Then
               'Modified by Lydia 2024/03/05 改模組
               'strExc(1) = strExc(1) & "、" & GetPrjPeople1(ChangeCustomerL(iTM(intI)))
               strExc(1) = strExc(1) & "、" & PUB_GetApplT_CNAME(ChangeCustomerL(iTM(intI)))
            End If
        Next intI
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
           " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','收據抬頭', " & CNULL(ChgSQL(strExc(1))) & ")"
         
        'Added by Lydia 2022/12/19 註冊證形式
        If strSrvDate(1) >= "20230101" Then
           If iCP10 = "314" Then '申請註冊證副本314: 註冊證形式
              '申請內容1
              ii = ii + 1
              strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','申請內容1','申請本件商標之紙本商標註冊證副本。')"
           End If
           If iTM(136) = "1" Then
              strExc(1) = "電子"
           ElseIf iTM(136) = "2" Then
              strExc(1) = "紙本"
           Else
              strExc(1) = "電子/紙本"
           End If
           ii = ii + 1
           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                 " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','註冊證形式','" & strExc(1) & "')"
        End If
        'end 2022/12/19
        'Added by Lydia 2024/09/12 補證103
        If iCP10 = "103" Then
            '申請內容1
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','申請內容1','1.申請補發本件註冊證。" & vbCrLf & "2.具結：本件註冊商標/標章註冊證確實遺失。')"
        End If
        'end 2024/09/12
        
        If iCP10 = "101" Then '申請
            '主張優先權
            'Modified by Lydia 2019/07/31 +NA72
            'Modified by Lydia 2020/02/19 商標申請書之主張優先權, 國家為239者改寫死帶 EU歐盟----阿蓮 (與專利案不同)
            'strExc(0) = "select sqldatet(PD05) as PD05 ,PD06,NA03,NA72,NVL(A1.TM01||A1.TM02||A1.TM03||A1.TM04,A2.TM01||A2.TM02||A2.TM03||A2.TM04) AS caseno,PD10 " & _
                     "from PRIDATE,NATION,TRADEMARK A1,TRADEMARK A2 " & _
                     "WHERE PD01='" & iTM(1) & "' AND PD02='" & iTM(2) & "' AND PD03='" & iTM(3) & "' AND PD04 ='" & iTM(4) & "' " & _
                     "AND PD06=A1.TM12(+) AND PD05=A1.TM11(+) AND PD07=A1.TM10(+) " & _
                     "AND PD06=A2.TM15(+) AND PD05=A2.TM11(+) AND PD07=A2.TM10(+) " & _
                     "AND PD07=NA01(+) " & _
                     "ORDER BY PD01,PD02,PD03,PD04"
            strExc(0) = "select sqldatet(PD05) as PD05 ,PD06,NA03,DECODE(PD07,'239','EU歐盟',NA72) NA72,NVL(A1.TM01||A1.TM02||A1.TM03||A1.TM04,A2.TM01||A2.TM02||A2.TM03||A2.TM04) AS caseno,PD10 " & _
                     "from PRIDATE,NATION,TRADEMARK A1,TRADEMARK A2 " & _
                     "WHERE PD01='" & iTM(1) & "' AND PD02='" & iTM(2) & "' AND PD03='" & iTM(3) & "' AND PD04 ='" & iTM(4) & "' " & _
                     "AND PD06=A1.TM12(+) AND PD05=A1.TM11(+) AND PD07=A1.TM10(+) " & _
                     "AND PD06=A2.TM15(+) AND PD05=A2.TM11(+) AND PD07=A2.TM10(+) " & _
                     "AND PD07=NA01(+) " & _
                     "ORDER BY PD01,PD02,PD03,PD04"
            intI = 1
            Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                With rsAD
                    .MoveFirst
                    jj = 1
                    strExc(1) = ""
                    Do While Not .EOF
                         'Modified by Lydia 2019/07/31 na03=>na72 IPO國籍代碼+中文國名
                         strExc(1) = strExc(1) & _
                                          "【主張優先權" & jj & "】  " & vbCrLf & _
                                          "　　【優先權日】　　　　　　　" & .Fields("pd05") & vbCrLf & _
                                          "　　【受理國家或地區】　　　　" & .Fields("na72") & vbCrLf & _
                                          "　　【申請案號】　　　　　　　" & .Fields("pd06") & vbCrLf
                         jj = jj + 1
                         .MoveNext
                    Loop
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                          " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','主張優先權','" & ChgSQL(strExc(1)) & "')"
                End With
            End If
        End If
        '商標顏色
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商標顏色','" & "墨色" & "')"
          
        '商品服務類別及名稱: 101申請,102延展
        'Modified by Lydia 2020/10/21 + 308.分割,313.減縮商品
        'Modified by Lydia 2022/09/15  配合FCT案：101申請,102延展改成即時抓基本檔資料; ex.FCT-49623有38類描述長度有5857
        'If iCP10 = "101" Or iCP10 = "102" Or iCP10 = "308" Or iCP10 = "313" Then
        If iCP10 = "308" Or iCP10 = "313" Then
            strExc(1) = "": strExc(2) = "": strExc(3) = ""
            strExc(0) = BeforePrintGetDBData("TMGoods:" & iTM(1) & "-" & iTM(2) & "-" & iTM(3) & "-" & iTM(4) & "-||區隔", True)
            strTmp = "" 'Added by Lydia 2020/10/26
            If Trim(strExc(0)) <> "" Then
                 tmpArr1 = Empty
                 tmpArr1 = Split(strExc(0), "||")
                 jj = 1
                 For intA = 0 To UBound(tmpArr1)
                     strExc(1) = Trim(tmpArr1(intA))
                     If strExc(1) <> "" Then
                          'Added by Lydia 2020/10/21 313.減縮商品
                          If iCP10 = "313" Then
                                strExc(2) = strExc(2) & _
                                                 "【擬減縮商品或服務名稱" & jj & "】  " & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Format(Mid(strExc(1), 1, InStr(strExc(1), "：") - 1), "000") & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & Mid(strExc(1), InStr(strExc(1), "：") + 1) & vbCrLf
                                strExc(3) = strExc(3) & _
                                                 "【減縮後指定商品或服務名稱" & jj & "】  " & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Format(Mid(strExc(1), 1, InStr(strExc(1), "：") - 1), "000") & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & vbCrLf
                          'Added by Lydia 2020/10/26  分割: 分割序號01=原類別名稱,分割序號=02 修改後的類別名稱(帶空白)
                          ElseIf iCP10 = "308" Then
                                strExc(2) = strExc(2) & _
                                                 "【分割後商品服務類別名稱或證明標的內容1】  " & vbCrLf & _
                                                 "　　【分割序號】　　　　　　　01" & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Mid(strExc(1), 1, InStr(strExc(1), "：") - 1) & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & Mid(strExc(1), InStr(strExc(1), "：") + 1) & vbCrLf
                          'end 2020/10/26
                          Else 'Added by Lydia 2020/10/21
                                'Modified by Lydia 2019/07/31 阿蓮:FCT不要組群代碼 ; 嘉雯: 內商的商品服務類別內容是用智慧局插件產生,代空白也可以,而外商多半無法用智慧局所以才用人工撰寫
                                'strExc(2) = strExc(2) & _
                                                 "【指定使用商品服務類別及名稱" & jj & "】  " & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Mid(strExc(1), 1, InStr(strExc(1), "：") - 1) & vbCrLf & _
                                                 IIf(iCP10 = "101", "　　【組群代碼】　　　　　　　" & vbCrLf, "") & _
                                                 "　　【商品服務名稱】　　　　　" & Mid(strExc(1), InStr(strExc(1), "：") + 1) & vbCrLf
                                strExc(2) = strExc(2) & _
                                                 "【指定使用商品服務類別及名稱" & jj & "】  " & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Mid(strExc(1), 1, InStr(strExc(1), "：") - 1) & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & Mid(strExc(1), InStr(strExc(1), "：") + 1) & vbCrLf
                          End If 'Added by Lydia 2020/10/21
                          jj = jj + 1
                          strTmp = strTmp & "," & Mid(strExc(1), 1, InStr(strExc(1), "：") - 1)   'Added by Lydia 2020/10/26
                     End If
                 Next intA
            'Added by Lydia 2019/07/31 阿蓮: 因為在產生申請書時才撰寫商品服務,所以依收文的類別產生
            'Memo by Lydia 2019/07/31 嘉雯: 內商的商品服務類別內容是用智慧局插件產生,代空白也可以,而外商多半無法用智慧局所以才用人工撰寫
            ElseIf iTM(9) <> "" Then
                 tmpArr1 = Empty
                 tmpArr1 = Split(iTM(9), ",")
                 jj = 1
                 For intA = 0 To UBound(tmpArr1)
                     strExc(1) = Trim(tmpArr1(intA))
                     If strExc(1) <> "" Then
                          'Added by Lydia 2020/10/21 313.減縮商品
                          If iCP10 = "313" Then
                                strExc(2) = strExc(2) & _
                                                 "【擬減縮商品或服務名稱" & jj & "】  " & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Format(strExc(1), "000") & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & vbCrLf
                                strExc(3) = strExc(3) & _
                                                 "【減縮後指定商品或服務名稱" & jj & "】  " & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & Format(strExc(1), "000") & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & vbCrLf
                          'Added by Lydia 2020/10/26  分割: 分割序號01=原類別名稱,分割序號=02 修改後的類別名稱(帶空白)
                          ElseIf iCP10 = "308" Then
                                strExc(2) = strExc(2) & _
                                                 "【分割後商品服務類別名稱或證明標的內容1】  " & vbCrLf & _
                                                 "　　【分割序號】　　　　　　　01" & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & strExc(1) & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & vbCrLf
                          'end 2020/10/26
                          Else 'Added by Lydia 2020/10/21
                                strExc(2) = strExc(2) & _
                                                 "【指定使用商品服務類別及名稱" & jj & "】  " & vbCrLf & _
                                                 "　　【類別】　　　　　　　　　" & strExc(1) & vbCrLf & _
                                                 "　　【商品服務名稱】　　　　　" & vbCrLf
                          End If 'Added by Lydia 2020/10/21
                          jj = jj + 1
                          strTmp = strTmp & "," & strExc(1) 'Added by Lydia 2020/10/26
                     End If
                 Next intA
            'end 2019/07/31
            Else
                'Added by Lydia 2020/10/21 313.減縮商品
                If iCP10 = "313" Then
                     strExc(2) = strExc(2) & _
                                      "【擬減縮商品或服務名稱1】  " & vbCrLf & _
                                      "　　【類別】　　　　　　　　　" & vbCrLf & _
                                      "　　【商品服務名稱】　　　　　" & vbCrLf
                     strExc(3) = strExc(3) & _
                                      "【減縮後指定商品或服務名稱1】  " & vbCrLf & _
                                      "　　【類別】　　　　　　　　　" & vbCrLf & _
                                      "　　【商品服務名稱】　　　　　" & vbCrLf
                'Added by Lydia 2020/10/26  分割: 分割序號01=原類別名稱,分割序號=02 修改後的類別名稱(帶空白)
                ElseIf iCP10 = "308" Then
                      strExc(2) = strExc(2) & _
                                       "【分割後商品服務類別名稱或證明標的內容1】  " & vbCrLf & _
                                       "　　【分割序號】　　　　　　　01" & vbCrLf & _
                                       "　　【類別】　　　　　　　　　" & vbCrLf & _
                                       "　　【商品服務名稱】　　　　　" & vbCrLf
                'end 2020/10/26
                Else 'Added by Lydia 2020/10/21
                     'Modified by Lydia 2019/07/31
                     'strExc(2) = "【指定使用商品服務類別及名稱1】  " & vbCrLf & _
                                      "　　【類別】　　　　　　　　　" & vbCrLf & _
                                     IIf(iCP10 = "101", "　　【組群代碼】　　　　　　　" & vbCrLf, "") & _
                                      "　　【商品服務名稱】　　　　　" & vbCrLf
                     strExc(2) = "【指定使用商品服務類別及名稱1】  " & vbCrLf & _
                                      "　　【類別】　　　　　　　　　" & vbCrLf & _
                                      "　　【商品服務名稱】　　　　　" & vbCrLf
                End If 'Added by Lydia 2020/10/21
                strTmp = strTmp & ",01"  'Added by Lydia 2020/10/26
            End If
            ii = ii + 1
            '申請
            If iCP10 <> "308" Then 'Added by Lydia 2020/10/26
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                      " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','指定使用商品服務類別及名稱','" & ChgSQL(strExc(2)) & "')"
            End If 'Added by Lydia 2020/10/26
            If iCP10 = "102" Then '延展
                 strTxt(ii) = Replace(strTxt(ii), "指定使用商品服務類別及名稱", "部分延展")
            'Added by Lydia 2020/10/21
            ElseIf iCP10 = "308" Then
                 'Modified by Lydia 2020/10/26 增加分割: 分割序號01=原類別名稱,分割序號=02 修改後的類別名稱(帶空白)
                 'strTxt(ii) = Replace(strTxt(ii), "指定使用商品服務類別及名稱", "分割後商品服務類別名稱")
                 tmpArr1 = Empty
                 tmpArr1 = Split(Mid(strTmp, 2), ",")
                 For intA = 0 To UBound(tmpArr1)
                    If Trim(tmpArr1(intA)) <> "" Then
                        strExc(2) = strExc(2) & _
                                         "【分割後商品服務類別名稱或證明標的內容2】  " & vbCrLf & _
                                         "　　【分割序號】　　　　　　　02" & vbCrLf & _
                                         "　　【類別】　　　　　　　　　" & Trim(tmpArr1(intA)) & vbCrLf & _
                                         "　　【商品服務名稱】　　　　　" & vbCrLf
                    End If
                 Next intA
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','分割後商品服務類別名稱','" & ChgSQL(strExc(2)) & "')"
                 ii = ii + 1
                 '因為無法抓子案的明確件數，所以預設1件
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','分割件數','1')"
                 'end 2020/10/26
                 ii = ii + 1
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','他案辦理日期','" & ChangeTStringToTDateString(strSrvDate(2)) & "')"
            ElseIf iCP10 = "313" Then
                 strTxt(ii) = Replace(strTxt(ii), "指定使用商品服務類別及名稱", "擬減縮商品服務類別名稱")
                 ii = ii + 1
                 strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','減縮後指定商品服務類別名稱','" & ChgSQL(strExc(3)) & "')"
            End If
            'end 2020/10/21
        'Added by Lydia 2020/10/21
        'Modified by Lydia 2024/08/06 +309中文證明書 => Or iCP10 = "309"
        ElseIf iCP10 = "304" Or iCP10 = "309" Then '英文證明書
             strExc(0) = BeforePrintGetDBData("TMGoods:" & iTM(1) & "-" & iTM(2) & "-" & iTM(3) & "-" & iTM(4) & "-中文", True)
             If strExc(0) <> "" Then
                 '單一類別的案件,開頭不顯示類別代號 (嘉雯&阿蓮的溝通結果)
                 If InStr(iTM(9), ",") = 0 Then
                      strExc(0) = Mid(strExc(0), InStr(strExc(0), "：") + 1)
                 End If
                 If Trim(strExc(0)) <> "" Then
                    ii = ii + 1
                    strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商品服務類別名稱中文','" & ChgSQL(strExc(0)) & "')"
                 End If
             End If
             If iCP10 = "304" Then  'Added by Lydia 2024/08/06
               strExc(1) = BeforePrintGetDBData("TMGoods:" & iTM(1) & "-" & iTM(2) & "-" & iTM(3) & "-" & iTM(4) & "-英文", True)
               If strExc(1) <> "" Then
                  '單一類別的案件,開頭不顯示類別代號 (嘉雯&阿蓮的溝通結果)
                  If InStr(iTM(9), ",") = 0 Then
                       strExc(1) = Mid(strExc(1), InStr(strExc(1), "：") + 1)
                  End If
                  If Trim(strExc(1)) <> "" Then
                      ii = ii + 1
                      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','商品服務類別名稱英文','" & ChgSQL(strExc(1)) & "')"
                  End If
               End If
             End If 'Added by Lydia 2024/08/06
        'end 2020/10/21
        'Add By Sindy 2020/9/28
        ElseIf iCP10 = "725" Then '代辦退費
            strTmp = "（　　）智商/慧商　　　　字第　　　　　　　　　　號函"
            If iCP08 <> "" Then strTmp = iCP08
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','機關文號','" & strTmp & "')"
            '抓電子收據資枓
            If m_CP43 <> "" Then
               strSql = "select cp09,cp64 from caseprogress where cp09='" & m_CP43 & "' and instr(cp64,'收據號碼:')>0"
               intA = 1
               Set rsAD = ClsLawReadRstMsg(intA, strSql)
               If intA = 1 Then
                  strTmp = ""
                  If InStr(rsAD.Fields("cp64"), "收據號碼:") > 0 Then
                     strTmp = Mid(rsAD.Fields("cp64"), InStr(rsAD.Fields("cp64"), "收據號碼:") + 5, 11)
                  End If
                  If strTmp <> "" Then
                     ii = ii + 1
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','電子收據號碼','" & strTmp & "')"
                     '電子收據紀錄檔
                     strSql = "select er01,er03 from ereceipt where er01='" & strTmp & "'"
                     intA = 1
                     Set rsAD = ClsLawReadRstMsg(intA, strSql)
                     If intA = 1 Then
                        strTmp = Val("" & rsAD.Fields("er03"))
                        ii = ii + 1
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','電子收據規費','" & strTmp & "')"
                     End If
                  End If
               End If
            End If
            '2020/9/28 END
        End If
        
        '附送書件
        'If iCP10 = "101" Or iCP10 = "102" Or iCP10 = "103" Then '申請,延展,補換發證書103
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-基本資料表', '" & iCaseNo & ".contact.pdf')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-委任書', '" & iCaseNo & ".poa.pdf')"
        'End If
        If iCP10 = "101" Then  '申請
            ii = ii + 1
            'Modified by Lydia 2020/07/16 更名:「.priority.pdf」改為「.PRI.pdf」
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-優先權證明文件', '" & iCaseNo & ".PRI.pdf')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-展覽會優先權證明文件', '" & iCaseNo & ".PRI.pdf')"
        End If
        If iCP10 = "102" Then  '延展
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-變更證明文件', '" & iCaseNo & ".change.pdf')"
        End If
        If iCP10 = "103" Then  '補換發證書103
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','附件-具結書', '" & iCaseNo & ".declaration.pdf')"
        End If
        'Added by Lydia 2023/11/30
        If m_CP10 = "306" Then  '自請撤回->非申請案
           ii = ii + 1
           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','指定名稱','" & m2_CP10ex & "')"
           ii = ii + 1
           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','指定名稱2','" & Replace(m2_CP10ex, "案", "") & "')"
        End If
        'end 2023/11/30
   End If
   
   'Added by Lydia 2025/01/24  商爭案件:異議(601)／評定(603)／廢止(605)
   'Move by Lydia 2025/05/14 從 'end 2023/11/30 因為申請書和基本資料表都需要，所以從下方移過來
   'Modified by Lydia 2025/05/14 + (tm(1) = "T" Or tm(1) = "FCT") And
   If (iTM(1) = "T" Or iTM(1) = "FCT") And InStr("601,603,605", iCP10) > 0 And Trim(iCP40 & iCP41) <> "" Then
       tmpArr1 = Empty
       tmpArr1 = Split(iCP40, "，")
       strExc(0) = UBound(tmpArr1) 'Added by Lydia 2025/05/14
       For jj = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(jj)) <> "" Then
             'Added by Lydia 2025/05/14 先以中文名稱數量增加序號
             ii = ii + 1
             strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','對造名稱" & jj + 1 & "-序號','♀')"
             'end 2025/05/14
             ii = ii + 1
             strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','對造名稱" & jj + 1 & "-中文名稱','" & tmpArr1(jj) & "')"
           End If
       Next jj
       tmpArr1 = Empty
       tmpArr1 = Split(iCP41, "，")
       For jj = 0 To UBound(tmpArr1)
           If Trim(tmpArr1(jj)) <> "" Then
             'Added by Lydia 2025/05/14 增加序號
             If jj > Val(strExc(0)) Then
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','對造名稱" & jj + 1 & "-序號','♀')"
             End If
             'end 2025/05/14
             ii = ii + 1
             strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                " VALUES ('" & iET01 & "','" & iCp09 & "','" & iET03 & "','" & strUserNum & "','對造名稱" & jj + 1 & "-英文名稱','" & tmpArr1(jj) & "')"
           End If
       Next jj
   End If
   'end 2025/01/24
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

'Add By Sindy 2020/11/17 寄發指示信
Private Sub cmdDataMail_Click()
On Error GoTo ErrHnd
   
   If PUB_T_AppFormSendMail(LBL1(3).Caption, LBL1(3).Caption, _
         m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, Me) = False Then
      Exit Sub
   End If
   
   Exit Sub
ErrHnd:
   If Err.Number > 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

'Added by Lydia 2021/06/25 檢查查名單對照檔和卷宗區TS.menu，有缺就自動補上資料
Private Sub ChkTMQmapData(ByVal pCP09 As String)
Dim strQ1 As String, intQ As Integer
Dim rsQD As New ADODB.Recordset
Dim strTmpA As String
    
    'Added by Lydia 2023/11/07 查名流水號修改為同一號;ex.T-246287(AB2042812)有112002286,112002287
    strQ1 = "select min(tqc01) mno from tmqcasemap where tqc02='" & pCP09 & "' "
    intQ = 1
    Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
    If intQ = 1 Then
       strSql = "Update TmqCaseMap set tqc01='" & rsQD.Fields("mno") & "' where tqc02='" & pCP09 & "' and tqc01<> '" & rsQD.Fields("mno") & "' "
       cnnConnection.Execute strSql
    End If
    '卷宗區有TS.menu，但是查名單對照檔沒有; ex T-233397
    strQ1 = "SELECT CPP01,SUBSTR(CPP02,INSTR(CPP02,'H'),9) as CPP02t,TQC02,TQC03 FROM CASEPAPERPDF A,TMQCASEMAP B " & _
                "WHERE CPP01='" & pCP09 & "' AND UPPER(CPP02) LIKE '%." & UCase(TMQ_查名作業 & ".menu") & "' AND CPP01=TQC02(+) AND SUBSTR(CPP02,INSTR(CPP02,'H'),9)=TQC03(+) AND TQC02 IS NULL "
    intQ = 1
    Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
    If intQ = 1 Then
        strQ1 = PUB_GetTMQCaseMapNo(pCP09) '取得流水號
        rsQD.MoveFirst
        Do While Not rsQD.EOF
            strTmpA = Left(Format(ServerTime, "000000"), 4)
            strSql = "insert into tmqcasemap (tqc01,tqc02,tqc03,tqc04,tqc05,tqc06) values ('" & strQ1 & "', '" & pCP09 & "', '" & rsQD.Fields("cpp02t") & "', '" & strUserNum & "', " & strSrvDate(1) & ", '" & strTmpA & "' ) "
            cnnConnection.Execute strSql
            rsQD.MoveNext
        Loop
    End If
    
    '有查名單對照檔，但是卷宗區沒有TS.menu
    strTmpA = PUB_CaseNo2FileName(m_CP01, m_CP02, m_CP03, m_CP04) & "." & m_CP10 & "."
    strQ1 = "SELECT TQC02,TQC03,CPP01,CPP02," & CNULL(strTmpA) & "||TQC03||" & CNULL("." & TMQ_查名作業 & ".menu") & " AS CPP02T " & _
                "FROM TMQCASEMAP A, CASEPAPERPDF B WHERE TQC02='" & pCP09 & "' AND TQC02=CPP01(+) " & _
                "AND " & CNULL(strTmpA) & "||TQC03||" & CNULL("." & TMQ_查名作業 & ".menu") & "=CPP02(+) AND CPP02 IS NULL "
    intQ = 1
    Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
    If intQ = 1 Then
        rsQD.MoveFirst
        Do While Not rsQD.EOF
            strTmpA = "" & rsQD.Fields("cpp02t")
            strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
                        " values('" & pCP09 & "','" & strTmpA & "',0,'" & strUserNum & "'," & _
                               strSrvDate(1) & ",to_char(sysdate,'hh24miss')," & _
                               strSrvDate(1) & ",to_char(sysdate,'hh24miss'),'Y')"
            cnnConnection.Execute strSql
            rsQD.MoveNext
        Loop
    End If
    
    'Added by Lydia 2024/12/05 T-251717在12/4已設定「無須查名/自行查名」，但齊備日被拿掉;推測是承辦人上「無須查名/自行查名」，同時程序人員已在分案作業的畫面仍是未齊備；先不動分案作業，直接在工作進度維護補齊備日
    strQ1 = "select cp09,cp143,cp64,count(tqc03) cnt From caseprogress, tmqcasemap " & _
            "where cp09='" & pCP09 & "' and instr(cp64,'查名備註:無須查名/自行查名') > 0 and instr(cp64,'查名備註:取消無須查名/自行查名') = 0 " & _
            "and cp09=tqc02(+) group by cp09,cp143,cp64 "
    intQ = 1
    Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
    If intQ = 1 Then
      If Val("" & rsQD.Fields("cp143")) = 0 And Val("" & rsQD.Fields("cnt")) = 0 Then
         intQ = InStr("" & rsQD.Fields("cp64"), "查名備註:無須查名/自行查名(")
         If intQ < 11 Then
            strTmpA = strSrvDate(1)
         Else
            strTmpA = DBDATE(Trim(Mid("" & rsQD.Fields("cp64"), intQ - 10, 10)))
         End If
         strSql = "Update caseprogress set cp143=" & strTmpA & ",cp64=replace(cp64,'查名備註:無須查名/自行查名(','查名備註:無須查名/自行查名(" & ChangeWStringToTDateString(strTmpA) & "修正，')  where cp09 = '" & pCP09 & "' "
         cnnConnection.Execute strSql
         txtCP64 = Replace(txtCP64, "查名備註:無須查名/自行查名(", "查名備註:無須查名/自行查名(" & ChangeWStringToTDateString(strTmpA) & "修正，")
         txtCP64.Tag = txtCP64
         textCP143 = TransDate(strTmpA, 1)
         textCP143.Tag = textCP143
      End If
    End If
    'end 2024/12/04
    Set rsQD = Nothing
End Sub

'Added by Lydia 2021/12/23
Private Sub txtEP12_GotFocus()
    TextInverse txtEP12
End Sub

Private Sub txtCP64_GotFocus()
    TextInverse txtCP64
End Sub
