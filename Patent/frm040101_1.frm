VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040101_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利分案"
   ClientHeight    =   6360
   ClientLeft      =   -396
   ClientTop       =   1596
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9156
   Begin VB.TextBox txtF0309 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7035
      Locked          =   -1  'True
      TabIndex        =   141
      Top             =   516
      Width           =   1665
   End
   Begin VB.CommandButton cmdCPP 
      Caption         =   "卷宗區"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   0
      TabIndex        =   54
      Top             =   60
      Width           =   705
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "檢視接洽單"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   721
      TabIndex        =   55
      Top             =   60
      Width           =   1065
   End
   Begin VB.TextBox txtEngGroup 
      Height          =   270
      Left            =   6795
      MaxLength       =   1
      TabIndex        =   125
      Top             =   945
      Width           =   285
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   1
      Left            =   945
      MaxLength       =   1
      TabIndex        =   0
      Top             =   1260
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   5
      Left            =   945
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1800
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   3
      Left            =   945
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1530
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   4
      Left            =   4410
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1530
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      Height          =   270
      Index           =   2
      Left            =   4410
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1260
      Width           =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "一案兩請資料"
      Height          =   400
      Index           =   6
      Left            =   2748
      TabIndex        =   57
      Top             =   45
      Width           =   1245
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4320
      Left            =   135
      TabIndex        =   64
      Top             =   2070
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   7620
      _Version        =   393216
      TabsPerRow      =   8
      TabHeight       =   529
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm040101_1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(43)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(41)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(40)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(30)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(29)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(37)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label3(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label3(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(12)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(13)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(14)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(15)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label3(11)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label3(12)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(18)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblDivCase"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(20)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(21)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(42)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(22)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(23)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblFeeYear"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(168)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(121)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label20"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(8)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(0)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(1)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(18)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(14)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(13)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(12)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(9)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(3)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(5)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text1(15)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text1(4)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(6)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text1(11)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text1(22)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text1(10)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text1(16)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text1(23)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Text1(2)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Text1(17)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Label1(16)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Text1(7)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "lblCP71"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Label1(39)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Text1(25)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "txtCP118"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "textPA4"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "textPA3"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txtDivCaseNo(1)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtDivCaseNo(2)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtDivCaseNo(3)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txtDivCaseNo(4)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "textPA2"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "textPA1"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txtFeeYear(1)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txtFeeYear(2)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txtFavDt"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "Combo3"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "CmdFav"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "FraLOS"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "Check11"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "txtCP71"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "MSHFlexGrid1"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "Frame1"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "Frame2"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).ControlCount=   75
      TabCaption(1)   =   "其他"
      TabPicture(1)   =   "frm040101_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPA178"
      Tab(1).Control(1)=   "txtCP97"
      Tab(1).Control(2)=   "txtPA161"
      Tab(1).Control(3)=   "txtCP147"
      Tab(1).Control(4)=   "txtCP98"
      Tab(1).Control(5)=   "lblPA178"
      Tab(1).Control(6)=   "Text1(21)"
      Tab(1).Control(7)=   "Text1(19)"
      Tab(1).Control(8)=   "Text1(20)"
      Tab(1).Control(9)=   "txtCP99"
      Tab(1).Control(10)=   "LblCP97"
      Tab(1).Control(11)=   "lblPA161"
      Tab(1).Control(12)=   "Label1(172)"
      Tab(1).Control(13)=   "Label5"
      Tab(1).Control(14)=   "Label1(19)"
      Tab(1).Control(15)=   "Label1(17)"
      Tab(1).Control(16)=   "Label1(11)"
      Tab(1).Control(17)=   "Label2(3)"
      Tab(1).Control(18)=   "Label2(2)"
      Tab(1).Control(19)=   "Label1(9)"
      Tab(1).Control(20)=   "Label1(10)"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "簽核"
      TabPicture(2)   =   "frm040101_1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdAddInfo"
      Tab(2).Control(1)=   "txtF0301"
      Tab(2).Control(2)=   "GRD1"
      Tab(2).Control(3)=   "Label67"
      Tab(2).Control(4)=   "txtNote"
      Tab(2).Control(5)=   "Label66"
      Tab(2).Control(6)=   "txtF0407"
      Tab(2).Control(7)=   "Label68"
      Tab(2).ControlCount=   8
      Begin VB.Frame Frame2 
         Height          =   396
         Left            =   4464
         TabIndex        =   151
         Top             =   2856
         Width           =   2220
         Begin VB.OptionButton Option1 
            Appearance      =   0  '平面
            BackColor       =   &H80000005&
            Caption         =   "當天"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   24
            TabIndex        =   154
            Top             =   144
            Width           =   660
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  '平面
            BackColor       =   &H80000005&
            Caption         =   "之前"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   744
            TabIndex        =   153
            Top             =   144
            Width           =   660
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  '平面
            BackColor       =   &H80000005&
            Caption         =   "之後"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1476
            TabIndex        =   152
            Top             =   144
            Width           =   660
         End
      End
      Begin VB.Frame Frame1 
         Height          =   396
         Left            =   900
         TabIndex        =   123
         Top             =   2850
         Width           =   3636
         Begin VB.OptionButton OptSendType 
            Caption         =   "指定日期"
            Height          =   180
            Index           =   3
            Left            =   1740
            TabIndex        =   36
            Top             =   135
            Width           =   1008
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "收款後"
            Height          =   180
            Index           =   2
            Left            =   888
            TabIndex        =   35
            Top             =   135
            Width           =   850
         End
         Begin VB.TextBox txtCP142 
            Height          =   270
            Left            =   2745
            MaxLength       =   7
            TabIndex        =   37
            Top             =   90
            Width           =   825
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "不限制"
            Height          =   180
            Index           =   1
            Left            =   24
            TabIndex        =   34
            Top             =   135
            Width           =   850
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1056
         Left            =   96
         TabIndex        =   65
         Top             =   3240
         Width           =   8700
         _ExtentX        =   15346
         _ExtentY        =   1863
         _Version        =   393216
         Cols            =   11
         FixedCols       =   0
         BackColorBkg    =   16772048
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         MergeCells      =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
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
         _Band(0).Cols   =   11
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.TextBox txtCP71 
         Height          =   300
         Left            =   3810
         MaxLength       =   7
         TabIndex        =   147
         Top             =   2310
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txtPA178 
         Height          =   300
         Left            =   -74100
         MaxLength       =   1
         TabIndex        =   51
         Top             =   3570
         Width           =   285
      End
      Begin VB.CheckBox Check11 
         Caption         =   "急件"
         ForeColor       =   &H00000000&
         Height          =   200
         Left            =   3780
         TabIndex        =   143
         Top             =   360
         Width           =   765
      End
      Begin VB.CommandButton CmdAddInfo 
         Caption         =   "補件完成"
         CausesValidation=   0   'False
         Height          =   400
         Left            =   -67440
         TabIndex        =   53
         Top             =   450
         Width           =   1200
      End
      Begin VB.TextBox txtF0301 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -73860
         Locked          =   -1  'True
         TabIndex        =   135
         Top             =   420
         Width           =   1215
      End
      Begin VB.Frame FraLOS 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   285
         Left            =   4770
         TabIndex        =   133
         Top             =   0
         Width           =   3375
         Begin VB.TextBox txtLOSagree 
            Height          =   270
            Left            =   1890
            MaxLength       =   1
            TabIndex        =   38
            Top             =   -8
            Width           =   405
         End
         Begin VB.Label LBL6 
            Caption         =   "是否需要法律所配合：　　　(Y: 配合) "
            Height          =   195
            Left            =   0
            TabIndex        =   134
            Top             =   30
            Width           =   3135
         End
      End
      Begin VB.CommandButton CmdFav 
         Caption         =   "優惠期事實發生日期"
         Height          =   270
         Left            =   1940
         TabIndex        =   132
         Top             =   1420
         Width           =   1815
      End
      Begin VB.TextBox txtCP97 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   -73800
         MaxLength       =   4
         TabIndex        =   50
         Top             =   3240
         Width           =   585
      End
      Begin VB.TextBox txtPA161 
         Height          =   300
         Left            =   -69570
         MaxLength       =   1
         TabIndex        =   49
         Top             =   2940
         Width           =   255
      End
      Begin VB.TextBox txtCP147 
         Height          =   300
         Left            =   -72975
         MaxLength       =   1
         TabIndex        =   48
         Top             =   2940
         Width           =   255
      End
      Begin VB.ComboBox Combo3 
         Height          =   276
         ItemData        =   "frm040101_1.frx":0054
         Left            =   3330
         List            =   "frm040101_1.frx":0061
         TabIndex        =   32
         Top             =   2595
         Width           =   1305
      End
      Begin VB.TextBox txtFavDt 
         Height          =   300
         Left            =   3780
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   18
         Top             =   1425
         Width           =   855
      End
      Begin VB.TextBox txtFeeYear 
         Height          =   300
         Index           =   2
         Left            =   8325
         TabIndex        =   23
         Top             =   1425
         Width           =   285
      End
      Begin VB.TextBox txtFeeYear 
         Height          =   300
         Index           =   1
         Left            =   7830
         TabIndex        =   22
         Top             =   1425
         Width           =   285
      End
      Begin VB.TextBox txtCP98 
         Height          =   300
         Left            =   -73560
         MaxLength       =   12
         TabIndex        =   46
         Top             =   2130
         Width           =   675
      End
      Begin VB.TextBox textPA1 
         Height          =   300
         Left            =   5700
         MaxLength       =   3
         TabIndex        =   9
         Top             =   615
         Width           =   495
      End
      Begin VB.TextBox textPA2 
         Height          =   300
         Left            =   6180
         MaxLength       =   6
         TabIndex        =   10
         Top             =   615
         Width           =   855
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   300
         Index           =   4
         Left            =   7725
         MaxLength       =   2
         TabIndex        =   42
         Top             =   2625
         Width           =   390
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   300
         Index           =   3
         Left            =   7365
         MaxLength       =   1
         TabIndex        =   41
         Top             =   2625
         Width           =   390
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   300
         Index           =   2
         Left            =   6660
         MaxLength       =   6
         TabIndex        =   40
         Top             =   2625
         Width           =   705
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   300
         Index           =   1
         Left            =   6270
         MaxLength       =   3
         TabIndex        =   39
         Top             =   2625
         Width           =   390
      End
      Begin VB.TextBox textPA3 
         Height          =   300
         Left            =   7020
         MaxLength       =   1
         TabIndex        =   11
         Top             =   615
         Width           =   255
      End
      Begin VB.TextBox textPA4 
         Height          =   300
         Left            =   7260
         MaxLength       =   2
         TabIndex        =   12
         Top             =   615
         Width           =   375
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm040101_1.frx":0085
         Height          =   1995
         Left            =   -70440
         TabIndex        =   136
         Top             =   2220
         Width           =   4215
         _ExtentX        =   7451
         _ExtentY        =   3535
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtCP118 
         Alignment       =   2  '置中對齊
         Height          =   300
         Left            =   3765
         MaxLength       =   1
         TabIndex        =   8
         Top             =   612
         Width           =   348
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   25
         Left            =   3780
         TabIndex        =   20
         Top             =   912
         Width           =   852
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公告日"
         Height          =   180
         Index           =   39
         Left            =   135
         TabIndex        =   150
         Top             =   2355
         Width           =   540
      End
      Begin VB.Label lblCP71 
         AutoSize        =   -1  'True
         Caption         =   "延緩月數/日期"
         Height          =   180
         Left            =   2595
         TabIndex        =   149
         Top             =   2355
         Visible         =   0   'False
         Width           =   1125
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   7
         Left            =   1035
         TabIndex        =   148
         Top             =   2310
         Width           =   855
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請案號"
         Height          =   180
         Index           =   16
         Left            =   6696
         TabIndex        =   146
         Top             =   3012
         Width           =   720
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   17
         Left            =   7428
         TabIndex        =   145
         Top             =   2940
         Width           =   1380
         VariousPropertyBits=   671107099
         MaxLength       =   20
         Size            =   "2434;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblPA178 
         AutoSize        =   -1  'True
         Caption         =   "證書形式      （1: 電子 2: 紙本）"
         Height          =   180
         Left            =   -74850
         TabIndex        =   144
         Top             =   3615
         Width           =   2475
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   2
         Left            =   1035
         TabIndex        =   31
         Top             =   2610
         Width           =   375
         VariousPropertyBits=   671107097
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "接洽單編號："
         Height          =   180
         Left            =   -74940
         TabIndex        =   140
         Top             =   420
         Width           =   1080
      End
      Begin MSForms.TextBox txtNote 
         Height          =   1200
         Left            =   -74940
         TabIndex        =   52
         Top             =   930
         Width           =   8745
         VariousPropertyBits=   -1466939365
         ScrollBars      =   3
         Size            =   "15425;2117"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "呈報內容："
         Height          =   180
         Left            =   -74940
         TabIndex        =   139
         Top             =   690
         Width           =   900
      End
      Begin MSForms.TextBox txtF0407 
         Height          =   1995
         Left            =   -74400
         TabIndex        =   138
         Top             =   2220
         Width           =   3885
         VariousPropertyBits=   -1466939361
         ScrollBars      =   3
         Size            =   "6853;3528"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "內容："
         Height          =   300
         Left            =   -74940
         TabIndex        =   137
         Top             =   2220
         Width           =   540
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   23
         Left            =   2925
         TabIndex        =   28
         Top             =   1995
         Width           =   1710
         VariousPropertyBits=   671107099
         Size            =   "3016;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   16
         Left            =   1035
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2025
         Width           =   855
         VariousPropertyBits=   671107097
         MaxLength       =   8
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   10
         Left            =   1260
         TabIndex        =   7
         Top             =   612
         Width           =   348
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "614;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   22
         Left            =   3420
         TabIndex        =   25
         Top             =   1695
         Width           =   1215
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "2143;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   11
         Left            =   5970
         TabIndex        =   26
         Top             =   1725
         Width           =   375
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "661;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   6
         Left            =   1260
         TabIndex        =   24
         Top             =   1695
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   4
         Left            =   1035
         TabIndex        =   19
         Top             =   1425
         Width           =   855
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   15
         Left            =   1035
         TabIndex        =   16
         Top             =   1155
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   12
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   5700
         TabIndex        =   21
         Top             =   1425
         Width           =   855
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   3
         Left            =   5700
         TabIndex        =   17
         Top             =   1155
         Width           =   855
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   9
         Left            =   7785
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   885
         Width           =   855
         VariousPropertyBits=   671107097
         MaxLength       =   7
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   12
         Left            =   5700
         TabIndex        =   14
         Top             =   885
         Width           =   855
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   21
         Left            =   -73560
         TabIndex        =   43
         Top             =   390
         Width           =   3975
         VariousPropertyBits=   671107099
         MaxLength       =   50
         Size            =   "7011;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   705
         Index           =   19
         Left            =   -73560
         TabIndex        =   44
         Top             =   690
         Width           =   7275
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "12832;1244"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   675
         Index           =   20
         Left            =   -73560
         TabIndex        =   45
         Top             =   1380
         Width           =   7275
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "12832;1191"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   13
         Left            =   1035
         TabIndex        =   13
         Top             =   885
         Width           =   495
         VariousPropertyBits=   671107099
         MaxLength       =   3
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   14
         Left            =   5715
         TabIndex        =   29
         Top             =   1995
         Width           =   900
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1587;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   18
         Left            =   7785
         TabIndex        =   30
         Top             =   1995
         Width           =   900
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1587;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   5700
         TabIndex        =   6
         Top             =   345
         Width           =   855
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   1035
         TabIndex        =   5
         Top             =   345
         Width           =   855
         VariousPropertyBits=   671107099
         MaxLength       =   6
         Size            =   "1508;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   8
         Left            =   5970
         TabIndex        =   33
         Top             =   2310
         Width           =   2655
         VariousPropertyBits=   671107099
         Size            =   "4683;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCP99 
         Height          =   495
         Left            =   -73560
         TabIndex        =   47
         Top             =   2400
         Width           =   7275
         VariousPropertyBits=   -1467987941
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "12832;873"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LblCP97 
         AutoSize        =   -1  'True
         Caption         =   "承辦人基數"
         Height          =   180
         Left            =   -74850
         TabIndex        =   131
         Top             =   3285
         Width           =   900
      End
      Begin VB.Label lblPA161 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司　                 (T:專利商標 J:智權公司 空白:系統預設)"
         Height          =   180
         Left            =   -71265
         TabIndex        =   130
         Top             =   2985
         Width           =   5055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "與他案合併計算結餘，請於案件備註欄註明""與某案號合併計算結餘""！"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   172
         Left            =   -72450
         TabIndex        =   129
         Top             =   2160
         Width           =   5550
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "是否為複雜或特殊案件         (Y:是)"
         Height          =   180
         Left            =   -74850
         TabIndex        =   128
         Top             =   2985
         Width           =   2670
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "是否電子送件         (Y:是)"
         Height          =   180
         Left            =   2655
         TabIndex        =   127
         Top             =   660
         Width           =   1950
      End
      Begin VB.Label Label1 
         Caption         =   "送件方式"
         Height          =   255
         Index           =   121
         Left            =   135
         TabIndex        =   122
         Top             =   2970
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件屬性"
         Height          =   180
         Index           =   168
         Left            =   2580
         TabIndex        =   121
         Top             =   2655
         Width           =   720
      End
      Begin VB.Label lblFeeYear 
         AutoSize        =   -1  'True
         Caption         =   "繳費年度：第         -         年"
         Height          =   180
         Left            =   6705
         TabIndex        =   120
         Top             =   1470
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦期限"
         Height          =   180
         Index           =   23
         Left            =   3012
         TabIndex        =   119
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PCT申請號"
         Height          =   180
         Index           =   22
         Left            =   2040
         TabIndex        =   118
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "取消收文日"
         Height          =   180
         Index           =   42
         Left            =   6705
         TabIndex        =   117
         Top             =   930
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "證書號數"
         Height          =   180
         Index           =   21
         Left            =   2610
         TabIndex        =   116
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PCT優先權日"
         Height          =   180
         Index           =   20
         Left            =   6705
         TabIndex        =   115
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "註記修改理由"
         Height          =   180
         Index           =   19
         Left            =   -74850
         TabIndex        =   114
         Top             =   2400
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人加乘註記"
         Height          =   240
         Index           =   17
         Left            =   -74850
         TabIndex        =   113
         Top             =   2130
         Width           =   1260
      End
      Begin VB.Label lblDivCase 
         AutoSize        =   -1  'True
         Caption         =   "分割母案本所案號"
         Height          =   180
         Left            =   4770
         TabIndex        =   112
         Top             =   2670
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利種類"
         Height          =   180
         Index           =   18
         Left            =   135
         TabIndex        =   111
         Top             =   2655
         Width           =   720
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   12
         Left            =   1485
         TabIndex        =   110
         Top             =   2655
         Width           =   1065
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
         Index           =   11
         Left            =   1590
         TabIndex        =   108
         Top             =   930
         Width           =   1350
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "2381;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "點數"
         Height          =   180
         Index           =   15
         Left            =   135
         TabIndex        =   107
         Top             =   2040
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "與國內                                   案號相同"
         Height          =   180
         Index           =   14
         Left            =   135
         TabIndex        =   106
         Top             =   1200
         Width           =   2835
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請國家"
         Height          =   180
         Index           =   13
         Left            =   135
         TabIndex        =   105
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PCT申請日"
         Height          =   180
         Index           =   12
         Left            =   4770
         TabIndex        =   104
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分所案號"
         Height          =   180
         Index           =   11
         Left            =   -74850
         TabIndex        =   91
         Top             =   390
         Width           =   720
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   1
         Left            =   6660
         TabIndex        =   86
         Top             =   390
         Width           =   1980
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "3492;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   180
         Index           =   0
         Left            =   1995
         TabIndex        =   85
         Top             =   390
         Width           =   1200
         VariousPropertyBits=   27
         Caption         =   "Label3"
         Size            =   "2117;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   3
         Left            =   -72000
         TabIndex        =   80
         Top             =   3630
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   2
         Left            =   -72120
         TabIndex        =   79
         Top             =   2490
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質"
         Height          =   180
         Index           =   37
         Left            =   4770
         TabIndex        =   78
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人"
         Height          =   180
         Index           =   0
         Left            =   135
         TabIndex        =   77
         Top             =   390
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "轉本所案號"
         Height          =   180
         Index           =   1
         Left            =   4770
         TabIndex        =   76
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限"
         Height          =   180
         Index           =   2
         Left            =   135
         TabIndex        =   75
         Top             =   1470
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "相關總收文號"
         Height          =   180
         Index           =   3
         Left            =   135
         TabIndex        =   74
         Top             =   1740
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否算案件數           (N : 不算)"
         Height          =   180
         Index           =   29
         Left            =   135
         TabIndex        =   72
         Top             =   660
         Width           =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收文日"
         Height          =   180
         Index           =   30
         Left            =   4770
         TabIndex        =   71
         Top             =   930
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限"
         Height          =   180
         Index           =   40
         Left            =   4770
         TabIndex        =   70
         Top             =   1470
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "卷宗性質                           (1.申請 2.異議 3.舉發)"
         Height          =   180
         Index           =   41
         Left            =   4770
         TabIndex        =   69
         Top             =   1200
         Width           =   3630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否取消閉卷             (Y : 取消閉卷)"
         Height          =   180
         Index           =   43
         Left            =   4770
         TabIndex        =   68
         Top             =   1770
         Width           =   2760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "進度備註"
         Height          =   180
         Index           =   9
         Left            =   -74850
         TabIndex        =   67
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件備註"
         Height          =   180
         Index           =   10
         Left            =   -74850
         TabIndex        =   66
         Top             =   1380
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號"
         Height          =   180
         Index           =   4
         Left            =   4770
         TabIndex        =   73
         Top             =   2355
         Width           =   1080
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一筆(&N)"
      Height          =   400
      Index           =   5
      Left            =   1802
      TabIndex        =   56
      Top             =   45
      Width           =   930
   End
   Begin VB.CommandButton Command2 
      Caption         =   "優先權資料"
      Height          =   400
      Index           =   4
      Left            =   4009
      TabIndex        =   58
      Top             =   45
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "擬制喪失新穎性關聯"
      Height          =   400
      Index           =   0
      Left            =   5090
      TabIndex        =   59
      Top             =   45
      Width           =   1425
   End
   Begin VB.CommandButton Command2 
      Caption         =   "案件進度"
      Height          =   400
      Index           =   1
      Left            =   6531
      TabIndex        =   60
      Top             =   45
      Width           =   885
   End
   Begin VB.CommandButton Command2 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   2
      Left            =   7432
      TabIndex        =   61
      Top             =   45
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   8250
      TabIndex        =   62
      Top             =   45
      Width           =   885
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   14
      Left            =   4608
      TabIndex        =   156
      Top             =   528
      Width           =   972
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "1714;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發文日"
      Height          =   180
      Index           =   25
      Left            =   4008
      TabIndex        =   155
      Top             =   528
      Width           =   540
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "目前表單狀態："
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   5790
      TabIndex        =   142
      Top             =   516
      Width           =   1230
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   24
      Left            =   2745
      TabIndex        =   90
      Top             =   480
      Visible         =   0   'False
      Width           =   855
      VariousPropertyBits=   671107099
      Size            =   "3413;476"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1080
      TabIndex        =   63
      Top             =   930
      Width           =   4215
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "9049;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   13
      Left            =   7200
      TabIndex        =   126
      Top             =   990
      Width           =   1545
      VariousPropertyBits=   27
      Caption         =   "lblEngGroup"
      Size            =   "2725;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工程師組別"
      Height          =   180
      Index           =   24
      Left            =   5805
      TabIndex        =   124
      Top             =   990
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label4"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3720
      TabIndex        =   109
      Top             =   744
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4"
      Height          =   180
      Index           =   35
      Left            =   3720
      TabIndex        =   103
      Top             =   1530
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3"
      Height          =   180
      Index           =   34
      Left            =   240
      TabIndex        =   102
      Top             =   1530
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2"
      Height          =   180
      Index           =   32
      Left            =   3720
      TabIndex        =   101
      Top             =   1260
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人1"
      Height          =   180
      Index           =   31
      Left            =   240
      TabIndex        =   100
      Top             =   1260
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5"
      Height          =   180
      Index           =   36
      Left            =   240
      TabIndex        =   99
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人"
      Height          =   180
      Index           =   38
      Left            =   3720
      TabIndex        =   98
      Top             =   1800
      Width           =   540
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   2
      Left            =   1260
      TabIndex        =   97
      Top             =   1260
      Width           =   2400
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "4233;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   3
      Left            =   4725
      TabIndex        =   96
      Top             =   1260
      Width           =   4050
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "7144;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   4
      Left            =   1260
      TabIndex        =   95
      Top             =   1530
      Width           =   2400
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "4233;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   5
      Left            =   4725
      TabIndex        =   94
      Top             =   1530
      Width           =   4050
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "7144;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   6
      Left            =   1260
      TabIndex        =   93
      Top             =   1800
      Width           =   2400
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "4233;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   7
      Left            =   4410
      TabIndex        =   92
      Top             =   1800
      Width           =   4380
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "7726;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   10
      Left            =   6750
      TabIndex        =   89
      Top             =   720
      Width           =   2010
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "3545;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   9
      Left            =   1080
      TabIndex        =   88
      Top             =   720
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2805;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label3 
      Height          =   180
      Index           =   8
      Left            =   1080
      TabIndex        =   87
      Top             =   510
      Width           =   1590
      VariousPropertyBits=   27
      Caption         =   "Label3"
      Size            =   "2805;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號"
      Height          =   180
      Index           =   5
      Left            =   240
      TabIndex        =   84
      Top             =   516
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號"
      Height          =   180
      Index           =   6
      Left            =   240
      TabIndex        =   83
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱"
      Height          =   180
      Index           =   7
      Left            =   240
      TabIndex        =   82
      Top             =   930
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員"
      Height          =   180
      Index           =   8
      Left            =   5805
      TabIndex        =   81
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "frm040101_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Morgan 2021/12/3 改成Form2.0 (Text1,Label3...)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'Modified by Morgan 2013/5/28 PPH(431)控制改為TW-SUPA(434)
'2005/7/5整理
Option Explicit

'Modify By Cheng 2002/07/08
'Dim StrTot1(0 To 500) As String, StrTot2(0 To 500) As String
Dim StrTot1(0 To 1023) As String, StrTot2(0 To 1023) As String
Dim IntNow As Integer, IntTot As Integer
Dim strReceiveNo As String
Dim cm(7) As String
'Modify by Morgan 2005/3/2
'Dim pa(T_PA) As String, cp(T_CP) As String
Dim pa() As String, cp() As String

Dim intWhere As Integer, intRow As Integer
'Modify by Amy 2023/01/05 設為Public改變數
'Dim strPriority(1 To 5) As String 'Modify by Amy 2014/06/10
Public strPrity1 As String, strPrity2 As String, strPrity3 As String, strPrity4 As String, strPrity5 As String
'end 2023/01/05
Dim strDateTmp(1 To 2) As String
' 案件性質
Dim m_CP10 As String
' 本所期限
Dim m_CP06 As String
' 法定期限
Dim m_CP07 As String
Dim m_CP31 As String 'Add By Sindy 2010/10/29
'Add By Cheng 2002/06/11
Dim m_strCP06 As String '記錄原始的本所期限
Dim m_strCP07 As String '記錄原始的法定期限
Dim m_strCP07_1 As String '記錄計算出來的原始法定期限  add by sonia 2018/2/21
Dim strFirstPriDate As String  '最早的優先權日期
Dim m_strCP06Update As String '更新後的本所期限
'Add by Morgan 2004/2/18
'若承辦人是王協理且未發文則要發EMail通知
Dim stCP09 As String, stCP14 As String, stCP27 As String
'Add by Morgan 2004/3/29
'實審通知日本所期限,法定期限,母案收文號,申請國家
Dim m_stVar(0 To 3) As String, m_stCP09 As String, m_stPA09 As String
'Add by Morgan 2004/6/7
'一案兩申請
Dim m_PA2 As String
Dim m_bolActive As Boolean
Dim m_stCP98 As String 'Add by Morgan 2005/3/4
Dim m_EP06 As String 'Add by Morgan 2005/3/29
'Add by Morgan 2006/4/25
Dim m_bol30xMail As Boolean '改請案是否Mail通知承辦
Dim m_bol30xMailDesc As String ''改請案通知承辦的Mail內容
'Add by Morgan 2006/5/23
Dim strPCTPriDate As String
'Add by Morgan 2006/9/7
Dim bol416Msg As Boolean '是否提醒程序下一程序實審已結
Dim bol416Mail As Boolean '是否發Mail通知智權人員實審做銷案
Dim bol414Rec As Boolean '是否PCT大陸發明案有收文恢復專利權
Dim strMail2FEngCP09 As String '轉案號發Mail給國外案工程師的收文號
Dim m_str605NP22 As String 'Add by Morgan 2006/10/26 取消閉卷恢復年費管制的NP22
Dim m_bolUpdCP27 As Boolean 'Add by Morgan 2006/12/27是否上發文日
Dim strPA14Msg As String 'Add by Morgan 2007/1/30 國內案已公告訊息
'2008/11/13 add by sonia 相關總收文號的資料
Dim m_CP43CP08 As String
Dim m_CP43CP64 As String
'2008/11/13 END
Dim m_bolFMP As Boolean 'Add by Morgan 2009/11/4 是否 FMP 案
Dim m_bolFMP2 As Boolean 'Added by Lydia 2021/05/31 是否為寰華案
'Add by Morgan 2010/1/20
Dim m_lngRefund As Long '未退預繳年費金額
Dim m_i605FromYear As Integer, m_i605ToYear As Integer '繳年費起迄年
Dim m_bolCP98Check As Boolean
Dim m_oldEP04 As String, m_newEP04 As String 'Add by Morgan 2010/6/18
Dim m_strOldCP10 As String, m_strOldPA09 As String 'Add by Morgan 2010/10/29
Dim m_CP30 As String 'Add by Morgan 2011/4/22
Dim m_RefCP53 As String 'Add by Morgan 2011/8/17 +核准函紀錄的繳費起始年度
Dim m_bolCtrl231 As Boolean '是否管制寄存證明 Added by Morgan 2012/7/2
Dim m_strPriType As String '主張優先權期限適用類別 1.發明或新型,2.設計 Added by Morgan 2012/7/4
Dim m_CP31isYGetCP05 As String 'Add By Sindy 2014/1/29
'Add by Amy 2015/01/22
Dim m_bolIsFirstKeyCP14 As Boolean '北所第一次輸承辦人
Dim m_bolChkCP14OK As Boolean
Dim bolCP14Mail As Boolean '承辦人員修改後是否發mail
'Removed by Morgan 2021/4/8
''Added by Lydia 2015/05/13 同時收文的案件性質
'Dim m_TogCP10 As String
'Dim m_bolTogCP10 As Boolean
'Dim m_bolUpdCP07 As Boolean '法限+在途15天
'end 2021/4/8
Dim m_bolAuto404 As Boolean '是否自動收文延期 Added by Morgan 2015/12/14
Dim m_DualCaseNo(1 To 4) As String 'Added by Lydia 2016/09/29 發明一案兩請案的新型案本所號
Dim m_bolAuto429 As Boolean 'Added by Lydia 2016/09/29 一案兩請的新型案是否自動收文放棄專利權
Dim m_bolTw_SUPA_LawDateChk As Boolean 'Added by Morgan 2016/10/7 TW-SUPA 是否檢查法限
Dim m_HKMemo As String 'Added by Morgan 2018/4/18 香港關聯CFP案原因
Dim m_str442DeadLine As String '加在途期間後之法限 Added by Morgan 2020/2/7
'Added by Lydia 2020/05/20 法律所案源收文
Dim m_LOS15 As String '案源單號
Dim m_LOS01 As String '案源總收文號
Dim m_LOS02 As String 'Added by Lydia 2020/06/09 案源案件類型
Dim m_LOS07 As String '放棄日期
Dim m_RefCP10 As String 'Added by Morgan 2021/10/5 相關案號案件性質
'Add by Amy 2022/10/17
Dim stCPM35 As String, stF0207_A6 As String, m_F0308 As String, m_F0309 As String, strUpdDate As String, strUpdTime As String, IsEConsultRec As Boolean
Dim stF0307_Now As String, stF0309_Now As String '登入時F030X值
Dim pSaveMsg As String 'Added by Lydia 2023/03/25
Dim m_CN307Updated As Boolean 'Added by Morgan 2023/5/19
Dim strMsgCloseCancel As String 'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之年費605、維持費606、延展費607，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。

'Modify By Sindy 2023/1/10
'Private Sub Process()
'Optional intRunType As Integer = 0 : 0.分案確定
'                                     1.補件完成,呼叫此函數為了更新畫面上修正的欄位值
Private Function Process(Optional intRunType As Integer = 0) As Boolean
'2023/1/10 END
Dim i As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strOfficeKind As String '所別
'Add by Morgan 2004/2/17
'是否已發Mail
Dim bolMail As Boolean
'Add by Morgan 2004/7/6
Dim stPS As String   'Mail 備註
'Add by Morgan 2004/7/20
Dim stAppNo As String   '未設定減免身分客戶代碼
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim bol106 As Boolean '是否有收文主張國際優先權
Dim bol106Mail As Boolean '是否通知智權人員收文主張國際優先權
Dim bol421Check1 As Boolean '是否同案收第2次以上技術報告
Dim bol421Check2 As Boolean, str421Check2 As String '是否他案已准技術報告
Dim bol421Check3 As Boolean, str421Check3 As String '是否他案已提技術報告

Dim m_strCP07Update As String '更新後的法定期限
Dim bolDiscount As Boolean 'Add by Morgan 2010/1/21 是否年費可減免
Dim oTopForm As Form 'Add by Morgan 2010/8/5
Dim strAnoPA() As String, strCaseNo As String, bolIsExistsCP10 As Boolean 'Add By Sindy 2014/7/8
'Added by Lydia 2016/01/28
Dim strPD As String '判斷國際或國內優先權
Dim tmpContent As String
Dim bolCancel As Boolean
Dim stVTable As String
Dim bolCheck As Boolean 'Added by Morgan 2019/12/26
Dim m_StrTo As String, m_StrSub As String, m_StrCont As String 'Added by Lydia 2022/05/30 整理frm880005改用寄信模組
Dim strCRC08 As String
   
   Process = False
   'Add by Amy 2014/11/19 +當程序(P12)操作台灣A類非電子收文時檢查接洽單
   'Modified by Morgan 2016/6/22 +非臺灣案
   'If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
   '    If Pub_StrUserSt03 = "P12" And pa(1) = "P" And pa(9) = 台灣國家代號 And Me.Label3(8) < "B" Then
   'Add By Sindy 2021/3/9 增加控管轉本所案號時,不檢查接洽單PDF檔
   If Not (textPA1 <> "" And textPA2 <> "") Then
   '2021/3/9 END
      If Pub_StrUserSt03 = "P12" And Me.Label3(8) < "B" Then
        'Modified by Morgan 2016/6/30 +排除FMP案
        'Removed by Morgan 2018/11/5 改FMP案也要檢查接洽單--玲玲
        'If (內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(cp(12), 1) <> "F") Or pa(9) = 台灣國家代號 Then
        'end 2018/11/5
        'end 2016/6/30
      'end 2016/6/22
            'Modify By Sindy 2022/12/16 電子收文不用檢查
            If Not (txtF0301 <> "" And Len(txtF0301) = 10) Then
            '2022/12/16 END
              If PUB_CheckPDF2(cp(9), 0, True, strExc(0), Text1(1)) = False Then
                  MsgBox "無接洽單PDF檔,不可分案!", vbCritical
                  Exit Function
              End If
            End If
          'End If 'Removed by Morgan 2018/11/5
      End If
      'end 2014/11/19
      'Add by Amy 2022/12/09 電子收文需檢查一案兩請是否有資料
      'Modify by Amy 2022/12/23 +案件性質新型申請(102)才彈
      'Modify by Sindy 2023/3/9 +案件性質新型申請(101)也要檢查
      If txtF0301 <> MsgText(601) And (Text1(13) = 台灣國家代號 Or Text1(13) = 大陸國家代號) And (Text1(1) = "102" Or Text1(1) = "101") Then
            ReDim Preserve strAnoPA(1 To TF_PA) As String
            If Pub_GetField("ConsultRecordList", "CRL01='" & txtF0301 & "'", "CRL67") <> MsgText(601) Then
                If PUB_IsDualApplyCom(pa, strAnoPA, strCaseNo) = False Then
                    MsgBox "一案兩請無資料,不可繼續!", vbCritical
                    Exit Function
                End If
            End If
      End If
      'end 2022/12/09
   End If
   
   'Added by Lydia 2023/12/14 檢查智財協作在分案時若未建立相關案號(caserelation1)時則跳提醒程序人員，但可選擇輸或不輸 !
   'Modified by Lydia 2023/12/15 PS及CPS之智財協作967，TT及S之智財協作737，L之智財協作7601，(也可用案件性質中文判斷)在分案時若未建立相關案號且為ACS且為TIPS的案件時，提醒文字：「案件性質為智財協作，請先依接洽單輸入相關卷號資料」。
   'If pa(1) = "PS" And Text1(1) = "967" Then
   '   If PUB_IfCaseRelation1Exists(pa(1), pa(2), pa(3), pa(4)) = False Then
   '      If MsgBox("案件性質為" & Label3(1).Caption & "，請確認接洽單是否有相關案號，是否補輸入？", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
   '         Exit Function
   '      End If
   '   End If
   'End If
   'end 2023/12/14
   If pa(1) = "PS" And InStr(Label3(1).Caption, "智財協作") > 0 Then
      'Modified by Lydia 2024/05/13 +113=智財布局分析，也為智財協作的案件性質之一
      If PUB_ChkACSforTIPS(pa(1) & pa(2) & pa(3) & pa(4), , True, "113") = False Then
         MsgBox "案件性質為" & Label3(1).Caption & "，請先依接洽單輸入相關卷號資料", vbExclamation
         Exit Function
      End If
   End If
   'end 2023/12/15
   
   pSaveMsg = "" 'Added by Lydia 2023/03/25
   
   '若非執行轉本所案號時
   If Me.textPA1.Text = "" Or Me.textPA2.Text = "" Then
      'Add by Amy 2018/04/09 P案年費移作次年時若下一程序有605~607期限則提醒並不可存檔
      If pa(1) = "P" And Text1(1) = "612" Then
         With MSHFlexGrid1
            For i = 1 To .Rows - 1
               'modify by sonia 2025/4/28 加入判斷是否已解除期限.TextMatrix(i, 6)
               If .TextMatrix(i, 0) = "" And .TextMatrix(i, 6) = "" And (.TextMatrix(i, 8) = "605" Or .TextMatrix(i, 8) = "606" Or .TextMatrix(i, 8) = "607") Then
                   MsgBox "有未收文之" & .TextMatrix(i, 1) & "期限，不可存檔!"
                   Exit Function
               End If
            Next i
         End With
      End If
      'end 2018/04/09
      '2011/11/14 ADD BY SONIA P-095813 舉發答辯檢查應有年費期限
      If Text1(1) = 舉發答辯 Then
         CheckOC3
         StrSQLa = "select cp09 from caseprogress where '" & pa(1) & "'=cp01(+) and '" & pa(2) & "'=cp02(+) and '" & pa(3) & "'=cp03(+) and '" & pa(4) & "'=cp04(+) and '605'=cp10(+) and cp27 is null and cp57 is null " & _
                   "union select np01 from nextprogress where '" & pa(1) & "'=np02(+) and '" & pa(2) & "'=np03(+) and '" & pa(3) & "'=np04(+) and '" & pa(4) & "'=np05(+) and np07 in ('605','606','607') and np06 is null"
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If AdoRecordSet3.RecordCount = 0 Then
            'Modified by Morgan 2016/9/20
            '改提醒但可繼續，因有可能年費已繳完(無下次繳費日) Ex.P-85361
            'MsgBox "案件性質為舉發答辯時，但此案目前無年費期限 !", vbCritical
            'Exit Sub
            If MsgBox("案件性質為舉發答辯，但此案目前無年費期限！是否確定要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            End If
            'end 2016/9/20
         End If
      End If
      '2011/11/14 END
      
      'Modify by Morgan 2004/8/16   新法延緩公告申請書不再需要相關總收文號
      'If (Text1(1) = 請求公告 Or Text1(1) = 延緩公告) And Text1(6) = "" Then
         'MsgBox "案件性質為請求公告或延緩公告時，相關總收文號不可空白 !", vbCritical
         
      If Text1(1) = 請求公告 And Text1(6) = "" Then
         MsgBox "案件性質為請求公告時，相關總收文號不可空白 !", vbCritical
         Exit Function
      End If

      '2005/8/30 ADD BY SONIA
      If Text1(13) = 台灣國家代號 And (Text1(1) = 修正 Or Text1(1) = 申復) Then
         If Text1(6) = "" Then
            strTit = "台灣修正或申復案件"
            strMsg = "請確認是否無來函相關總收文號?"
            nResponse = MsgBox(strMsg, vbYesNo + vbDefaultButton2, strTit)
            If nResponse = vbNo Then
               Exit Function
            End If
         End If
      End If
      '2005/8/30 END
      
      '2010/2/4 ADD BY SONIA
      If Left(Text1(1), 1) = "3" And Text1(1) <> "307" Then
         If Text1(4) = "" And Text1(5) = "" Then
            strTit = "改請案件性質"
            strMsg = "改請案件性質，請確認是否輸入期限 ? "
            nResponse = MsgBox(strMsg, vbYesNo + vbDefaultButton1, strTit)
            If nResponse = vbYes Then
               If Text1(4) = "" Then
                  Text1(4).SetFocus
               End If
               Exit Function
            End If
         End If
      End If
      '2010/2/4 END
   End If
    
   If CheckDataValid() = False Then
      Exit Function
   End If
   
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Function
   
   If textPA1 <> "" And textPA2 <> "" Then
       strExc(1) = textPA1
       strExc(2) = textPA2
       strExc(3) = textPA3
       If strExc(3) = "" Then strExc(3) = "0"
       strExc(4) = textPA4
       If strExc(4) = "" Then strExc(4) = "00"
       strExc(5) = Text1(1).Text '案件性質
       strExc(6) = Label3(1) '案件性質名稱
       strExc(7) = Text1(12) '收文日
       strExc(8) = Label3(8) '總收文號
       strExc(9) = pa(26)
       'edit by nickc 2007/02/05 不用 dll 了
       'If Not objLawDll.ChkSameCase(strExc) Then Exit Sub
       If Not ClsLawChkSameCase(strExc) Then Exit Function
       'Added by Lydia 2020/08/18 更新相關卷號前,先檢查是否有重複
       If m_CP31 = "Y" Then
          If PUB_ChkUpdCR(pa(1), pa(2), pa(3), pa(4), strExc(1), strExc(2), strExc(3), strExc(4)) = False Then
              Exit Function
          End If
       End If
       'end 2020/08/18
    End If
    
    If Me.textPA1.Text <> "" And Me.textPA2.Text <> "" Then
       MsgBox "本次只更新轉本所案號的資料，不更新其他欄位的修改!!!", vbCritical
      'Add by Morgnan 2006/9/20 關聯案提醒
      Me.Tag = ""
      If InStr(CaseMapIn, Text1(1)) > 0 Then
         Set frm1104_1.m_form = Me
         frm1104_1.m_CP09 = cp(9)
         frm1104_1.m_CaseNoBefore = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
         frm1104_1.m_CaseNoAfter = textPA1 & "-" & textPA2 & "-" & Right("0" & textPA3, 1) & "-" & Right("00" & textPA4, 2)
         If frm1104_1.GetRelation = True Then
            frm1104_1.Show vbModal
            If Me.Tag = "0" Then
               Exit Function
            End If
         End If
      End If
      'end 2006/9/20
       
    'Add by Morgan 2004/3/17
    Else
       '分割案
       If Text1(1) = "307" Then
          Erase m_stVar(): m_stCP09 = ""
          If txtDivCaseNo(1).Text = "" Or txtDivCaseNo(2).Text = "" Then
             'Modify by Morgan 2004/11/30 所有都要控制<--AMY
             If pa(1) = "P" Then
                MsgBox "分割案必須輸入分割母案本所案號！", vbCritical
                txtDivCaseNo_GotFocus 1
                txtDivCaseNo(1).SetFocus
                Exit Function
             End If
             'End 2004/11/30
             
          '檢查母案本所案號是否存在
          'Modified by Morgan 2012/11/8 改呼叫公用函數檢查
          'ElseIf CheckDivCase(m_stPA09) = False Then
          ElseIf PUB_CheckDivCase(txtDivCaseNo, pa) = False Then
             txtDivCaseNo_GotFocus 1
             txtDivCaseNo(1).SetFocus
             Exit Function
             
'Removed by Morgan 2011/12/7 P案收文不必管控實審期限(發文管控的才是真的法限),而且母案應該都會有期限會更新到分割--郭 Ex.P-100262
'          'FCP及P的國內案件,當專利種類為'發明'且案件性質為'分割'時,若無收文未取消收文之'實體審查'則顯示'此分割案尚未收文實體審查，期限為XXXXXX，請提醒智權人員 !!
'          ElseIf pa(1) = "P" And Text1(2) = "1" And Text1(13) = "000" Then
'
'             Dim strTmp1(0 To 4) As String, strTmp(1 To 3) As String, bolMsg As Boolean
'             For i = 1 To 4
'                strTmp1(i) = txtDivCaseNo(i)
'             Next
'             '讀取實體審查得法定期限
'             If GetMoneyDate(4, m_stPA09, strTmp1, strTmp(1), strTmp(2), strTmp(3)) = True Then
'                If strTmp(3) <> "" Then
'                   strTmp(3) = CompDate(2, 1, strTmp(3))
'                   '法定期限
'                   m_stVar(3) = PUB_Get416LawLimit(Text1(12), strTmp(3))
'                   '本所期限= 法定期限-4天
'                   m_stVar(0) = PUB_GetWorkDay1(CompDate(2, -4, m_stVar(3)), True)
'                   'Add by Morgan 2004/6/7
'                   '本所期限在三個月內才提示
'                   If m_stVar(0) > strTmp(3) Or CompDate(1, 3, m_stVar(0)) > strTmp(3) Then
'                      bolMsg = True
'                   Else
'                      bolMsg = False
'                   End If
'                   '檢查有收'實體審查'否，有則抓收文號-->m_stCP09
'                   If PUB_Get416CP09(m_stCP09, ChangeWStringToTString(m_stVar(0)), pa(), bolMsg) = False Then
'                      Exit Sub
'                   End If
'                Else
'                   MsgBox "無法讀取實體審查的法定期限！", vbCritical
'                   Exit Sub
'                End If
'             Else
'                MsgBox "無法讀取實體審查的法定期限！", vbCritical
'                Exit Sub
'             End If

          End If
          
          'Added by Morgan 2017/4/5
          If pa(9) = "020" And cp(27) = "" Then
            m_CN307Updated = False 'Added by Morgan 2023/5/19
            'Modified by Morgan 2019/12/26 +bolCheck 母案是分割案時才要提醒 Ex:P-124176--郭
            bolCheck = False
            PUB_Get307CtrlDate txtDivCaseNo(1), txtDivCaseNo(2), txtDivCaseNo(3), txtDivCaseNo(4), , strExc(1), bolCheck
            '最原始的母案之領證期限已過
            If Val(strExc(1)) > 0 And strExc(1) < strSrvDate(1) Then
               '母案為分割案
               If bolCheck = True Then 'Added by Morgan 2019/12/26
                  'Modified by Morgan 2023/5/19 大陸分割期限控制--陳玲玲
                  'If MsgBox("分案期限(法限：" & ChangeWStringToTDateString(strExc(1)) & ")已過！請確認此次分案是否依據審查意見通知書指示辦理？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                  '   Exit Function
                  'End If
                  strExc(0) = "select cp06,cp07,cpm04 From caseprogress,casepropertymap" & _
                     " where cp01='" & txtDivCaseNo(1) & "' and cp02='" & txtDivCaseNo(2) & "' and cp03='" & txtDivCaseNo(3) & "' and cp04='" & txtDivCaseNo(4) & "'" & _
                     " and cp10 in ('1202','1307') and cpm01(+)=cp01 and cpm02(+)=cp10 order by cp07 desc"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  '分割母案有收到審查意見通知或分案通知
                  If intI = 1 Then
                     '是否依指示辦理
                     If MsgBox("分案期限(法限：" & ChangeWStringToTDateString(strExc(1)) & ")已過！請確認此次分案是否依據審查意見通知書指示辦理？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                        Exit Function
                     '未過期
                     ElseIf RsTemp("cp07") > strSrvDate(1) Then
                        m_CN307Updated = True
                        Text1(5) = TransDate(RsTemp("cp07"), 1)
                        Text1(4) = TransDate(RsTemp("cp06"), 1)
                        MsgBox "期限已設定為分割母案" & strExc(2) & "期限！", vbExclamation
                     '已過期
                     Else
                        strExc(2) = RsTemp("cpm04")
                        strExc(3) = RsTemp("cp07")
                        strExc(0) = "select np08,np09,cp27 From nextprogress,caseprogress" & _
                           " where np02='" & txtDivCaseNo(1) & "' and np03='" & txtDivCaseNo(2) & "' and np04='" & txtDivCaseNo(3) & "' and np05='" & txtDivCaseNo(4) & "'" & _
                           " and np07='601' and cp01(+)=np02 and cp02(+)=np03 and cp03(+)=np04 and cp04(+)=np05 and cp10(+)=np07"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        '分割母案有領證期限
                        If intI = 1 Then
                           '未過期
                           If RsTemp("np09") > strSrvDate(1) Then
                              m_CN307Updated = True
                              Text1(5) = TransDate(RsTemp("np09"), 1)
                              Text1(4) = TransDate(RsTemp("np08"), 1)
                              MsgBox "期限已設定為分割母案領證期限！", vbExclamation
                           '已過期
                           Else
                              MsgBox "分割母案領證期限(法限：" & ChangeWStringToTDateString(RsTemp("np09")) & ")已過！", vbExclamation
                              Exit Function
                           End If
                        '分割母案無領證期限
                        Else
                           MsgBox "分割母案" & strExc(2) & "期限(法限：" & ChangeWStringToTDateString(strExc(3)) & ")已過且尚無領證期限！", vbExclamation
                           Exit Function
                        End If
                     End If
                  '分割母案未收到審查意見通知或分案通知
                  Else
                     MsgBox "分割母案未收到審查意見通知或分案通知，本案無法辦理[分案]申請", vbExclamation
                     Exit Function
                  End If
                  'end 2023/5/19
               End If
            End If
          End If
          'end 2017/4/5
          
         'Added by Morgan 2023/11/28
         'Modified by Morgan 2023/11/30 先限台灣案,大陸案要再確認--郭 Ex:P116808
         If pa(9) = "000" And cp(27) = "" Then
            strExc(0) = "select * from casemap where cm10='3' and cm01='" & txtDivCaseNo(1) & "'" & " and cm02='" & txtDivCaseNo(2) & "' and cm03='" & txtDivCaseNo(3) & "' and cm04='" & txtDivCaseNo(4) & "'" & _
               "union select * from casemap where cm10='3' and cm05='" & txtDivCaseNo(1) & "'" & " and cm06='" & txtDivCaseNo(2) & "' and cm07='" & txtDivCaseNo(3) & "' and cm08='" & txtDivCaseNo(4) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "母案為一案兩請，請確認分割案是否援用一案兩請。", vbExclamation
            End If
         End If
         'end 2023/11/28
          
       '台灣新型(核准日>930701)收鑑定報告時，若未收技術報告時，發e-mail通知智權人員。
       ElseIf Text1(1) = "906" Then
         If Text1(13) = "000" And Text1(2) = "2" And Val(Text1(12)) >= 930701 And Val(pa(20)) >= 930701 Then
            If Chk421Exist(pa, Text1(3)) = False Then
               stPS = "※注意，本案尚未收技術報告！"
            End If
         End If
       End If
       
       'Add by Morgan 2004/6/8
       'P、CFP一案二申請於分案時建立關聯；提醒條件：同一申請人同一天收文同一申請國家同一案件名稱但不同專利種類時，若未建立關聯則提醒使用者。
       If InStr("101,102,103", Text1(1)) > 0 Then
         m_PA2 = ""
         If PUB_DualCaseExist(pa, m_PA2) = True Then
            If PUB_DualCaseRelationExist(pa) = False Then
               If MsgBox("本案與 " & m_PA2 & " 案可能為一案兩申請且尚未建立關聯，確定要繼續？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Function
            End If
         End If
      End If
      
      'add by sonia 2015/9/15 澳門發明案一定要輸大陸案關聯
      If Text1(13) = "044" And Text1(1) = "101" Then
         If Text1(15) = "" Then
            MsgBox "澳門的發明案，國內案號不可空白！", vbInformation
            Text1(15).SetFocus
            Text1_GotFocus 15
            Exit Function
         Else
            CheckOC3
            ChgCaseNo Text1(15).Text, strExc
            StrSQLa = "select pa09 from patent where pa01='" & strExc(1) & "' and pa02='" & strExc(2) & "' and pa03='" & strExc(3) & "' and pa04='" & strExc(4) & "' "
            AdoRecordSet3.CursorLocation = adUseClient
            AdoRecordSet3.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If AdoRecordSet3.RecordCount <> 0 Then
               '檢查是否為大陸案
               If CheckStr(AdoRecordSet3.Fields("pa09")) <> "020" Then
                  MsgBox "澳門的發明，國內案號應填入申請國家是大陸的案號！", vbInformation
                  Text1(15).SetFocus
                  Text1_GotFocus 15
                  CheckOC3
                  Exit Function
               End If
            Else
               MsgBox "此案號不存在！", vbInformation
               Text1(15).SetFocus
               Text1_GotFocus 15
               CheckOC3
               Exit Function
            End If
            CheckOC3
         End If
      End If
      'end 2015/9/15
      
      m_HKMemo = "" 'Added by Morgan 2018/4/18
      'add by nickc 2005/06/07 加入檢查若是香港，則要有關聯到大陸的案子
      'Modify by Morgan 2006/8/31 標準專利紀錄請求(110)才要
      'If text1(13).Text = "013" Then
      If Text1(13).Text = "013" And Text1(1) = "110" Then
         If Text1(15) = "" Then
               If Text1(2) = "1" Then
                  bolCancel = False 'Added by Morgan 2018/4/18
                  'Modified by Morgan 2017/12/20
                  'MsgBox "香港的發明案，國內案號不可空白！", vbInformation
                  If MsgBox("香港的發明案，尚未輸入國內案號！是否確認無關聯案號？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                  'Modified by Morgan 2018/4/18 改提示輸入基礎案的公開日,由系統自動算本所期限及法定期限,(公開日加六個月為法定期限)--玲玲
                  '   If Text1(4) = "" Or Text1(5) = "" Then
                  '      MsgBox "請輸入本所期限及法定期限！", vbExclamation
                  '      If Text1(4) = "" Then Text1(4).SetFocus: Exit Sub
                  '      If Text1(5) = "" Then Text1(5).SetFocus: Exit Sub
                  '   End If
                  'Else
                     If cp(27) = "" Then
                        strExc(0) = InputBox("基礎案公開日：" & vbCrLf & vbCrLf & "(格式: yyyymmdd, 例: 20180101)", "請輸入基礎案公開日！")
                        If strExc(0) = "" Then
                           bolCancel = True
                        ElseIf ChkDate(strExc(0)) = False Then
                           bolCancel = True
                        Else
                           m_HKMemo = ChangeWStringToWDateString(strSrvDate(1)) & " 基礎案公開日：" & strExc(0) & "; "
                           strExc(1) = ""
                           Text1(5) = TransDate(CompDate(1, 6, strExc(0)), 1)
                           Text1(4) = TransDate(PUB_GetWorkDay1(CompDate(2, -5, CompDate(1, -1, Text1(5))), True), 1)
                        End If
                     End If
                  'Added by Morgan 2023/2/10
                  Else
                     bolCancel = True
                  'end 2023/2/10
                  End If
                  If bolCancel Then
                  'end 2018/4/18
                  'end 2017/12/20
                     Text1(15).SetFocus
                     Text1_GotFocus 15
                     Exit Function
                  End If 'Added by Morgan 2017/12/20
               End If
         'Modified by Morgan 2018/4/18
         'Else
         ElseIf Text1(15).Tag <> Text1(15) Then
         'end 2018/4/18
               CheckOC3
               ChgCaseNo Text1(15).Text, strExc
               'Modify by Morgan 2007/4/26
               'StrSQLa = "select * from patent where pa01='" & strExc(1) & "' and pa02='" & strExc(2) & "' and pa03='" & strExc(3) & "' and pa04='" & strExc(4) & "' "
               StrSQLa = "select p1.pa09, p2.pa09 pa09x from patent p1,patent p2 where p1.pa01='" & strExc(1) & "' and p1.pa02='" & strExc(2) & "' and p1.pa03='" & strExc(3) & "' and p1.pa04='" & strExc(4) & "' and p2.pa01(+)=p1.pa01 and p2.pa02(+)=p1.pa02 and p2.pa03(+)=p1.pa03 and p2.pa09(+)=201"
               AdoRecordSet3.CursorLocation = adUseClient
               AdoRecordSet3.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If AdoRecordSet3.RecordCount <> 0 Then
                  If Text1(2) = "1" Then
                     bolCancel = False
                     '檢查是否為大陸案
                     'Modify by Morgan 2007/4/26 香港也可和EPC、英國關聯
                     'If CheckStr(AdoRecordSet3.Fields("pa09")) <> "020" Then
                     '   MsgBox "香港的發明，國內案號應填入申請國家是大陸的案號！", vbInformation
                     If (AdoRecordSet3("pa09") <> "020" And AdoRecordSet3("pa09") <> "221" And AdoRecordSet3("pa09") <> "201") Then
                        MsgBox "香港的標準專利只能與大陸、EPC或英國案件關聯！", vbInformation
                        bolCancel = True
                     ElseIf AdoRecordSet3("pa09") = "221" Then
                        If IsNull(AdoRecordSet3("pa09x")) Then
                           MsgBox "該關聯案為EPC案但未指定英國！", vbInformation
                           bolCancel = True
                        End If
                     End If
                     'Added by Morgan 2018/4/18
                     '若相關案為CFP案但又有相關大陸案時彈訊息提醒並確認
                     If bolCancel = False And strExc(1) = "CFP" Then
                        stVTable = PUB_GetRefCaseMapSQL(strExc, False)
                        strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNNo" & _
                           " from (" & stVTable & ") X,patent where pa01(+)=C01 and pa02(+)=C02 and pa03(+)=C03 and pa04(+)=C04 and pa09='020'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If MsgBox("本案相關案是否確認為 " & strExc(1) & "-" & strExc(2) & IIf(strExc(3) & strExc(4) = "000", "", "-" & strExc(3) & "-" & strExc(4)) & " 無誤？" & vbCrLf & "(該案有一相關大陸案 " & RsTemp("CNNo") & ")", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                              m_HKMemo = InputBox("原因：", "香港相關案為CFP案時需輸入原因")
                              If m_HKMemo = "" Then
                                 bolCancel = True
                              Else
                                 m_HKMemo = ChangeWStringToWDateString(strSrvDate(1)) & " 關聯至CFP案原因：" & m_HKMemo & "; "
                              End If
                           Else
                              bolCancel = True
                           End If
                        End If
                     End If
                     'end 2007/4/26
                     If bolCancel Then
                        Text1(15).SetFocus
                        Text1_GotFocus 15
                        CheckOC3
                        Exit Function
                     End If
                  Else
                     '檢查是否為台灣或是大陸
                     If CheckStr(AdoRecordSet3.Fields("pa09")) <> "020" And CheckStr(AdoRecordSet3.Fields("pa09")) <> "000" Then
                        MsgBox "香港的新型、設計，國內案號應填入申請國家是大陸或台灣的案號！", vbInformation
                        Text1(15).SetFocus
                        Text1_GotFocus 15
                        CheckOC3
                        Exit Function
                     End If
                  End If
               Else
                  MsgBox "此案號不存在！", vbInformation
                  Text1(15).SetFocus
                  Text1_GotFocus 15
                  CheckOC3
                  Exit Function
               End If
               CheckOC3
         End If
         
      End If
      
      '台灣案檢查減免身分(異議，舉發不用)
      'Modify by Morgan 2007/8/30 加第三人申請技術報告
      'If Text1(13).Text = "000" And InStr("801,803", Text1(1).Text) = 0 Then
      If Text1(13).Text = "000" And InStr("801,803,807", Text1(1).Text) = 0 Then
         For i = 1 To 5
            If txtAD(i).Enabled = True Then
               If txtAD(i).Text = "" Then
                  MsgBox "申請人減免身分不可空白", vbInformation
                  txtAD(i).SetFocus
                  txtAD_GotFocus i
                  Exit Function
               '公司可減免
               'Modify by Morgan 2004/7/29
               '學校不需證明
               'ElseIf (txtAD(i).Text = "2" Or txtAD(i).Text = "3") Then
               '學校
               ElseIf (txtAD(i).Text = "2") Then
                  '變更
                  If (txtAD(i).Tag <> "2" And txtAD(i).Tag <> "") Then
                     If MsgBox("確定要變更申請人【" & pa(25 + i) & "-" & Label3(i + 1) & "】減免身分為【學校】？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                        txtAD(i).SetFocus
                        txtAD_GotFocus i
                        Exit Function
                     End If
                  End If
               '公司
               ElseIf (txtAD(i).Text = "3") Then
                  '新增或變更
                  If (txtAD(i).Tag <> "3") Then
                     If MsgBox("申請人【" & pa(25 + i) & "-" & Label3(i + 1) & "】的減免身分將設定為【中小企業】，確定有【證明文件】存放於本卷？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                        txtAD(i).SetFocus
                        txtAD_GotFocus i
                        Exit Function
                     End If
                  End If
               '不可減免
               ElseIf (txtAD(i).Text = "N") Then
                  '身分變更
                  If (txtAD(i).Tag <> "N" And txtAD(i).Tag <> "") Then
                     If MsgBox("確定要變更申請人【" & pa(25 + i) & "-" & Label3(i + 1) & "】減免身分為【不可減免】？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
                        txtAD(i).SetFocus
                        txtAD_GotFocus i
                        Exit Function
                     End If
                  End If
               End If
            End If
         Next
      End If
      
      '大陸新申請案若未建國內外關聯時提醒並確認不建
      '2008/3/20 modify by sonia 改控制非台灣案之新申請案
      'If ((Val(Text1(1)) >= 101 And Val(Text1(1)) <= 105) Or (Val(Text1(1)) >= 108 And Val(Text1(1)) <= 112) Or Val(Text1(1)) = 117) And (Text1(13) = "020" Or Text1(13) = "056") Then
      'Modify by Morgan 2011/6/10 香港111(標準專利批准紀錄請求)除外
      'If ((Val(Text1(1)) >= 101 And Val(Text1(1)) <= 105) Or (Val(Text1(1)) >= 108 And Val(Text1(1)) <= 112) Or Val(Text1(1)) = 117) And Text1(13) <> "000" Then
      If ((Val(Text1(1)) >= 101 And Val(Text1(1)) <= 105) Or (Val(Text1(1)) >= 108 And Val(Text1(1)) <= 112 And Val(Text1(1)) <> 111) Or Val(Text1(1)) = 117) And Text1(13) <> "000" Then
         'Modified by Morgan 2017/12/20 香港標準專利批准紀錄另有例外控制
         'If cm(5) = "" And Text1(15).Text = "" Then
         If cm(5) = "" And Text1(15).Text = "" And Not (Text1(13).Text = "013" And Text1(1) = "110" And Text1(2) = "1") Then
            If MsgBox("本案尚未建國內外關聯，是否要建關聯？", vbYesNo + vbDefaultButton1) = vbYes Then
               Text1(15).SetFocus
               Exit Function
            End If
         End If
      End If
      
      'Added by Lydia 2016/09/29 一案兩請案件大陸發明案收文陳述意見在分案時,若相關新型案件下一程序有掛一道放棄專利權,請出現訊息告知user「是否作內部收文PXXXXX(新型案)放棄專利權」,USER選擇「是」系統直接作內部收文,承辦人掛分案的程序人員。
      m_bolAuto429 = False
      If pa(1) = "P" And pa(9) <> "000" And pa(8) = "1" And Text1(1) = "205" Then
         If PUB_IsDualApply(pa, m_DualCaseNo) Then
            strExc(0) = "select nvl(pa57||pa108,'N') pa57,nvl(np22,0) i429 from patent,nextprogress " & _
                        "where pa01='" & m_DualCaseNo(1) & "' and pa02='" & m_DualCaseNo(2) & "' and pa03='" & m_DualCaseNo(3) & "' and pa04='" & m_DualCaseNo(4) & "' " & _
                        "and pa01=np02(+) and pa02=np03(+) and pa03=np04(+) and pa04=np05(+) and np07='429' and np06 is null "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields("pa57") = "N" And RsTemp.Fields("i429") > 0 Then
                  If MsgBox("是否作內部收文" & m_DualCaseNo(1) & "-" & m_DualCaseNo(2) & IIf(m_DualCaseNo(3) & m_DualCaseNo(4) = "000", "", m_DualCaseNo(3) & "-" & m_DualCaseNo(4)) & "(新型案)放棄專利權？", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
                     m_bolAuto429 = True
                  End If
               End If
            End If
         End If
      End If
      'end 2016/09/29
      
      '判斷是否上發文日
      m_bolUpdCP27 = False
      If Text1(0) <> "" And cp(27) = "" Then
         'Add by Morgan 2005/5/26
         If Left(GetStaffDepartment(Text1(0)), 1) = "S" Then
            m_bolUpdCP27 = True
            If MsgBox("因承辦人為智權人員系統將自動上發文日，確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
               Exit Function
            End If
            
         'Modify by Morgan 2005/12/26
         '若案件性質為調卷(904)或後金(909)或列印專利資料(905)或諮詢(912)
         'Modified by Morgan 2020/12/3 912 諮詢／評估分析 改不自動發文--玲玲
         'ElseIf Text1(1).Text = "909" Or Text1(1).Text = "905" Or Text1(1).Text = "912" Then
         ElseIf Text1(1).Text = "909" Or Text1(1).Text = "905" Then
         'end 2020/12/3
            m_bolUpdCP27 = True
                  
         'Add by Morgan 2005/12/26 調卷(904)加判斷程序確認
         ElseIf Text1(1) = 調卷 Then
            m_bolUpdCP27 = True
            If Left(GetStaffDepartment(Text1(0)), 1) = "P" Then
               If MsgBox("因調卷承辦人為專利處人員，請問是否要上發文日？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
                  m_bolUpdCP27 = False
               End If
            End If
         
         'Add by Morgan 2006/12/27 超頁、超項費(917)若新案已發文則詢問是否上發文日(預設是)
         'Modify by Morgan 2007/4/27 加急件費(920) -- 玲玲
         'Modify by Morgan 2009/5/13 +補收款(911) -- 敏惠
         'ElseIf Text1(1) = "917" Then
         'Modify by Morgan 2010/1/6 +938,939
         'Modified by Lydia 2016/07/07 改成模組 PUB_ProcessPchk
'         ElseIf Text1(1) = "917" Or Text1(1) = "920" Or Text1(1) = "911" Or Text1(1) = "938" Or Text1(1) = "939" Then
'            strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' AND CP10 IN (" & CaseMapIn & ") and cp27>0"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' AND CP10 IN (" & CaseMapIn & ") and cp27>0"
'               intI = 1
'               'modify by sonia 2014/5/15 台灣案已收再審案則不必詢問,P-090238
'               'If MsgBox("新申請案已發文，請問本程序是否要上發文日？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
'               '   m_bolUpdCP27 = True
'               'End If
'               If Text1(13) <> "000" Then
'                  If MsgBox("新申請案已發文，請問本程序是否要上發文日？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
'                     m_bolUpdCP27 = True
'                  End If
'               Else
'                  strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' AND CP10='107' and nvl(cp57,0)=0"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 0 Then
'                     If MsgBox("新申請案已發文，請問本程序是否要上發文日？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
'                        m_bolUpdCP27 = True
'                     End If
'                  End If
'               End If
'            End If
         End If
         'Added by Lydia 2016/07/07
         'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
         If PUB_ProcessPchk(cp(1), cp(2), cp(3), cp(4), Text1(1).Text, pa(9), m_bolUpdCP27) Then
         End If
      End If
      
      'Add by Morgan 2006/9/7
      '大陸發明PCT案
      '1.若已收文主張國際優先權則PCT優先權日一定要輸
      '2.沒收文主張國際優先權且沒輸PCT優先權日時要提醒確認
      '3.沒收文主張國際優先權但有輸PCT優先權日時發Mail通知智權人員補收文
      If Text1(13) = "020" And Text1(14) <> "" Then
         Select Case Text1(1)
            'Modify by Morgan 2010/6/15 +新型申請
            Case "101", "102" '發明申請
               bol106 = False: bol106Mail = False
               strPD = "": tmpContent = "" 'Added by Lydia 2016/01/28
               'Modified by Lydia 2016/01/28 以優先權檔判斷是否要主張國際優先權(106)或國內優先權(121)
               'strExc(0) = "SELECT * FROM CASEPROGRESS WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP10='106' AND CP57 IS NULL"
               'intI = 1
               'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               ''有收106
               'If intI = 1 Then
               '   bol106 = True
               ''沒收
               'Else
               '   bol106 = False
               'End If
               'Modified by Lydia 2016/02/18 改成共用模組,另外增加106,121的判斷
               'Modify by Amy 2023/01/05 原: strPriority(1)
               bol106 = PUB_CheckPDMsg(cp(1), cp(2), cp(3), cp(4), Text1(1), pa(9), tmpContent, strPrity1)
                
               '有輸
               If Text1(18) <> "" Then
                  strFirstPriDate = PUB_GetFirstPriDate(cp)
                  If strFirstPriDate <> "" Then
                     If strFirstPriDate <> Text1(18) Then
                        If MsgBox("PCT優先權日與最早優先權日不符，請確認是否無誤！", vbYesNo + vbDefaultButton2) = vbNo Then
                           Exit Function
                        End If
                     End If
                  End If
                  If bol106 = False Then
                     bol106Mail = True
                  End If
               '沒輸
               Else
                  If bol106 = True Then
                     'Modified by Lydia 2016/01/28
                     'MsgBox "本案為PCT案且已收文主張國際優先權，PCT優先權日不可空白！"
                     MsgBox "本案為PCT案且已收文" & IIf(tmpContent = "", "主張國際優先權", tmpContent) & "，PCT優先權日不可空白！"
                     Exit Function
                  Else
                     'Modified by Lydia 2016/01/28
                     'If MsgBox("本案是否有主張國際優先權？", vbYesNo + vbDefaultButton1) = vbYes Then
                     If MsgBox("本案是否有" & IIf(tmpContent = "", "主張國際優先權", tmpContent) & "？", vbYesNo + vbDefaultButton1) = vbYes Then
                        Exit Function
                     End If
                  End If
               End If
            'Modified by Lydia 2016/02/18 + 121
            Case "106", "121" '主張國際優先權,主張國內優先權
               'Added by Lydia 2016/02/18 判斷優先權資料是否存在
               'Modify by Amy 2023/01/05  原:strPriority(1)
               bol106 = PUB_CheckPDMsg(cp(1), cp(2), cp(3), cp(4), Text1(1), pa(9), tmpContent, strPrity1)
               If bol106 = False Then
                  MsgBox tmpContent & "尚未輸入優先權資料！"
                  Exit Function
               End If
               'end 2016/02/18
               strPCTPriDate = PUB_GetPCTPriDate(pa(91))
               'Modify by Amy 2023/01/05  原:strPriority(2)
               strFirstPriDate = PUB_GetFirstPriDate2(strPrity2)
               If strPCTPriDate = "" Then
                  MsgBox "本案未記錄PCT優先權日，請檢查申請案期限是否正確！"
               ElseIf strFirstPriDate <> strPCTPriDate Then
                  If MsgBox("PCT優先權日與最早優先權日不符，請確認是否無誤！", vbYesNo + vbDefaultButton2) = vbNo Then
                     Exit Function
                  End If
               End If
               
            Case "414" '恢復權利
               If Text1(4) = "" Or Text1(5) = "" Then
                  MsgBox "無法計算期限，請先分案發明申請！"
                  Exit Function
               Else
                  MsgBox "本程序期限資料將會回寫到申請程序！"
               End If
         End Select
      End If
      'end 2006/9/7
      
      'Added by Lydia 2017/05/09 主張國內優先權提醒及期限管制
      If Text1(1) = "121" Then
         strExc(0) = "select pa01,pa02,pa03,pa04,pa14,pa57,pa59,c1.cp27,c2.cp71 from pridate,patent,caseprogress c1,caseprogress c2 " & _
                    "where pd01='" & pa(1) & "' and pd02='" & pa(2) & "' and pd03='" & pa(3) & "' and pd04='" & pa(4) & "' " & _
                    "and pd06=pa11(+) and pd05=pa10(+) and pd07=pa09(+) and pd07='" & pa(9) & "' " & _
                    "and pa01=c1.cp01(+) and pa02=c1.cp02(+) and pa03=c1.cp03(+) and pa04=c1.cp04(+) and c1.cp10(+)='601' and c1.cp158(+)>0 " & _
                    "and pa01=c2.cp01(+) and pa02=c2.cp02(+) and pa03=c2.cp03(+) and pa04=c2.cp04(+) and c2.cp10(+)='412' and c2.cp159(+)=0 " & _
                    "order by pd05 asc "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               If "" & RsTemp.Fields("cp27") <> "" Then '若基礎案領證(601)已發文
                  If Text1(13) = "000" Then '台灣案
                     If "" & RsTemp.Fields("pa14") = "" Then  '尚未公告
                        PUB_Get605NP pa(1), "" & RsTemp.Fields("cp27"), 0, strExc, "" & RsTemp.Fields("cp71") '基礎案的預定公告日
                        '若後案所主張之國內優先權基礎案領證已發文，則後案提申的法定期限請帶入基礎案的預定公告日，本所期限則提前2個工作天。
                        Text1(5) = TransDate(strExc(3), 1)
                        Text1(4) = TransDate(CompWorkDay(3, strExc(3), 1), 1)
                        If MsgBox("基礎案已領證，請再確認是否繼續?", vbYesNo + vbDefaultButton2) = vbNo Then
                            Exit Function
                        End If
                     Else
                        MsgBox "基礎案已公告!"
                        Exit Function
                     End If
                  ElseIf Text1(13) = "020" Then '大陸案
                     MsgBox "基礎案已領證，無法主張國內優先權，請告知業務並取消收文主張國內優先權!"
                     Exit Function
                  End If
               End If
               RsTemp.MoveNext
            Loop
         End If
      End If
      'end 2017/05/09
      
      'Add by Morgan 2006/10/26 '取消閉卷恢復年費管制提醒
      m_str605NP22 = ""
      'Modify by Morgan 2006/12/14 加控制案件性質非年費的才要
      If Text1(1) <> "605" And pa(57) = "Y" And Text1(11) = "Y" And Left(cp(12), 1) <> "F" Then
         strExc(0) = "select max(np09||np08||np22) from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07 in (605,606,607)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(RsTemp(0)) Then 'Add by Morgan 2007/1/30 有年費期限的才要
               If Val("" & Left(RsTemp(0), 8)) >= Val(strSrvDate(1)) Then
                  MsgBox "年費將恢復管制，本所期限為 " & Format(Mid(RsTemp(0), 9, 8) - 19110000, "### 年 ## 月 ## 日") & "(法定期限為 " & Format(Left(RsTemp(0), 8) - 19110000, "### 年 ## 月 ## 日)") & ")！"
                  m_str605NP22 = Mid(RsTemp(0), 17)
               'Modify by Morgan 2006/12/6 還要多加一天 --- 敏惠
               'ElseIf CompDate(1, 6, Val("" & Left(rsTemp(0), 8))) >= Val(strSrvDate(1)) Then
               ElseIf Val(CompDate(2, 1, CompDate(1, 6, Val("" & Left(RsTemp(0), 8))))) >= Val(strSrvDate(1)) Then
                  MsgBox "年費期限已逾期但未超過6個月，將恢復管制，本所期限為 " & Format(Mid(RsTemp(0), 9, 8) - 19110000, "### 年 ## 月 ## 日") & "(法定期限為 " & Format(Left(RsTemp(0), 8) - 19110000, "### 年 ## 月 ## 日") & ")！"
                  m_str605NP22 = Mid(RsTemp(0), 17)
               Else
                  'Modified by Morgan 2020/11/30 催年費會排除已通知過的年費,改提醒要新增年費期限--玲玲
                  'MsgBox "年費期限已逾期請更新後再分案！本所期限為 " & Format(Mid(RsTemp(0), 9, 8) - 19110000, "### 年 ## 月 ## 日") & "(法定期限為 " & Format(Left(RsTemp(0), 8) - 19110000, "### 年 ## 月 ## 日") & ")"
                  MsgBox "年費期限已逾期請【新增】下次年費期限後再分案！"
                  'end 2020/11/30
                  Exit Function
               End If
            End If
         End If
      End If
      'end 2006/10/26
   
      'Added by Morgan 2012/7/2 102/1/1 專利新法
      '台灣發明生醫案,若未收文申請寄存,詢問是否產生B類寄存證明
      m_bolCtrl231 = False
      If pa(1) = "P" And Text1(13) = "000" And Text1(1) = "101" And Left(Combo3, 1) = "3" Then
         If PUB_ChkCPExist(pa, "108") = False Then
            If PUB_ChkCPExist(pa, "231") = False Then
               strExc(0) = MsgBox("本案為化學生醫案且未收文申請寄存，是否要產生B類寄存證明管制期限？", vbYesNoCancel + vbQuestion + vbDefaultButton3)
               If strExc(0) = vbCancel Then
                  Exit Function
               ElseIf strExc(0) = vbYes Then
                  m_bolCtrl231 = True
               End If
            End If
         End If
      End If
      '回復優先權主張要檢查申請日不可晚於主張優先權期限
      If pa(1) = "P" And Text1(13) = "000" And Text1(1) = "124" Then
         'Modify by Amy 2023/01/05  strPriority原陣列,改變數
         'If GetAppDateLimit(Text1(2), strPriority, m_strPriType) < DBDATE(pa(10)) Then
         If GetAppDateLimit(Text1(2), strPrity2, strPrity4, m_strPriType) < DBDATE(pa(10)) Then
            MsgBox "本案申請日已逾主張優先權期限不可再收文" & Label3(1) & "!", vbExclamation
            Exit Function
         End If
      End If
      'end 2012/7/2
   
      'Add by Morgan 2010/1/20
      '台灣年費或退費分案檢查是否有預繳年費可退
      m_i605FromYear = 0: m_i605ToYear = 0
      If pa(9) = "000" And (Text1(1) = "908" Or Text1(1) = "605") Then
         If PUB_ChkRefund(pa, m_lngRefund) Then
            '是否可減免
            If txtAD(1) = "N" Or txtAD(2) = "N" Or txtAD(3) = "N" Or txtAD(4) = "N" Or txtAD(5) = "N" Then
               bolDiscount = False
            Else
               bolDiscount = True
            End If
            '退費
            If Text1(1) = "908" Then
               If MsgBox("本案尚有預繳年費可退是否要移作次年？", vbYesNo + vbDefaultButton2) = vbYes Then
                  If Not InputYear(m_i605FromYear, m_i605ToYear) Then
                     Exit Function
                  Else
                     strExc(1) = PUB_GetYearFee(pa(8), m_i605FromYear, m_i605ToYear, bolDiscount)
                     If Not CompRefund(m_lngRefund, Val(strExc(1)), Text1(1), Val(cp(17))) Then
                        Exit Function
                        
                     'Added by Morgan 2011/12/7 退費期限要設定為年費期限 Ex.P-59979
                     Else
                        strExc(1) = CompDate(0, m_i605FromYear - 1, pa(14))
                        strExc(2) = CompDate(2, -1, strExc(1))
                        'Added by Morgan 2014/10/28
                        If Text1(13) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                           strExc(3) = PUB_GetOurDeadline(strExc(2))
                        Else
                        'end 2014/10/28
                           strExc(3) = CompDate(2, -2, strExc(2))
                           strExc(3) = PUB_GetWorkDay1(strExc(3), True)
                        End If 'Added by Morgan 2014/10/28
                        
                        Text1(5) = TransDate(strExc(2), 1)
                        Text1(4) = TransDate(strExc(3), 1)
                        MsgBox "期限已設定為第 " & m_i605FromYear & " 年年費期限！" & vbCrLf & vbCrLf & "( 本所:" & Text1(4) & ", 法定:" & Text1(5) & " )", vbInformation
                     'end 2011/12/7
                     End If
                  End If
               End If
            '年費
            Else
               '收文有輸入起迄年
               'Modify by Morgan 2011/9/30 改抓畫面上輸入的年度
               'If Val(cp(53)) > 0 Then
               '   m_i605FromYear = Val(cp(53))
               '   m_i605ToYear = Val(cp(54))
               If Val(txtFeeYear(1)) > 0 And Val(txtFeeYear(2)) > 0 Then
                  m_i605FromYear = Val(txtFeeYear(1))
                  m_i605ToYear = Val(txtFeeYear(2))
               
               '未輸入
               ElseIf Not InputYear(m_i605FromYear, m_i605ToYear) Then
                  Exit Function
               End If
               strExc(1) = PUB_GetYearFee(pa(8), m_i605FromYear, m_i605ToYear, bolDiscount)
               If Not CompRefund(m_lngRefund, Val(strExc(1)), Text1(1), Val(cp(17))) Then
                  Exit Function
               End If
            End If
         End If
      End If
   End If
    'Added by Lydia 2020/01/21 所有系統別分案，若該收文號存在於下一程序(NP24)，若修改案件性質則彈訊息
    If Text1(1).Tag <> Text1(1).Text Then
        If Pub_CheckNP24Exists(Label3(8).Caption) = True Then
        End If
    End If
    'end 2020/01/21
         
   'Added by Lydia 2020/06/19 法律所案源收文：C類案源的案件性質若 "是否需要法律所配合"設定與來不同時提醒。
   If strSrvDate(1) >= 法律所案源收文啟用日 And pa(1) = "P" Then
        strExc(1) = "" 'Added by Lydia 2021/07/12 清空預設
        'Modified by Lydia 2020/07/23 重新整理: 因為案源收文已設定不可變更案件性質和申請國家,所以只要判斷非案源收文
        If m_LOS02 = "" And m_CP10 <> Text1(1).Text Then
           strExc(1) = PUB_GetLOSkind(pa(1), Text1(1), Text1(13))
           strExc(1) = Replace(strExc(1), "P", "")
           '準備程序在輸入接洽單已決定是否為案源的補收款, 所以不用另外判斷
           If strExc(1) <> "" Then
                 MsgBox "收文不可修改為法務案源的案件性質！", vbCritical
                 Exit Function
           End If
        End If
        'end 2020/07/23
          
        If m_LOS01 = "" And m_LOS07 = "" And FraLOS.Visible = True Then
            If (Left(strExc(1), 1) = "C" And m_LOS15 = "" And txtLOSagree = "Y") Or (Left(strExc(1), 1) = "C" And m_LOS15 <> "" And txtLOSagree <> "Y") _
               Or (strExc(1) = "" And Left(m_LOS02, 1) = "C" And m_LOS15 <> "" And txtLOSagree <> "Y") Then
               If MsgBox(" ""是否需要法律所配合"" 與接洽單不同，是否繼續作業？", vbExclamation + vbYesNo + vbDefaultButton2, "檢核案源單號") = vbNo Then
                   txtLOSagree.SetFocus
                   txtLOSagree_GotFocus
                   Exit Function
               End If
            End If
        End If
   End If
   'end 2020/06/19
   
   'Add by Amy 2023/01/03 從Command2搬過來,接洽單電子收文後玲玲反應,太早關
   If PUB_CheckFormExist("frm090801_Q") = True Then
        Unload frm090801_Q
   End If
   
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   If FormSave = False Then
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Function
   End If
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
    
   'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之年費605、維持費606、延展費607，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。
   If strMsgCloseCancel <> "" Then
      MsgBox "已還原「" & strMsgCloseCancel & "」期限", vbInformation, "取消閉卷"
   End If
         
   'Add by Morgan 2008/1/11
   '當電腦中心人員做新案的分案時需詢問是否發Mail通知工程師
   If Pub_StrUserSt03 = "M51" And Text1(0) <> "" And cp(31) = "Y" Then
      PUB_M51Mail cp(1) & cp(2) & cp(3) & cp(4), Text1(0).Text
   End If
   'end 2008/1/11
   
   'Add By Sindy 2014/7/8
   'Modify By Sindy 2022/11/14 一案兩請需同一申請人控管-新增控管案件性質401變更
   If cp(10) = 讓與 Or cp(10) = 專利權讓與 Or cp(10) = 變更 Then
      ReDim Preserve strAnoPA(1 To TF_PA) As String
      If PUB_IsDualApplyCom(pa, strAnoPA, strCaseNo, , , , "701,708,401", bolIsExistsCP10) = True Then
         If bolIsExistsCP10 = False Then
            'Modify By Sindy 2022/11/14 + IIf(cp(10) = 變更, "變更", "讓與")
            If MsgBox("本案為一案兩請案件，欲辦理" & IIf(cp(10) = 變更, "變更申請人名稱", "讓與") & "須發明及新型兩案同時辦理！" & vbCrLf & _
                      "是否發E-Mail通知智權同仁？", vbYesNo + vbDefaultButton2) = vbYes Then
               '發Mail
               'Modify By Sindy 2022/11/14 + IIf(cp(10) = 變更, "變更申請人名稱", "讓與")
               PUB_SendMail strUserNum, cp(13), "", _
                            pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)) & "為一案兩請案件，欲辦理" & IIf(cp(10) = 變更, "變更申請人名稱", "讓與") & "須發明及新型兩案同時辦理", _
                            "本所案號：" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)) & vbCrLf & _
                            "案件名稱：" & Me.Combo1.Text & vbCrLf & _
                            "申 請 人：" & Label3(2) & vbCrLf & _
                            "一案兩請另一案：" & strCaseNo & vbCrLf & vbCrLf & vbCrLf & _
                            "請再補收文！"
            End If
         End If
      End If
   End If
   '2014/7/8 END
   
   'add by Toni 2008/10/24
   '2008/11/26 MODIFY BY SONIA
   'If Text1(1) = 準備程序 Or Text1(1) = 言詞辯論 Then
   'Modify By Sindy 2023/3/28 控管台灣的才發Mail ex:TF-000870-1-06
   If (Text1(1) = 準備程序 Or Text1(1) = 言詞辯論) And Text1(4) <> "" And Text1(5) <> "" _
      And Text1(13) = "000" Then
   '2008/11/26 END
      '取得更新後的本所期限
      m_strCP06Update = GetCP06(Me.Label3(8).Caption)
      '取得更新後的法定期限
      m_strCP07Update = GetCP07(Me.Label3(8).Caption)
    
      If (Text1(0).Text <> Text1(0).Tag) Or (Text1(4).Text <> Text1(4).Tag) Or (Text1(5).Text <> Text1(5).Tag) Then
         '2008/11/13 ADD BY SONIA
         strSql = "select CP08,CP64 from CASEPROGRESS where CP09='" & Text1(6) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount > 0 Then
            m_CP43CP08 = CheckStr(adoRecordset.Fields(0))
            m_CP43CP64 = CheckStr(adoRecordset.Fields(1))
         End If
         '2008/11/13 END
         
         'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
         'Load frm880005
         'frm880005.txtEmail(0).Text = Pub_GetSpecMan("Q") & ";" & cp(13)
         '''2008/11/13 modify by sonia 再抓時間地點,法院案號
         'frm880005.txtEmail(1).Text = "開庭通知--分案案件：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
         'frm880005.txtEmail(2).Text = "本所案號：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & vbCrLf & _
                                       "案件名稱：" & Me.Combo1.Text & vbCrLf & _
                                       "案件性質：" & Me.Label3(1).Caption & vbCrLf & _
                                       "申請人　：" & Me.Label3(2).Caption & vbCrLf & _
                                       "承辦人　：" & Me.Label3(0).Caption & vbCrLf & _
                                       "智權人員　：" & Me.Label3(10).Caption & vbCrLf & _
                                       "法定期限：" & DBYEAR(m_strCP07Update) - 1911 & " 年 " & DBMONTH(m_strCP07Update) & " 月 " & DBDAY(m_strCP07Update) & " 日 " & vbCrLf & _
                                       "時間地點：" & m_CP43CP64 & vbCrLf & _
                                       "法院案號：" & m_CP43CP08
         'frm880005.Form_Activate: DoEvents
         'frm880005.cmdOK_Click 0: DoEvents
         'Modified by Lydia 2022/08/15 加發承辦人Text1(0)
         'Modify By Sindy 2023/12/8 法律所調整內專行政訴訟開庭通知之系統通知信也請一併轉陳亮之; 商標一併調整
         'Modified by Lydia 2024/10/30 串法律所案源資料，抓出法律所案件有承辦人且收文日最大的進度，抓承辦人及所有出庭律師。
         'm_StrTo = Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & cp(13) & IIf(Text1(0) <> "", ";" & Text1(0), "")
         m_StrTo = PUB_GetLosCL02list(cp(1), cp(2), cp(3), cp(4))
         m_StrTo = IIf(m_StrTo <> "", m_StrTo & ";", "") & Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & cp(13) & IIf(Text1(0) <> "", ";" & Text1(0), "")
         'end 2024/10/30
         
         m_StrSub = "開庭通知--分案案件：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
         m_StrCont = "本所案號：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & vbCrLf & _
                                       "案件名稱：" & Me.Combo1.Text & vbCrLf & _
                                       "案件性質：" & Me.Label3(1).Caption & vbCrLf & _
                                       "申請人　：" & Me.Label3(2).Caption & vbCrLf & _
                                       "承辦人　：" & Me.Label3(0).Caption & vbCrLf & _
                                       "智權人員　：" & Me.Label3(10).Caption & vbCrLf & _
                                       "法定期限：" & DBYEAR(m_strCP07Update) - 1911 & " 年 " & DBMONTH(m_strCP07Update) & " 月 " & DBDAY(m_strCP07Update) & " 日 " & vbCrLf & _
                                       "時間地點：" & m_CP43CP64 & vbCrLf & _
                                       "法院案號：" & m_CP43CP08
         PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
         'end 2022/05/30
      End If
   End If
   'end 2008/10/24
   
    'Add by Morgan 2006/9/7
    If bol416Msg = True Then
      MsgBox "本案已由發明改請為" & Label3(12) & "，下一程序實審期限已自動取消！"
    End If
    If bol416Mail = True Then
      Call PUB_SendMail(strUserNum, cp(13), cp(9), cp(1) & cp(2) & cp(3) & cp(4) & " 案已由發明改請為" & Label3(12) & "，請將已收文之實體審查程序銷案！", " ")
    End If
    If bol106Mail = True Then
      'Modified by Lydia 2016/01/28
      'Call PUB_SendMail(strUserNum, cp(13), cp(9), cp(1) & cp(2) & cp(3) & cp(4) & " 案為大陸PCT" & Label3(12) & "案且有輸[PCT優先權日]但未收文[主張國際優先權]，請補收文該程序！", " ")
      Call PUB_SendMail(strUserNum, cp(13), cp(9), cp(1) & cp(2) & cp(3) & cp(4) & " 案為大陸PCT" & Label3(12) & "案且有輸[PCT優先權日]但未收文[" & IIf(tmpContent = "", "主張國際優先權", tmpContent) & "]，請補收文該程序！", " ")
    End If
    'end 2006/9/7
    
    If IsEmptyText(textPA1) = False And IsEmptyText(textPA2) = False Then
       strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(Label3(9))
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
       If RsTemp.Fields(0) < 1 Then
          MsgBox "原本所案號 " & pa(1) & pa(2) & pa(3) & pa(4) & "已無案件進度資料，請通知收文人員刪號！", vbInformation
       Else
          MsgBox "原本所案號為 " & pa(1) & pa(2) & pa(3) & pa(4) & "，請自行更新原本所案號之下一程序資料 !", vbInformation
       End If
    End If
    
   '若是執行分案(非轉本所案號)
   If Me.textPA1.Text = "" Or Me.textPA2.Text = "" Then
       '承辦人不為北所程序及工程師時提醒目次
       'Modified by Morgan 2013/10/23 考慮程序新人
       'StrSQLa = "Select  *  From EngineerProgress, Staff, Caseprogress Where EP05=ST01(+) And EP02=CP09(+) And EP02='" & Me.Label3(8).Caption & "' And (EP05<>'81002' And EP05<>'73017') And CP01<>'PS' And ST06<>'1' "
       StrSQLa = "Select  *  From EngineerProgress, Staff, Caseprogress Where EP05=ST01(+) And EP02=CP09(+) And EP02='" & Me.Label3(8).Caption & "' And NVL(ST05,' ')<>'75' And CP01<>'PS' And ST06<>'1' "
       'end 2013/10/23
       rsA.CursorLocation = adUseClient
       rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
       If rsA.RecordCount > 0 Then
            MsgBox "此案之承辦人為<" & rsA.Fields("ST02").Value & ">，目次為<" & rsA.Fields("EP01").Value & ">!!!", vbInformation + vbOKOnly
       End If
       If rsA.State <> adStateClosed Then rsA.Close
       Set rsA = Nothing
       
       'Add by Morgan 2005/4/15
       StrSQLa = "Select EP06 From EngineerProgress Where EP02='" & Me.Label3(8).Caption & "'"
       rsA.CursorLocation = adUseClient
       rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
       If rsA.RecordCount > 0 Then
         m_EP06 = "" & rsA.Fields("EP06")
       Else
         m_EP06 = ""
       End If
       If rsA.State <> adStateClosed Then rsA.Close
       Set rsA = Nothing
       '2005/4/15 end
       
       '若本所期限為當日或假日期限, 則發E-Mail給承辦人
       If Me.Text1(0).Text <> "" Then
           '取得更新後的本所期限
           m_strCP06Update = GetCP06(Me.Label3(8).Caption)
           If WorkDayCheck = True Then
                  '取消詢問是否發E-Mail
'Modified by Morgan 2018/8/16 承辦人若為F編號會有問題(內部郵件收件員工編號有虛建編號) Ex:P-92596口頭審理F5704
'                   Load frm880005
'                   strOfficeKind = PUB_GetST06(strUserNum)
'                   '若使用者為北所人員, 則E-Mail後面不加@taie.com.tw
'                   If strOfficeKind = "1" Then
'                       frm880005.txtEmail(0).Text = Me.Text1(0).Text
'                   '若使用者非北所人員, 則E-Mail後面加@taie.com.tw
'                   Else
'                       '分所
'                       frm880005.txtEmail(0).Text = Me.Text1(0).Text & "@taie.com.tw"
'                   End If
'                   frm880005.txtEmail(1).Text = "本所期限到期通知"
'                   frm880005.txtEmail(2).Text = "收文號：" & Me.Label3(8).Caption & vbCrLf & _
                                                               "本所案號：" & Me.Label3(9).Caption & vbCrLf & _
                                                               "案件名稱" & Me.Combo1.Text & vbCrLf & _
                                                               "案件性質：" & Me.Text1(1).Text & " " & Me.Label3(1).Caption & vbCrLf & _
                                                               "收文日：" & DBYEAR(Me.Text1(12).Text) - 1911 & " 年 " & DBMONTH(Me.Text1(12).Text) & " 月 " & DBDAY(Me.Text1(12).Text) & " 日 " & vbCrLf & _
                                                               "承辦人：" & Me.Text1(0).Text & " " & Me.Label3(0).Caption & vbCrLf & _
                                                               "本所期限：" & DBYEAR(m_strCP06Update) - 1911 & " 年 " & DBMONTH(m_strCP06Update) & " 月 " & DBDAY(m_strCP06Update) & " 日 " & vbCrLf & vbCrLf & _
                                                               "※本所期限為當日期限或假日期限!!!"
'                   frm880005.Form_Activate: DoEvents
'                   frm880005.cmdok_Click 0: DoEvents
'                   bolMail = True 'Add by Morgan 2004/2/18
                   strExc(1) = "本所期限到期通知"
                   strExc(2) = "收文號：" & Me.Label3(8).Caption & vbCrLf & _
                                                               "本所案號：" & Me.Label3(9).Caption & vbCrLf & _
                                                               "案件名稱" & Me.Combo1.Text & vbCrLf & _
                                                               "案件性質：" & Me.Text1(1).Text & " " & Me.Label3(1).Caption & vbCrLf & _
                                                               "收文日：" & DBYEAR(Me.Text1(12).Text) - 1911 & " 年 " & DBMONTH(Me.Text1(12).Text) & " 月 " & DBDAY(Me.Text1(12).Text) & " 日 " & vbCrLf & _
                                                               "承辦人：" & Me.Text1(0).Text & " " & Me.Label3(0).Caption & vbCrLf & _
                                                               "本所期限：" & DBYEAR(m_strCP06Update) - 1911 & " 年 " & DBMONTH(m_strCP06Update) & " 月 " & DBDAY(m_strCP06Update) & " 日 " & vbCrLf & vbCrLf & _
                                                               "※本所期限為當日期限或假日期限!!!"
                  PUB_SendMail strUserNum, Me.Text1(0).Text, "", strExc(1), strExc(2)
                  bolMail = bolMailSendOk
'end 2018/8/16
           End If
       End If
       'Add by Morgan 2004/2/18
       '若承辦人是王協理則要發EMail通知
       stCP14 = Me.Text1(0).Text
       'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
       If bolMail = False And stCP14 = "99050" Then
           stCP09 = Me.Label3(8).Caption
           Call PUB_SendMail(strUserNum, "99050", stCP09, "分案通知", "", stPS)
       'end 2024/07/16
       'Add by Morgan 2004/7/6
       '若有通知事項時發Mail通知承辦工程師
       ElseIf stPS <> "" Then
           stCP09 = Me.Label3(8).Caption
           Call PUB_SendMail(strUserNum, stCP14, stCP09, "分案通知", "", stPS)
       End If
       
       'Modify by Morgan 2005/4/29 有承辦人才發
       If Text1(0).Text <> "" Then
         'Add by Morgan 2005/3/29 若承辦人變更時通知智權人員，並註明文件齊備日
         'Modify by Morgan 2008/7/25 B類收文不用通知智權人員
         'If Text1(0).Tag <> "" And Text1(0).Text <> Text1(0).Tag Then
         'Modify by Amy 2015/01/22
         'If Text1(0).Tag <> "" And Text1(0).Text <> Text1(0).Tag And Left(cp(9), 1) <> "B" Then
         If bolCP14Mail = True And Left(cp(9), 1) <> "B" Then
           stPS = "原承辦人：" & Text1(0).Tag & " " & GetStaffName(Text1(0).Tag) & vbCrLf
           'Added by Lydia 2017/05/25 若承辦人之部門為程序(P12),不要帶文件齊備日的訊息(如最後一行),因為程序一定不會輸文件齊備日欄,程序承辦的案件性質也不一定有文件的問題.
           If PUB_GetStaffST15(Text1(0).Tag, "1") <> "P12" Then
              If m_EP06 <> "" Then
                 stPS = stPS & "文件齊備日：" & ChangeWStringToTDateString(m_EP06) & vbCrLf
              Else
                 stPS = stPS & "文件齊備日：未齊備" & vbCrLf
              End If
           End If 'end 2017/05/25
           
           'Added by Morgan 2023/8/14
           PUB_SetEngInform cp(9), Text1(0).Tag, False, strExc(0)
           If strExc(0) <> "" Then stPS = stPS & vbCrLf & strExc(0)
           'end 2023/8/14
           
           'Modified by Lydia 2021/04/19 收件人為智權人員+承辦人，另外CC給原承辦人
           'Call PUB_SendMail(strUserNum, cp(13), cp(9), "承辦人變更通知", "", stPS)
           'Modified by Lydia 2021/11/02 收件人和CC排除舜禹F5588
           'Call PUB_SendMail(strUserNum, cp(13) & ";" & Text1(0).Text, cp(9), "承辦人變更通知", "", stPS, , , , , Text1(0).Tag)
           'Modified by Lydia 2025/03/13 改用模組取得
           'Call PUB_SendMail(strUserNum, cp(13) & IIf(Text1(0) <> "F5588", ";" & Text1(0), ""), cp(9), "承辦人變更通知", "", stPS, , , , , IIf(Text1(0).Tag <> "F5588", Text1(0).Tag, ""))
           Call PUB_SendMail(strUserNum, cp(13) & IIf(InStr(Pub_SetF51Order("F", ""), Text1(0)) = 0, ";" & Text1(0), ""), cp(9), "承辦人變更通知", "", stPS, , , , , IIf(InStr(Pub_SetF51Order("F", ""), Text1(0).Tag) = 0, Text1(0).Tag, ""))
         End If
         
         'Added by Lydia 2021/04/16 副本收件人;P台灣案分案時,倘若是由分所工程師承辦不同所智權同仁之案件時,發MAIL通知智權同仁時,請一併通知該分所工程師。
         strExc(1) = cp(13)
         If Text1(13) = "000" Then
             'Modified by Lydia 2021/04/21 限制分所工程師
             'If PUB_GetST06(Text1(0).Text) <> PUB_GetST06(cp(13)) Then
             strExc(2) = PUB_GetST06(Text1(0).Text)
             If strExc(2) <> "1" And strExc(2) <> PUB_GetST06(cp(13)) Then
             'end 2021/04/21
                 strExc(1) = strExc(1) & ";" & Text1(0)
             End If
         End If
         'end 2021/04/16
         '2011/5/24 add by sonia 杜副總要求台灣新申請案第一次分案時通知智權人員
         If Text1(13) = "000" And InStr(NewCasePtyList, Text1(1)) > 0 And Text1(0).Tag = "" And Text1(0).Text <> "" And Left(cp(9), 1) <> "B" Then
            'Modified by Lydia 2021/04/16 收件人+承辦人
            'Call PUB_SendMail(strUserNum, cp(13), cp(9), "台灣專利新申請案分案通知", "", "")
            Call PUB_SendMail(strUserNum, strExc(1), cp(9), "台灣專利新申請案分案通知", "", "")
         'Add By Sindy 2012/4/3 台灣案除領證,年費外,分案或修改承辦人時,且承辦人為P10或P11部門者,E-Mail給收文智權人員
         ElseIf Text1(13) = "000" And Text1(1) <> "601" And Text1(1) <> "605" And Text1(0).Tag = "" And Text1(0).Text <> "" And Left(cp(9), 1) <> "B" Then
           stPS = ""
           'Added by Morgan 2023/8/14
           PUB_SetEngInform cp(9), Text1(0).Tag, False, strExc(0)
           If strExc(0) <> "" Then stPS = stPS & vbCrLf & strExc(0)
           'end 2023/8/14
           
            'Modify By Sindy 2012/4/9 承辦人為P10或P11部門者(刪除101/4/5)
            'If GetStaffDepartment(Text1(0)) = "P10" Or GetStaffDepartment(Text1(0)) = "P11" Then
               'Modified by Lydia 2021/04/16 收件人+承辦人
               'Call PUB_SendMail(strUserNum, cp(13), cp(9), "台灣專利案分案通知", "", "")
               Call PUB_SendMail(strUserNum, strExc(1), cp(9), "台灣專利案分案通知", "", stPS)
            'End If
         End If
         '2011/5/24 end
         
         'Add By Sindy 2022/12/27 承辦人為P12.專利處程序時,要發分案通知 (And PUB_GetST03(Text1(0).Text) = "P12")
         If strSrvDate(1) >= 接洽單電子收文啟用日 Then
            If Text1(0).Text <> Text1(0).Tag And Text1(0).Tag = "" Then
'               '接洽單第一筆案件性質,才發
'               If Pub_ConIsFirstCRC(txtF0301, cp(9)) = True Then
'                  Call PUB_SendMail(strUserNum, Text1(0).Text, cp(9), "分案通知", "")
'               End If
               strSql = "Select CRC01,cp10,cp31,cp14" & _
                        " From consultrecCMP,caseprogress" & _
                        " Where CRC01='" & txtF0301 & "' and CRC08=cp09(+) and cp157 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If RsTemp.RecordCount = 0 Then '全部分案完成
                  strSql = "Select cp14,cp09,CRC08" & _
                           " From consultrecCMP,caseprogress,staff" & _
                           " Where CRC01='" & txtF0301 & "' and CRC08=cp09(+) and cp14=st01(+) and st03='P12'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     RsTemp.MoveFirst
                     Do While Not RsTemp.EOF '幾筆程序人員
                        strCRC08 = RsTemp.Fields("CRC08")
                        strSql = "Select CRC02,cp09,cp10,cp01,cp02,cp03,cp04,cp31,cp14" & _
                                 " From consultrecCMP,caseprogress" & _
                                 " Where CRC01='" & txtF0301 & "' and CRC08=cp09(+) and CRC08='" & strCRC08 & "'" & _
                                 " order by CRC02 asc"
                        intI = 1
                        Set rsA = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           'P-129669領證+變更(1120003831)
                           If Not ((rsA.Fields("cp10") = "601" Or rsA.Fields("cp10") = "605") And _
                                    "" & rsA.Fields("cp31") <> "Y") Then
                              strExc(10) = rsA.Fields("cp14")
                              Call PUB_SendMail(strUserNum, strExc(10), rsA.Fields("cp09"), "分案通知（本所案號：" & rsA.Fields("cp01") & "-" & rsA.Fields("cp02") & IIf(rsA.Fields("cp03") & rsA.Fields("cp04") = "000", "", "-" & rsA.Fields("cp03") & "-" & rsA.Fields("cp04")) & "）", "", IIf(Check11.Value = 1, "注意：急件！", ""))
                           End If
                        End If
                        RsTemp.MoveNext
                     Loop
                  End If
                  
                  'Add By Sindy 2024/11/14 P的分案通知，
                  '   增加新申請案的案件性質若承辦人為'F'字頭時，發信給系統特殊設定人員「H」，目前是陳品薇98012
                  strSql = "Select cp14,cp09,CRC08" & _
                           " From consultrecCMP,caseprogress" & _
                           " Where CRC01='" & txtF0301 & "' and CRC08=cp09(+) and substr(cp14,1,1)='F'" & _
                           " and InStr('" & NewCasePtyList & "', '" & Text1(1) & "') > 0"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     RsTemp.MoveFirst
                     Do While Not RsTemp.EOF
                        strCRC08 = RsTemp.Fields("CRC08")
                        strSql = "Select CRC02,cp09,cp10,cp01,cp02,cp03,cp04,cp31,cp14" & _
                                 " From consultrecCMP,caseprogress" & _
                                 " Where CRC01='" & txtF0301 & "' and CRC08=cp09(+) and CRC08='" & strCRC08 & "'" & _
                                 " order by CRC02 asc"
                        intI = 1
                        Set rsA = ClsLawReadRstMsg(intI, strSql)
                        If intI = 1 Then
                           strExc(10) = Pub_GetSpecMan("H")
                           Call PUB_SendMail(strUserNum, strExc(10), rsA.Fields("cp09"), "分案通知（本所案號：" & rsA.Fields("cp01") & "-" & rsA.Fields("cp02") & IIf(rsA.Fields("cp03") & rsA.Fields("cp04") = "000", "", "-" & rsA.Fields("cp03") & "-" & rsA.Fields("cp04")) & "）", "", IIf(Check11.Value = 1, "注意：急件！", ""))
                        End If
                        RsTemp.MoveNext
                     Loop
                  End If
                  '2024/11/14 END
               End If
            End If
         End If
         '2022/12/27 END
         
         'Add by Morgan 2005/4/21 發明改請判斷若未公開則發Mail通知承辦人發函智慧局說明暫不公開,不分國家都要 --郭
         If m_bol30xMail = True Then
            Call PUB_SendMail(strUserNum, Text1(0), cp(9), m_bol30xMailDesc, " ")
         End If
      
         'Add by Morgan 2007/8/31 檢查是否有關係企業申請技術報告
         If Text1(1) = "421" Or Text1(1) = "807" Then
            If Text1(1) = "421" Then
               strExc(1) = pa(11)
            Else
               strExc(1) = Text1(17)
            End If
            If strExc(1) <> "" Then
               strExc(0) = "select pa01,pa02,pa26,cp10,cp24" & _
                  " from patent,caseprogress where pa01 in ('P','FCP')" & _
                  " and pa11='" & strExc(1) & "' and substr(pa26,1,6)='" & Left(pa(26), 6) & "'" & _
                  " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
                  " and cp10='" & Text1(1) & "' and cp57 is null and cp09<>'" & cp(9) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
               With RsTemp
                  Do While Not .EOF
                     If .Fields("pa01") = pa(1) And .Fields("pa02") = pa(2) Then
                        bol421Check1 = True
                     Else
                        If IsNull(.Fields("cp24")) Then
                           bol421Check2 = True
                           If InStr(str421Check2, .Fields("pa01") & "-" & .Fields("pa02")) = 0 Then
                              str421Check2 = str421Check2 & IIf(str421Check2 <> "", ",", "") & .Fields("pa01") & "-" & .Fields("pa02")
                           End If
                        Else
                           bol421Check3 = True
                           If InStr(str421Check3, .Fields("pa01") & "-" & .Fields("pa02")) = 0 Then
                              str421Check3 = str421Check3 & IIf(str421Check3 <> "", ",", "") & .Fields("pa01") & "-" & .Fields("pa02")
                           End If
                        End If
                     End If
                     
                     .MoveNext
                  Loop
               End With
                  strExc(1) = ""
                  If bol421Check1 = True Then
                     strExc(1) = strExc(1) & "本案已申請過【" & Label3(1) & "】，特此告知！"
                  End If
                  
                  If bol421Check3 = True Then
                     strExc(1) = strExc(1) & "本申請號已有另案( " & str421Check2 & " )提出【" & Label3(1) & "】申請並已完成，請確認是否提出影印申請即可！"
                  ElseIf bol421Check2 = True Then
                     strExc(1) = strExc(1) & "本申請號已有另案( " & str421Check2 & " )提出【" & Label3(1) & "】申請，請確認是否要再次提出申請！"
                  End If
                  
                  If strExc(1) <> "" Then
                     strExc(2) = cp(13)
                     If Text1(0) <> "" Then
                        strExc(2) = strExc(2) & ";" & Text1(0)
                     End If
                     Call PUB_SendMail(strUserNum, strExc(2), cp(9), strExc(1))
                  End If
               End If
            End If
         End If
      End If
      'end 2007/8/31
      
      '2010/2/4 ADD BY SONIA
      '改請獨立306且案號未合併時,發MAIL通知秀玲
      If Text1(1) = "306" Then
         strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' AND CP10 IN (" & CaseMapIn & ") and cp27>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then  '案號未合併
            'add by sonia 2022/12/29
            If cp(3) <> "0" Then
               m_StrTo = "83002"
               m_StrSub = "改請獨立分案通知：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
               m_StrCont = "請先至自動編號檔抓一個P案新案流水號，再執行 案件改號作業，將此案號改為新案號。" & vbCrLf & _
                           "執行完請通知專業部程序、工程師及智權人員(若收據已印，確認是否重開收據)！!"
               PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
            End If
            'end 2022/12/29
        Else
            'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
            'frm880005.txtEmail(0).Text = "83002"
            'frm880005.txtEmail(1).Text = "改請獨立分案通知：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
            '''Modified by Morgan 2012/12/19 +/衍生設計
            'frm880005.txtEmail(2).Text = "請詢問專業部原聯合/衍生設計案案號，先檢查原聯合/衍生設計案案號與改請獨立案號的基本資料，除案件名稱欄外將資料改成一致後，" & vbCrLf & _
                                         "再執行原聯合/衍生設計案案號的 案件改號作業(原聯合/衍生設計案案號要刪除)！" & vbCrLf
            '''2010/11/3 add by sonia
            'frm880005.txtEmail(2).Text = frm880005.txtEmail(2).Text & vbCrLf & "執行完改案號後要重新分案, 點選下一程序期限以消期限並串相關總收文號 !"
            ''2010/11/3 end
            'frm880005.Form_Activate: DoEvents
            'frm880005.cmdOK_Click 0: DoEvents
            m_StrTo = "83002"
            m_StrSub = "改請獨立分案通知：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
            m_StrCont = "請詢問專業部原聯合/衍生設計案案號，先檢查原聯合/衍生設計案案號與改請獨立案號的基本資料，除案件名稱欄外將資料改成一致後，" & vbCrLf & _
                                         "再執行原聯合/衍生設計案案號的 案件改號作業(原聯合/衍生設計案案號要刪除)！" & vbCrLf & _
                                         "執行完改案號後要重新分案, 點選下一程序期限以消期限並串相關總收文號 !" & vbCrLf & _
                                         "再通知專業部程序、工程師及智權人員(若收據已印，確認是否重開收據)"
            PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
            'end 2022/05/30
         End If
      End If
      'ADD BY SONIA 2016/3/29 +改請衍生設計
      If Text1(1) = "305" Or Text1(1) = "308" And cp(3) = "0" Then
         'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
         'frm880005.txtEmail(0).Text = "83002"
         'frm880005.txtEmail(1).Text = Label3(1) & "分案通知：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & Label3(1) & ", 請詢問專業部獨立案案號 !"
         'frm880005.txtEmail(2).Text = "請依主旨執行 案件改號作業 ，執行完請通知專業部！" & vbCrLf & vbCrLf
         'frm880005.Form_Activate: DoEvents
         'frm880005.cmdOK_Click 0: DoEvents
         m_StrTo = "83002"
         m_StrSub = Label3(1) & "分案通知：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & Label3(1) & ", 請詢問專業部欲轉入之獨立案案號 !"
         m_StrCont = "請依主旨執行 案件改號作業，將此案號改為該獨立案號之子案案號。" & vbCrLf & _
                           "執行完請通知專業部程序、工程師及智權人員(若收據已印，確認是否重開收據)！" & vbCrLf & vbCrLf
         PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
         'end 2022/05/30
      End If
      '2016/3/29 END
      
      '2012/7/19 add by sonia
      '若由紙本送件改為電子送件時,mail給智權人員
      'Modified by Morgan 2013/9/12 +中間程序也可設定電子送,改判斷新申請案
      'If Me.txtCP118.Text <> Me.txtCP118.Tag Then
      'Modified by Morgan 2015/6/17 +改請,分割
      'Modified by Morgan 2024/1/30 +台灣(因增加大陸案也可設定電子送件)
      If Me.txtCP118.Text <> Me.txtCP118.Tag And (InStr(NewCasePtyList, Text1(1)) > 0 Or Left(Text1(1), 1) = "3") And Text1(13) = "000" Then
         'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
         'frm880005.txtEmail(0).Text = cp(13)
         m_StrTo = cp(13)
         If Me.txtCP118.Text = "Y" Then
            'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
            'frm880005.txtEmail(1).Text = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & " 可改為電子送件之通知，可減免規費600元！請務必要做後續處理！"
            '''Modified by Morgan 2013/5/15
            '''frm880005.txtEmail(2).Text = "　申請人１：" & Label3(2) & vbCrLf & _
                                         "　案件名稱：" & pa(5) & vbCrLf & vbCrLf & _
                                         "　本案收文費用如下：" & vbCrLf & _
                                         "　　　收文費用：" & Format(cp(16), DDollar) & "     規費：" & Format(cp(17), DDollar) & "     點數：" & Format(cp(18), "###,###.#") & vbCrLf & vbCrLf & _
                                         "　後續處理方式：" & vbCrLf & _
                                         "　●若不同意改為電子送件，請立即通知專利處程序人員取消電子送件之設定，否則將以電子送件方式處理。" & vbCrLf & _
                                         "　●若同意改為電子送件：" & vbCrLf & _
                                         "　　1. 若已開立收據則請將原收據註明費用及規費如何修改後，退回財務處重開；" & vbCrLf & _
                                         "　　2. 若尚未開立收據亦請通知財務處費用及規費如何修改。" & vbCrLf
            'frm880005.txtEmail(2).Text = "　申請人１：" & Label3(2) & vbCrLf & _
                                         "　案件名稱：" & pa(5) & vbCrLf & vbCrLf & _
                                         "　本案收文費用如下：" & vbCrLf & _
                                         "　　　收文費用：" & Format(cp(16), DDollar) & "     規費：" & Format(cp(17), DDollar) & "     點數：" & Format(cp(18), "###,###.#") & vbCrLf & vbCrLf & _
                                         "　後續處理方式：" & vbCrLf & _
                                         "　●若不同意改為電子送件，請立即通知專利處程序人員取消電子送件之設定，否則將以電子送件方式處理。" & vbCrLf & _
                                         "　●若同意改為電子送件（若未回覆，則視為同意），收據將在接獲前述E-MAIL後三個工作天內開立，費用及" & vbCrLf & _
                                         "　　規費將會自動扣除600元(即收文總費用會減少600元)。" & vbCrLf
            m_StrSub = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & " 可改為電子送件之通知，可減免規費600元！請務必要做後續處理！"
            m_StrCont = "　申請人１：" & Label3(2) & vbCrLf & _
                                         "　案件名稱：" & pa(5) & vbCrLf & vbCrLf & _
                                         "　本案收文費用如下：" & vbCrLf & _
                                         "　　　收文費用：" & Format(cp(16), DDollar) & "     規費：" & Format(cp(17), DDollar) & "     點數：" & Format(cp(18), "###,###.#") & vbCrLf & vbCrLf & _
                                         "　後續處理方式：" & vbCrLf & _
                                         "　●若不同意改為電子送件，請立即通知專利處程序人員取消電子送件之設定，否則將以電子送件方式處理。" & vbCrLf & _
                                         "　●若同意改為電子送件（若未回覆，則視為同意），收據將在接獲前述E-MAIL後三個工作天內開立，費用及" & vbCrLf & _
                                         "　　規費將會自動扣除600元(即收文總費用會減少600元)。" & vbCrLf
            'end 2022/05/30
         Else
            'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
            'frm880005.txtEmail(1).Text = "通知 " & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & " 已取消電子送件之設定！"
            'frm880005.txtEmail(2).Text = "　申請人１：" & Label3(2) & vbCrLf & _
                                         "　案件名稱：" & pa(5) & vbCrLf & vbCrLf & _
                                         "　本案前已通知可改電子送件，但您不同意，現已取消電子送件之設定，謝謝。" & vbCrLf
            m_StrSub = "通知 " & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & " 已取消電子送件之設定！"
            m_StrCont = "　申請人１：" & Label3(2) & vbCrLf & _
                                         "　案件名稱：" & pa(5) & vbCrLf & vbCrLf & _
                                         "　本案前已通知可改電子送件，但您不同意，現已取消電子送件之設定，謝謝。" & vbCrLf
            'end 2022/05/30
         End If
         'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
         'frm880005.Form_Activate: DoEvents
         'frm880005.cmdOK_Click 0: DoEvents
         PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
         'end 2022/05/30
         
         '2013/9/13 add by sonia 中所工程師案件也要發給工程師,但因會有請假職代代收問題故另發e-mail給工程師
         If PUB_GetST06(cp(14)) = "2" Then
            'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
            'frm880005.txtEmail(0).Text = cp(14)
            m_StrTo = cp(14)
            If Me.txtCP118.Text = "Y" Then
               'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
               'frm880005.txtEmail(1).Text = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & " 可改為電子送件之通知，可減免規費600元！已通知智權人員做後續處理！"
               'frm880005.txtEmail(2).Text = "　申請人１：" & Label3(2) & vbCrLf & _
                                            "　案件名稱：" & pa(5) & vbCrLf & vbCrLf
               m_StrSub = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & " 可改為電子送件之通知，可減免規費600元！已通知智權人員做後續處理！"
               m_StrCont = "　申請人１：" & Label3(2) & vbCrLf & _
                                            "　案件名稱：" & pa(5) & vbCrLf & vbCrLf
               'end 2022/05/30
            Else
               'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
               'frm880005.txtEmail(1).Text = "通知 " & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & " 已取消電子送件之設定！"
               'frm880005.txtEmail(2).Text = "　申請人１：" & Label3(2) & vbCrLf & _
                                            "　案件名稱：" & pa(5) & vbCrLf & vbCrLf & _
                                            "　本案前已通知可改電子送件，但智權人員不同意，現已取消電子送件之設定。" & vbCrLf
               m_StrSub = "通知 " & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & " 已取消電子送件之設定！"
               m_StrCont = "　申請人１：" & Label3(2) & vbCrLf & _
                                            "　案件名稱：" & pa(5) & vbCrLf & vbCrLf & _
                                            "　本案前已通知可改電子送件，但智權人員不同意，現已取消電子送件之設定。" & vbCrLf
               'end 2022/05/30
            End If
            'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
            'frm880005.Form_Activate: DoEvents
            'frm880005.cmdOK_Click 0: DoEvents
            PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
            'end 2022/05/30
         End If
         '2013/9/13 end
      End If
      '2012/7/19 end
   
   Else    '若是執行轉本所案號
      'Add by Morgan 2006/9/20
      '新案改案號時若有國外案且已分案則發Mail通知國外案工程師
      If strMail2FEngCP09 <> "" Then Mail2Eng
      
      '2010/2/4 ADD BY SONIA
      '改請聯合305時,發MAIL通知秀玲
      'Modified by Morgan 2012/12/19 +改請衍生設計
      If Text1(1) = "305" Or Text1(1) = "308" Then
         'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
         'frm880005.txtEmail(0).Text = "83002"
         'frm880005.txtEmail(1).Text = Label3(1) & "分案通知：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & Label3(1) & "至" & Me.textPA1.Text & "-" & Me.textPA2.Text & "-" & Me.textPA3.Text & "-" & Me.textPA4.Text
         'frm880005.txtEmail(2).Text = "請依主旨執行 案件改號作業(原案號要刪除) ，執行完請通知專業部！" & vbCrLf & vbCrLf
         m_StrTo = "83002"
         m_StrSub = Label3(1) & "分案通知：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & Label3(1) & "至" & Me.textPA1.Text & "-" & Me.textPA2.Text & "-" & Me.textPA3.Text & "-" & Me.textPA4.Text
         m_StrCont = "請依主旨執行 案件改號作業(原案號要刪除) ，執行完請通知專業部！" & vbCrLf & vbCrLf
         'end 2022/05/30
         If cp(31) = "Y" Then
            'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
            'frm880005.txtEmail(2).Text = frm880005.txtEmail(2).Text & "原本所案號" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & "已無進度資料，" & vbCrLf & _
                                         "請詢問專業部原獨立案案號，先檢查獨立案案號與改請後案號的基本資料，除案件名稱欄外將資料改成一致後，" & vbCrLf & _
                                         "再執行獨立案案號的 案件改號作業(原獨立案號要刪除)！" & vbCrLf & vbCrLf & "執行完請檢查下一程序期限是否消除, 若未消除請通知程序處理 !"
            m_StrCont = m_StrCont & "原本所案號" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & "已無進度資料，" & vbCrLf & _
                                         "請詢問專業部原獨立案案號，先檢查獨立案案號與改請後案號的基本資料，除案件名稱欄外將資料改成一致後，" & vbCrLf & _
                                         "再執行獨立案案號的 案件改號作業(原獨立案號要刪除)！" & vbCrLf & vbCrLf & "執行完請檢查下一程序期限是否消除, 若未消除請通知程序處理 !" & vbCrLf & _
                                         "執行完請通知專業部程序、工程師及智權人員(若收據已印，確認是否重開收據)！" & vbCrLf
         End If
         'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
         'frm880005.Form_Activate: DoEvents
         'frm880005.cmdOK_Click 0: DoEvents
         PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
         'end 2022/05/30
      End If
      '2010/2/4 END
   End If
   
   'Added by Lydia 2023/03/25 從PUB_UpdCfpDate1回傳訊息
   If pSaveMsg <> "" Then
       MsgBox pSaveMsg, vbInformation + vbOKOnly, "主張國內優先權提申期限以母案之審查意見通知之期限為準"
   End If
   'end 2023/03/25
   
   'Add By Sindy 2023/1/10
   Process = True
   If intRunType = 1 Then Exit Function '補件完成呼叫
   '2023/1/10 END
      
   If IntNow <> IntTot Then
       'Modify By Cheng 2001/12/25
'            GetData IntNow
       Text1(0).SetFocus
       GetData IntNow
       Set oTopForm = Screen.ActiveForm 'Add by Morgan 2010/8/5
       'Add by Morgan 2004/7/21
       m_bolActive = False
       Form_Activate
       SSTab1.Tab = 0   '2005/7/7 ADD BY SONIA
       oTopForm.SetFocus 'Add by Morgan 2010/8/5
    Else
       For i = 0 To IntTot
          frm040101.strSave = frm040101.strSave & "," & StrTot2(i)
       Next
       Unload frm040101_1
       ' 設定滑鼠游標為等待狀態
       Screen.MousePointer = vbHourglass
       ' 90.07.06 modify by louis
       frm040101.RefreshData
       ' 設定滑鼠游標為預設
       Screen.MousePointer = vbDefault
       frm040101.Show
    End If
End Function

'Added by Morgan 2013/3/28
Private Sub SettxtCP147()
   '若案件屬性更改時重新預設複雜或特殊案件
   If Combo3.Tag <> "" Or (Combo3.Tag = "" And pa(158) <> Left(Combo3, 1)) Then
      If Combo3.Tag <> Left(Combo3, 1) Then
         txtCP147 = GetCP147Default()
      End If
   End If
   Combo3.Tag = Left(Combo3, 1)
End Sub

'Add by Amy 2022/10/07 補件完成
Private Sub CmdAddInfo_Click()
    Dim oTopForm As Form
    Dim i As Integer
        
On Error GoTo ErrHand
    
    If txtNote = MsgText(601) Then
        MsgBox "呈報內容不可為空！", vbExclamation
        Exit Sub
    End If
    
    'If Process(1) = False Then Exit Sub 'Add By Sindy 2023/1/10 玲玲她們要修改分案作業上的欄位值改為正確才能呈報
    
    Screen.MousePointer = vbHourglass
    
    m_F0308 = stF0307_Now '登入時的上一處理人員
    m_F0309 = Flow_補件完成
    strUpdDate = strSrvDate(1)
    strUpdTime = Right("000000" & ServerTime, 6)
   
    cnnConnection.BeginTrans
    '簽核檔
    strSql = "update FLOW002 set " & _
               "F0205='" & strUpdDate & "'" & _
               ",F0206='" & strUpdTime & "'" & _
               ",F0207='5',F0204='" & strUserNum & "'" & _
               " where F0201='" & txtF0301 & "' and F0202='A7'  and F0207 is null "
    cnnConnection.Execute strSql
    
    '退回程序時再新增2筆待簽核的記錄
    Call SetConultRecPrePerson_Flow002(Me.Name, txtF0301, "A6") '主管
    Call SetConultRecPrePerson_Flow002(Me.Name, txtF0301, "A7") '程序
        
    '表單主檔
    strSql = "update FLOW003 set " & _
            "F0307='A7'" & _
            ",F0308='" & m_F0308 & "'" & _
            ",F0309='" & m_F0309 & "'" & _
            " where F0301='" & txtF0301 & "' "
    cnnConnection.Execute strSql
    
    strSql = GetInsertFLOW004Sql(txtF0301, strUserNum, strUpdDate, strUpdTime, m_F0309, ChgSQL(Trim(txtNote.Text)), "A7", "A6")
    cnnConnection.Execute strSql
    
    cnnConnection.CommitTrans
    
    Screen.MousePointer = vbDefault
    
    If IntNow <> IntTot Then
        Text1(0).SetFocus
        GetData IntNow
        Set oTopForm = Screen.ActiveForm '避免有呼叫其他表單,Active不在此表單上
        m_bolActive = False
        Form_Activate
        SSTab1.Tab = 0
        oTopForm.SetFocus
    Else
        For i = 0 To IntTot
            frm040101.strSave = frm040101.strSave & "," & StrTot2(i)
        Next
        Unload frm040101_1
        ' 設定滑鼠游標為等待狀態
        Screen.MousePointer = vbHourglass
        frm040101.RefreshData
        ' 設定滑鼠游標為預設
        Screen.MousePointer = vbDefault
        frm040101.Show
    End If
    
    Exit Sub
    
    
ErrHand:
    Screen.MousePointer = vbDefault
    cnnConnection.RollbackTrans
    MsgBox "補件失敗！" & vbCrLf & Err.Description
End Sub

'Add By Sindy 2022/10/31 該文號的卷宗區
Private Sub cmdCPP_Click()
    'Mark by Amy 2022/12/23 卷宗區與接洽單可同時開
'   'Add by Amy 2022/11/16
'   If PUB_CheckFormExist("frm090801_Q") = True Then
'        Unload frm090801_Q
'   End If
   Screen.MousePointer = vbHourglass
   frm100101_L.m_CP09 = Label3(8)
   frm100101_L.m_strKey = Label3(8) 'lblCaseNo.Caption
   frm100101_L.SetParent Me
   If frm100101_L.QueryData = True Then
      frm100101_L.Show
      Me.Hide
   Else
      Unload frm100101_L
   End If
   Screen.MousePointer = vbDefault
End Sub

'Add by Amy 2022/10/17 檢視接洽單
Private Sub cmdFile_Click()
    frm090801_Q.SetParent Me
    frm090801_Q.m_blnCallPrint = True
    frm090801_Q.Text5 = txtF0301
    Call frm090801_Q.cmdok_Click(4)
    frm090801_Q.Show 'Add by Amy 2022/11/15
End Sub

'Added by Morgan 2013/3/28
Private Sub Combo3_Click()
   SettxtCP147
End Sub

'Added by Morgan 2013/3/28
Private Sub Combo3_LostFocus()
   SettxtCP147
End Sub

'Add By Sindy 2010/10/29
Private Sub Combo3_Validate(Cancel As Boolean)
   If Combo3 <> "" Then
      Combo3 = Left(Combo3, 1) + "." + PUB_GetCaseAttributeName(Left(Combo3, 1))
      If Combo3 = Left(Combo3, 1) + "." Then
         Combo3 = Left(Combo3, 1)
         Cancel = True
         Combo3.SetFocus
      End If
   End If
   SettxtCP147 'Added by Morgan 2013/3/28
End Sub
'2010/10/29 End

Private Sub Command2_Click(Index As Integer)
Dim i As Integer
   'Add by Amy 2022/11/15 按 確定/回前畫面/下一筆,若接洽單已開需關閉
   'Modify by Amy 2023/01/03 拿掉 Index = 2(確定鈕)改至檢查完存檔前關
   If Index = 3 Or Index = 5 Then
        If PUB_CheckFormExist("frm090801_Q") = True Then
             Unload frm090801_Q
        End If
   End If
   'Add by Amy 2023/01/05 一案兩請和擬制喪失新穎性關聯,開過再按放最上層
   If Index = 0 Or Index = 6 Then
        If PUB_CheckFormExist("frm040109_1") = True Then
            frm040109_1.ZOrder 0
        End If
   End If
   
   Select Case Index
      'Modified by Morgan 2015/9/14
      'Case 0 '相關案號
      '   Where1103ComeFrom Me, pa(1), pa(2), pa(3), pa(4)
      Case 0 '擬制喪失新穎性關聯
         Load frm040109_1  'Add by Morgan 2004/10/28
         Set frm040109_1.frmParent = Me
         frm040109_1.m_CM10 = "6"
         frm040109_1.txtCode(0) = pa(1)
         frm040109_1.txtCode(1) = pa(2)
         frm040109_1.txtCode(2) = pa(3)
         frm040109_1.txtCode(3) = pa(4)
         frm040109_1.txtCode(8) = "1"
         If frm040109_1.ChkExist = True Then
            frm040109_1.Move frm040109_1.Left, frm040109_1.Top - 550
            If MsgBox("該案已建立擬制喪失新穎性關聯是否修改？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
               Unload frm040109_1
            Else
               Me.Hide
               frm040109_1.txtCode(8) = "2"
               frm040109_1.Show
               frm040109_1.cmdOK(0).Value = True
            End If
         'Add By Sindy 2022/11/23
         ElseIf txtF0301 <> MsgText(601) And cp(157) = "" Then
            strExc(10) = Pub_GetCRLCaseMap(txtF0301, "6", "P", pa(1), pa(2), pa(3), pa(4))
            If strExc(10) <> "" Then
               frm040109_1.txtCode(4) = SystemNumber(strExc(10), 1)
               frm040109_1.txtCode(5) = SystemNumber(strExc(10), 2)
               frm040109_1.txtCode(6) = SystemNumber(strExc(10), 3)
               frm040109_1.txtCode(7) = SystemNumber(strExc(10), 4)
            End If
            '2022/11/23 END
         End If
      'end 2015/9/14
      
      Case 1 '案件進度
         frm040101_2.iGo = 4
         frm040101_2.Show
         Me.Hide
      Case 2 '確定
         Process

      Case 3 '回前畫面
         For i = IntNow - 1 To IntTot
            frm040101.Tag = frm040101.Tag & "," & StrTot2(i)
         Next
         For i = 0 To IntNow - 2
            frm040101.strSave = frm040101.strSave & "," & StrTot2(i)
         Next
         frm040101.Show
         Unload frm040101_1
            
      Case 4 '優先權資料
         'Add by Amy 2023/01/05 此支進優先權表單改不是強制表單,故進入時畫面鎖住
         mdiMain.Enabled = False
         Me.Enabled = False
         'Modify by Amy 2014/06/10 + strPriority(5)
         'Modify by Amy 2023/01/05 strPriority原陣列,改變數,並加表單名
         'ModifyPriority strPriority(1), strPriority(2), strPriority(3), pa(8), , pa(1) & pa(2) & pa(3) & pa(4), pa(9), , strPriority(4), strPriority(5)
         ModifyPriority strPrity1, strPrity2, strPrity3, pa(8), , pa(1) & pa(2) & pa(3) & pa(4), pa(9), , strPrity4, strPrity5, , Me
         
      Case 5 '下一筆
         If IntNow <> IntTot Then
            'Modify By Cheng 2001/12/25
'            GetData IntNow
            Me.Text1(0).SetFocus
            GetData IntNow
            m_bolActive = False
            Form_Activate
         Else
            Unload frm040101_1
            frm040101.Show
         End If
         
      'Add by Morgan 2004/6/14   一案兩請資料
      Case 6
         m_bolCP98Check = True
         Load frm040109_1  'Add by Morgan 2004/10/28
         Set frm040109_1.frmParent = Me
         frm040109_1.txtCode(0) = pa(1)
         frm040109_1.txtCode(1) = pa(2)
         frm040109_1.txtCode(2) = pa(3)
         frm040109_1.txtCode(3) = pa(4)
         frm040109_1.txtCode(8) = "1"
         If frm040109_1.ChkExist = True Then
            frm040109_1.Move frm040109_1.Left, frm040109_1.Top - 550
            If MsgBox("該案已建立一案兩請關聯是否修改？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
               Unload frm040109_1
            Else
               Me.Hide
               frm040109_1.txtCode(8) = "2"
               frm040109_1.Show
               frm040109_1.cmdOK(0).Value = True 'Added by Morgan 2015/9/14
            End If
         'Add By Sindy 2022/11/23
         ElseIf txtF0301 <> MsgText(601) And cp(157) = "" Then
            strExc(10) = Pub_GetCRLCaseMap(txtF0301, "3", "P", pa(1), pa(2), pa(3), pa(4))
            If strExc(10) <> "" Then
               frm040109_1.txtCode(4) = SystemNumber(strExc(10), 1)
               frm040109_1.txtCode(5) = SystemNumber(strExc(10), 2)
               frm040109_1.txtCode(6) = SystemNumber(strExc(10), 3)
               frm040109_1.txtCode(7) = SystemNumber(strExc(10), 4)
            End If
            '2022/11/23 END
         End If
   End Select
End Sub

Private Function FormSave() As Boolean
   Dim varTmp As Variant
   Dim strTxt(1 To 20) As String, strSql As String, intStep As Integer, strTmp(1 To 3) As String
   Dim i As Integer
   Dim StrSQLa As String, WorkDate1 As String, WorkDate2 As String
   Dim rsA As New ADODB.Recordset
   'edit by nickc 2007/02/02
   'Dim sPA(1 To T_PA) As String
   'Dim sSP(1 To T_SP) As String
   Dim sPA() As String
   Dim sSP() As String
   ReDim sPA(1 To TF_PA) As String
   ReDim sSP(1 To tf_SP) As String
   'Add by Morgan 2004/6/3
   Dim stCP71 As String '延緩月數
   'Add by Morgan 2004/7/21
   Dim stAD03 As String, stAD10 As String
   'Add by Morgan 2005/1/18
   Dim stCP33 As String, stCP34 As String
   'add by nickc 2005/06/07
   Dim stCM10 As String
   'Add by Morgan 2007/1/30
   Dim stInCNo(1 To 4) As String '國內案案號
   Dim stInPA08 As String '國內案專利種類
   Dim stInPA09 As String '國內案申請國家
   Dim stInPA10 As String '國內案申請日
   Dim stInPA14 As String '國內案公告日或預估公告日
   strPA14Msg = ""
   'end 2007/1/30
   Dim st307Msg As String '分割案提醒訊息
   
   Dim strCP48 As String '承辦期限
   'Add by Morgan 2010/5/7
   Dim bolInsCM As Boolean '是否新增國內外關聯
   Dim bolChgNoMail As Boolean '是否通知轉案號 Add by Morgan 2010/7/22
   Dim stCP141 As String, stCP142 As String 'Add by Morgan 2010/12/29
   Dim m_list() As String '相關案 Added by Morgan 2012/3/26
   Dim stDate(3) As String 'Added by Morgan 2012/6/28
   Dim stCP118 As String 'Added by Morgan 2013/5/14
   Dim strCP29SQL As String 'add by sonia 2015/11/25
   Dim strDivCaseNo(1 To 4) As String 'Added by Lydia 2021/07/19 分割案之母案案號
   Dim stCP122 As String 'Add by Amy 2022/11/15
   Dim douStPrice As Double, douLowPrice As Double
   Dim stCP164 As String 'Added by Morgan 2023/8/29
   
On Error GoTo ErrorHandler
   
   cnnConnection.BeginTrans
   
   intStep = 1

   '若有輸入轉本所案號
   If Me.textPA1.Text <> "" And Me.textPA2.Text <> "" Then
   
      If cp(31) = "Y" Then bolChgNoMail = True 'Add by Morgan 2010/7/22
      
      textPA3 = Right("0" & textPA3, 1)
      textPA4 = Right("00" & textPA4, 2)
      
      'Modify by Morgan 2010/12/28 要先新增基本檔,否則紀錄原FC代理人的 Trigger 會錯
      '判斷是否新增專利或服務業務基本案
      Select Case pa(1)
         Case "P", "CFP", "FCP":
            StrSQLa = "SELECT * FROM PATENT WHERE " & ChgPatent(Me.textPA1.Text & Me.textPA2.Text & Me.textPA3.Text & Me.textPA4.Text)
         Case Else:
            StrSQLa = "SELECT * FROM SERVICEPRACTICE WHERE " & ChgService(Me.textPA1.Text & Me.textPA2.Text & Me.textPA3.Text & Me.textPA4.Text)
      End Select
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, StrSQLa)
      If intI = 0 Then
         bolChgNoMail = True 'Add by Morgan 2010/7/22
         Select Case pa(1)
            Case "P", "CFP", "FCP":
               If PUB_ReadPatentData(sPA(), pa(1), pa(2), pa(3), pa(4)) Then
                  sPA(1) = Me.textPA1.Text
                  sPA(2) = Me.textPA2.Text
                  sPA(3) = Left(Me.textPA3.Text & "0", 1)
                  sPA(4) = Left(Me.textPA4.Text & "00", 2)
                  If PUB_AddNewPatent(sPA()) = False Then
                    GoTo ErrorHandler
                  End If
               End If
            Case Else:
               If PUB_ReadServicePracticeData(sSP(), pa(1), pa(2), pa(3), pa(4)) Then
                  sSP(1) = Me.textPA1.Text
                  sSP(2) = Me.textPA2.Text
                  sSP(3) = Left(Me.textPA3.Text & "0", 1)
                  sSP(4) = Left(Me.textPA4.Text & "00", 2)
                  If PUB_AddNewServicePractice(sSP()) = False Then
                    GoTo ErrorHandler
                  End If
               End If
         End Select
      End If
      
      'Modify by Morgan 2004/6/8 新案旗標要清除 CP31=NULL
      '2005/7/14 MODIFY BY SONIA 若該案號無基本檔則 CP31='Y' 否則 NULL
      'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP01='" & Me.textPA1.Text & "',CP02='" & Left(Me.textPA2.Text & "000000", 6) & "' ,CP03='" & Left(Me.textPA3.Text & "0", 1) & "' ,CP04='" & Left(Me.textPA4.Text & "00", 2) & "' ,CP43='', CP31=NULL WHERE CP09='" & Me.Label3(8).Caption & "'"
      strTxt(intStep) = "UPDATE CASEPROGRESS a SET CP01='" & Me.textPA1.Text & "',CP02='" & Left(Me.textPA2.Text & "000000", 6) & "' ,CP03='" & Left(Me.textPA3.Text & "0", 1) & "' ,CP04='" & Left(Me.textPA4.Text & "00", 2) & "' ,CP43=''"
      'Modify by Moragn 2010/7/22 改判斷進度檔
      'If textPA1 = "P" Then
      '   strTxt(intStep) = strTxt(intStep) & ", CP31=(select DECODE(MAX(1),NULL,'Y',NULL) from PATENT where PA01='" & Me.textPA1.Text & "' AND PA02='" & Me.textPA2.Text & "' AND PA03='" & Me.textPA3.Text & "' AND PA04='" & Me.textPA4.Text & "')"
      'Else
      '   strTxt(intStep) = strTxt(intStep) & ", CP31=(select DECODE(MAX(1),NULL,'Y',NULL) from servicepractice where SP01='" & Me.textPA1.Text & "' AND SP02='" & Me.textPA2.Text & "' AND SP03='" & Me.textPA3.Text & "' AND SP04='" & Me.textPA4.Text & "')"
      'End If
      strTxt(intStep) = strTxt(intStep) & ", CP31=(select DECODE(MAX(1),NULL,'Y',NULL) from CASEPROGRESS b where b.cp01='" & textPA1 & "' AND b.cp02='" & textPA2 & "' AND b.cp03='" & textPA3 & "' AND b.cp04='" & textPA4 & "')"
      'end 2010/7/22
      strTxt(intStep) = strTxt(intStep) & " WHERE CP09 = '" & Me.Label3(8).Caption & "'"
      '2005/7/14 END
      'Add By Cheng 2002/11/05
      cnnConnection.Execute strTxt(intStep), intI
      intStep = intStep + 1
            
      'Add by Morgan 2006/9/7
      '更正財務相關資料
      PUB_UpdateAccData cp(9), cp(1) & cp(2) & cp(3) & cp(4)
      
      'Add by Morgan 2006/9/20
      '要轉關聯
      strMail2FEngCP09 = ""
      If Me.Tag = "1" Then
         strExc(1) = textPA1
         strExc(2) = textPA2
         strExc(3) = textPA3
         strExc(4) = textPA4
         strExc(0) = "select cp09 from casemap,caseprogress where cm10='0' and cm05='" & cp(1) & "' and cm06='" & cp(2) & "' and cm07='" & cp(3) & "' and cm08='" & cp(4) & "' and cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03 and cp04(+)=cm04 and cp27 is null and cp57 is null and cp14 is not null and cp10 in (" & CaseMapOut & ")"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Do While Not RsTemp.EOF
               strMail2FEngCP09 = strMail2FEngCP09 & IIf(strMail2FEngCP09 <> "", ",", "") & RsTemp.Fields(0)
               RsTemp.MoveNext
            Loop
         End If
         PUB_UpdateCaseRelation cp, strExc
         
      End If
      'end 2006/9/20
            
'cancel by sonia 2024/11/26 已不立卷不必再通知分所收文人員
'      'Add by Morgan 2010/7/22 若為分所收文案件則發Mail通知收文人員
'      'If bolChgNoMail Then 'Remove by Morgan 2010/8/6
'         strExc(0) = PUB_GetST06(cp(65))
'         If strExc(0) > "1" Then
'            strExc(1) = "原本所案號 " & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
'            strExc(1) = strExc(1) & " 已更改為 " & textPA1 & "-" & textPA2 & IIf(textPA3 & textPA4 = "000", "", "-" & textPA3 & "-" & textPA4) & " 。"
'            '2010/12/2 modify by sonia
'            'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'               " values ('" & strUserNum & "','" & cp(65) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'               ",'" & ChgSQL(strExc(1)) & "','如旨' )"
'            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'               " values ('" & strUserNum & "','" & cp(65) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'               ",'" & ChgSQL(strExc(1)) & "','總收文號：" & Label3(8) & " 改本所案號如主旨')"
'            cnnConnection.Execute strSql, intI
'         End If
'      'End If 'Remove by Morgan 2010/8/6
'      'end 2010/7/22
'end 2024/11/26
      
   '若未輸入轉本所案號
   Else
            
      'Add by Morgan 2004/7/21
      '設定客戶減免身分
      If Text1(13).Text = "000" Then
         For i = 1 To 5
            If txtAD(i).Enabled = True Then
               '身分有變更才要做
               If txtAD(i).Tag <> txtAD(i).Text Then
                  '不可減免
                  If txtAD(i).Text = "N" Then
                     strSql = PUB_GetADSQL(pa(25 + i), Text1(13).Text, "N")
                  '自然人
                  'Modify by Morgan 2006/3/14
                  '學校也不用證明
                  'ElseIf txtAD(i).Text = "1" Then
                  ElseIf (txtAD(i).Text = "1" Or txtAD(i).Text = "2") Then
                     strSql = PUB_GetADSQL(pa(25 + i), Text1(13).Text, "Y", txtAD(i).Text)
                  '公司
                  Else
                     '原來沒有減免資料或不可減免
                     If txtAD(i).Tag = "" Or txtAD(i).Tag = "N" Then
                        strSql = PUB_GetADSQL(pa(25 + i), Text1(13).Text, "Y", txtAD(i).Text, pa(1), pa(2), pa(3), pa(4))
                     '修改減免身分別
                     Else
                        strSql = PUB_GetADSQL(pa(25 + i), Text1(13).Text, "Y", txtAD(i).Text)
                     End If
                  End If
                  cnnConnection.Execute strSql
               End If
            End If
         Next
      End If
      'end 2004/7/21
      
      'Add by Morgan 2006/11/3
      '大陸發明申請案若為PCT案或有主張優先權時需掛實審期限
      If Text1(13) = "020" And Text1(1) = "101" Then
         strExc(10) = ""
         If Text1(18) <> "" Then
            strExc(10) = Text1(18)
         ElseIf Text1(14) <> "" Then
            strExc(10) = Text1(14)
         Else
            'Modify by Amy 2023/01/05 原:strPriority(2)
            strExc(10) = PUB_GetFirstPriDate2(strPrity2)
         End If
         If strExc(10) <> "" Then
            strExc(0) = "select na27 from nation where na01='020'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(9) = CompDate(1, Val("" & RsTemp(0)), strExc(10)) '法限
               
               'Add by Morgan 2009/11/4
               'FMP 案實審的所限=法限-10天
               'Modified by Morgan 2018/10/3 非FMP也改10天
               'If m_bolFMP Then
                  'Added by Lydia 2025/10/29
                  strExc(1) = ""
                  If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                     strExc(0) = PUB_GetPOurDeadline(strExc(9), Text1(13), strExc(1), pa(1), Text1(1))
                  Else
                  'end 2025/10/29
                     strExc(0) = CompDate(2, -10, strExc(9))
                  End If 'Added by Lydia 2025/10/29
               'Else
               'end 2009/11/4
               '   strExc(0) = ""
               '   strExc(1) = cp(1)
               '   strExc(2) = Text1(13)
               '   strExc(3) = strExc(9)
               '   GetCtrlDT strExc
               'End If
               'end 2018/10/3
               
               strExc(8) = PUB_GetWorkDay1(strExc(0), True) '所限
               If PUB_ChkCPExist(cp, "416", 0, strExc(7)) Then
                  strSql = "Update CaseProgress Set CP06=" & strExc(8) & ",CP07=" & strExc(9) & " Where CP09='" & strExc(7) & "' and cp27 is null"
               ElseIf PUB_ChkNPExist(cp, "416", 0, strExc(6), strExc(5)) Then
                  'Modified by Lydia 2025/10/29 +NP23
                  strSql = "Update NextProgress Set NP08=" & strExc(8) & ",NP09=" & strExc(9) & ",NP23=" & IIf(strExc(1) = "", "NP23", strExc(1)) & "  Where NP22=" & strExc(6) & " and NP01='" & strExc(5) & "'"
               Else
                  strSql = "declare intMax number;begin select max(np22)+1 into intMax from nextprogress;IF intMax IS NULL THEN intMax:=1; END IF;"
                  'Modified by Lydia 2025/10/29 +NP23
                  strSql = strSql & "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22,NP23) " & _
                     " Values ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','416'," & strExc(8) & "," & strExc(9) & ",'" & cp(13) & "',intMax," & CNULL(strExc(1), True) & "); "
                  strSql = strSql & " end;"
               End If
               cnnConnection.Execute strSql, intI
            End If
         End If
      End If
      'end 2006/11/3
            
      'Add by Morgan 2004/3/22
      '若有輸入分割母案本所案號則更新 DIVISIONCASE
      If Text1(1) = "307" Then
         'Added by Lydia 2021/07/19
         strDivCaseNo(1) = txtDivCaseNo(1)
         strDivCaseNo(2) = txtDivCaseNo(2)
         strDivCaseNo(3) = txtDivCaseNo(3)
         strDivCaseNo(4) = txtDivCaseNo(4)
         'end 2021/07/19
         If (txtDivCaseNo(1) <> txtDivCaseNo(1).Tag Or txtDivCaseNo(2) <> txtDivCaseNo(2).Tag Or txtDivCaseNo(3) <> txtDivCaseNo(3).Tag Or txtDivCaseNo(4) <> txtDivCaseNo(4).Tag) Then
            '若原先有建立關聯則更新，否則新增
            If txtDivCaseNo(1).Tag <> "" Then
               strTxt(intStep) = " UPDATE DIVISIONCASE SET DC05='" & txtDivCaseNo(1) & "', DC06='" & txtDivCaseNo(2) & "', DC07='" & txtDivCaseNo(3) & "', DC08='" & txtDivCaseNo(4) & "'" & _
                  " , DC12='" & strUserNum & "', DC13=TO_CHAR(SYSDATE,'YYYYMMDD'), DC14=TO_CHAR(SYSDATE,'HHMISS')" & _
                  " WHERE DC01='" & pa(1) & "' AND DC02='" & pa(2) & "' AND DC03='" & pa(3) & "' AND DC04='" & pa(4) & "'"
            Else
               strTxt(intStep) = " INSERT INTO DIVISIONCASE (DC01, DC02, DC03, DC04, DC05, DC06, DC07, DC08, DC09, DC10, DC11 )" & _
                  " VALUES('" & pa(1) & "' ,'" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'" & _
                  " ,'" & txtDivCaseNo(1) & "', '" & txtDivCaseNo(2) & "', '" & txtDivCaseNo(3) & "', '" & txtDivCaseNo(4) & "'" & _
                  " ,'" & strUserNum & "', TO_CHAR(SYSDATE,'YYYYMMDD'), TO_CHAR(SYSDATE,'HHMISS'))"
            End If
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
         End If
         
'Removed by Morgan 2011/12/7 P案收文不必管控實審期限((發文管控的才是真的法限)),而且母案應該都會有期限會更新到分割--郭 Ex.P-100262
'         'FCP及P的國內案件，若專利種類為'發明'且案件性質為'分割'時，實審期限=母案之申請日＋3年<=分割案之收文日＋1月
'         '若有收文未取消收文之'實體審查'，則更新該筆'實體審查'之期限，若無則新增下一程序'實體審查'期限，並顯示'此分割案尚未收文實體審查，期限為XXXXXX，請提醒智權人員 !!'
'         If pa(1) = "P" And Text1(2) = "1" And Text1(13) = "000" Then
'            '有收'實體審查'
'            If m_stCP09 <> "" Then
'               strTxt(intStep) = " UPDATE CASEPROGRESS SET CP06=" & m_stVar(0) & ",CP07=" & m_stVar(3) & " WHERE CP09='" & m_stCP09 & "'"
'            '沒有收'實體審查'
'            Else
'               strTxt(intStep) = _
'                  " DECLARE" & _
'                     " V_NP22 NUMERIC(10,0);" & _
'                     " V_NP02 VARCHAR2(9);" & _
'                  " BEGIN" & _
'                     " SELECT MAX(NP02) INTO V_NP02 FROM NEXTPROGRESS WHERE NP01='" & Label3(8) & "' AND NP07='416';" & _
'                     " IF V_NP02 IS NULL THEN" & _
'                        " SELECT NVL(MAX(NP22),0)+1 INTO V_NP22 FROM NEXTPROGRESS;" & _
'                        " INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22)" & _
'                        " VALUES ('" & Label3(8) & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'" & _
'                        " ,'416'," & m_stVar(0) & "," & m_stVar(3) & ",'" & cp(13) & "',V_NP22);" & _
'                     " ELSE" & _
'                        " UPDATE NEXTPROGRESS SET NP08=" & m_stVar(0) & ",NP09=" & m_stVar(3) & " WHERE NP01='" & Label3(8) & "' AND NP07='416';" & _
'                     " END IF;" & _
'                  " END;"
'            End If
'            cnnConnection.Execute strTxt(intStep)
'            intStep = intStep + 1
'         End If

      End If
      'end 2004/3/22
      
      '更新基本檔
      strTmp(1) = ""
      Select Case pa(1)
         Case "P"
            '專利種類
            strTmp(1) = strTmp(1) & ",PA08='" & Me.Text1(2).Text & "'"
            '閉卷
            If Text1(11) = "Y" Then
               strTmp(1) = strTmp(1) & ",PA57=NULL,PA58=NULL,PA59=NULL"
            End If
            
            'Modify by Morgan 2007/8/30 加第三人申請技術報告807
            'If Text1(1) = 異議_專 Or Text1(1) = 舉發 Or Text1(1) = 鑑定報告 Then
            If Text1(1) = 異議_專 Or Text1(1) = 舉發 Or Text1(1) = 鑑定報告 Or Text1(1) = "807" Then
               'Add by Morgan 2010/3/5
               '若原來是申請案且無對造資料時改更新基本檔資料到進度檔(收錯案件性質)
               If pa(23) = "1" Then
                  '檢查原對造案件名稱空白才做
                  If cp(37) & cp(38) & cp(39) = "" Then
                     cp(37) = pa(5)
                     cp(38) = pa(6)
                     cp(39) = pa(7)
                     strSql = "update caseprogress set cp37='" & ChgSQL(cp(37)) & "',cp38='" & ChgSQL(cp(38)) & "',cp39='" & ChgSQL(cp(39)) & "' where cp09='" & cp(9) & "'"
                     cnnConnection.Execute strSql, intI
                  End If
               End If
               
               '公告日
               strTmp(1) = strTmp(1) & ",PA14=" & CNULL(DBDATE(Text1(7)), True)
               'Add by Morgan 2004/7/28
               'If Text1(1) = 異議_專 Or Text1(1) = 舉發 Then
               If Text1(1) = 異議_專 Or Text1(1) = 舉發 Or Text1(1) = "807" Then
                  '案件名稱(PA05,PA06,PA07),申請案號(PA11)
                  strTmp(1) = strTmp(1) & ",PA05='" & ChgSQL(cp(37)) & "',PA06='" & ChgSQL(cp(38)) & "',PA07='" & ChgSQL(cp(39)) & "',PA11='" & ChgSQL(Text1(17).Text) & "'"
                  '證書號
                  If Text1(1) = "807" Then
                     strTmp(1) = strTmp(1) & ",PA22=" & CNULL(Text1(22))
                  End If
               End If
            End If
            'end 2007/8/30
                     
            'Modify by Morgan 2006/5/23 加PCT資料回寫
            'Modify by Morgan 2006/7/19 PCT進入國家階段可請發明或新型
            'If Text1(13) = 大陸國家代號 And Text1(1) = "101" Then
            If Text1(13) = 大陸國家代號 And (Text1(1) = "101" Or Text1(1) = "102") Then
               '更新案件備註
               'Modify by Morgan 2009/11/30 +PCT申請號
               'Text1(20) = PUB_GetNewCaseMemo(Text1(20).Text, Text1(18).Text)
               Text1(20) = PUB_GetNewCaseMemo(Text1(20).Text, Text1(18).Text, Text1(23).Text)
               '是否PCT案,申請日
               If Text1(14) <> "" Then
                  strTmp(1) = strTmp(1) & ",PA46='Y',PA10=" & TransDate(Text1(14), 2)
                  'Modify by Amy 2013/12/11 Mark 2013/08/28程式 改由AutoBatchDay做
'                  'Add by Amy 2013/08/28 PCT進入國家階段之發明或新型需將PCT申請案下一程序「進入國家階段」期線上「Y」若未算過結餘自動上結餘
'                  strExc(0) = "Select pa01||pa02||pa03||pa04,pa01,pa02,pa03,pa04 From Patent Where pa11='" & Text1(23) & "' and pa09='056' "
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI > 0 Then
'                       strExc(0) = "Update NextProgress set NP06='Y' Where np07='119' and " & ChgNextProgress(RsTemp.Fields(0))
'                       cnnConnection.Execute strExc(0), intI
'                       bolEndModCash = True
'                       Pub_UpdateEndModCash RsTemp.Fields("pa01"), RsTemp.Fields("pa02"), RsTemp.Fields("pa03"), RsTemp.Fields("pa04")
'                  End If
'                 'end 2013/08/28

               Else
                  strTmp(1) = strTmp(1) & ",PA46=Null"
               End If
            End If
            'end 2006/5/23
            
            '是否有救濟程序
            If Left(Text1(1), 1) = 5 Then
               strTmp(1) = strTmp(1) & ",PA18='Y'"
            End If
            '是否有爭議程序
            If Left(Text1(1), 1) = 8 Then
               strTmp(1) = strTmp(1) & ",PA19='Y'"
            End If

            'Add by Morgan 2010/3/17
            If txtFavDt.Visible = True Then
               'Modified by Morgan 2012/3/22 改放 PA140
               'strExc(0) = PUB_GetFavorDate(Text1(20))
               'If strExc(0) <> "" Then
               '   Text1(20) = Replace(Text1(20), "新穎性優惠期日期" & strExc(0) & ";", "")
               'End If
               'Text1(20) = "新穎性優惠期日期" & txtFavDt.Text & ";" & Text1(20)
               strTmp(1) = strTmp(1) & ",PA140=" & CNULL(DBDATE(txtFavDt), True)
               'end 2012/3/22
            End If

            'Added by Morgan 2012/9/5
            'Modify by Amy 2016/08/29
            If txtPA161.Visible = True And txtPA161.Tag <> txtPA161 Then
               strTmp(1) = strTmp(1) & ",PA161='" & txtPA161 & "'"
            End If
            'end 2012/9/5
            
            'Added by Morgan 2022/12/28
            If txtPA178.Visible And txtPA178 <> pa(178) Then
               strTmp(1) = strTmp(1) & ",PA178='" & txtPA178 & "'"
            End If
            'end 2022/12/28
            
            'Modify By Sindy 2010/10/29 增加更新PA158
            'Modified by Morgan 2012/3/12 +PA150
            strTxt(intStep) = "UPDATE PATENT SET PA09=" & CNULL(Text1(13)) & _
               ",PA23=" & CNULL(Text1(3)) & ",PA47=" & CNULL(Text1(21)) & ",PA48=" & CNULL(Text1(8)) & _
               ",PA91=" & CNULL(ChgSQL(Text1(20))) & strTmp(1) & _
               ",PA158=" & CNULL(Left(Combo3, 1)) & ",PA150='" & txtEngGroup & "'" & _
               " WHERE " & ChgPatent(Label3(9))
            Pub_SeekTbLog strTxt(intStep)  'Added by Morgan 2024/10/7
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
            
            'Added by Lydia 2021/07/19 FMP分案(307)新增分割母案案號時，將母案之案件名稱、發明人、代表人資料、優先權資料、個案地址複製到子案。
            If Text1(1) = "307" And m_bolFMP = True And txtDivCaseNo(1).Tag = "" And strDivCaseNo(1) <> "" And strDivCaseNo(2) <> "" Then
                Call PUB_FCPCopyDataToCase(pa(), strDivCaseNo())
                '直接帶母案的優先權資料; 因為後面ClsPDSavePriority預設會清除優先權資料
                'Modify by Amy 2023/01/05 strPriority原陣列,改變數
                If Not ClsPDReadPriority(strDivCaseNo, strPrity1, strPrity2, strPrity3, strPrity4, strPrity5) Then
                End If
            End If
            'end 2021/07/19
            
            'Added by Lydia 2018/08/31 FMP案輸入分割母案本所案號，將母案名稱代入子案 '8/30 問郭經理P案不用
            'Memo by Lydia 2021/07/19 在修改分割母案本所案號，將母案名稱代入子案
            If Text1(1) = "307" And m_bolFMP = True And txtDivCaseNo(1).Tag <> "" And _
                     (txtDivCaseNo(1) <> txtDivCaseNo(1).Tag Or txtDivCaseNo(2) <> txtDivCaseNo(2).Tag Or txtDivCaseNo(3) <> txtDivCaseNo(3).Tag Or txtDivCaseNo(4) <> txtDivCaseNo(4).Tag) Then
                 strExc(1) = ""
                 If ClsPDCheckCaseCodeIsExist(txtDivCaseNo(1), txtDivCaseNo(2), txtDivCaseNo(3), txtDivCaseNo(4), strExc(5), strExc(6), strExc(7)) = True Then
                     If pa(5) <> strExc(5) Then strExc(1) = strExc(1) & ", PA05=" & CNULL(ChgSQL(strExc(5)))
                     If pa(6) <> strExc(6) Then strExc(1) = strExc(1) & ", PA06=" & CNULL(ChgSQL(strExc(6)))
                     If pa(7) <> strExc(7) Then strExc(1) = strExc(1) & ", PA07=" & CNULL(ChgSQL(strExc(7)))
                     If strExc(1) <> "" Then
                          strSql = "Update Patent set " & Mid(strExc(1), 2) & "  where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "' "
                          'Modified by Lydia 2018/10/22 +詳細記錄(True)
                          'Modified by Lydia 2025/10/30 改用模組判斷
                          'Pub_SeekTbLog strSql, , True
                          Pub_SeekTbLog strSql, , PUB_FilterSeekSQL(strSql)
                          cnnConnection.Execute strSql
                     End If
                 End If
            End If
            'end 2018/08/31
         Case "PS"
            If Text1(11) = "Y" Then
               strTmp(1) = strTmp(1) & ",SP15=NULL,SP16=NULL,SP17=NULL"
            End If
            'Add by Amy 2017/07/13 回存基本檔
            If txtPA161.Visible = True And txtPA161.Tag <> txtPA161 Then
                strTmp(1) = strTmp(1) & ",SP85='" & txtPA161 & "'"
            End If
            
            'Modified by Morgan 2012/3/12 +SP79
            strTxt(intStep) = "UPDATE SERVICEPRACTICE SET SP09=" & CNULL(Text1(13)) & _
               ",SP29=" & CNULL(Text1(8)) & ",SP28=" & CNULL(Text1(21)) & _
               ",SP18=" & CNULL(Text1(20)) & strTmp(1) & ",SP79='" & txtEngGroup & "'" & _
               " WHERE " & ChgService(Label3(9))
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
            
      End Select
      
      'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之年費605、維持費606、延展費607，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。
      If Text1(11) = "Y" Then
         strMsgCloseCancel = PUB_GetCaseCloseCancel(pa(1), pa(2), pa(3), pa(4), pa(9))
      End If
      
      '更新進度檔
      strTmp(2) = ""
      
      'Modify by Morgan 2007/8/30 加第三人申請技術報告807
      'If Text1(1) = 異議_專 Or Text1(1) = 舉發 Or Text1(1) = 鑑定報告 Then
      If Text1(1) = 異議_專 Or Text1(1) = 舉發 Or Text1(1) = 鑑定報告 Or Text1(1) = "807" Then
         strTmp(2) = strTmp(2) & ",CP36=" & CNULL(Text1(17))
      End If
      
      'Add by Morgan 2004/6/3 加延緩月數
      If txtCP71.Visible = True Then
         strTmp(2) = strTmp(2) & ",CP71=" & "'" & Val(txtCP71) & "'"
      End If
      
      'Add by Morgan 2010/3/15
      If Text1(25).Enabled = True Then
         strTmp(2) = strTmp(2) & ",CP48=" & CNULL(DBDATE(Text1(25)), True)
      End If
      
      'Add by Morgan 2010/3/17
      'Modify by Morgan 2010/12/29 改用 option
      'If txtTakeCtrl.Text = "Y" Then
      If cp(27) = "" Then
      
'Removed by Morgan 2024/2/22 已收款通知已改為有設收款後送件時依台灣(PS1)或非台灣(PS2)通知特殊設定人員
'         If OptSendType(2).Value = True Then '收款後送件
'            If Text1(0) <> "" Then
'               'Added by Morgan 2023/11/10
'               'Modified by Morgan 2024/2/2
'               'strExc(1) = GetStaffDepartment(Text1(0))
'               'If strExc(1) = "P12" Then
'               '   strExc(2) = Text1(0)
'               If Val(cp(16)) > 0 And Val(cp(79)) > 0 Then
'                  If pa(9) = "000" Then
'                     strExc(2) = Pub_GetSpecMan("PS1")
'                  Else
'                     strExc(2) = Pub_GetSpecMan("PS2")
'                  End If
'               'end 2023/11/10
'                  strSql = "update UndeliveredRec set UD04='" & strExc(2) & "' where UD01='" & cp(9) & "' and UD02=" & strSrvDate(1)
'                  cnnConnection.Execute strSql, intI
'                  If intI = 0 Then
'                     strSql = "insert into UndeliveredRec(UD01,UD02,UD03,UD04) VALUES('" & cp(9) & "'," & strSrvDate(1) & ",'1','" & strExc(2) & "')"
'                     cnnConnection.Execute strSql, intI
'                  End If
'               End If 'Added by Morgan 2023/11/10
'            End If
'         Else
'            strSql = "delete UndeliveredRec where UD01='" & cp(9) & "'"
'            cnnConnection.Execute strSql, intI
'         End If
         
      'Add by Morgan 2010/12/29
         intI = Abs(OptSendType(1).Value * 1) + Abs(OptSendType(2).Value * 2) + Abs(OptSendType(3).Value * 3)
         If intI > 0 Then
            stCP141 = intI
         End If
         If txtCP142.Text <> "" Then
            stCP142 = DBDATE(txtCP142.Text)
         End If
         stCP141 = CNULL(stCP141)
         stCP142 = CNULL(stCP142, True)
      Else
         stCP141 = "CP141"
         stCP142 = "CP142"
      'end 2010/12/29
      End If
      
      'Added by Morgan 2023/8/29
      stCP164 = "CP164"
      'Added by Morgan 2023/8/29
      'Modified by Morgan 2025/1/23
      If Frame2.Visible = True And Frame2.Enabled = True Then
         stCP164 = CNULL(IIf(Option1(0).Value = True, "1", IIf(Option1(1).Value = True, "2", IIf(Option1(2).Value = True, "3", ""))))
      End If
      'end 2023/8/29
      
      'Memo by Lydia 2016/07/07 與工程師主管分案(frm060117) 相同
      'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
      If m_bolUpdCP27 = True Then
         strTmp(2) = strTmp(2) & ",CP27=" & strSrvDate(1)
      End If
      
      'Add By Sindy 2022/12/15 有修改案件性質
      If Text1(1).Tag <> Text1(1).Text Then
         'Modify By Sindy 2023/9/22 + , Text1(1).Tag
         If PUB_ModCrLCRCData(Label3(8), cp(140), Text1(1).Text, Text1(1).Tag, pa(9), Text1(19)) = False Then
            GoTo ErrorHandler
         End If
      End If
      '2022/12/15 END
      
      '標準價,底價
      'Modify by Morgan 2010/10/29 改申請國家或案件性質時重抓(若改標準價原收文資料要維持不變)--秀玲
      If m_strOldCP10 <> Text1(1) Or m_strOldPA09 <> Text1(13) Then
         'Modify By Sindy 2022/12/15
'         strExc(0) = "SELECT CF13,CF14 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & Text1(1) & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            stCP33 = Val("" & RsTemp.Fields(0))
'            stCP34 = Val("" & RsTemp.Fields(1))
'            strTmp(2) = strTmp(2) & ",CP33=" & stCP33
'            strTmp(2) = strTmp(2) & ",CP34=" & stCP34
'         End If
         'If ClsPDGetCaseLowPrice(pa(1), pa(9), Text1(1), douStPrice, douLowPrice, Text1(2), Left(Combo3, 1), cp(140)) = 1 Then
         If ClsPDGetCaseLowPrice(pa(1), Text1(13), Text1(1), douStPrice, douLowPrice, Text1(2), Left(Combo3, 1), cp(140)) = 1 Then
            strTmp(2) = strTmp(2) & ",CP33=" & douStPrice
            strTmp(2) = strTmp(2) & ",CP34=" & douLowPrice
         End If
         '2022/12/15 END
      End If
      
      'ADD BY SONIA 2014/5/28 台灣年費之回復原狀414自動掛相關總收文號為年費,且本所及法定設定同年費
      If Text1(13).Text = "000" And Text1(1) = "414" Then
         strExc(0) = "SELECT CP09,CP06,CP07 FROM CASEPROGRESS WHERE " & ChgCaseprogress(Label3(9)) & " AND CP10='605' AND CP27 IS NULL AND CP57 IS NULL"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Text1(6) = "" & RsTemp.Fields(0)
            Text1(4) = TransDate(Val("" & RsTemp.Fields(1)), 1)
            Text1(5) = TransDate(Val("" & RsTemp.Fields(2)), 1)
         End If
      End If
      'END 2014/5/28
      
      'Added by Morgan 2023/7/5
      'A/B類的擇一申復也請調整為收文日起算7個工作日--郭
      'Removed by Morgan 2023/8/1 取消-郭
      'If pa(1) = "P" And Text1(1) = "239" And cp(27) = "" Then
      '   Text1(4) = PUB_GetPBRecCP06(Text1(4).Text, pa(1), Text1(1).Text, m_bolFMP, Text1(12).Text, True)
      'End If
      'end 2023/7/5
      
      '本所期限
      'Remove by Morgan 2007/6/6 改由Trigger控制
      '小於系統日時設為系統日
      'Modify by Morgan 2007/8/23 恢復控制但加判斷承辦人無到有
      'If Text1(4).Text <> "" And Val(Text1(4)) < strSrvDate(2) Then
      If Text1(4).Text <> "" And Val(Text1(4)) < Val(strSrvDate(2)) And Text1(0).Tag = "" And Text1(0) <> "" Then
         strTmp(2) = strTmp(2) & ",CP06=" & strSrvDate(1)
      Else
         strTmp(2) = strTmp(2) & ",CP06=" & CNULL(DBDATE(Text1(4)), True)
      End If
      'end 2007/6/6
      
      'Add by Morgan 2010/1/21
      If m_lngRefund > 0 Then
         '繳年費起迄年
         If m_i605FromYear > 0 Then
            strTmp(2) = strTmp(2) & ",CP53='" & m_i605FromYear & "'"
            strTmp(2) = strTmp(2) & ",CP54='" & m_i605ToYear & "'"
         '考慮退費若重分案且不移作次年時要清除
         ElseIf Text1(1) = "908" Then
            strTmp(2) = strTmp(2) & ",CP53=NULL,CP54=NULL"
         End If
      'Add by Morgan 2010/3/16
      ElseIf txtFeeYear(1).Visible = True Then
         strTmp(2) = strTmp(2) & ",CP53='" & txtFeeYear(1) & "'"
         strTmp(2) = strTmp(2) & ",CP54='" & txtFeeYear(2) & "'"
         
      End If
      'end 2010/1/21
      
      'Add by Morgan 2011/4/22 延期要紀錄NP22
      If Text1(1) = "404" Then
         strTmp(2) = strTmp(2) & ",CP30='" & m_CP30 & "'"
      End If
      
      'Added by Morgan 2013/5/15
      If txtCP118.Enabled = True Then
         stCP118 = txtCP118
         'Modified by Morgan 2024/1/30 +台灣(因增加大陸案也可設定電子送件)
         If stCP118 = "Y" And Text1(13) = "000" Then
            If cp(118) = "" Then
               'Modified by Morgan 2013/9/12 +中間程序也可設定電子送,改判斷新申請案
               'stCP118 = "W"
               'Modified by Morgan 2015/6/17 +改請,分割
               If InStr(NewCasePtyList, Text1(1)) > 0 Or Left(Text1(1), 1) = "3" Then
                  stCP118 = "W"
               End If
            Else
               stCP118 = cp(118)
            End If
         End If
      Else
         stCP118 = cp(118)
      End If
      'end 2013/5/15

      '2012/7/19 modify by sonia 加是否電子送件欄cp118
      'Modified by Morgan 2012/7/20 +CP147
      'Modify by Amy 2015/01/22 北所第一次輸入承辦人若北所分案日為null則更新為系統日
      strTxt(intStep) = "": bolCP14Mail = False
      If Text1(0).Tag = "" Then
        If Text1(0).Tag <> Trim(Text1(0).Text) Then
            '原承辦人沒值,且有修改
            strTxt(intStep) = strTxt(intStep) & ",CP14=" & CNULL(Text1(0))
            If m_bolIsFirstKeyCP14 And cp(157) = "" Then
                strTxt(intStep) = strTxt(intStep) & ",CP157=" & strSrvDate(1)
            End If
        End If
      Else
        If Trim(Text1(0)) = "" Then
            '原承辦人有值,但畫面上為空
            strExc(1) = ClsPDGetStaff(Text1(0).Tag, strExc(0))
            strExc(0) = "此程序原承辦人為 " & strExc(0) & " 是否取消原承辦人?" & vbCrLf & vbCrLf & _
                            "是:取消原承辦人 / 否:保留原承辦人"
            If MsgBox(strExc(0), vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                strTxt(intStep) = strTxt(intStep) & ",CP14=null,CP157=null"
                bolCP14Mail = True
            Else
                If m_bolIsFirstKeyCP14 And Text1(0) <> "" And cp(157) = "" Then
                    strTxt(intStep) = strTxt(intStep) & ",CP157=" & strSrvDate(1)
                End If
            End If
        ElseIf Text1(0).Text <> Text1(0).Tag Then
            '原承辦人有值,且有修改
            strTxt(intStep) = strTxt(intStep) & ",CP14=" & CNULL(Text1(0))
            If m_bolIsFirstKeyCP14 And Text1(0) <> "" And cp(157) = "" Then
                strTxt(intStep) = strTxt(intStep) & ",CP157=" & strSrvDate(1)
            End If
            bolCP14Mail = True
        Else
             If m_bolIsFirstKeyCP14 And Text1(0) <> "" And cp(157) = "" Then
                strTxt(intStep) = strTxt(intStep) & ",CP157=" & strSrvDate(1)
            End If
        End If
      End If
      
      'add by sonia 2015/11/25 承辦人若為專利處繪圖部門P13人員且該筆進度尚無繪圖人員時,自動更新CP29繪圖人員為承辦人
      strCP29SQL = ""
      If Trim(Text1(0)) <> "" Then
         strExc(0) = "SELECT CP09 FROM CASEPROGRESS,STAFF WHERE CP09='" & Label3(8) & "' AND CP29 IS NULL AND '" & Text1(0) & "'=ST01(+) AND 'P13'=ST03"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strCP29SQL = ",CP29=" & CNULL(Text1(0))
         End If
      End If
      'end 2015/11/25
     
     'modify by sonia 2015/11/25 加入strCP29SQL語法
     'Modified by Morgan 2018/4/18 +m_HKMemo
     'Modified by Lydia 2023/12/14 因為欄位只有顯示功能不會修改，所以拿掉",CP18=" & Val(Text1(16))  --- Sindy
     strTxt(intStep) = "UPDATE CASEPROGRESS SET CP05=" & DBDATE(Text1(12)) & _
         ",CP07=" & CNULL(DBDATE(Text1(5)), True) & _
         ",CP10=" & CNULL(Text1(1)) & ",CP13=" & CNULL(Text1(24)) & _
         strTxt(intStep) & ",CP26=" & CNULL(Text1(10)) & ",CP43=" & CNULL(Text1(6)) & _
         ",CP57=" & CNULL(TransDate(Text1(9), 2)) & _
         ",CP64=" & CNULL(ChgSQL(IIf(m_HKMemo <> "", m_HKMemo, "") & Text1(19))) & strTmp(2) & _
         ",CP118=" & CNULL(stCP118) & _
         ",CP141=" & stCP141 & ",CP142=" & stCP142 & ",CP147='" & txtCP147 & "',CP164=" & stCP164 & strCP29SQL
      'Modify By Sindy 2018/10/25 分案時若設定特殊出名代理人T(專利商標),請將案件進度資料的出名代理人設定為:林景郁
      'Modified by Lydia 2020/03/31 +strSrvDate(1) < 事務所合併日
      If strSrvDate(1) < 事務所合併日 And cp(27) = "" Then '未發文
         If txtPA161.Visible = True And txtPA161.Tag <> txtPA161 Then '有異動欄位值
            If txtPA161 = "T" Then
               strTxt(intStep) = strTxt(intStep) & ",CP110='94007'"
            ElseIf txtPA161.Tag = "T" Then
               strTxt(intStep) = strTxt(intStep) & ",CP110=null"
            End If
         End If
      End If
      strTxt(intStep) = strTxt(intStep) & " WHERE CP09='" & Label3(8) & "'"
      '2018/10/25 END
      'end 2015/01/22
      
      'Add by Morgan 2005/3/29 若承辦人變更時紀錄異動人員日期時間
      'Modify by Amy 2015/01/22 修改判斷條件
      'If Text1(0).Tag <> "" And Text1(0).Text <> Text1(0).Tag Then
      'Memo by Lydia 2016/07/07 與工程師主管分案(frm060117) 相同
      'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
      If bolCP14Mail = True Then
         'Add By Sindy 2013/11/28
         Call PUB_ChgEmpUpdEEP05(Label3(8), Text1(0).Tag, Text1(0).Text, "1")
         '2013/11/28 END
         
         'Modified by Morgan 2012/4/30 要寫 Log 並改觸發 Trigger 更新修改人員日期時間
         'strSql = "UPDATE CASEPROGRESS SET CP68='" & strUserNum & "',CP69=to_number(to_char(sysdate,'YYYYMMDD')),CP70=to_number(to_char(sysdate,'HH24MI')) WHERE CP09='" & cp(9) & "'"
         'cnnConnection.Execute strSql
         Pub_SeekTbLog strTxt(intStep)
         strTxt(intStep) = "begin user_data.user_enabled:=1; " & strTxt(intStep) & "; end;"
         'end 2012/4/30
      
      'Added by Morgan 2024/2/6 送件方式修改也要紀錄 還要再調整測試(單引號..
'      ElseIf cp(27) = "" And ((cp(141) <> stCP141 And UCase(stCP141) <> "CP141" And Not (cp(141) = "" And UCase(stCP141) = "NULL")) _
'         Or (cp(142) <> stCP142 And UCase(stCP142) <> "CP142" And Not (cp(142) = "" And UCase(stCP142) = "NULL")) _
'         Or (cp(164) <> stCP164 And UCase(stCP164) <> "CP164" And Not (cp(164) = "" And UCase(stCP164) = "NULL"))) Then
'         Pub_SeekTbLog strTxt(1)
'         strTxt(1) = "begin user_data.user_enabled:=1; " & strTxt(1) & "; end;"
         
      End If
      
      cnnConnection.Execute strTxt(intStep), intI
      intStep = intStep + 1
      
      'Added by Morgan 2023/8/31
      'FMP案:承辦期限<=指定送件日期-2工作天
      If m_bolFMP Then
         If OptSendType(3).Value = True And txtCP142 <> "" And Option1(2).Value = False Then
            strCP48 = CompWorkDay(3, DBDATE(txtCP142), 1)
            If strCP48 < strSrvDate(1) Then strCP48 = strSrvDate(1)
            strSql = "update caseprogress set cp48=" & strCP48 & " where cp09='" & cp(9) & "' and (cp48 is null or cp48>" & strCP48 & ")"
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2023/8/31
      
      'Added by Morgan 2023/8/11
      '若更換工程師或分案予不同的工程師之系統通知--杜燕文
      If Text1(0).Text <> "" Then
         If InStr("P10,P11", GetST15(Text1(0).Text)) > 0 Then
            '分案/改承辦人
            If m_bolIsFirstKeyCP14 Or Text1(0).Text <> Text1(0).Tag Then
               '排除已有發Mail的條件
               If Not (bolCP14Mail = True And Left(cp(9), 1) <> "B") And Not (Text1(13) = "000" And Text1(1) <> "601" And Text1(1) <> "605" And Text1(0).Tag = "" And Left(cp(9), 1) <> "B") Then
                  PUB_SetEngInform cp(9), Text1(0).Tag, True
               End If
            End If
         End If
      End If
      'end 2023/8/11
      
      'Add by Morgan 2010/6/30
      '異議答辯、舉發答辯更新對造號數名稱為被異議(舉發)之C類來函資料
      If Text1(1) = "802" Or Text1(1) = "804" Then
         If Text1(6) <> "" And cp(43) <> Text1(6) Then
            strSql = "update caseprogress a set (cp36,cp37,cp38,cp39,cp40,cp41,cp42)=(select b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42 from caseprogress b where b.cp09=a.cp43) where CP09='" & Label3(8) & "' and exists(select * from caseprogress b where b.cp09=a.cp43 and b.cp10 in ('1801','1802'))"
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2010/6/30
                     
      'Added by Lydia 2020/05/20 法律所案源收文：如果案件性質或申請國家有變化,則需要對應分案; 5/28 +配合開庭
      If strSrvDate(1) >= 法律所案源收文啟用日 And pa(1) = "P" And m_LOS07 = "" Then '排除已放棄的案源
          'Modified by Lydia 2020/07/23 重新整理: 因為案源收文已設定不可變更案件性質和申請國家,所以只要判斷有案源
          'If Text1(13).Text <> Text1(13).Tag Or m_CP10 <> Text1(1).Text Or (m_LOS15 = "" And txtLOSagree = "Y") Then
          '    Call PUB_UpdateCP10toPT(pa(1), pa(2), pa(3), pa(4), Label3(8), m_CP10, Text1(13).Tag, Text1(1).Text, Text1(13).Text, Text1(4).Text, cp(13), pa(26), IIf(m_LOS15 = "" And txtLOSagree = "Y", True, False))
          'End If
          '
          'If Text1(0).Tag = "" And Text1(0).Text <> "" Then
          '    strSql = PUB_GetLOSkind(pa(1), Text1(1).Text, Text1(13).Text)
          '    'Modified by Lydia 2020/06/09 判斷是否為補收文
          '    'If strSql <> "" Then
          '    strExc(1) = ""
          '    If m_LOS15 <> "" And strSql = "" Then strExc(1) = PUB_GetLOSplus(pa(1), pa(2), pa(3), pa(4), Text1(1).Text, Text1(13).Text, m_LOS02)
          '    If strSql <> "" Or strExc(1) <> "" Then
          '    'end 2020/06/09
          '       Call PUB_UpdateLOS01(pa(1), pa(2), pa(3), pa(4), Label3(8), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30), txtLOSagree)
          '    End If
          'End If
          If m_LOS15 <> "" And Text1(0).Tag = "" And Text1(0).Text <> "" Then
              Call PUB_UpdateLOS01(pa(1), pa(2), pa(3), pa(4), Label3(8), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30), txtLOSagree)
          End If
          'end 2020/07/23
      End If
      'end 2020/05/20
      
      '更新下一程序
      With MSHFlexGrid1
         For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "v" Then
               strTmp(1) = .TextMatrix(i, 7)
               strTmp(2) = .TextMatrix(i, 8)
               strTmp(3) = .TextMatrix(i, 9)
               'Modified by Lydia 2021/08/31 +更新NP24
               strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y', NP24='" & cp(9) & "' WHERE NP01='" & strTmp(1) & "' AND " & _
                  "NP07=" & strTmp(2) & " AND NP22=" & strTmp(3)
                'Add By Cheng 2002/11/05
                Pub_SeekTbLog strTxt(intStep) 'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業，若畫面勾選下一程序期限且存檔有上續辦Y的都寫Log以便事後能追蹤
                cnnConnection.Execute strTxt(intStep)
               intStep = intStep + 1
            End If
         Next
      End With
            
      '7
      'Modify by Morgan 2007/12/11 加判斷新案案件性質,原來有國內案且國內案或申請國家有改的才刪
      'If Text1(15) <> "" Then
      If InStr(CaseMapIn, Text1(1)) > 0 Then
         If Text1(15).Tag <> "" And (Text1(15).Text <> Text1(15).Tag Or Text1(13).Text <> Text1(13).Tag) Then
            'edit by nickc 2005/06/07
            'strTxt(intStep) = "DELETE FROM CASEMAP WHERE CM01='" & cm(0) & "' AND CM02='" & cm(1) & _
               "' AND CM03='" & cm(2) & "' AND CM04='" & cm(3) & "' AND CM05='" & cm(4) & _
               "' AND CM06='" & cm(5) & "' AND CM07='" & cm(6) & "' AND CM08='" & cm(7) & _
               "' AND CM10='0'"
            'Modified by Lydia 2015/09/09 +5.澳門發明案與大陸案之關聯
            strTxt(intStep) = "DELETE FROM CASEMAP WHERE CM01='" & cm(0) & "' AND CM02='" & cm(1) & _
               "' AND CM03='" & cm(2) & "' AND CM04='" & cm(3) & "' AND CM05='" & cm(4) & _
               "' AND CM06='" & cm(5) & "' AND CM07='" & cm(6) & "' AND CM08='" & cm(7) & _
               "' AND (CM10='0' or cm10='4' or cm10='5') "
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
         End If
         'end 2007/12/11
      End If
      
      If Text1(15) <> "" Then
         ChgCaseNo Text1(15).Text, stInCNo
         
         'add by nickc 2005/06/07
         stCM10 = "0"
         StrSQLa = "select * from patent where pa01='" & stInCNo(1) & "' and pa02='" & stInCNo(2) & "' and pa03='" & stInCNo(3) & "' and pa04='" & stInCNo(4) & "' "
         CheckOC3
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If AdoRecordSet3.RecordCount <> 0 Then
         
'Modify by Morgan 2006/6/9 國內大陸國外香港時cm10=4
'            If Text1(2) = "1" Then
'               '檢查是否為大陸案
'               'Modify by Morgan 2006/6/
'               If CheckStr(AdoRecordSet3.Fields("pa09")) = "020" Then
'                  stCM10 = "4"
'               End If
'            Else
'               '檢查是否為台灣或是大陸
'               If CheckStr(AdoRecordSet3.Fields("pa09")) = "020" Then
'                  stCM10 = "4"
'               ElseIf CheckStr(AdoRecordSet3.Fields("pa09")) = "000" Then
'                  stCM10 = "0"
'               End If
'            End If

'         Else
'            MsgBox "此案號不存在！", vbInformation
'            Text1(15).SetFocus
'            Text1_GotFocus 15
'            CheckOC3
'            Exit Sub

            'Modify by Morgan 2007/4/26
            'If "" & AdoRecordSet3.Fields("pa09") = "020" And Text1(13) = "013" Then
            '2011/8/23 MODIFY BY SONIA 加判斷專利種類為發明者
            'Modified by Lydia 2016/07/07 改模組
            'If ("" & AdoRecordSet3.Fields("pa09") = "020" Or "" & AdoRecordSet3.Fields("pa09") = "221" Or "" & AdoRecordSet3.Fields("pa09") = "201") And Text1(13) = "013" And "" & AdoRecordSet3.Fields("pa08") = "1" Then
            '   stCM10 = "4"
            'End If
            ''Added by Lydia 2015/09/09 增加澳門發明案與大陸案之關聯
            'If "" & AdoRecordSet3.Fields("pa09") = "020" And Text1(13) = "044" And "" & AdoRecordSet3.Fields("pa08") = "1" Then
            '   stCM10 = "5"
            'End If
            'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
            stCM10 = PUB_GetPcm10(stCM10, "" & AdoRecordSet3.Fields("pa09"), Text1(13), "" & AdoRecordSet3.Fields("pa08"))
            'end 2016/07/07
            
            'Modify by Morgan 2007/3/16 非PCT案才要
            If Text1(14) = "" Then
               'Add by Morgan 2007/1/30 國內案為台灣案且已公告時要彈訊息 -- 郭
               stInPA08 = "" & AdoRecordSet3.Fields("pa08")
               stInPA09 = "" & AdoRecordSet3.Fields("pa09")
               stInPA10 = "" & AdoRecordSet3.Fields("pa10")
               stInPA14 = "" & AdoRecordSet3.Fields("pa14")
               'Modified by Morgan 2012/8/24
               '改判斷所有相同案最早公告者
               'If stInPA09 = "000" And stInPA14 <> "" Then
               '   strPA14Msg = "國內案已於 " & Format(Val(stInPA14) - 19110000, "###/##/##") & " 公告！"
               'End If
               strExc(0) = "SELECT PA14,PA01,PA02,PA03,PA04 FROM casemap,PATENT WHERE cm05='" & stInCNo(1) & "' and cm06='" & stInCNo(2) & "' and cm07='" & stInCNo(3) & "' and cm08='" & stInCNo(4) & "' and cm10='0'" & _
                  " and pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04 AND PA14>0 order by 1"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strPA14Msg = "相同案 " & RsTemp(1) & "-" & RsTemp(2) & "-" & RsTemp(3) & "-" & RsTemp(4) & " 已於 " & Format(RsTemp(0) - 19110000, "##/##/##") & " 公告！"
               End If
               'end 2007/1/30
            End If
            'end 2007/3/16
         End If
         CheckOC3
         
         'edit by nickc 2005/06/07
'         strTxt(intStep) = "INSERT INTO CASEMAP (CM01,CM02,CM03,CM04,CM05,CM06,CM07,CM08,CM10) " & _
            "VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & _
            stInCNo (1) & "','" & stInCNo(2) & "','" & stInCNo(3) & "','" & stInCNo(4) & "','0')"
         
         'Modify by Morgan 2007/12/11 加判斷國內案或申請國家有改的才新增
         bolInsCM = False
         If (Text1(15).Text <> Text1(15).Tag Or Text1(13).Text <> Text1(13).Tag) Then
            strTxt(intStep) = "INSERT INTO CASEMAP (CM01,CM02,CM03,CM04,CM05,CM06,CM07,CM08,CM10) " & _
               "VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & _
               stInCNo(1) & "','" & stInCNo(2) & "','" & stInCNo(3) & "','" & stInCNo(4) & "','" & stCM10 & "')"
        
            cnnConnection.Execute strTxt(intStep), intI
            intStep = intStep + 1
            bolInsCM = True
            
            'Add by Morgan 2011/5/31
            '若國外案尚無繪圖人員時設定為國內案繪圖人員
            'Modified by Morgan 2013/7/30 不必考慮是否原先有繪圖人員一律更新--瓊玉 (Ex.P105970 分案後又改工程師且關聯後建導致與國內案的繪圖人員不同)
            'If cp(29) = "" Then
            'Modified by Lydia 2021/11/10 衍生的香港或澳門案分案先抓大陸案之中英文Title
            'If Text1(1) <> "110" Then  'Added by Morgan 2015/6/26 排除標準專利記錄請求110 --品薇  (111非新案案件性質不必控制)
            If Text1(1) <> "110" And Not (((Text1(13) = "013" And Text1(1) = "110") Or (Text1(13) = "044" And Text1(1) = "101")) And m_bolFMP = True) Then
               'Modified by Morgan 2016/8/5 繪圖人員離職時不更新--游經理
               strExc(0) = "select cp29 from caseprogress,staff" & _
                  " where cp01='" & stInCNo(1) & "' and cp02='" & stInCNo(2) & "'" & _
                  " and cp03='" & stInCNo(3) & "' and cp04='" & stInCNo(4) & "'" & _
                  " and cp10 in ('" & Replace(NewCasePtyList, ",", "','") & "') and cp29 is not null and st01(+)=cp29 and st04='1'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  'Modified by Morgan 2013/7/30
                  'strSql = "update caseprogress a set cp29='" & RsTemp(0) & "'" & _
                     " where cp09='" & cp(9) & "' and cp29 is null"
                  strSql = "update caseprogress a set cp29='" & RsTemp(0) & "'" & _
                     " where cp09='" & cp(9) & "'"
                  'end 2013/7/30
                  cnnConnection.Execute strSql, intI
               End If
            'Added by Lydia 2018/06/27 待FMP案的香港案和母案大陸案建立關連後，香港案的專利名稱直接抓母案大陸案的專利名稱
            'Modified by Lydia 2021/11/10 衍生的香港或澳門案分案先抓大陸案之中英文Title。
            'ElseIf Text1(13) = "013" And m_bolFMP = True Then
            ElseIf ((Text1(13) = "013" And Text1(1) = "110") Or (Text1(13) = "044" And Text1(1) = "101")) And m_bolFMP = True Then
                strExc(0) = "select pa05,pa06,pa07 from patent where pa09='020' and pa01='" & stInCNo(1) & "' and pa02='" & stInCNo(2) & "' and pa03='" & stInCNo(3) & "' and pa04='" & stInCNo(4) & "' "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                     strSql = "Update Patent Set PA05=" & CNULL(ChgSQL("" & RsTemp.Fields("pa05"))) & ",PA06=" & CNULL(ChgSQL("" & RsTemp.Fields("pa06"))) & ", PA07=" & CNULL(ChgSQL("" & RsTemp.Fields("pa07"))) & " " & _
                                 "where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "' "
                     Pub_SeekTbLog strSql
                     cnnConnection.Execute strSql, intI
                End If
            'end 2018/06/27
            End If 'Added by Morgan 2015/6/26
            'End If
         End If
         'end 2007/12/11
         
         'Added by Lydia 2022/06/08 大陸衍生澳門案在分案確定時，請帶大陸案之發明人、代表人至衍生澳門案之發明人、代表人資料、個案地址
         If Text1(13) = "044" And Text1(1) = "101" And m_bolFMP = True Then
             '香港案一直都是內專處理，故不管
             strExc(0) = Text1(15)
             Call ChgCaseNo(strExc(0), strExc)
             Call PUB_FCPCopyDataToCase(pa(), strExc())
             '刪除優先權資料
             If ClsPDDeletePriority(pa) = True Then
             End If
         End If
         'end 2022/06/08
         
         'Modify by Morgan 2007/3/16 非PCT案,香港案才要
         'Modify by Lydia 2015/09/09 +澳門案
         'If Text1(14) = "" And stCM10 <> "4" Then
         If Text1(14) = "" And Not (stCM10 = "4" Or stCM10 = "5") Then
            'Add by Morgan 2005/1/25 若國內案已發文則文件齊備日上系統日
            'Remove by Lydia 2016/07/07 改成模組PUB_SavePtoUpd1 'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
            'strExc(0) = "SELECT CP09 FROM CASEPROGRESS" & _
               " WHERE CP01='" & stInCNo(1) & "' AND CP02='" & stInCNo(2) & "'" & _
               " AND CP03='" & stInCNo(3) & "' AND CP04='" & stInCNo(4) & "'" & _
               " AND CP10 in (" & CaseMapIn & ") AND CP27>0"
               
            'intI = 1
            'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            'If intI = 1 Then
            '   strTxt(intStep) = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & _
            '      " WHERE EP02='" & Label3(8) & "' AND EP06 IS NULL"
            '   cnnConnection.Execute strTxt(intStep), intI
            '   intStep = intStep + 1
            '   'Add by Morgan 2005/2/15 重新計算承辦期限
            '   'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
            '   'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP48=" & PUB_PomoteDate(Label3(8), TransDate(Text1(4), 2)) & _
            '   '   " WHERE CP09='" & Label3(8) & "'"
            '   'cnnConnection.Execute strTxt(intStep)
            '   If intI = 1 Then '齊備日有更新才做
            '      If m_bolFMP Or PUB_IfSetCP48() Then 'Add by Morgan 2010/10/1
            '         strCP48 = Pub_GetHandleDay(pa(1), Text1(13), Text1(1), , TransDate(Text1(4), 2), Label3(8))
            '         If strCP48 <> "" Then
            '            strTxt(intStep) = "UPDATE CASEPROGRESS SET CP48=" & strCP48 & _
            '               " WHERE CP09='" & Label3(8) & "'"
            '            cnnConnection.Execute strTxt(intStep), intI
            '            intStep = intStep + 1
            '         End If
            '      End If 'Add by Morgan 2010/10/1
            '   End If
            '   'end 2007/10/15
            'End If
            ''2005/1/25

            ''Add by Morgan 2005/3/15 若國外案為大陸外觀設計且無繪圖人員時帶國內案繪圖人員
            ''2005/4/19 MODIFY BY SONIA 再控制國外案為大陸外觀設計有國內案時,草圖及墨圖都不計件
            ''2009/2/2 modify by sonia 瓊玉說不應限制設計P-090127,只要是大陸新申請案無繪圖人員都要帶國內案繪圖人員且草墨圖都不計件
            ''If Text1(13) = "020" And Text1(1) = "103" Then
            ''Remove by Lydia 2016/07/07 改成模組PUB_SavePtoUpd1 'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
            'If Text1(13) = "020" And InStr(CaseMapIn, Text1(1)) > 0 Then
            '   Call PUB_UpdateEP13(Label3(8), stInCNo())
            'End If
            ''2005/3/15 END
         
            'Add by Morgan 2007/1/30 國內案為台灣案且未公告時抓預估公告日或公開日更新大陸案期限
            'Modified by Morgan 2017/5/8 排除有主張優先權者--郭
            'If stInPA09 = "000" And stInPA14 = "" Then
            'Modified by Morgan 2017/10/2 新案翻譯201不必更新 Ex.P116855 --玲玲
            If stInPA09 = "000" And stInPA14 = "" And Text1(1) <> "201" Then
               strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='106' and cp57 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
            'end 2017/5/8
                  stInPA14 = PUB_GetPrePA14(stInCNo)
                  strExc(3) = "期限來源:" & Right("  " & stInCNo(1), 3) & "-" & stInCNo(2) & "-" & stInCNo(3) & "-" & stInCNo(4) & "(公告日/預定公告日);" 'Added by Morgan 2016/12/26
                  '發明公開
                  If stInPA08 = "1" Then
                     strExc(0) = "SELECT PD05 FROM PRIDATE WHERE PD01='" & stInCNo(1) & "' AND PD02='" & stInCNo(2) & "' AND PD03='" & stInCNo(3) & "' AND PD04='" & stInCNo(4) & "' ORDER BY PD05 ASC"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strExc(1) = "" & RsTemp("PD05")
                     Else
                        strExc(1) = stInPA10
                     End If
                     '法定期限=預估公開日=申請日(最早優先權日)+18個月
                     strExc(2) = CompDate(1, 18, TransDate(strExc(1), 2))
                     If stInPA14 = "" Or Val(strExc(2)) < Val(stInPA14) Then
                        stInPA14 = strExc(2)
                        strExc(3) = "期限來源:" & Right("  " & stInCNo(1), 3) & "-" & stInCNo(2) & "-" & stInCNo(3) & "-" & stInCNo(4) & "(預定公開日);" 'Added by Morgan 2016/12/26
                     End If
                     'end 2007/1/30
                  End If
                  If stInPA14 <> "" Then
                     'Add by Morgan 2008/8/15 若期限已過則提醒
                     If stInPA14 < strSrvDate(1) Then
                        strPA14Msg = "國內案之預定公告日/公開日( " & Format(Val(stInPA14) - 19110000, "###/##/##") & ")已過！"
                     Else
                        '本所期限=法定期限-10天
                        strExc(7) = stInPA14
                        'Added by Lydia 2025/10/29
                        If strSrvDate(1) >= 內專本所約定期限啟用日 Then
                           strExc(6) = PUB_GetPOurDeadline(strExc(7), pa(9), , pa(1), Text1(1))
                        Else
                        'end 2025/10/29
                           strExc(6) = PUB_GetWorkDay1(CompDate(2, -10, strExc(7)), True)
                        End If 'Added by Lydia 2025/10/29
                        If strExc(6) < strSrvDate(1) Then strExc(6) = strSrvDate(1)
                        'Modified by Morgan 2016/12/26 +CP64
                        strSql = "UPDATE CASEPROGRESS SET CP06=" & Val(strExc(6)) & ",CP07=" & Val(strExc(7)) & ",CP64=SUBSTR(CP64,1,INSTR(CP64,'期限來源:')-1)||'" & strExc(3) & "'||SUBSTR(CP64,INSTR(CP64,';',instr(CP64,'期限來源:'))+1) WHERE CP09='" & cp(9) & "' AND CP27 IS NULL AND ( CP07 IS NULL OR CP07>" & Val(strExc(7)) & " )"
                        adoTaie.Execute strSql, intI
                     End If
                  End If
               End If 'Added by Morgan 2017/5/8
            End If
            'end 2007/1/30
         End If
         'end 2007/3/16
      End If
      
      '若更改承辦人欄位, 則更新ENG核稿人
      If Me.Text1(0).Text <> Me.Text1(0).Tag Then
         'Modify By Sindy 2016/5/24
         '無完稿日時,清掉核稿人及判發人
         strExc(0) = "SELECT ep02 FROM EngineerProgress WHERE ep02='" & Label3(8) & "' and ep09 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "update EngineerProgress set ep04=null,ep40=null WHERE ep02='" & Label3(8) & "'"
            cnnConnection.Execute strSql
         Else
            '檢查是否有送判或判發
            If PUB_ChkEmpFlowExists(Label3(8), EMP_送判) = False And _
               PUB_ChkEmpFlowExists(Label3(8), EMP_判發) = False Then
               strSql = "update EngineerProgress set ep40=null WHERE ep02='" & Label3(8) & "'"
               cnnConnection.Execute strSql
            End If
         End If
'          'edit by nickc 2007/08/16 修正更新欄位
'          'strTxt(intStep) = "UPDATE ENGINEERPROGRESS SET EP03=(" & _
'          "SELECT PP04 FROM CASEPROGRESS,PROMOTERPROOFREADER WHERE CP09='" & Label3(8) & _
'          "' AND CP01=PP01(+) AND '" & Me.Text1(0).Text & "'=PP02(+) AND CP10=PP03(+)) WHERE EP02='" & Label3(8) & "'"
'          strTxt(intStep) = "UPDATE ENGINEERPROGRESS SET EP04=(" & _
'          "SELECT PP04 FROM CASEPROGRESS,PROMOTERPROOFREADER WHERE CP09='" & Label3(8) & _
'          "' AND CP01=PP01(+) AND '" & Me.Text1(0).Text & "'=PP02(+) AND CP10=PP03(+)) WHERE EP02='" & Label3(8) & "'"
'         cnnConnection.Execute strTxt(intStep)
'          intStep = intStep + 1
         '2016/5/24 END
      End If
      'Modified by Lydia 2015/09/09 改共用模組- 依大陸案更新香港澳門期限
      'add by nickc 2005/07/05 以下香港案標準專利紀錄請求做
'      If Text1(13) = "013" And Text1(1) = "110" Then
'         Dim tmpCp07 As String
'         Dim tmpCp06 As String
'         Dim tmpCP48 As String
'         'Modify by Morgan 2007/9/7 要抓香港的工作天數CF02='013',否則抓不到資料
'         'strExc(0) = "SELECT PA.*,NVL(CF04,0) CF04 FROM patent PA,Casefee WHERE pa01='" & strExc(1) & "' AND pa02='" & strExc(2) & _
'            "' AND pa03='" & strExc(3) & "' AND pa04='" & strExc(4) & "' and PA01=cf01(+) and pa09=cf02(+) and '110'=cf03 "
'         'Modified by Lydia 2015/09/09 +大陸案是否公開(PA09,PA08,PA12,PA16||PA21||PA108||PA136)
'         'strExc(0) = "SELECT PA12,NVL(CF04,0) CF04 FROM patent,Casefee WHERE pa01='" & stInCNo(1) & "' AND pa02='" & stInCNo(2) & "' AND pa03='" & stInCNo(3) & "' AND pa04='" & stInCNo(4) & "' and cf01(+)=pa01 and cf02(+)='013' and cf03(+)='110'"
'         'end 2007/9/7
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If Not IsNull(RsTemp.Fields("PA12").Value) Then
'               strTxt(intStep) = "UPDATE engineerPROGRESS SET ep06=" & strSrvDate(1) & " WHERE ep02='" & Label3(8) & "' "
'               cnnConnection.Execute strTxt(intStep)
'               intStep = intStep + 1
'
'               tmpCp07 = CompDate(1, 6, RsTemp.Fields("PA12").Value)
'               tmpCp06 = PUB_GetWorkDay1(CompDate(2, -5, CompDate(1, -1, tmpCp07)), True)
'
'               'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
'               'tmpCP48 = CompWorkDay(RsTemp.Fields("CF04").Value, strSrvDate(1), 0)
'               'If Val(tmpCP48) > Val(tmpCp06) Then tmpCP48 = "" & Val(tmpCp06)
'
'               'Modify by Morgan 2010/10/19
'               'tmpCP48 = Pub_GetHandleDay(pa(1), "013", "110", , tmpCp06, Label3(8))
'               If m_bolFMP Or PUB_IfSetCP48() Then
'                  tmpCP48 = Pub_GetHandleDay(pa(1), "013", "110", , tmpCp06, Label3(8))
'               Else
'                  tmpCP48 = "CP48"
'               End If
'               'end 2010/10/19
'
'               'end 2007/10/11
'
'               strTxt(intStep) = "UPDATE casePROGRESS SET cp06=" & tmpCp06 & ",cp07=" & tmpCp07 & ",cp48=" & tmpCP48 & " WHERE CP09='" & Label3(8) & "' "
'               cnnConnection.Execute strTxt(intStep)
'                intStep = intStep + 1
'            End If
'         End If
'
'         'Add by Morgan 2009/11/6
'         '更新大陸案的標準專利紀錄請求期限(NP)為續辦
'         strSql = "Update nextprogress set np06='Y' where np02='" & stInCNo(1) & "' and np03='" & stInCNo(2) & "' and np04='" & stInCNo(3) & "' and np05='" & stInCNo(4) & "' and np06 is null and np07='110'"
'         cnnConnection.Execute strSql, intI
'         'end 2009/11/6
'      End If

      'Remove by Lydia 2016/07/07 改成模組PUB_SavePtoUpd1
      ''依大陸案更新香港澳門期限
      'If Text1(15) <> "" Then
      '   '大陸-香港案
      '   If Text1(13) = "013" And Text1(1) = "110" Then
      '      Call PUB_UpdCP07by020(stInCNo, m_bolFMP, "4", strSrvDate(1))
      '      '更新大陸案的標準專利紀錄請求期限(NP)為續辦
      '      strSql = "Update nextprogress set np06='Y' where np02='" & stInCNo(1) & "' and np03='" & stInCNo(2) & "' and np04='" & stInCNo(3) & "' and np05='" & stInCNo(4) & "' and np06 is null and np07='110'"
      '      cnnConnection.Execute strSql, intI
      '   End If
      '   '大陸-澳門案
      '   If Text1(13) = "044" And Text1(1) = "101" Then
      '      Call PUB_UpdCP07by020(stInCNo, m_bolFMP, "5")
      '   End If
      'End If
      'Added by Lydia 2016/07/07 內專分案(工程師主管分案frm060117)-更新關聯案
      'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
      Call PUB_SavePtoUpd1(m_bolFMP, pa, Label3(8), Text1(1).Text, Text1(4).Text, Text1(14).Text, Text1(15), stCM10)
'end 2015/09/09

      Select Case Me.Text1(1).Text
         Case "601" '領證
            '法定期限
            If Val(Me.Text1(5).Text) > 0 Then
               '小於收文日
               If Val(Me.Text1(5).Text) < Val(Me.Text1(12).Text) Then
                  '抓案件收費表的下次期限
                  If GetCF12(pa(1), pa(9), Me.Text1(1).Text) <> 0 Then
                     m_strCP07 = DBDATE(CompDate(2, (GetCF12(pa(1), pa(9), Me.Text1(1).Text)), Format(m_strCP07)))
                  Else
                     m_strCP07 = DBDATE(CompDate(1, (GetCF28(pa(1), pa(9), Me.Text1(1).Text)), Format(m_strCP07)))
                     'add by sonia 2018/2/21 原法定若為月底則延期後也要是月底 FCP-042474
                     PUB_LastDayConvert DBDATE(Me.Text1(5).Tag), m_strCP07
                  End If
                  '本所期限=系統日
                  strSql = "UPDATE CASEPROGRESS SET CP06=" & strSrvDate(1) & ",CP07=" & CNULL(m_strCP07) & _
                           " WHERE CP09='" & Label3(8) & "'"
                  cnnConnection.Execute strSql
               End If
            End If
            
         Case "605" '繳年費
            If pa(72) <> "" Then 'Added by Morgan 2018/10/1 香港案第1年無法計算 Ex:P-106893
               'modify by sonia 2018/2/21 重算原法定期限,以免下一程序已改為逾期6個月的期限
               'If Val(Me.Text1(5).Text) < Val(Me.Text1(12).Text) Then
               m_strCP07_1 = PUB_GetNextFeeDate(pa)
               If DBDATE(Text1(12)) > Val(m_strCP07_1) Then
               'end 2018/2/21
                  If GetCF12(pa(1), pa(9), Me.Text1(1).Text) <> 0 Then
                     m_strCP07 = DBDATE(CompDate(2, (GetCF12(pa(1), pa(9), Me.Text1(1).Text)), Format(m_strCP07_1)))
                  Else
                     m_strCP07 = DBDATE(CompDate(1, (GetCF28(pa(1), pa(9), Me.Text1(1).Text)), Format(m_strCP07_1)))
                     'add by sonia 2018/2/21 原法定若為月底則延期後也要是月底 FCP-042474
                     PUB_LastDayConvert m_strCP07_1, m_strCP07
                  End If
                  'Added by Morgan 2012/10/15
                  '102新法:年費過期6個月(跨102年)自動設法限=原法限+18月,所限=系統日
                  If pa(9) = "000" And Val(m_strCP07) >= 20130101 And m_strCP07 < DBDATE(Text1(12)) Then
                     'modify by sonia 2018/2/21 改以計算出來的原始法定期限計算,並考慮月底問題
                     'm_strCP07 = CompDate(1, 18, cp(7))
                     m_strCP07 = CompDate(1, 18, m_strCP07_1)
                     PUB_LastDayConvert m_strCP07_1, m_strCP07
                     'end 2018/2/21
                  End If
                  'end 2012/10/15
                  
                  strSql = "UPDATE CASEPROGRESS SET CP06=" & strSrvDate(1) & ",CP07=" & CNULL(m_strCP07) & _
                           " WHERE CP09='" & Label3(8) & "'"
                  cnnConnection.Execute strSql
               End If
            End If
            
      End Select

      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.SavePriority(pA, strPriority(1), strPriority(2), strPriority(3)) Then
      'Modify by Morgan 2007/4/24 加strPriority(4)
      'Modify by Amy 2014/06/10 +strPriority(5)
      'Modify by Amy 2023/01/05 strPriority原陣列,改變數
      If ClsPDSavePriority(pa, strPrity1, strPrity2, strPrity3, strPrity4, strPrity5) Then
         strFirstPriDate = PUB_GetFirstPriDate(cp) 'Add by Morgan 2006/5/12 重讀最早優先權日
      End If
   
      '91.11.3 ADD BY SONIA 更新主張優先權期限
      'Modify by Morgan 2004/9/22 加 主張國內優先權 121
      If Me.Text1(1).Text = 主張優先權 Or Me.Text1(1).Text = "121" Then
'Modify by Morgan 2010/8/9 改呼叫共用函式
'         'Modify by Morgan 2006/7/6 判斷非PCT案才要
'         '(PCT案的主張優先權不需掛期限，收文是為了要輸優先權資料但不需提出申請，且國外部有請款需要。)
'         'If strFirstPriDate <> "" Then
'         If strFirstPriDate <> "" And pa(46) <> "Y" Then
'            'Add by Morgan 2009/8/10
'            '分割案不必控管
'            If PUB_ChkCPExist(cp, "307") = False Then
'            'end 2009/8/10
'
'               'Modify by Moragn 2007/4/25
'               '主張一個以上優先權的固定為最早優先權日+6個月,另外主張或被主張的若為設計時也是+6個月,剩下的則為+12個月
'               ''發明或新型之主張優先權法定期限為最早優先權日+12月, 設計為最早優先權日+6月
'               'If Text1(2) = "3" Then
'               '   strExc(10) = 6
'               'Modify by Morgan 2007/5/9 加控制主張多個優先權且有設計時才為早優先權日+6個月
'               'If InStr(strPriority(1), "，") > 0 Then
'               '   strExc(10) = 6
'               'ElseIf Text1(2) = "3" Or InStr(strPriority(4), "3") > 0 Then
'               If Text1(2) = "3" Or InStr(strPriority(4), "3") > 0 Then
'               'end 2007/5/9
'                  strExc(10) = 6
'               'end 2007/4/25
'               Else
'                  strExc(10) = 12
'               End If
'               WorkDate2 = DBDATE(CompDate(1, strExc(10), Format(strFirstPriDate)))
'               '國內案本所期限=法定期限-2天, 非國內案本所期限=法定期限-7天
'               'Modify by Morgan 2007/3/2
'               'If Text1(13) <> 台灣國家代號 Then
'               '   WorkDate1 = DBDATE(CompDate(2, -7, Format(WorkDate2)))
'               If Text1(13) = 台灣國家代號 Then
'                  WorkDate1 = DBDATE(CompDate(2, -2, Format(WorkDate2)))
'               'end 2007/3/2
'               Else
'                  'Add by Morgan 2010/1/22
'                  If m_bolFMP Then
'                     WorkDate1 = PUB_GetDeadLine(DBDATE(cp(5)), WorkDate2, 2)
'                  Else
'                     WorkDate1 = DBDATE(CompDate(2, -7, Format(WorkDate2)))
'                  End If
'               End If
'               '若本所期限非工作天則抓最近的工作天
'               WorkDate1 = PUB_GetWorkDay1(WorkDate1, True)
'
'               '2005/12/28 ADD BY SONIA 同時更新新案之期限(發現CFP程式於92.5.27已加入)
'
'               strExc(1) = WorkDate2
'               strExc(2) = WorkDate1
'
'               'Remove by Morgan 2010/1/22 期限同優先權
'               ''Add by Morgan 2009/11/4
'               ''FMP期限抓設定
'               'If m_bolFMP Then
'               '   strExc(1) = PUB_GetDeadLine(DBDATE(cp(5)), WorkDate2, 1)
'               '   strExc(2) = PUB_GetDeadLine(DBDATE(cp(5)), WorkDate2, 2)
'               'End If
'               ''end 2009/11/4
'
'               strSql = "UPDATE CASEPROGRESS SET CP06=" & CNULL(strExc(2)) & ",CP07=" & CNULL(strExc(1)) & _
'                        " WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP31='Y' and CP57 IS NULL"
'               strSql = strSql & " and (CP07 IS NULL OR CP07>" & strExc(1) & ")" 'Add by Morgan 2010/3/17 期限大的才要更新(若有主張新穎性優惠時也會有期限所以要比較)
'               cnnConnection.Execute strSql, intI
'               '2005/12/28 END
'
'               'Modify by Morgan 2010/3/17 改用新案期限更新
'               'strSql = "UPDATE CASEPROGRESS SET CP06=" & CNULL(WorkDate1) & ",CP07=" & CNULL(WorkDate2) & _
'                        " WHERE CP09='" & cp(9) & "'"
'               strSql = "UPDATE CASEPROGRESS a SET (CP06,CP07)=(select b.CP06,b.CP07 from caseprogress b" & _
'                  " WHERE b.CP01=a.CP01 AND b.CP02=a.CP02 AND b.CP03=a.CP03 AND b.CP04=a.CP04" & _
'                  " AND b.CP31='Y' and b.cp57 is null) WHERE CP09='" & cp(9) & "'"
'               'END 2010/3/17
'               cnnConnection.Execute strSql, intI
'
'               'Add by Morgan 2010/3/17
'               '同時更新新穎性優惠期限
'               strSql = "UPDATE CASEPROGRESS a SET (CP06,CP07)=(select b.CP06,b.CP07 from caseprogress b" & _
'                  " WHERE b.CP01=a.CP01 AND b.CP02=a.CP02 AND b.CP03=a.CP03 AND b.CP04=a.CP04" & _
'                  " AND b.CP31='Y' and b.cp57 is null) " & _
'                  " WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "'" & _
'                  " AND cp10='123' and cp57 is null"
'               cnnConnection.Execute strSql, intI
'            End If
'         End If
         'Modified by Lydia 2023/03/25 +pSaveMsg
         PUB_UpdCfpDate1 pa(1), pa(2), pa(3), pa(4), , pSaveMsg
      End If
      '91.11.3 END
      
      'Added by Morgan 2012/7/3
      '回復優先權主張法定期限=最早優先權日+16個月(發明,新型),+10個月(設計)
      If Text1(13) = "000" And Text1(1) = "124" And cp(27) = "" Then
         strFirstPriDate = PUB_GetFirstPriDate(pa, m_strPriType)
         If strFirstPriDate <> "" Then
            strExc(2) = ""
            If Text1(2) = "3" Or m_strPriType = "2" Then
               strExc(3) = CompDate(1, 10, strFirstPriDate)
            Else
               strExc(3) = CompDate(1, 16, strFirstPriDate)
            End If
            stDate(0) = ""
            stDate(1) = pa(1)
            stDate(2) = pa(9)
            stDate(3) = strExc(3)
            GetCtrlDT stDate
            '本所期限
            strExc(2) = PUB_GetWorkDay1(stDate(0), True)
            strSql = "update caseprogress set cp06=" & CNULL(strExc(2), True) & ",cp07=" & CNULL(strExc(3), True) & " where cp09='" & cp(9) & "' and cp27 is null"
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2012/7/3
      
      'Add by Morgan 2010/3/17
      '新穎性優惠期要同時更新優先權及申請程序的期限
      If Text1(1) = "123" And Text1(5) <> "" Then
         strExc(1) = DBDATE(Text1(5))
         strExc(2) = PUB_GetWorkDay1(Text1(4), True)
         '新案
         strSql = "UPDATE CASEPROGRESS SET CP06=" & strExc(2) & ",CP07=" & strExc(1) & _
                  " WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "'" & _
                  " and cp57 is null AND CP31='Y' and (cp07 is null or cp07>" & strExc(1) & ")"
         cnnConnection.Execute strSql, intI
         
         '主張優先權
         strSql = "UPDATE CASEPROGRESS a SET (CP06,CP07)=(select b.CP06,b.CP07 from caseprogress b" & _
            " WHERE b.CP01=a.CP01 AND b.CP02=a.CP02 AND b.CP03=a.CP03 AND b.CP04=a.CP04" & _
            " AND b.CP31='Y' and b.cp57 is null) " & _
            " WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "'" & _
            " and cp57 is null AND cp10 in ('106','121')"
         cnnConnection.Execute strSql, intI
         
         '本程序
         strSql = "UPDATE CASEPROGRESS a SET (CP06,CP07)=(select b.CP06,b.CP07 from caseprogress b" & _
            " WHERE b.CP01=a.CP01 AND b.CP02=a.CP02 AND b.CP03=a.CP03 AND b.CP04=a.CP04" & _
            " AND b.CP31='Y' and b.cp57 is null) WHERE CP09='" & cp(9) & "'"
         cnnConnection.Execute strSql, intI
      End If
      
      'Added by Lydia 2015/07/21 台灣P案主張大陸優先權之保密審查控管
      If cp(1) = "P" And Text1(13) = "000" And Text1(1) = 主張優先權 Then
          '台灣案若主張大陸優先權且大陸案非本所辦理,請產生下一程序控管「保密審查」(430)(本所期限設定一個月),並發E-MAIL通知業務同仁
          'Added by Lydia 2015/10/26 +排除大陸外觀設計案(PD08=3)
          'Modified by Morgan 2019/1/21 +有申請日條件(優先權號為本所案號且無優先權日者為之未送件本所案) Ex:P-121949
          strExc(0) = "select a.*,b.pa10,b.pa11 from pridate a,patent b where pd01='" & cp(1) & "' AND pd02='" & cp(2) & "' AND pd03='" & cp(3) & "' AND pd04='" & cp(4) & "' " & _
                      "and pd07='020' and pd06=pa11(+) and pd07=pa09(+) and pd08=pa08(+) and (pd08 is null or pd08<> 3) and pd05>0"
          intI = 1
          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
          If intI = 1 Then
             RsTemp.MoveFirst
             Do While Not RsTemp.EOF
                 strExc(1) = Trim("" & RsTemp.Fields("pa11"))
                 If Len(strExc(1)) = 0 Then
                    '重新分案不變更保密審查
                    strExc(0) = "select * from nextprogress where np02='" & cp(1) & "' AND np03='" & cp(2) & "' AND np04='" & cp(3) & "' AND np05='" & cp(4) & "' and np07='430' "
                    
                    Set rsA = New ADODB.Recordset
                    rsA.CursorLocation = adUseClient
                    rsA.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
                    If rsA.RecordCount = 0 Then
                       strExc(0) = CompDate(1, 1, strSrvDate(1))
                       strExc(8) = PUB_GetWorkDay1(strExc(0), True) '所限
                       
                        'Added by Morgan 2020/9/25
                        '若晚於新申請案本所期限往前7個工作日(若早於系統日則設定為系統日)則設較早者
                        strExc(0) = "select cp06 from caseprogress where cp01='" & cp(1) & "' AND cp02='" & cp(2) & "' AND cp03='" & cp(3) & "' AND cp04='" & cp(4) & "' AND CP31='Y' and cp57 is null and cp06>0"
                        intI = 1
                        Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                            strExc(1) = CompWorkDay(7, rsA("cp06"), 1)
                            If strExc(1) < strSrvDate(1) Then
                               strExc(1) = strSrvDate(1)
                            End If
                            
                            If strExc(8) > strExc(1) Then
                              strExc(8) = strExc(1)
                            End If
                        End If
                        rsA.Close
                        'end 2020/9/25
    
                       strSql = "declare intMax number;begin select max(np22)+1 into intMax from nextprogress;IF intMax IS NULL THEN intMax:=1; END IF;"
                       strSql = strSql & "Insert Into NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP10,NP22) " & _
                            " Values ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "','430'," & strExc(8) & ",'" & cp(13) & "',intMax); "
                       strSql = strSql & " end;"
                       cnnConnection.Execute strSql, intI
                       
                       'Added by Morgan 2025/1/24
                       If strSrvDate(1) >= P業務區劃分啟用日 Then
                        strExc(2) = PUB_GetPHandler(cp(1) & cp(2) & cp(3) & cp(4))
                       Else
                       'end 2025/1/24
                           strExc(2) = GetStaffName(Pub_GetSpecMan("A1")) 'P案分案人(P案程序)
                           
                        End If 'Added by Morgan 2025/1/24
                     
                       'Modified by Lydia 2015/10/20 改信件內容
'                       strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " 本案主張大陸優先權，且大陸案非本所辦理請確認是否發明人為大陸人，如「是」大陸案必須申請「保密審查」，請告知客戶，且台灣案將於「保密審查」核准後方可辦理，並請儘速回覆" & strExc(2) & "。"
'                       strExc(3) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " 主張大陸優先權之保密審查控管,詳情請見內容"
                       strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " 本案主張中國大陸優先權，惟該大陸案並非本所辦理，若該發明是在中國大陸完成研發，則大陸案必須申請「保密審查」，且台灣案須於「保密審查」核准後方可提申，請告知客戶並請儘速回覆" & strExc(2) & "。"
                       strExc(3) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " 主張中國大陸優先權之保密審查控管,詳情請見內容"
                       strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                            " values ('" & strUserNum & "','" & cp(13) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                            ",'" & ChgSQL(strExc(3)) & "','" & ChgSQL(strExc(1)) & "')"
                       cnnConnection.Execute strSql, intI
                    End If
                 End If
               RsTemp.MoveNext
             Loop
          End If
      End If
      
      'Remove by Morgan 2011/3/31
      '100/4/1 以後多國案草墨圖改要計件(草齊日=墨齊日)
      ''Add by Morgan 2004/12/13 若有收文日較早(<)的國外案時草圖是否計件上'N'
      ''2008/12/4 modify by sonia 加入繪圖未分案確認條件 CFP-021680
      ''If PUB_IfCaseMapExist(cp, IIf(pa(9) = "000", "0", "1")) = True Then
      ''Modify by Morgan 2010/12/27 +控制新申請案的案件性質才要做質--秀玲
      'If InStr(NewCasePtyList, Text1(1)) > 0 Then
      '   If PUB_IfCaseMapExist(cp, IIf(pa(9) = "000", "0", "1")) = True And cp(107) = "" Then
      '      strSql = "Update EngineerProgress SET EP20='N' Where EP02='" & cp(9) & "'"
      '      cnnConnection.Execute strSql
      '   End If
      'End If
      ''2004/12/3 end
      'end 2011/3/31
      
      'Add by Morgan 2005/3/4 更新計件值加乘註記
      'Call PUB_UpdateCaseValue(cp(9))
      '第一次分案
      If m_stCP98 = "" Then
         strSql = "UPDATE CASEPROGRESS SET CP98=" & txtCP98 & ",CP99='" & txtCP99 & "',CP101=" & cp(101) & ",CP104=" & cp(104) & _
            " WHERE CP09='" & cp(9) & "'"
         cnnConnection.Execute strSql
      ElseIf (cp(98) <> txtCP98 Or cp(99) <> txtCP99) Then
         '若加乘註記資料有異動時寫LOG
         Call PUB_InsFlagStory(cp(9), "1", cp(98), txtCP98, txtCP99)
         strSql = "UPDATE CASEPROGRESS SET CP98=" & txtCP98 & ",CP99='" & txtCP99 & "' WHERE CP09='" & cp(9) & "'"
         cnnConnection.Execute strSql
      End If
      '2005/3/4 end
      
      'Add by Morgan 2006/4/25 發明改請判斷若未公開則新增B類其他,並於備註說明函知IPO發明暫不公開
      'Modified by Morgan 2024/10/22 備註內容修改 "請函知智慧局發明暫不公開"-->"請向智慧局確認本發明案之公開日期，若是早於審查意見之期限，則須修改本案之期限。"
      m_bol30xMail = False
      m_bol30xMailDesc = Empty
      If Text1(1) = "302" Or Text1(1) = "303" Then
         If Text1(0) <> "" Then
            '原先為發明案且未公開
            If Text1(2).Tag = "1" And pa(12) = "" Then
               m_bol30xMail = True
               m_bol30xMailDesc = "原發明案【" & pa(1) & pa(2) & pa(3) & pa(4) & "】未公開且已收文改請" & IIf(Text1(1) = "302", "新型", "設計") & "，請向智慧局確認本發明案之公開日期，若是早於審查意見之期限，則須修改本案之期限！"
               strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
                  "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP64) VALUES " & _
                  "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & _
                  ",'" & AutoNo("B", 6) & "','" & 其他 & "','90'," & CNULL(cp(12)) & "," & CNULL(cp(13)) & _
                  ",'" & Text1(0) & "','N','N','N','" & cp(9) & "','請向智慧局確認本發明案之公開日期，若是早於審查意見之期限，則須修改本案之期限。') "
               cnnConnection.Execute strSql
            End If
         End If
         'Add by Morgan 2006/9/7
         '發明改請時檢查下一程序有實審未續辦時上"N"並彈訊息,若已收文則發Mail通知智權人員銷案
         If Text1(2).Tag = "1" Then
            strSql = "UPDATE NEXTPROGRESS SET NP06='N',NP15=NP15||';因改請取消實審期限' WHERE NP02='" & cp(1) & "' AND NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP07='416' AND NP06 IS NULL"
            adoTaie.Execute strSql, intI
            If intI > 0 Then
               bol416Msg = True
            Else
               strExc(0) = "SELECT * FROM CASEPROGRESS WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP10='416' AND CP27 IS NULL AND CP57 IS NULL"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  bol416Mail = True
               End If
            End If
         End If
      End If
      '2006/4/25 end
      
      'Add by Morgan 2006/9/8
      '大陸PCT案恢復權利期限要回寫申請案
      'Modify by Morgan 2009/11/6 改大陸案的恢復權利都回寫相關總收文號
      'If Text1(13) = "020" And Text1(14) <> "" And Text1(1) = "414" Then
      '   strSQL = "Update Caseprogress Set CP06=" & CNULL(TransDate(Text1(4), 2)) & ",CP07=" & CNULL(TransDate(Text1(5), 2)) & _
            " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'" & _
            " and cp10 in ('101','102') and cp27 is null and cp57 is null"
      If Text1(13) = "020" And Text1(1) = "414" And Text1(6) <> "" And Text1(4) <> "" And Text1(5) <> "" Then
         strSql = "Update Caseprogress Set CP06=" & CNULL(TransDate(Text1(4), 2)) & ",CP07=" & CNULL(TransDate(Text1(5), 2)) & _
            " where cp09='" & Text1(6) & "' and cp27 is null"
      'end 2009/11/6
         cnnConnection.Execute strSql, intI
      End If
      
      '閉卷恢復年費管制
      If m_str605NP22 <> "" Then
         'Modify by Morgan 2006/12/14 NP11,NP12不必清除
         strSql = "Update Nextprogress set NP06=null" & _
            " where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'" & _
            " and np22=" & m_str605NP22
         cnnConnection.Execute strSql, intI
      End If
      
      If m_bolFMP Then
         'Add by Morgan 2006/6/26
         '收文智權人員為'F'字頭者,新申請案期限為收文日+1個月(沒期限才更新)

         'Added by Lydia 2015/09/04 剔除有大陸案關聯之香港或澳門案
         'If cp(27) = "" And Text1(4) = "" And InStr(NewCasePtyList, Text1(1)) > 0 Then
         'Remove by Lydia 2016/07/07 改成模組PUB_SavePtoUpd2 'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
         'If cp(27) = "" And Text1(4) = "" And InStr(NewCasePtyList, Text1(1)) > 0 And Not (stCM10 = "4" Or stCM10 = "5") Then
         '   strExc(0) = PUB_GetDeadLine(DBDATE(Text1(12)), "", 2)
         '   'Modify by Morgan 2009/11/5
         '   'FMP 案非PCT且未主張優先權則不掛法限(所限=收文日+1個月)
         '   'strSQL = "update caseprogress set cp06=" & Val(strExc(0)) & ",cp07=" & Val(strExc(0)) & _
         '      " where cp09='" & cp(9) & "' and cp06 is null"
         '   strExc(0) = PUB_GetWorkDay1(strExc(0), True)
         '   strSql = "update caseprogress set cp06=" & Val(strExc(0)) & _
         '      " where cp09='" & cp(9) & "' and cp06 is null"
         '   'end 2009/11/5
         '   intI = 0
         '   cnnConnection.Execute strSql, intI
         'End If
      
         'Add by Morgan 2007/8/21
         '外專收文翻譯案件，若承辦人為外專工程師時核稿人設為自己
         'Modify by Morgan 2009/11/13 +209,210 也要
         'If Text1(1) = "201" Then
         'Modify by Morgan 2011/9/1 +942
         'Modified by Morgan 2013/11/6 +235核對中說格式
         'Remove by Lydia 2016/07/07 改成模組PUB_SavePtoUpd2 'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
         'If Text1(1) = "201" Or Text1(1) = "209" Or Text1(1) = "235" Or Text1(1) = "210" Or Text1(1) = "942" Then
         '   'Add by Morgan 2009/11/5
         '   '更新承辦期限、核稿期限
         '   '抓有所限的新案(若無法限時也會設所限)
         '   strExc(0) = "select cp05,cp07 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('101','102','103') and cp06>0"
         '   intI = 1
         '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '   If intI = 1 Then
         '      '承辦期限
         '      strExc(1) = PUB_GetDeadLine(RsTemp.Fields("cp05"), "" & RsTemp.Fields("cp07"), 3)
         '      strSql = "update caseprogress set cp48=" & CNULL(strExc(1), True) & " where cp09='" & cp(9) & "'"
         '      cnnConnection.Execute strSql, intI
         '
         '      'Modified by Morgan 2015/3/11 只有 201 要預設核稿人及期限  --靜芳
         '      If Text1(1) = "201" Then
         '         '核稿期限
         '         strExc(2) = PUB_GetDeadLine(RsTemp.Fields("cp05"), "" & RsTemp.Fields("cp07"), 4)
         '         strSql = "update engineerprogress set ep08=" & CNULL(strExc(2), True) & " where ep02='" & cp(9) & "'"
         '         cnnConnection.Execute strSql, intI
         '      End If
         '   End If
         '   'end 2009/11/5
         '
         '   'Modified by Morgan 2015/3/11 只有 201 要預設核稿人及期限  --靜芳
         '   'If Text1(0) <> "" Then
         '   If Text1(0) <> "" And Text1(1) = "201" Then
         '   'end 2015/3/11
         '
         '      '2008/4/8 MODIFY BY SONIA加ST03=F81
         '      'strExc(0) = "select 1 from staff_idmap,staff where '" & Text1(0) & "' in (sim01,sim02)" & _
         '      '   " and st01(+)=sim01 and st03='F21' and st04='1'"
         '      'Modify by Morgan 2008/9/18 改抓ST15='F21'的
         '      strExc(0) = "select 1 from staff_idmap,staff where '" & Text1(0) & "' in (sim01,sim02)" & _
         '         " and st01(+)=sim01 and ST15='F21' and st04='1'"
         '      intI = 1
         '      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '      If intI = 1 Then
         '         strSql = "update engineerprogress set ep04='" & Text1(0) & "' where ep02='" & cp(9) & "'"
         '         cnnConnection.Execute strSql, intI
         '      End If
         '   End If
         'End If
         'end 2007/8/21
         'Added by Lydia 2016/07/07 內專分案(工程師主管分案frm060117)-更新寰華案
         'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
         Call PUB_SavePtoUpd2(m_bolFMP, pa, Text1(0), Label3(8), Text1(1).Text, Text1(12).Text, Text1(4).Text, cp(27), stCM10)
         
         'Added by Morgan 2012/3/19 輸入工程師組別同時設定該案未分案之新案翻譯,製作中說,檢視中說,901告知代理人的承辦人為該組管制人
         'Modified by Lydia 2023/02/21 改成模組
'         If pa(150) <> txtEngGroup And txtEngGroup <> "" Then
'            'Modifed by Morgan 2012/7/4 +942,203
'            'Modified by Morgan 2013/11/6 +235核對中說格式
'            strExc(0) = "select cp09,cp10 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp14 is null and cp57 is null and cp10 in ('201','209','235','210','901','942','203')"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               With RsTemp
'               Do While Not .EOF
'                  'Modify by Amy 2015/04/07 自動上cp157 (北所分案日)
'                  strSql = "update caseprogress set cp14=(select oMan from SetSpecMan where OCODE=decode('" & txtEngGroup & "','1','T','2','R','3','S','4','T1')),cp157=" & strSrvDate(1) & _
'                     " where cp09='" & .Fields("cp09") & "'"
'                  cnnConnection.Execute strSql, intI
'
'                  'Modified by Morgan 2013/11/6 +235核對中說格式
'                  'Modified by Morgan 2015/3/11 只有 201 要預設核稿人及期限  --靜芳
'                  'If .Fields("cp10") = "201" Or .Fields("cp10") = "209" Or .Fields("cp10") = "235" Or .Fields("cp10") = "210" Or .Fields("cp10") = "942" Then
'                  If .Fields("cp10") = "201" Then
'                  'end 2015/3/11
'                     strSql = "update engineerprogress set ep04=(select cp14 from caseprogress where cp09=ep02) where ep02='" & .Fields("cp09") & "'"
'                     cnnConnection.Execute strSql, intI
'                  End If
'                  .MoveNext
'               Loop
'               End With
'            End If
'         End If
'         'end 2012/3/19
'
'         'Added by Lydia 2018/05/22 FMP案-命名作業分工程師組別,通知相關人員
'         'Modifiedby Lydia 2021/04/08 判斷有改組別; ex.P-126940進來改AB0010843的承辦人,清空命名記錄
'         'If cp(31) = "Y" Then
'         If cp(31) = "Y" And pa(150) <> txtEngGroup Then
'            strExc(1) = ""
'            Select Case txtEngGroup.Text
'                Case "1": strExc(1) = Pub_GetSpecMan("T")
'                Case "2": strExc(1) = Pub_GetSpecMan("R")
'                Case "3": strExc(1) = Pub_GetSpecMan("S")
'                Case "4": strExc(1) = Pub_GetSpecMan("T1")
'                Case Else: strExc(1) = "B"
'            End Select
'            'Added by Lydia 2022/10/12 特殊情況之指定職代
'             strExc(1) = PUB_GetStateForMan(strExc(1))
'
'            strSql = "update transcasetitle set TCT16=" & CNULL(pa(5)) & ",TCT17=" & CNULL(pa(6)) & ", TCT04=" & IIf(strExc(1) <> "B", CNULL(strExc(1)), "Null") & ",TCT05=NULL, TCT06=NULL, TCT07=NULL,TCT08=NULL,TCT09=NULL,TCT10=NULL,TCT11=NULL,TCT12=NULL, " & _
'                        " TCT112='" & strUserNum & "', TCT113=" & strSrvDate(1) & ", TCT114=" & Mid(Format(ServerTime, "000000"), 1, 4) & " where tct01='" & Me.Label3(8).Caption & "' "
'            cnnConnection.Execute strSql, intI
'            If intI > 0 Then
'                '更改分案組別 , 通知雙方
'                If pa(150) <> "" Then
'                        '重新分案 , 清除卷宗區記錄, 直到新組別主管確認再次產生
'                        strSql = "delete from casepaperpdf where cpp01='" & Me.Label3(8).Caption & "' and instr(cpp02,'" & FCP命名記錄 & "') > 0 "
'                        cnnConnection.Execute strSql, intI
'                        '命名作業的主管確認自動掛承辦人=命名人員並且上已分案,所以改工程師組別時一併清空,直到下次主管確認一併更新
'                        strSql = "Update caseprogress set cp14=null, cp122=null where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                                    " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' and cp10 in ('902','203') "
'                        cnnConnection.Execute strSql, intI
'                        '清空-工程師收告代901和主動修正203
'                        strSql = "Update caseprogress set cp14=null, cp122=null,cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & "修改工程師組別：" & PUB_GetFCPGrpName(pa(150)) & "->" & PUB_GetFCPGrpName(txtEngGroup.Text) & ";'||cp64 " & _
'                                    " where cp09 in (select cp09 from caseprogress,staff where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
'                                    " and cp158=0 and cp159=0 and substr(cp09,1,1)='B' and cp10 in ('901','203') and cp65=st01(+) and st03='F21' )  "
'                        cnnConnection.Execute strSql, intI
'
'                        strExc(2) = pa(150) & "-" & txtEngGroup.Text
'                        strExc(3) = ""
'                        Select Case pa(150) 'CC:原工程師主管
'                            Case "1": strExc(3) = Pub_GetSpecMan("T")
'                            Case "2": strExc(3) = Pub_GetSpecMan("R")
'                            Case "3": strExc(3) = Pub_GetSpecMan("S")
'                            Case "4": strExc(3) = Pub_GetSpecMan("T1")
'                            Case Else: strExc(3) = "B"
'                        End Select
'                        If PUB_GetTCTmail(True, 2, pa(1), pa(2), pa(3), pa(4), Me.Label3(8).Caption, "", strExc(1), strExc(2), , , strExc(3)) Then
'                        End If
'                Else
'                        If PUB_GetTCTmail(True, 1, pa(1), pa(2), pa(3), pa(4), Me.Label3(8).Caption, "", strExc(1)) Then
'                        End If
'                End If
'            End If
'         End If 'If cp(31) = "Y" And pa(150) <> txtEngGroup Then
'         'Added by Lydia 2021/05/31 內專程序分案後，系統自動發email通知
'         'Modified by Lydia 2021/11/17 排除香港案(確定要排除母案為寰華案衍生的香港案通知，且包含所有香港案收文之分案 by Phoebe)
'         If Text1(0).Tag = "" And Text1(0).Tag <> Text1(0).Text And Text1(13) <> "013" Then
'            'Modified by Lydia 2022/06/30 分案(307)改通知工程師
'            'If cp(31) = "Y" Then '新案:通知程序, CC:程序主管(二級主管)
'            If cp(31) = "Y" And Text1(1) <> "307" Then
'               strExc(1) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
'               strExc(2) = PUB_GetFCPProSup(strExc(1))
'            Else  '中間程序:通知輸入的承辦人員
'               strExc(1) = Trim(Text1(0))
'               strExc(3) = PUB_GetST03(strExc(1))
'               'CC: 日文組工程師為副理+協理; 英文組工程師為副理; 其餘部門為人員主管(二級主管)
'               If strExc(3) = "F21" Then
'                   strExc(2) = PUB_GetFCPEngSup(strExc(1), True)
'                   If txtEngGroup.Text = "3" Then
'                       strExc(4) = Pub_GetSpecMan("S")
'                       If InStr(strExc(1) & ";" & strExc(2), strExc(4)) = 0 Then
'                           strExc(2) = strExc(2) & IIf(strExc(2) <> "", ";", "") & strExc(4)
'                       End If
'                   End If
'               Else
'                   'FMP案只通知工程師，寰華案全部都通知
'                   If m_bolFMP2 = False Then
'                       strExc(1) = ""
'                   Else
'                       strExc(2) = PUB_GetFCPProSup(strExc(1))
'                   End If
'               End If
'            End If
'            If strExc(1) <> "" Then
'                'Added by Lydia 2022/06/30 分案(307)改通知工程師,並CC程序
'                If cp(31) = "Y" And Text1(1) = "307" Then
'                     strExc(3) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
'                     strExc(2) = strExc(2) & ";" & strExc(3)
'                End If
'                'end 2022/06/30
'                strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'                   " values( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd')" & _
'                   ",to_char(sysdate,'hh24miss'),'" & pa(1) & pa(2) & IIf(pa(3) <> "0", pa(3), "") & IIf(pa(4) <> "00", pa(4), "") & "【" & Trim(Label3(1).Caption) & "已分案】請處理後續！" & "','同主旨','" & strExc(2) & "' ) "
'                cnnConnection.Execute strSql
'            End If
'         End If
'         'end 2021/05/31
         Call PUB_SavePtoUpd4(pa, Text1(13).Text, txtEngGroup.Text, pa(150), cp(31), Label3(8).Caption, Text1(1).Text, Text1(0).Text, Text1(0).Tag, pa(5), pa(6))
         'end 2023/02/21
      End If  'If m_bolFMP Then
      
      'Add by Morgan 2009/7/15 大陸分割案期限控管
      If Text1(1).Text = "307" And cp(27) = "" Then
         If Text1(13).Text = "020" Then
            If Not m_CN307Updated Then 'Added by Morgan 2023/5/19
               st307Msg = PUB_Update307Ref(cp(9))
            End If
         'Added by Morgan 2011/12/6 +台灣 Ex.P-100262
         ElseIf Text1(13).Text = "000" Then
            st307Msg = PUB_Update307RefTw(cp(9))
            
         End If
      End If
      
      '2009/10/16 ADD BY SONIA FMP案件以"FMP"+申請國家+案件性質重新抓費用,規費
      'Modify by Morgan 2009/12/30 改案件性質才要重抓
      'If Left(cp(12), 1) = "F" Then
      If Left(cp(12), 1) = "F" And Text1(1) <> cp(10) Then
         strExc(0) = "SELECT NVL(CF06,0),NVL(CF08,0) FROM CASEFEE WHERE CF01='FCP' AND CF02='" & pa(9) & "' AND CF03='" & Text1(1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "update caseprogress set cp16=" & Val(RsTemp.Fields(0)) & ",cp17=" & Val(RsTemp.Fields(1)) & ",cp18=" & (Val(RsTemp.Fields(0)) - Val(RsTemp.Fields(1))) / 1000 & _
               " where cp09='" & cp(9) & "' "
            intI = 0
            cnnConnection.Execute strSql, intI
         End If
      End If
      '2009/10/16 END
      
      
      'Added by Morgan 2012/7/2 102/1/1 專利新法
      If pa(9) = "000" Then
         '新增B類寄存證明(231)
         If m_bolCtrl231 = True Then
            strExc(2) = ""
            strExc(3) = PUB_GetFirstPriDate(pa)
            '法限=最早優先權日起16個月
            If strExc(3) <> "" Then
               strExc(3) = CompDate(1, 16, strExc(3))
               stDate(0) = ""
               stDate(1) = pa(1)
               stDate(2) = pa(9)
               stDate(3) = strExc(3)
               GetCtrlDT stDate
               '本所期限
               strExc(2) = PUB_GetWorkDay1(stDate(0), True)
            End If
            
            strExc(1) = AutoNo("B", 6)
            strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07" & _
               ",CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43) " & _
               " select cp01,cp02,cp03,cp04,to_char(sysdate,'yyyymmdd')," & CNULL(strExc(2), True) & "," & CNULL(strExc(3), True) & _
               ",'" & strExc(1) & "','231','90',cp12,cp13,cp14,'N','N','N',cp09" & _
               " from caseprogress where cp09='" & cp(9) & "'"
            cnnConnection.Execute strSql, intI
            
         End If
      End If
      'end 2012/7/2
   
      'Added by Morgan 2012/9/12
      '新案若設公司別與已開收據不同時發Mail通知財務處及智權人員
      'If txtPA161.Visible = True And txtPA161.Tag <> txtPA161 And Left(cp(60), 1) = "E" Then
      If txtPA161.Visible = True And Left(cp(60), 1) = "E" Then
         'Modify By Sindy 2013/12/15
         'If strSrvDate(1) >= InvoiceStartDate Then
         
         'Modified by Morgan 2015/5/12 PA161 Y 都改 T
         'If Val(m_CP31isYGetCP05) >= Val(InvoiceStartDate) Then
         '   strExc(0) = "select a0k01||DECODE(A0K32,'N','(暫不印)','Y','(待列印)') from acc0j0,acc0k0 where a0j01='" & cp(9) & "' and a0k01(+)=a0j13 and a0k11<>'" & IIf(txtPA161 = "T", "1", IIf(txtPA161 = "J", "J", "2")) & "'"
         'Else
         ''2013/12/15 END
         '   strExc(0) = "select a0k01||DECODE(A0K32,'N','(暫不印)','Y','(待列印)') from acc0j0,acc0k0 where a0j01='" & cp(9) & "' and a0k01(+)=a0j13 and a0k11<>'" & IIf(txtPA161 = "Y", "1", "2") & "'"
         'End If
         strExc(0) = "select a0k01||DECODE(A0K32,'N','(暫不印)','Y','(待列印)') from acc0j0,acc0k0 where a0j01='" & cp(9) & "' and a0k01(+)=a0j13 and a0k11<>'" & IIf(txtPA161 = "T", "1", IIf(txtPA161 = "J", "J", "2")) & "'"
         'end 2015/5/12
         
         strExc(1) = ""
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Do While Not RsTemp.EOF
               If strExc(1) = "" Then
                  strExc(1) = RsTemp(0)
               Else
                  strExc(1) = strExc(1) & "," & RsTemp(0)
               End If
               RsTemp.MoveNext
            Loop
            'Modify By Sindy 2013/12/15
            'If strSrvDate(1) >= InvoiceStartDate Then
            
            'Modified by Morgan 2015/5/12 PA161 Y 都改 T
            'If Val(m_CP31isYGetCP05) >= Val(InvoiceStartDate) Then
            '   strExc(1) = "專利案 " & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " 設定為以" & IIf(txtPA161 = "T", "專利商標", IIf(txtPA161 = "J", "台一智權", "專利法律")) & "出名與收據 " & strExc(1) & " 的公司別不同，請更正！"
            'Else
            ''2013/12/15 END
            '   strExc(1) = "專利案 " & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " 設定為以專利" & IIf(txtPA161 = "Y", "商標", "法律") & "出名與收據 " & strExc(1) & " 的公司別不同，請更正！"
            'End If
            'Modified by Lydia 2020/03/31 智慧所更名日
            'strExc(1) = "專利案 " & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " 設定為以" & IIf(txtPA161 = "T", "專利商標", IIf(txtPA161 = "J", "台一智權", "專利法律")) & "出名與收據 " & strExc(1) & " 的公司別不同，請更正！"
            ''end 2015/5/12
            If txtPA161 = "T" Then
                strExc(2) = CompNameQuery("1", "4")
            ElseIf txtPA161 = "J" Then
                strExc(2) = CompNameQuery("J", "4")
            Else
                strExc(2) = CompNameQuery("2", "4")
            End If
            strExc(1) = "專利案 " & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " 設定為以" & strExc(2) & "出名與收據 " & strExc(1) & " 的公司別不同，請更正！"
            'end 2020/03/31
            'Modify by Amy 2024/05/15 財務2個特殊設定拆成3個
            If Val(strSrvDate(1)) >= Val(財務拆總帳出納國內應收啟用日) Then
                strExc(3) = Pub_GetSpecMan("財務處應收處理人員")
            Else
               strExc(3) = Pub_GetSpecMan("財務處總帳人員")
            End If
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " values ('" & strUserNum & "','" & strExc(3) & ";" & cp(13) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & ChgSQL(strExc(1)) & "','如旨')"
            'end 2024/05/15
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2012/9/12
               
      'Add by Morgan 2010/5/7
      '新增關聯且國內案已發文
      If bolInsCM Then
         'Modified by Morgan 2014/6/3 改承辦人若為程序時更新草墨齊、自動確認繪圖分案並發Mail通知,工程師則都不要
         If GetStaffDepartment(Text1(0)) = "P12" Then
            '若是國內案已發文後才收文大陸案件,則請在建關聯後發mail通知繪圖人員製作pdf檔給大陸工程師
            'Modify by Morgan 2010/5/24 會稿完成也可
            'Modified by Morgan 2011/10/26 +PCT Ex.P100033
            'If Text1(13) = "020" Then
            'Modified by Morgan 2015/6/30 +香港短期及外觀
            If Text1(13) = "020" Or Text1(13) = "056" Or (Text1(13) = "013" And Text1(2) <> "1") Then
               'Modify by Morgan 2011/7/5 改回只判斷國內案已發文
               'Modified by Morgan 2016/2/24 再改加判斷會稿完成 --郭雅娟
               strExc(0) = "select cp27,ep08 from caseprogress,engineerprogress where ep02(+)=cp09 and cp01='" & stInCNo(1) & "'" & _
                  " and cp02='" & stInCNo(2) & "' and cp03='" & stInCNo(3) & "' and cp04='" & stInCNo(4) & "'" & _
                  " AND CP10 in (" & CaseMapIn & ") and (cp27>0 or ep08>0)"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  
                  'Removed by Morgan 2013/10/23 取消,承辦單已電子化--瓊玉
                  'Modified by Morgan 2014/6/3
                  'Removed by Morgan 2016/2/24 移到下面(比照會稿完成改為通知上墨並新增歷程)
                  'strExc(1) = stInCNo(1) & "-" & stInCNo(2) & IIf(stInCNo(3) & stInCNo(4) = "000", "", "-" & stInCNo(3) & "-" & stInCNo(4))
                  'strExc(2) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
                  'strExc(3) = "國內案 " & strExc(1)
                  'If RsTemp(0) > 0 Then
                  '   strExc(3) = strExc(3) & " 已發文"
                  'Else
                  '   strExc(3) = strExc(3) & " 已會稿完成"
                  'End If
                  ''strExc(3) = "'" & strExc(3) & ",請製作" & Label3(11) & "案 " & strExc(2) & " pdf 檔之圖式交承辦人" & "'||st02||'('||cp14||')'"
                  'strExc(3) = "'" & strExc(3) & ",請製作" & Label3(11) & "案 " & strExc(2) & " pdf 檔之圖式'"
      
                  'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                     " select '" & strUserNum & "',cp29,to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                     "," & strExc(3) & ",'如旨' from caseprogress,staff" & _
                     " where cp09='" & cp(9) & "' and cp57 is null and cp27 is null and cp29 is not null and st01(+)=cp14"
                  'cnnConnection.Execute strSql, intI
                  'end 2013/10/23
                  'end 2016/2/24
                  
                  '草墨都不計件的案子，更新草墨圖要計件,加乘註記=0.2,張數=0(國外計件值會變成0.6)
                  '草齊日=墨齊日=文齊日
                  strExc(0) = "select cp09,ep20,ep29,ep17,cp29 from caseprogress,engineerprogress" & _
                     " where cp09='" & cp(9) & "' and cp57 is null and cp27 is null and ep02(+)=cp09 and cp29 is not null"
                  intI = 1
                  Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     'Modified by Morgan 2014/6/3 舊規則程式不用再留
   '                  If AdoRecordSet3("ep20") = "N" And AdoRecordSet3("ep29") = "N" Then
   '                     strSql = "update caseprogress set cp101=0.2,cp104=0.2,cp107=decode(cp29,null,null,'Y') where cp09='" & AdoRecordSet3(0) & "'"
   '                     cnnConnection.Execute strSql, intI
   '                     'Modify by Morgan 2010/6/30 草齊墨齊改上系統日(因為可能工程師還沒有輸文齊日)
   '                     'strSql = "update engineerprogress set ep20=null,ep29=null,ep16=0,ep19=0,ep14=ep06,ep17=ep06 where ep02='" & AdoRecordSet3(0) & "'"
   '                     strSql = "update engineerprogress set ep20=null,ep29=null,ep16=0,ep19=0,ep14=nvl(ep14," & strSrvDate(1) & "),ep17=nvl(ep17," & strSrvDate(1) & ") where ep02='" & AdoRecordSet3(0) & "'"
   '                     cnnConnection.Execute strSql, intI
   '
   '                  'Add by Morgan 2011/5/16
   '                  '新規則改要計件
   '                  Else
                        strSql = "update caseprogress set cp107='Y' where cp09='" & AdoRecordSet3(0) & "'"
                        cnnConnection.Execute strSql, intI
                     'end 2014/6/3
                     
                        strSql = "update engineerprogress set ep14=nvl(ep14,decode(ep20,null,to_char(sysdate,'yyyymmdd')))" & _
                           ",ep17=nvl(ep17,decode(ep29,null,to_char(sysdate,'yyyymmdd')))" & _
                           " where ep02='" & AdoRecordSet3(0) & "'"
                        cnnConnection.Execute strSql, intI
                     'End If 'Removed by Morgan 2014/6/3
                     
                     'Added by Morgan 2016/2/24
                     If IsNull(AdoRecordSet3("ep29")) And IsNull(AdoRecordSet3("ep17")) Then
                        strExc(1) = stInCNo(1) & "-" & stInCNo(2) & IIf(stInCNo(3) & stInCNo(4) = "000", "", "-" & stInCNo(3) & "-" & stInCNo(4))
                        strExc(2) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
                        strExc(3) = "國內案 " & strExc(1) & " 已" & IIf(RsTemp("cp27") > 0, "發文", "會稿完成") & "," & IIf(Text1(13) = "020", "大陸", Label3(11)) & "案 " & strExc(2) & " 請上墨處理！"
                        strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                           " values('" & strUserNum & "','" & AdoRecordSet3("cp29") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                           ",'" & strExc(3) & "','如旨')"
                        cnnConnection.Execute strSql, intI
                        strSql = "insert into empelectronprocess(eep01,eep02,eep03,eep04,eep05,eep06,eep07,eep08) select '" & AdoRecordSet3("cp09") & "',nvl(max(eep02),0)+1" & _
                           ",'QPGMR','22','" & AdoRecordSet3("cp29") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'國內案已" & IIf(RsTemp("cp27") > 0, "發文", "會稿完成") & "，請上墨處理！'" & _
                           " FROM empelectronprocess WHERE eep01='" & AdoRecordSet3("cp09") & "'"
                        cnnConnection.Execute strSql, intI
                     End If
                     'end 2016/2/24
                  End If
               End If
            End If
         End If 'Added by Morgan 2014/6/3
      End If
      'end 2010/5/7
   
      '2012/3/13 add by sonia 案件性質 941分析, 分案時自動上齊備日
      'Remove by Lydia 2016/07/07 改成模組PUB_SavePtoUpd3 'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
      'If Text1(1).Text = "941" Then
      '   strTxt(intStep) = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & Label3(8) & "' AND EP06 IS NULL"
      '   cnnConnection.Execute strTxt(intStep), intI
      '   intStep = intStep + 1
      'End If
      ''2012/3/13 end
   
      'Add by Morgan 2010/6/17
      '若已開請款單則換承辦人或核稿人時發Mail通知靜芳
      'Remove by Lydia 2016/07/07 改成模組PUB_SavePtoUpd3
      'If cp(60) > "X" Then
      '   m_newEP04 = ""
      '   If Text1(1) = "201" Then
      '      strExc(0) = "select ep04 from engineerprogress where ep02='" & cp(9) & "'"
      '      intI = 1
      '      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      '      If intI = 1 Then
      '         m_newEP04 = "" & RsTemp(0)
      '      End If
      '   End If
      '   PUB_PointReAssignInform Label3(9), cp(60), cp(14), Text1(0), m_oldEP04, m_newEP04
      'End If
      'Added by Lydia 2016/07/07 內專分案(工程師主管分案frm060117)-更新其他
      Call PUB_SavePtoUpd3(pa, Text1(0), Label3(8), Text1(1).Text, cp(14), cp(60), m_oldEP04)
      
   'Remove by Morgan 2011/4/29 取消
   'Add by Morgan 2011/4/26
   'A類收文的配合開庭要檢查若該案已有其他A類收文的配合開庭且不請款時該收文也設定為不請款
   '   If cp(1) = "P" And cp(9) < "B" And Text1(1) = "226" And Val(cp(16)) = 0 Then
   '      strSql = "update caseprogress a set cp20='N' where cp09='" & cp(9) & "'" & _
   '         " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02" & _
   '         " and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp10=a.cp10 and b.cp09<'B'" & _
   '         " and b.cp57||b.cp20='N' and nvl(b.cp16,0)=0 and b.cp09<>a.cp09)"
   '      cnnConnection.Execute strSql, intI
   '   End If
      
      'Added by Morgan 2012/3/26 主張新穎性優惠期要更新相關美國發明案提申期限
      If txtFavDt.Visible = True And txtFavDt <> "" Then
         If PUB_GetRefCaseList(cp(), m_list()) = True Then
            If UBound(m_list, 2) > 1 Then
               PUB_UpdateUsPatent m_list
            End If
         End If
      End If
      'end 2012/3/26
   
      'Add by Amy 2014/09/05 承辦人計件值若修改則更新並寫log-玲玲
      If Val(txtCP97) <> Val(cp(97)) Then
           'Modify by Amy 2014/09/09 改寫至例外TB
           'strSql = "UPDATE CaseProgress Set CP97=" & Val(txtCP97) & " Where CP09='" & Label3(8).Caption & "' "
           If ExistCheck("EXVALUE", "EV01", Label3(8).Caption, strExc(0), False) = True Then
               strSql = "UPDATE EXVALUE Set EV02=" & Val(txtCP97) & " Where EV01='" & Label3(8).Caption & "' "
           Else
               strSql = "INSERT INTO EXVALUE (EV01,EV02) VALUES ('" & Label3(8).Caption & "','" & Val(txtCP97) & "') "
           End If
           'end 2014/09/09
           Pub_SeekTbLog strSql '記錄修改log
           cnnConnection.Execute strSql
      End If
      'end 2014/09/05
   
'Removed by Morgan 2021/4/8 使用在途期限已改需另外收文在途期限間(442)會於分案時自動發文並更新相關號期限
'      'Added by Lydia 2015/05/13 進度備註+法定期限已加在途期限15天
'      If m_bolUpdCP07 = True Then
'         strExc(5) = GetCaseProData(cp(9), "CP64")
'         strExc(5) = Trim(strExc(5))
'         strExc(5) = IIf(Len(strExc(5)) > 0, strExc(5) + ",法定期限已加在途期限15天", "法定期限已加在途期限15天")
'         strSql = " UPDATE caseprogress set cp64=" & CNULL(strExc(5)) & " where cp09=" & CNULL(cp(9))
'         cnnConnection.Execute strSql
'      End If
'end 2021/4/8


      'Added by Morgan 2015/10/13
      '審查意見或核駁修改承辦人時一併修改相關收文號之告代承辦人
      'Remove by Lydia 2016/07/07 改成模組PUB_SavePtoUpd2 'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
      'If m_bolFMP And (Text1(0) <> "" And (Text1(1) = "1202" Or Text1(1) = "1002")) Then
      '   strSql = "update caseprogress set  cp14='" & Text1(0) & "' where cp43='" & Label3(8) & "' and cp10='901' and cp27 is null"
      '   cnnConnection.Execute strSql, intI
      'End If
      'end 2015/10/13
      
      'Added by Morgan 2015/12/15
      If m_bolAuto404 Then
         strExc(1) = AutoNo("B", 6)
         'Modified by Morgan 2025/7/17 承辦人改掛負責該區的程序
         strExc(2) = PUB_GetPHandler(cp(1) & cp(2) & cp(3) & cp(4))
         strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07" & _
            ",CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP64) " & _
            " select cp01,cp02,cp03,cp04," & strSrvDate(1) & ",cp06,cp07" & _
            ",'" & strExc(1) & "','404','90',cp12,cp13,'" & strExc(2) & "','N','N','N',cp09,CP64" & _
            " from caseprogress where cp09='" & cp(9) & "'"
         cnnConnection.Execute strSql, intI
         
         'Added by Morgan 2025/7/17
         strExc(3) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & "內部收文(延期-" & Label3(1) & ")分案通知!!"
         strSql = "insert into mailcache a(mc01,mc02,mc03,mc04,mc07,mc08,mc13)" & _
               " values ('" & strUserNum & "','" & strExc(2) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & strExc(3) & "','','" & strExc(1) & "')"
            cnnConnection.Execute strSql, intI
      End If
      'end 2015/12/15
      
      'Added by Lydia 2016/09/29 一案兩請的新型案做自動收文放棄專利權
      If m_bolAuto429 Then
         strExc(0) = "select np08,np09 from nextprogress where np02='" & m_DualCaseNo(1) & "' and np03='" & m_DualCaseNo(2) & "' and np04='" & m_DualCaseNo(3) & "' and np05='" & m_DualCaseNo(4) & "' and np07='429' and np06 is null "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
             strExc(3) = "" & RsTemp(0)
             strExc(4) = "" & RsTemp(1)
         End If
         
         strSql = "Update nextprogress set np06='Y' where np02='" & m_DualCaseNo(1) & "' and np03='" & m_DualCaseNo(2) & "' and np04='" & m_DualCaseNo(3) & "' and np05='" & m_DualCaseNo(4) & "' and np07='429' and np06 is null "
         cnnConnection.Execute strSql, intI
         
         strExc(1) = PUB_GetAKindSalesNo(m_DualCaseNo(1), m_DualCaseNo(2), m_DualCaseNo(3), m_DualCaseNo(4))
         strExc(2) = PUB_GetStaffST15(strExc(1), 1)
         strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07," & _
                "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP64) VALUES " & _
                "('" & m_DualCaseNo(1) & "','" & m_DualCaseNo(2) & "','" & m_DualCaseNo(3) & "','" & m_DualCaseNo(4) & "'," & strSrvDate(1) & "," & strExc(3) & "," & strExc(4) & _
                ",'" & AutoNo("B", 6) & "','429','90'," & CNULL(strExc(2)) & "," & CNULL(strExc(1)) & _
                ",'" & strUserNum & "','N','N','N','" & cp(9) & "','') "
         cnnConnection.Execute strSql, intI
      End If
      'end 2016/09/29
      
      'Added by Morgan 2017/4/10
      '大陸案承辦人為程序時自動設不會稿
      If Text1(13) = "020" And Text1(0) <> "" Then
         If GetStaffDepartment(Text1(0)) = "P12" Then
            strSql = "update engineerprogress set ep34='N' where ep02='" & cp(9) & "' and ep34 is null"
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2017/4/10
      
      'Modify By Sindy 2016/4/13 抽出來變共用Func
      Call PUB_UpdRelationCaseFixEP(cp(1), cp(2), cp(3), cp(4), Text1(1), Label3(1))
      '2016/4/13 END
      
      'Added by Morgan 2018/3/26
      'P案主張國內優先權在分案時,若該先申請案之關聯CFP案尚未發文者,請發EMAIL通知該案之工程師,並副本給王副總。
      '"CFPXXX之關聯案PXXX己被主張國內優先權,請確認是否要改鍵關聯,並回覆給王副總。"
      'Modified by Morgan 2020/6/23 取消未發文條件，排除已取消收文--郭
      If Me.Text1(1).Text = "121" Then
         strExc(0) = "select pd01||'-'||pd02||decode(pd03||pd04,'000','','-'||pd03||'-'||pd04) CNo" & _
            ",pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) PNo" & _
            ",cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CFPNo" & _
            ",nvl(cp14,'79075') cp14 from pridate,patent,casemap,caseprogress" & _
            " where pd01='" & cp(1) & "' and pd02='" & cp(2) & "' and pd03='" & cp(3) & "' and pd04='" & cp(4) & "'" & _
            " and pa11(+)=pd06 and pa01(+)=pd01 and pa09(+)=pd07" & _
            " and cm05(+)=pa01 and cm06(+)=pa02 and cm07(+)=pa03 and cm08(+)=pa04 and cm10='0' and cm01='CFP'" & _
            " and cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03 and cp04(+)=cm04" & _
            " and cp10 in (" & CaseMapOut & ") and cp159=0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Do While Not RsTemp.EOF
            'Modified by Lydia 2023/04/24 修改王副總退休之相關控制
            'strExc(1) = RsTemp("CFPNo") & " 之關聯案 " & RsTemp("PNo") & " 已被 " & RsTemp("CNo") & " 主張國內優先權,請確認是否要改鍵關聯,並回覆給王副總。"
            If strSrvDate(1) >= "20230511" Then
                strExc(2) = "99050"
                strExc(3) = "李經理"
            ElseIf strSrvDate(1) >= "20230501" Then
                strExc(2) = "71011;99050"
                strExc(3) = "王副總"
            Else
                strExc(2) = "71011"
                strExc(3) = "王副總"
            End If
            strExc(1) = RsTemp("CFPNo") & " 之關聯案 " & RsTemp("PNo") & " 已被 " & RsTemp("CNo") & " 主張國內優先權,請確認是否要改鍵關聯,並回覆" & strExc(3) & "。"
            'end 2023/04/24
            'Modified by Lydia 2023/04/24 '71011' => CNULL(strExc(2))
            strSql = "insert into mailcache a(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
               " values ('" & strUserNum & "','" & RsTemp("cp14") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & ChgSQL(strExc(1)) & "','如旨'," & CNULL(strExc(2)) & ")"
            cnnConnection.Execute strSql, intI
               RsTemp.MoveNext
            Loop
         End If
      End If
      'end 2018/3/26
      
      'Added by Morgan 2020/2/7
      '大陸「在途期間」分案後自動上發文日並更新期限
      If Text1(1) = "442" Then
         strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & cp(9) & "'"
         cnnConnection.Execute strSql, intI
         
         If m_str442DeadLine <> "" Then
            PUB_442UpdateDeadline Text1(6), m_str442DeadLine, m_bolFMP
         End If
      End If
      'end 2020/2/7
      
      'Added by Morgan 2020/3/16
      'P台灣申請中案件，最後發文之A或B類(不續辦、閉卷、取消收文除外)之出名代理人有76012桂所長的案件，不論申請人是否有總委案件，存檔時都自動做內部收文變更401
      'Modified by Morgan 2020/4/17 941(分析)除外--玲玲
      If pa(1) = "P" And Text1(13) = "000" And Text1(1) <> "941" And cp(27) = "" Then
         'Modified by Morgan 2020/3/18 改寫函數共用
'         strExc(3) = ""
'         If PUB_ChkIsGuiCase(pa(1), pa(2), pa(3), pa(4), strExc(3)) Then
'            strSql = "update caseprogress set cp64=cp64 where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='401' and cp43='" & cp(9) & "' and cp64='變更代理人為閻+林'"
'            cnnConnection.Execute strSql, intI
'            If intI = 0 Then
'               '變更承辦人
'               strExc(2) = Pub_GetSpecMan("PS1")
'               '內部收文變更401(費用:0,規費:300,點數:-0.3)
'               strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
'                  "CP09,CP10,CP11,CP12,CP13,CP14,CP16,CP17,CP18,CP20,CP26,CP32,CP43,CP64) VALUES " & _
'                  "('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & _
'                  ",'" & AutoNo("B", 6) & "','401','90'," & CNULL(cp(12)) & "," & CNULL(cp(13)) & _
'                  ",'" & strExc(2) & "',0,300,-0.3,'N','N','N','" & cp(9) & "','變更代理人為閻+林') "
'               cnnConnection.Execute strSql, intI
'
'               'E-MAIL通知承辦人
'               strExc(1) = "本所案號 " & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
'               strExc(1) = strExc(1) & " 請辦理變更代理人為「閻�K泰、林景郁」。"
'               If strExc(3) <> "" Then
'                  strExc(1) = strExc(1) & "總委任書存於:" & strExc(3)
'               End If
'
'               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                  " values ('" & strUserNum & "','" & strExc(2) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'                  ",'" & ChgSQL(strExc(1)) & "','" & ChgSQL(strExc(1)) & "')"
'               cnnConnection.Execute strSql, intI
'            End If
'         End If
         PUB_Add401ForGuiCase cp(1), cp(2), cp(3), cp(4), cp(9), cp(12), cp(13)
'end 2020/3/18
      End If
      'end 2020/3/16
      
      'Added by Morgan 2023/12/28
      If strSrvDate(1) >= 指定日期啟用日 And Text1(0) <> "" And cp(27) = "" Then
         If OptSendType(3).Value And (Option1(0).Value Or Option1(2).Value) Then
            If GetStaffDepartment(Text1(0)) <> "P12" Then 'Added by Morgan 2024/1/22 程序不必通知--郭
               strExc(0) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & _
                  "分案提醒 「本案需於指定日" & ChangeTStringToTDateString(txtCP142) & IIf(Option1(2).Value, "之後", "") & "方可送件」，請留意承辦時間！"
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc13)" & _
                  " values('" & strUserNum & "','" & Text1(0) & "',to_char(sysdate,'yyyymmdd')" & _
                  ",to_char(sysdate,'hh24miss'),'" & strExc(0) & "','" & cp(9) & "')"
               cnnConnection.Execute strSql, intI
            End If
         End If
      End If
      'end 2023/12/28
   
   End If
   
   'Add by Amy 2022/10/17 +接洽單電子化
   'Modify by Amy 2022/11/15 +急件
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
        If Check11.Visible = True Then
            stCP122 = "N"
            If Check11.Value = 1 Then stCP122 = "Y"
            If cp(122) <> stCP122 Then
                strSql = "Update CaseProgress Set CP122=" & CNULL(stCP122) & " Where cp09='" & cp(9) & "' "
                cnnConnection.Execute strSql
            End If
        End If
        If txtF0301 <> MsgText(601) And stF0309_Now <> Flow_已分案 Then
   'end 2022/11/15
            'Add By Sindy 2022/11/22 檢查接洽單全部案件性質是否全部分案完成
            If PUB_GetCP140CP157IsOK(cp(9)) = True Then
            '2022/11/22 END
                m_F0309 = Flow_已分案
                strUpdDate = strSrvDate(1)
                strUpdTime = Right("000000" & ServerTime, 6)
            
                '簽核檔(已處理)
                strSql = "update FLOW002 set " & _
                       "F0205='" & strUpdDate & "'" & _
                       ",F0206='" & strUpdTime & "'" & _
                       ",F0207='3',F0204='" & strUserNum & "'" & _
                       " where F0201='" & txtF0301 & "' and F0202='A7'  and F0207 is null "
                cnnConnection.Execute strSql
                '表單主檔
                strSql = "update FLOW003 set " & _
                        "F0309=" & CNULL(m_F0309) & _
                        " where F0301='" & txtF0301 & "' "
                cnnConnection.Execute strSql
            End If
        End If
   End If
   'end 2022/10/17

   cnnConnection.CommitTrans
   FormSave = True
   
   'Add by Morgan 2008/8/15
   If strPA14Msg <> "" Then
      MsgBox strPA14Msg
   End If
   
   'Add by Morgan 2009/7/15
   If st307Msg <> "" Then
      MsgBox st307Msg
   End If
   
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description
   End If
End Function

'Added by Morgan 2012/6/18
'重讀資料
'Modify by Amy 2022/12/26 +intOpen090801
Public Sub QueryMainFile(Optional ByVal intOpen090801 As Integer = 0)
   IntNow = IntNow - 1
   GetData IntNow, True, intOpen090801
End Sub

'Modify by Amy 2022/12/26 intOpen090801-0:第一次開接洽單/1-從其他表單關接洽單(不再開啟)
Private Sub GetData(intSitu As Integer, Optional bolNoMainForm As Boolean, Optional intOpen090801 As Integer = 0)
   Dim stTmpDate As String 'Add by Morgan 2004/9/27
   Dim rsTmp1 As New ADODB.Recordset, i As Integer, txt, Lbl
   Dim strCP09 As String ' 90.10.18 modify by louis (記錄總收文號)
   Dim ii As Integer '回圈流水號'Add By Cheng 2001/12/25
   Dim arrTmp 'Add by Amy 2022/10/17
   
   'Add By Cheng 2002/06/11
   Me.Text1(2).Enabled = False
   
   textPA1 = Empty
   textPA2 = Empty
   textPA3 = Empty
   textPA4 = Empty
   stF0307_Now = Empty: stF0309_Now = Empty 'Add by Amy 2022/10/17
   
   For Each txt In Text1
      txt.Text = ""
   Next
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   
   '總收文號
   Label3(8) = StrTot2(intSitu)
   ' 90.10.18 modify by louis (記錄總收文號)
   strCP09 = StrTot2(intSitu)
   '本所案號
   Label3(9) = StrTot1(intSitu)
   i = Len(Label3(9)) - 9
   pa(1) = Left(Label3(9), i)
   pa(2) = Mid(Label3(9), i + 1, 6)
   pa(3) = Mid(Label3(9), i + 7, 1)
   pa(4) = Right(Label3(9), 2)
   Combo1.Clear
   
   If pa(1) = "P" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then
         'If pa(23) = "1" Then          '2008/10/27 cancel by sonia P-083188準備程序
            Combo1.AddItem "中 : " & pa(5)
            Combo1.AddItem "英 : " & pa(6)
            Combo1.AddItem "日 : " & pa(7)
            Combo1.ListIndex = 0
         'End If
         Text1(3) = pa(23)
         Text1(8) = pa(48)
         If pa(57) = "Y" Then
            Label4 = "已閉卷"
            Label1(43).Visible = True
            Text1(11).Visible = True
            'Added by Morgan 2024/11/20 +檢查一案兩請新型案的退費是否不可取消閉卷--玲玲
            If pa(9) = 台灣國家代號 And cp(10) = "908" Then
               If PUB_ChkDualCase(pa()) = True Then
                  Text1(11).Enabled = False
               End If
            End If
            'end 2024/11/20
         Else
            Label4 = ""
            Label1(43).Visible = False
            Text1(11).Visible = False
         End If
         If pa(9) <> "" Then
            Text1(13) = pa(9)
            ChgType 13
         End If
         'Modify by Morgan 2006/5/23
         'Text1(14) = pA(46)
         If pa(9) <> "056" Then
            If pa(46) = "Y" Then
               'PCT申請日
               Text1(14) = TransDate(pa(10), 2)
               'PCT優先權日
               Text1(18) = PUB_GetPCTPriDate(pa(91))
               'Add by Morgan 2009/11/30
               'PCT優先權號
               Text1(23) = PUB_GetPCTPriNo(pa(91))
            End If
         End If
         'end 2006/5/23
         
         For i = 26 To 30
            If pa(i) <> "" Then ChgType i
         Next
         If pa(75) <> "" Then ChgType 75
         Text1(7) = pa(14)
         Text1(22) = pa(22) 'Add by Morgan 2007/8/30
         Text1(20) = pa(91)
         Text1(21) = pa(47)
         'Add By Cheng 2002/06/10
         'Me.Text1(2).Enabled = True 'Removed by Morgan 2014/11/3 移到下面
         If pa(8) <> "" Then Me.Text1(2).Text = pa(8)
         
         'Add by Morgan 2005/11/25 從下面搬上來，要跟服務業務分開
         GetCustom 26
         GetCustom 27
         GetCustom 28
         GetCustom 29
         GetCustom 30
         GetFagent 75
         
         'Add By Sindy 2010/10/29
         If pa(158) = "" Then
            Combo3 = ""
         Else
            'Modified by Morgan 2018/9/18 設計案會錯 Ex:P-119545
            'Combo3 = pa(158) + "." + PUB_GetCaseAttributeName(pa(158))
            Combo3 = pa(158) + "." + PUB_GetCaseAttributeName(pa(158), pa(8))
            'end 2018/9/18
         End If
         '2010/10/29 End
            
         'Added by Morgan 2017/11/20 設計案不可輸入案件屬性--陳玲玲
         If Text1(2) = "3" Then
            Combo3.Enabled = False
         Else
            Combo3.Enabled = True
         End If
         'end 2017/11/20
         
      End If
      txtEngGroup = pa(150) 'Added by Morgan 2012/3/9
   ElseIf pa(1) = "PS" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.ReadServicePracticeDatabase(pA(), intWhere) Then
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
         AddCboName Combo1, pa(5), pa(6), pa(7)
'            Text1(3) = pa(23)
         Text1(8) = pa(29)
         If pa(15) = "Y" Then
            Label4 = "已閉卷"
            Label1(43).Visible = True
            Text1(11).Visible = True
         Else
            Label4 = ""
            Label1(43).Visible = False
            Text1(11).Visible = False
         End If
         If pa(8) <> "" Then ChgType 8
         If pa(9) <> "" Then Text1(13) = pa(9): ChgType (13)
         If pa(26) <> "" Then ChgType 26
         Text1(20) = pa(18)
         Text1(21) = pa(28)
         
         'Add by Morgan 2005/11/25
         GetCustom 8
         GetCustom 58
         GetCustom 59
         GetFagent 26
      End If
      
      txtEngGroup = pa(79) 'Added by Morgan 2012/3/9
   End If
   txtEngGroup.Tag = txtEngGroup.Text 'Added by Lydia 2023/04/26
   m_strCP06 = ""
   m_strCP07 = ""
   
   '2007/8/13 ADD BY SONIA銷卷提醒
   CheckCaseDestroy pa(1), pa(2), pa(3), pa(4)
   '2007/8/13 END
   
   'Modify By Sindy 2014/1/29
   m_CP31isYGetCP05 = GetCP31isY_CP05(pa(1), pa(2), pa(3), pa(4)) '取得本所案號新案件的收文日
   'Added by Lydia 2020/03/31 事務所合併日起新案只能空白或J，不可輸T
   If Val(m_CP31isYGetCP05) >= 事務所合併日 Then
       lblPA161.Caption = "特殊出名公司                     (J:智權公司 空白:系統預設)"
   'end 2020/03/31
   
   'Add By Sindy 2013/12/16
   'If strSrvDate(1) < InvoiceStartDate Then
   'Modify by Amy 2016/08/12 +台灣判斷-秀玲
   'Modify by Amy 2017/07/13 服務業務 新案且非台灣 才可輸J或空白-秀玲
   'Modified by Lydia 2020/03/31 +Else
   ElseIf Val(m_CP31isYGetCP05) < Val(InvoiceStartDate) Then
      If pa(9) = "000" Then
        'Modified by Morgan 2015/5/12
        'lblPA161.Caption = "是否以專利商標出名         (Y:是)"
        lblPA161.Caption = "特殊出名公司                     (T:專利商標 空白:系統預設)"
        'end 2015/5/12
      End If
   ElseIf pa(1) = "PS" And pa(9) <> "000" Then
        lblPA161.Caption = "特殊出名公司                     (J:智權公司 空白:系統預設)"
   End If
   'end 2017/07/13
   '2013/12/16 END
   
   cp(9) = StrTot2(intSitu)
   strDateTmp(1) = ""
   strDateTmp(2) = ""
   'Modify by Morgan 2005/3/2 不再Call dll
   'If objPublicData.ReadCaseProgressDatabase(cp(), intWhere) Then
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      'Added by Morgan 2014/11/3
      'Modified by Morgan 2014/12/11 +3字頭
      'modify by sonia 2024/7/18 +101~103
      If cp(1) = "P" And (cp(31) = "Y" Or Left(cp(10), 1) = "3" Or (cp(10) >= "101" And cp(10) <= "103")) Then
         Me.Text1(2).Enabled = True
      End If
      'end 2014/11/3
      
      'Added by Morgan 2012/3/9
      Label3(13).Visible = False
      txtEngGroup.Visible = False
      Label1(24).Visible = False
      If Left(cp(12), 1) = "F" Then
         Label3(13).Visible = True
         txtEngGroup.Visible = True
         Label1(24).Visible = True
      End If
      'end 2012/3/9
      
      'Add by Morgan 2009/11/4
      '設定是否 FMP 案
      If Left(cp(12), 1) = "F" And pa(9) <> "000" Then
         m_bolFMP = True
      Else
         m_bolFMP = False
      End If
      
      'Added by Lydia 2021/05/31 是否為寰華案
      m_bolFMP2 = False
      If m_bolFMP = True Then
          'Modified by Lydia 2023/02/21 改成模組
          'm_bolFMP2 = PUB_FMPtoCheck(0, 2, Pub_StrUserSt03, pa(1), pa(2), pa(3), pa(4)) '參考用;因為分案時有可能新案尚未發文,所以拿掉發文條件
          'strExc(0) = "select CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12 from caseprogress  where CP01='" & pa(1) & "' and CP02='" & pa(2) & "' and CP03='" & pa(3) & "' and CP04='" & pa(4) & "' and CP01 in ('P','PS') and CP31='Y' and SUBSTR(CP12,1,1) = 'F' and CP44='Y53374000' "
          'intI = 1
          'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
          'If intI = 1 Then
          '    m_bolFMP2 = True
          'End If
          m_bolFMP2 = PUB_GetFMP2toP(pa(1), pa(2), pa(3), pa(4))
          'end 2023/02/21
      End If
      'end 2021/05/31
      
      OptSendType(1).Caption = PUB_GetCP114Opt1Desc(cp(1), cp(10)) 'Added by Morgan 2024/1/19
      
      'Added by Morgan 2023/8/29
      If m_bolFMP Then
         OptSendType(1).Enabled = False
         OptSendType(2).Enabled = False
      Else
         OptSendType(1).Enabled = True
         OptSendType(2).Enabled = True
      End If
      If cp(27) = "" Then
         Frame1.Enabled = True
         Frame2.Enabled = True 'Added by Morgan 2024/1/5
      Else
         Frame1.Enabled = False
         Frame2.Enabled = False 'Added by Morgan 2024/1/5
      End If
      Option1(0).Enabled = False
      Option1(1).Enabled = False
      Option1(2).Enabled = False
      'end 2023/8/29
      
      'Added by Morgan 2025/1/23
      Option1(0).Value = False
      Option1(1).Value = False
      Option1(2).Value = False
      'end 2025/1/23
      
      'Added by Morgan 2023/8/31
      Frame2.Visible = True
      If Not m_bolFMP Then
         If strSrvDate(1) < 指定日期啟用日 Then 'Added by Morgan 2023/12/28
            Frame2.Visible = False
         End If
      End If
      'end 2023/8/31
      
      'Add by Morgan 2010/12/29
      txtCP142 = "" 'Added by Morgan 2014/7/17
      Select Case cp(141)
         Case "1"
            OptSendType(1).Value = True
         Case "2"
            OptSendType(2).Value = True
         Case "3"
            OptSendType(3).Value = True
            txtCP142.Text = TransDate(cp(142), 1)
            'Added by Morgan 2023/8/29
            Option1(0).Enabled = True
            Option1(1).Enabled = True
            Option1(2).Enabled = True
            If cp(164) = "1" Then
               Option1(0).Value = True
            ElseIf cp(164) = "2" Then
               Option1(1).Value = True
            ElseIf cp(164) = "3" Then
               Option1(2).Value = True
            End If
            'end 2023/8/29
         'Add by Morgan 2011/2/8 要清除否則按下一筆分案會預設前筆狀態
         Case Else
            OptSendType(1).Value = False
            OptSendType(2).Value = False
            OptSendType(3).Value = False
      End Select
      'end 2010/12/29
      
      If cp(13) <> "" Then Text1(24) = cp(13): ChgType (24)
      
      'Modify by Amy 2015/01/22 操作人員為北所且北所分案日期有值才顯示辦人
       'Add By Cheng 2003/10/29記錄原承辦人欄位值
       Me.Text1(0).Tag = cp(14)
       'End 2003/10/29
      If pub_strUserOffice = "1" And cp(157) = "" Then
            m_bolIsFirstKeyCP14 = True
            'Modify by Amy 2022/11/03 +if 接洽單電子收文上線後直接顯示(cp14=cra09),不需再輸
            If strSrvDate(1) >= 接洽單電子收文啟用日 Then
               Me.Text1(0).Tag = "" 'Add By Sindy 2022/12/27 因為電子收文上線畫面上承辦人會顯示出來,反過來Tag=空白,後續程式在比對承辦人才會是不一致的
            Else
               cp(14) = ""
            End If
      End If
      'end 2015/01/22
     
      If cp(14) <> "" Then
         Text1(0) = cp(14)
      'Added by Morgan 2025/5/23
      'P/CFP 設定為程序承辦且不需專業部主管分案的性質，承辦人都自動設預設為程序人員，若有需要，分案人員再自行修改，但CFP實體審查除外--郭
      'Modified by Morgan 2025/5/29 排除寰華案及FMP的實審
      'Modified by Lydia 2025/06/19 改模組名稱
      'ElseIf PUB_GetCPM35byCP10(cp(1), cp(10)) = "2" And m_bolFMP2 = False And Not (m_bolFMP And cp(10) = "416") Then
      ElseIf PUB_GetCPMbyCP10(cp(1), cp(10), "cpm35") = "2" And m_bolFMP2 = False And Not (m_bolFMP And cp(10) = "416") Then
         Text1(0) = PUB_GetPHandler(cp(1) & cp(2) & cp(3) & cp(4))
      'end 2025/5/23
      End If
            
      If Text1(0) <> "" Then ChgType (0)
      
      Text1(10) = cp(26) 'Modified by Morgan 2013/3/14 從下面移上來,必須放在案件性質前否則是否計件預設後會被蓋掉(Ex.延期)
      If cp(10) <> "" Then
         Text1(1) = cp(10)
         ChgType 1
      End If
      Text1(4) = cp(6): If cp(6) <> "" Then m_strCP06 = Val(cp(6)) + 19110000
      strDateTmp(1) = cp(6)
      Text1(5) = cp(7): If cp(7) <> "" Then m_strCP07 = Val(cp(7)) + 19110000
      '2008/10/24 add by Toni
      Text1(4).Tag = cp(6): If cp(6) <> "" Then m_strCP06 = Val(cp(6)) + 19110000
      Text1(5).Tag = cp(7): If cp(7) <> "" Then m_strCP07 = Val(cp(7)) + 19110000
      '2008/10/24
      strDateTmp(2) = cp(7)
      Text1(6) = cp(43)
      Text1(9) = cp(57)
      Text1(12) = cp(5)
      Text1(19) = cp(64)
      Text1(16) = cp(18)
      Text1(25) = TransDate(cp(48), 1) 'Add by Morgan 2010/3/15
      'Remove by Morgan 2007/8/30 改抓基本檔
      'Text1(17) = cp(36)
      Text1(17) = pa(11)
      'end 2007/8/30
      m_CP06 = cp(6)
      m_CP07 = cp(7)
      m_CP10 = cp(10)
      m_CP30 = cp(30) 'Add by Morgan 2011/4/22
      m_CP31 = cp(31) 'Add By Sindy 2010/10/29
      'Add by Morgan 2010/3/17
      If cp(10) = "123" Then
        ' lblFavDt.Visible = True
         txtFavDt.Visible = True
         'Modified by Morgan 2012/3/22 改抓 pa140
         'txtFavDt.Text = TransDate(PUB_GetFavorDate(pa(91)), 1)
         txtFavDt.Text = TransDate(pa(140), 1)
         CmdFav.Visible = True 'Add by Lydia 2015/02/02
      Else
        ' lblFavDt.Visible = False
         txtFavDt.Visible = False
         CmdFav.Visible = False 'Add by Lydia 2015/02/02
      End If
     
      '2012/7/19 ADD BY SONIA 記錄是否電子送件欄位值
      'Modified by Morgan 2013/5/14 電子送件有可能為 W 或 Y
      'Me.txtCP118.Tag = cp(118)
      'Me.txtCP118.Text = cp(118)
      If cp(118) <> "" Then
         txtCP118 = "Y"
      Else
         txtCP118 = ""
      End If
      txtCP118.Tag = txtCP118
      'end 2013/5/14
      
      Label3(14) = cp(27) 'Added by Morgan 2023/8/30
      
      Me.txtCP118.Enabled = False
      Me.txtCP118.Locked = True
      
      'Modified by Morgan 2013/5/5/15 已開收據不可變更是否電子送件
      'If pa(9) = "000" And InStr(NewCasePtyList, Text1(1)) > 0 And cp(27) = "" Then
      'Modified by Morgan 2013/9/12 台灣案除領證年費外皆可設定電子送件--郭(電子承辦單會議)
      'If pa(9) = "000" And InStr(NewCasePtyList, Text1(1)) > 0 And cp(27) = "" And cp(60) = "" Then
      'Modified by Morgan 2013/10/23 +232 補優先權證明--陳玲玲
      'Modified by Morgan 2014/10/17 +421 申請技術報告,807 申請第三人技術報告
      'Modified by Morgan 2015/1/16 +941分析
      'Modfiy by Amy 2016/06/28 +501 訴願/503 行政訴訟/803 舉發/804 舉發答辯
      'Modified by Morgan 2018/1/4 -232 補優先權證明--陳玲玲
      'Modified by Morgan 2020/4/13 -601,605--陳玲玲
      'Modified by Morgan 2020/7/28 -421 申請技術報告--潘韻丞
      
      '********** 若案件性質有調整時CASEPROGRESS_AFTER1(Trigger)也要同步修改-寫於此位置不要改 **********
      If pa(9) = "000" And Text1(1) <> "807" And Text1(1) <> "941" _
        And Text1(1) <> "501" And Text1(1) <> "503" And Text1(1) <> "803" And Text1(1) <> "804" And cp(27) = "" Then
        
         If InStr(NewCasePtyList, Text1(1)) > 0 Then
            If cp(60) = "" Then
               Me.txtCP118.Enabled = True
               Me.txtCP118.Locked = False
            End If
         Else
            Me.txtCP118.Enabled = True
            Me.txtCP118.Locked = False
            If txtCP118 = "" Then
               'Added by Morgan 2020/4/14 台灣領證年費預設電子送件--陳玲玲
               If Text1(1) = "601" Or Text1(1) = "605" Then
                  MsgBox Label3(1) & "將預設為電子送件！", vbExclamation
                  txtCP118 = "Y"
               Else
               'end 2020/4/14
               
                  strExc(0) = "select 1 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP10 IN (" & NewCasePtyList & ") and cp118 is not null"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     MsgBox "本案為電子送件案，本程序將預設為電子送件！", vbExclamation
                     txtCP118 = "Y"
                  End If
               End If 'Added by Morgan 2020/4/14
            End If
         End If
         'end 2013/9/12
      'Added by Morgan 2024/1/30 開放大陸案也可設定電子送件(收文時Trigger會預設Y)
      ElseIf pa(9) = "020" And cp(27) = "" Then
         txtCP118.Enabled = True
         txtCP118.Locked = False
      'end 2024/1/30
      End If
      '2012/7/19 end
   End If
   
   'Modify by Morgan 2007/8/30 加第三人申請技術報告807
   Text1(22).Enabled = False
   'If cp(10) <> 異議_專 And cp(10) <> 舉發 And cp(10) <> 鑑定報告 Then
   If cp(10) <> 異議_專 And cp(10) <> 舉發 And cp(10) <> 鑑定報告 And cp(10) <> "807" Then
      Text1(7).Enabled = False
      Text1(17).Enabled = False
      'Modify by Morgan 2007/8/31 不必清，存檔時控制
      'Text1(7) = ""
      'Text1(17) = ""
      'end 2007/831
      'Add by Morgan 2004/6/3
      '延緩公告延緩月數輸入控制
      CP71Switch Text1(13).Text, Text1(1).Text
   Else
      Text1(7).Enabled = True
      Text1(17).Enabled = True
      'Add by Morgan 2007/8/30 加證書號數
      If cp(10) = "807" Then
         Text1(22).Enabled = True
      End If
      'end 2007/8/30
   End If
   
   For i = 1 To 4
      cm(i - 1) = pa(i)
   Next
   '**************************************************************************************
   '2006/3/7 加註 BY SONIA,此處增加案件性質時,下列程式也要加
   'frm040104_1,frm040104_1_1,frm040104_3及
   'basPublic的ReadCaseRelationRst,InsertCaseRelationData,GetCaseRelationDataIn,
   '           GetCaseRelationDataOut
   '**************************************************************************************
   'edit by nickc 2005/06/07 加入案件性質 110 & 112
   'If (Text1(1) = "101" Or Text1(1) = "102" Or Text1(1) = "103" Or Text1(1) = "104" Or Text1(1) = "105" Or Text1(1) = "201" Or Text1(1) = "109") Then
   'Modify by Morgan 2006/5/9 案件性質改用常數控制
   'If (Text1(1) = "101" Or Text1(1) = "102" Or Text1(1) = "103" Or Text1(1) = "104" Or Text1(1) = "105" Or Text1(1) = "201" Or Text1(1) = "109" Or Text1(1) = "110" Or Text1(1) = "112") Then
   'Remove by Lydia 2016/07/07 改成模組 PUB_GetPcmNo
   'If InStr(CaseMapIn, Text1(1)) > 0 Then
   '   'edit by nickc 2007/02/05 不用 dll 了
   '   'If obj003.GetCaseMap(cm) = True Then Text1(15) = cm(4) & cm(5) & cm(6) & cm(7)
   '   If Cls003GetCaseMap(cm) = True Then Text1(15) = cm(4) & cm(5) & cm(6) & cm(7)
   '   'Add by Morgan 2006/6/9
   '   If Text1(15) = "" And Text1(13) = "013" Then
   '      'edit by nickc 2007/02/05 不用 dll 了
   '      'If obj003.GetCaseMap(cm, 4) = True Then Text1(15) = cm(4) & cm(5) & cm(6) & cm(7)
   '      If Cls003GetCaseMap(cm, 4) = True Then Text1(15) = cm(4) & cm(5) & cm(6) & cm(7)
   '   End If
   '   'Added by Lydia 2015/09/09 +澳門發明案與大陸案之關聯
   '   If Text1(15) = "" And Text1(13) = "044" Then
   '      If Cls003GetCaseMap(cm, 5) = True Then Text1(15) = cm(4) & cm(5) & cm(6) & cm(7)
   '   End If
   'End If
   'Added by Lydia 2016/07/07
   'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
   Call PUB_GetPcmNo(Text1(15), cm, Text1(13).Text, Text1(1).Text)
   
   'Remove by Morgan 2006/6/9 移到上面
'   'add by nickc 2005/06/07 若是沒有，抓 cm10='4' 的案子
'   If Text1(15) = "" Then
'      If obj003.GetCaseMap(cm, 4) = True Then Text1(15) = cm(4) & cm(5) & cm(6) & cm(7)
'   End If
   
   GetGrid StrTot2(intSitu), 0
   
   If pa(1) = "P" Then
      i = 專利
   Else
      i = 0
   End If
   
   'edit by nickc 2007/02/02 不用 dll 了
   'If Not objPublicData.ReadPriority(pA, strPriority(1), strPriority(2), strPriority(3)) Then
   'Modify by Morgan 2007/4/24 加strPriority(4)
   'Modify by Amy 2014/06/10 +strPriority(5)
   'Modify by Amy 2023/01/05 strPriority原陣列,改變數
   If Not ClsPDReadPriority(pa, strPrity1, strPrity2, strPrity3, strPrity4, strPrity5) Then
      
   End If

   IntNow = IntNow + 1
   frm040101_1.Refresh
   
   'Modify By Cheng 2002/01/24
   '若卷宗性質為申請(1), 才要顯示專利基本檔畫面
   If Me.Text1(3).Text = "1" Then
   
      If bolNoMainForm = False Then 'Added by Moran 2012/6/18
         'Modify by Morgan 2004/10/14    加 121-主張國內優先權
         'Modify by Morgan 2007/4/27 加急件費(920)--玲玲
         'If cp(10) <> 實體審查 And cp(10) <> 翻譯 And cp(10) <> 其他 And cp(10) <> 專利調查 And cp(10) <> 補收款 And cp(10) <> 主動修正 And cp(10) <> 主張優先權 And cp(10) <> 提早公開 And cp(10) <> 調卷 And cp(10) <> 回覆代理人 And cp(10) <> "307" And cp(10) <> "916" And cp(10) <> "917" And cp(10) <> "117" And cp(10) <> "405" And cp(10) <> "121" Then
         'Modify by Morgan 2009/12/23 +936,404
         'Modify by Morgan 2010/1/6 +938,939
         '2010/2/4 MODIFY BY SONIA 加改請獨立306
         If cp(10) <> 實體審查 And cp(10) <> 翻譯 And cp(10) <> 其他 And cp(10) <> 專利調查 And cp(10) <> 補收款 And cp(10) <> 主動修正 And cp(10) <> 主張優先權 And cp(10) <> 提早公開 And cp(10) <> 調卷 And cp(10) <> 回覆代理人 And cp(10) <> "307" And cp(10) <> "916" And cp(10) <> "917" And cp(10) <> "117" And cp(10) <> "405" And cp(10) <> "121" And cp(10) <> "920" And cp(10) <> "936" And cp(10) <> "404" And cp(10) <> "938" And cp(10) <> "939" And cp(10) <> "306" Then
            ShowMaintainForm strCP09, , , Me
         End If
         '91.10.31 END
      End If
      
   '卷宗性質不為申請時案件名稱抓對造資料
   Else
      '2008/10/27 add by sonia 再加案件性質條件,P-083188之準備程序
      If Text1(1) = 異議_專 Or Text1(1) = 舉發 Or Text1(1) = "807" Then
      '2008/10/27 end
         Combo1.Clear
         Combo1.AddItem "中 : " & cp(37)
         Combo1.AddItem "英 : " & cp(38)
         Combo1.AddItem "日 : " & cp(39)
         Combo1.ListIndex = 0
      End If
   End If
   
   'Add by Morgan 2003/12/07
   Call PUB_CheckSales(cp(1), cp(2), cp(3), cp(4), cp(5), cp(13), Label3(10))
   'End 2003/12/07
   
   'Add by Morgan 2004/3/29
   Call GetDivCase
   
   'Added by Lydia 2016/10/13 分割案若有主張優先權於分案時，不必輸入優先權資料，請直接帶母案的優先權資料P-114427(母案P-104573)
   'Modify by Amy 2023/01/05 strPriority(1)
   If strPrity1 = "" And txtDivCaseNo(1) & txtDivCaseNo(2) <> "" And Me.Text1(1) = 主張優先權 Then
        strExc(1) = txtDivCaseNo(1)
        strExc(2) = txtDivCaseNo(2)
        strExc(3) = txtDivCaseNo(3)
        strExc(4) = txtDivCaseNo(4)
        'Modify by Amy 2023/01/05 strPriority原陣列,改變數
        If Not ClsPDReadPriority(strExc, strPrity1, strPrity2, strPrity3, strPrity4, strPrity5) Then
        End If
   End If
   'end 2016/10/13
   
   'edit by nickc 2005/07/05 只適用於未准駁前之案件
   'If cp(10) = "203" And pa(9) = "000" And pa(8) = "2" Then  '主動修正
   'Modify by Morgan 2006/5/3
   '台灣發明or新型的主動修正期限為申請日(最早優先權日)起算
   '大陸新型or設計的主動修正期限為申請日起2個月內
   'If cp(10) = "203" And pA(9) = "000" And pA(8) = "2" And pA(16) = "" Then  '主動修正
   If cp(10) = "203" Then
      If Text1(4).Text = "" And Text1(5).Text = "" Then
         If pa(9) = "000" Then
         
'Removed by Morgan 2014/10/30
'            'Added by Morgan 2012/7/10 102新法
'            '1020101以後無期限
'            If Val(strSrvDate(1)) < 20130101 Then
'            'end 2012/7/10
'
'               '台灣新型
'               If pa(8) = "2" Then
'                  '申請日起兩個月
'                  stTmpDate = DBDATE(CompDate(1, 2, pa(10)))
'                  Text1(5).Text = ChangeWStringToTString(stTmpDate)
'                  '本所=法定-2天
'                  Text1(4).Text = ChangeWStringToTString(PUB_GetWorkDay1(CompDate(2, -2, stTmpDate), True))
'               '台灣發明
'               ElseIf pa(8) = "1" Then
'                  '申請日(最早優先權日)起15個月內
'                  strFirstPriDate = PUB_GetFirstPriDate(cp) 'Modify by Morgan 2006/5/12 改Call共用Function
'                  If strFirstPriDate <> "" Then
'                     stTmpDate = CompDate(1, 15, strFirstPriDate)
'                  Else
'                     stTmpDate = CompDate(1, 15, pa(10))
'                  End If
'                  Text1(5).Text = TransDate(stTmpDate, 1)
'                  '所限
'                  strExc(1) = pa(1)
'                  strExc(2) = pa(9)
'                  strExc(3) = stTmpDate
'                  GetCtrlDT strExc
'                  Text1(4).Text = TransDate(PUB_GetWorkDay1(strExc(0), True), 1)
'               End If
'
'            'Added by Morgan 2012/7/10 102新法
'               '法定期限為1020101以後者也不用設
'               If Val(DBDATE(Text1(5))) > 20130000 Then
'                  Text1(5) = ""
'                  Text1(4) = ""
'               End If
'            End If
'            'end 2012/7/10
            
         '大陸
         ElseIf pa(9) = "020" Then
                  
            '新型or設計
            If (pa(8) = "2" Or pa(8) = "3") Then
               'Modified by Morgan 2017/3/6 內部收文也要用改寫函數
'               '法限=申請日起兩個月
'               stTmpDate = CompDate(1, 2, pa(10))
'
'               'Modify by Morgan 2006/7/4
'               '法限=所限=申請日起兩個月-10天(非假日)
'               'Text1(5).Text = TransDate(stTmpDate, 1)
'               ''所限
'               'strExc(1) = pA(1)
'               'strExc(2) = pA(9)
'               'strExc(3) = stTmpDate
'               'GetCtrlDT strExc
'               'Text1(4).Text = TransDate(strExc(0), 1)
'               stTmpDate = CompDate(2, -10, stTmpDate)
'               stTmpDate = PUB_GetWorkDay1(stTmpDate, True)
'               Text1(5).Text = TransDate(stTmpDate, 1)
'               Text1(4).Text = Text1(5).Text
               If PUB_GetNon101CN203Date(pa(10), strExc(0), stTmpDate) Then
                  Text1(5).Text = TransDate(stTmpDate, 1)
                  Text1(4).Text = TransDate(strExc(0), 1)
               End If
               'end 2017/3/6
            End If
            'Add by Morgan 2010/6/2
            '大陸主動修正若無期限時要以通知進入實審(1214,1215,1204)的官方發文日+3月計算,所限=法限-7天
            If Text1(5) = "" Then '法限
               'Modified by Lyda 2015/09/09 +CP04
               strExc(0) = "select CP133 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP10 IN ('1214','1215','1204') ORDER BY CP05 DESC"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If Not IsNull(RsTemp(0)) Then
                     strExc(1) = CompDate(1, 3, RsTemp(0))
                     'Added by Lydia 2025/10/29
                     If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                        strExc(2) = PUB_GetPOurDeadline(strExc(1), pa(9))
                     Else
                     'end 2025/10/29
                        strExc(2) = PUB_GetWorkDay1(CompDate(2, -7, strExc(1)), True)
                     End If 'Added by Lydia 2025/10/29
                     If Val(strExc(2)) < Val(strSrvDate(1)) Then
                        strExc(2) = strSrvDate(1)
                     End If
                     Text1(5) = TransDate(strExc(1), 1)
                     Text1(4) = TransDate(strExc(2), 1)
                  End If
               'Added by Lydia 2015/09/09 官方未發文
               Else  '大陸主動修正的分案時，該案不管進度檔或下一程序檔有實審期限時，則同時將實審期限更新至主動修正的期限。
                    strExc(0) = "select CP06,CP07 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP10 ='416' and cp57 is null " & _
                         "UNION select NP08,NP09 from nextprogress where NP02='" & cp(1) & "' and NP03='" & cp(2) & "' AND NP04='" & cp(3) & "' AND NP05='" & cp(4) & "' AND NP07 ='416' and (np06='Y' or np06 is null) "
                    intI = 1
                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                    If intI = 1 Then
                       RsTemp.MoveFirst
                       Do While Not RsTemp.EOF
                          If Not IsNull(RsTemp.Fields(0)) And Not IsNull(RsTemp.Fields(1)) Then
                             Text1(4) = TransDate(RsTemp.Fields(0), 1)
                             Text1(5) = TransDate(RsTemp.Fields(1), 1)
                             Exit Do
                          End If
                          RsTemp.MoveNext
                       Loop
                    End If
               'end 2015/09/09
               End If
            End If
         End If
      End If
   End If
   '2006/5/3 end

   'Add by Morgan 2005/3/4
   '若加乘註記空白則帶預設
   m_stCP98 = cp(98)
   If m_stCP98 = "" Then
      If PUB_GetFlagValue(cp(9), cp(98), cp(101), cp(104)) = False Then
         MsgBox "無法讀取取得加乘註記預設值！", vbCritical
      End If
   End If
   txtCP98 = cp(98)
   txtCP99 = cp(99)
   '2005/3/4 end
   
   'Add by Morgan 2010/3/16
   'Modify by Morgan 2011/8/18
   '非台灣案不預設
   If pa(9) = "000" Then
      txtFeeYear(1) = cp(53)
      txtFeeYear(2) = cp(54)
   Else
      txtFeeYear(1) = ""
      txtFeeYear(2) = ""
   End If
   'end 2010/3/16
   
   'Add by Morgan 2007/12/11
   Text1(13).Tag = Text1(13)
   Text1(15).Tag = Text1(15)
   'end 2007/12/11

   '2008/11/27 ADD BY SONIA 226配合開庭不可改案件性質,鎖住本所期限,本所=法定-2天
   If cp(10) = "226" Then
      Text1(1).Enabled = False
      Text1(4).Enabled = False
   Else
      Text1(1).Enabled = True
      Text1(4).Enabled = True
   End If
   '2008/11/27 END
   
   'Add by Morgan 2010/10/29
   m_strOldCP10 = Text1(1)
   m_strOldPA09 = Text1(13)
   'end 2010/10/29

   
   'Add by Morgan 2009/12/23 延期不可改案件性質
   MSHFlexGrid1.Enabled = True
   'Modified by Morgan 2020/2/7 +442 在途期限
   If cp(10) = "404" Or cp(10) = "442" Then
      'Text1(1).Enabled = False
      'Text1(10) = "N" 'Added by Morgan 2013/3/14 要預設不計件
      If cp(27) <> "" Then
         MSHFlexGrid1.Enabled = False '已發文不可再點選
      End If
   Else
      'Text1(1).Enabled = True   '2011/3/14 CANCEL BY SONIA 因為上面配合開庭226已鎖住
   End If
   
   '2010/2/4 ADD BY SONIA
   '改請聯合305時提醒應收文原本所案號,轉入新案號後系統自動發MAIL通知電腦中心合併案號
   If cp(10) = "305" And cp(3) = "0" Then
      MsgBox "改請聯合，請確認是否收在原本所案號！" & vbCrLf & "並請先將此程序轉至聯合案案號，以便系統發MAIL通知電腦中心合併案號！", vbExclamation, "改請聯合分案提醒"
   End If
   'Added by Morgan 2012/12/19 +改請衍生設計
   If cp(10) = "308" And cp(3) = "0" Then
      MsgBox "改請衍生設計，請確認是否收在原本所案號！" & vbCrLf & "並請先將此程序轉至衍生設計案案號，以便系統發MAIL通知電腦中心合併案號！", vbExclamation, "改請衍生設計分案提醒"
   End If
   
'cancel by sonia 2014/7/28 已於2010/2/4加於Process
'   '改請獨立306且案號未合併時提醒人工發MAIL通知電腦中心合併案號
'   If cp(10) = "306" Then
'      strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' AND CP10 IN (" & CaseMapIn & ") and cp27>0"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'      If intI = 1 Then  '案號未合併
'      Else
'         'Modified by Morgan 2012/12/19 +/衍生設計
'         MsgBox "請人工發MAIL通知電腦中心原聯合/衍生設計案案號及改請獨立新案號，以便做案號合併處理！", vbExclamation, "改請獨立分案提醒"
'      End If
'   End If
'   '2010/2/4 END
'end 2014/7/28
   
   'Added by Morgan 2012/11/14
   'Modified by Morgan 2020/2/3 +只限台灣案--陳玲玲 Ex:P119886
   If cp(10) = "413" And cp(27) = "" And pa(9) = "000" Then
      strExc(1) = PUB_GetFirstPriDate(cp)
      If strExc(1) = "" Then strExc(1) = DBDATE(pa(10))
      strExc(2) = CompDate(1, 15, strExc(1))
      If strSrvDate(1) > strExc(2) Then
         MsgBox "已超過申請日(優先權日)起算15個月！", vbExclamation
      End If
   End If
   'end 2012/11/14
         
'Removed by Morgan 2024/2/22 接洽單已可設定送件方式，此處無須再自動預設，且已收款通知已改規則
'   'Add by Morgan 2010/3/17
'   '收款後送件控制有費用才要紀錄
'   If Val(cp(16)) > 0 And Val(cp(79)) > 0 And cp(27) = "" Then
'      'Modified by Morgan 2013/10/23 考慮程序新人
'      'If cp(14) = "81002" Or cp(14) = "73017" Then
'      'Modified by Morgan 2015/6/3 81002 權限已改但仍會承辦
'      If cp(14) = "81002" Or PUB_GetST05(cp(14)) = "75" Then
'      'end 2013/10/23
'         'Modify by Morgan 2010/12/29 改用 Option
'         'txtTakeCtrl.Locked = False
'         Frame1.Enabled = True
'         strExc(0) = "select * from UndeliveredRec where ud01='" & cp(9) & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            'Modify by Morgan 2010/12/29 改用 Option
'            'txtTakeCtrl.Text = "Y"
'            OptSendType(2).Value = True
'         End If
'
'      'Modified by Morgan 2023/8/31
'      'Else
'      ElseIf Not m_bolFMP Then
'      'end 2023/8/31
'         'txtTakeCtrl.Locked = True
'         If strSrvDate(1) < 指定日期啟用日 Then 'Added by Morgan 2023/12/28
'            Frame1.Enabled = False
'         End If
'      End If
'
'   'Modified by Morgan 2023/8/31
'   'Else
'   ElseIf Not m_bolFMP Then
'   'end 2023/8/31
'      'Modify by Morgan 2010/12/29 改用 Option
'      'txtTakeCtrl.Enabled = False
'      If strSrvDate(1) < 指定日期啟用日 Then 'Added by Morgan 2023/12/28
'         Frame1.Enabled = False
'      End If
'   End If
'end 2024/2/22
   
   'Add by Morgan 2010/6/18
   '若已請款則紀錄原核稿人以便判斷是否要通知重新分配點數
   If cp(60) > "X" And cp(10) = "201" Then
      strExc(0) = "select ep04 from engineerprogress where ep02='" & cp(9) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_oldEP04 = "" & RsTemp(0)
      End If
   End If
         
   'Add by Morgan 2010/8/18
   '分析預設本所期限=承辦期限=收文日+1週
   If cp(10) = "941" Then
      'Modified by Morgan 2017/6/29 改原來無承辦人時不管原來有無期限都預設為收文日+5個工作天(不含收文日)
      'If Text1(4) = "" Or Text1(25) = "" Then
      '   strExc(2) = PUB_GetWorkDay1(CompDate(2, 7, cp(5)), True)
      If Text1(0).Tag = "" And cp(27) = "" Then
         strExc(2) = CompWorkDay(5, CompDate(2, 1, cp(5)))
      'end 2017/6/29
         If Val(strExc(2)) < Val(strSrvDate(1)) Then
            strExc(2) = strSrvDate(1)
         End If
         strExc(2) = TransDate(strExc(2), 1)
         
         'Modified by Morgan 2017/6/29
         'If Text1(4) = "" Then
         '   Text1(4) = strExc(2)
         'End If
         Text1(4) = strExc(2)
         'end 2017/6/29
         
         If PUB_IfSetCP48(cp(9)) Then 'Add by Morgan 2010/10/1
            If Text1(25) = "" Then
               Text1(25) = strExc(2)
               cp(48) = DBDATE(strExc(2))
            End If
         End If 'Add by Morgan 2010/10/1
      End If
   End If
   'end 2010/8/18
   
   'Add by Morgan 2011/8/17
   '大陸領證抓相關收文號的繳費起始日
   m_RefCP53 = ""
   If pa(9) = "020" And cp(10) = "601" And cp(43) <> "" Then
      strExc(0) = "select cp53 from caseprogress where cp09='" & cp(43) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_RefCP53 = "" & RsTemp(0)
      End If
   End If
   
   Set434Date 'Added by Morgan 2012/3/30
   'Added by Lydia 2020/05/20 法律所案源收文
   Call ReadLOS
   Call SetLOSagree
   
   'Add by Morgan 2004/2/9
   '取消指定國家
   'Add by Morgan 2006/4/21
   '紀錄原案件性質,專利種類
   Text1(1).Tag = Text1(1)
   Text1(2).Tag = Text1(2)
   '2006/4/21 end
   
   'Add by Morgan 2010/4/12
   '年費已發文則繳費年度不可改
   If cp(27) <> "" Then
      txtFeeYear(1).Enabled = False
      txtFeeYear(2).Enabled = False
   End If
      
   'Added by Morgan 2012/6/18
   '若為併號請以連絡單通知電腦中心處理
   If cp(27) <> "" Or strReceiveNo > "C" Then
      textPA1.Enabled = False
      textPA2.Enabled = False
      textPA3.Enabled = False
      textPA4.Enabled = False
   Else
      textPA1.Enabled = True
      textPA2.Enabled = True
      textPA3.Enabled = True
      textPA4.Enabled = True
   End If
   'end 2012/6/18
   
   'Added by Morgan 2012/7/20
   '是否為複雜或特殊案件
   txtCP147 = cp(147)
   txtCP147.Tag = txtCP147
   If cp(14) = "" And txtCP147 = "" Then txtCP147 = GetCP147Default()
   'end 2012/7/20
   
   'Added by Morgan 2012/9/12
   '新案可設定是否以專利商標出名欄位
   'Modified by Morgan 2013/5/16 +改請案也要可以選(3字頭,307是新案也適用)
   'If cp(1) = "P" And cp(31) = "Y" Then
   'Modify by Amy 2017/07/13 服務業務 新案且非台灣 就顯示專利商標出名欄位-秀玲
   If (cp(1) = "P" And (cp(31) = "Y" Or (pa(9) = "000" And Left(cp(10), 1) = "3"))) Or (cp(1) = "PS" And cp(31) = "Y" And pa(9) <> "000") Then
      lblPA161.Visible = True
      txtPA161.Visible = True
      'Modify by Amy 2017/07/13 開放服務業務也顯示顯示專利商標出名欄位
      If cp(1) = "P" Then
        txtPA161 = pa(161)
      Else
        txtPA161 = pa(85)
      End If
      'Modify by Amy 2016/08/29
      txtPA161.Tag = txtPA161
      'Add by Amy 2016/08/12 +客戶檔收據公司別
      'Mark by Amy 2017/11/24 CFP-29915 申請人出名公司為J公司,個案改開專利法律(pa161空白),因個案為空白會預設申請人出名公司-秀玲:拿掉
'      If txtPA161 = MsgText(601) Then
'        If cp(1) = "P" Then
'            txtPA161 = GetReceiptCmp(Left(GetNewFagent(pa(26)), 8), Mid(GetNewFagent(pa(26)), 9, 1), cp(1), pa(9))
'        ElseIf pa(8) <> MsgText(601) Then
'            txtPA161 = GetReceiptCmp(Left(GetNewFagent(pa(8)), 8), Mid(GetNewFagent(pa(8)), 9, 1), cp(1), pa(9))
'        End If
'      End If
      'end 2016/08/12
      'end 2017/07/13
       'Add by Amy 2018/07/03 收文日在一個月之內才可修改 特殊出名公司
      txtPA161.Enabled = False
      If Val(strSrvDate(1)) >= Val(cp(5)) + 19110000 And Val(strSrvDate(1)) <= Val(DBDATE(DateAdd("m", 1, Format(Val(cp(5)) + 19110000, "####/##/##")))) Then
        txtPA161.Enabled = True
      End If
   Else
      lblPA161.Visible = False
      txtPA161.Visible = False
   End If
   'end 2012/9/12
  
   txtCP97 = cp(97) 'Add by amy 2014/09/05 增加承辦人計件值欄位讓user修改-玲玲
   
   Set30xDate 'Added by Morgan 2012/10/4
   
   txtPA178 = pa(178) 'Added by Morgan 2022/12/28
   
'Removed by Morgan 2021/4/8 使用在途期限已改需另外收文在途期限間(442)會於分案時自動發文並更新相關號期限
'  'Added by Lydia 2015/05/13 P案申請國家非台灣,若收文時期限已過法定期限之控制
'   m_TogCP10 = "": m_bolTogCP10 = False: m_bolUpdCP07 = False
'   'Added by Lydia 2015/08/07 已逾法定期限再秀出是否加在途期限十五天
'    If pa(9) <> 台灣國家代號 And InStr("107,205,204", cp(10)) > 0 And Text1(5) < strSrvDate(2) Then
'      'Modified by Lydia 2015/08/07
'      'If CheckCPExists("404", cp(9)) Then
'      If CheckCPExists("404", cp(9), cp(43)) Then
'         m_TogCP10 = "404": m_bolTogCP10 = True
'      End If
'      If InStr("205,204", cp(10)) > 0 And m_bolTogCP10 = True Then
'         '若收文(205)陳述意見、(204)補正,同時收文(404)延期,於分案時法定期限小於本所期限仍可進行分案
'      Else
'        If m_bolTogCP10 = False Then
'          '若收文(107)復審(不能延期),(205)陳述意見、(204)補正,未收文延期,於分案時秀訊息告知user
'            strExc(0) = "select CP133,NVL(CP134,0) CP134 from caseprogress where cp09='" & cp(43) & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               If Not IsNull(RsTemp(0)) And RsTemp(1) > 0 Then
'                  '在途期限=>官方來函日加15天加官方給的期限(月份)
'                  strExc(1) = CompDate(2, 15, RsTemp(0))
'                  strExc(1) = CompDate(1, RsTemp(1), strExc(1))
'                  strExc(1) = ChangeWStringToTString(strExc(1))
'                  If MsgBox("法定期限加在途期限15天為" & strExc(1) & " ,若欲更新請再輸入法定期限.", vbYesNo + vbDefaultButton1, "更新法定期限") = vbYes Then
'                     Text1(5) = strExc(1): m_bolUpdCP07 = True
'                  End If
'               Else
'                  MsgBox "相關總收文號無官方來函日或期限(月份) ,請先回到該官方來函輸入.", vbExclamation
'               End If
'            End If
'        End If
'        m_bolTogCP10 = False
'      End If
'   End If
'   'end 2015/05/13
   
   'add by sonia 2016/4/29
   If Text1(1) = "421" Then
      If strExc(1) <> "" Then
         strExc(0) = "select pa01,pa02,pa26,cp10,cp24" & _
            " from patent,caseprogress where pa11='" & pa(11) & "'" & _
            " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
            " and cp10='" & Text1(1) & "' and cp57 is null and cp09<>'" & cp(9) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "本案已申請過【" & Label3(1) & "】，請與智權同仁確認是否再申請一次！", vbExclamation
         End If
      End If
   End If
   'end 2016/4/29
   
   'Added by Lydia 2021/05/31 FMP案和寰華案之中間程序在分案時，有C類相關總收文並且承辦人為工程師，預設承辦人為相關總收文的工程師，若人員離職改為主管（副理）
   If m_bolFMP = True And cp(43) <> "" And Mid(cp(43), 1, 1) = "C" And Text1(0).Text = "" Then
      strExc(0) = "select cp14,st04 from caseprogress,staff where cp09='" & cp(43) & "' and cp14=st01(+) and st03='F21' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
          If "" & RsTemp.Fields("st04") <> "1" Then
              'Modified by Lydia 2023/01/12 日文組只需單純回傳離職人員的主管
              'Text1(0).Text = PUB_GetFCPEngSup("" & RsTemp.Fields("cp14"), True)
              Text1(0).Text = PUB_GetFCPEngSup("" & RsTemp.Fields("cp14"), True, True)
          Else
              Text1(0).Text = "" & RsTemp.Fields("cp14")
          End If
          Call Text1_Validate(0, False)
      End If
   End If
   'end 2021/05/31
   
   'Add by Amy 2022/10/17 +簽核頁籤,接洽單電子收文才顯示「檢視接洽單」鈕
   cmdFile.Visible = False
   Me.cmdCPP.Visible = False 'Add By Sindy 2022/11/1
   SSTab1.TabVisible(2) = False
   txtF0301 = Pub_GetIsFlowCP140(cp(9))
   'Add by Amy 2022/11/15
   Label24.Visible = False: txtF0309.Visible = False '目前狀態
   Check11.Visible = False '急件
   Check11.Value = 0 'Add By Sindy 2023/1/10 要先清欄位值,再後續判斷是否急件
   'end 2022/11/15
   
   'Modify by Amy 2023/01/03 +Len(txtF0301) = 10,8碼(結案單)不可開接洽單會錯
   If strSrvDate(1) >= 接洽單電子收文啟用日 And txtF0301 <> MsgText(601) And Len(txtF0301) = 10 Then
        cmdFile.Visible = True
        Me.cmdCPP.Visible = True 'Add By Sindy 2022/11/1
        '補件完成 欄-案件表單流程備註檔屬於分案作業相關資訊
        SetFlow004TextBox txtF0407, txtF0301, " And F0408 in('A5','A6','A7') And F0409 in('A5','A6','A7') "
        '案件表單簽核檔
        strSql = "SELECT ST02||nvl(F0208,'') 簽核人員,decode(F0202," & ShowFlow簽核人員身份 & ") 身份,sqldateT(F0205) 日期,sqltime6(F0206) 時間,decode(F0207," & ShowFlow簽核結果 & ") 簽核結果,F0204 FROM FLOW002,Staff WHERE F0201='" & txtF0301 & "' and F0204=ST01(+) order by decode(F0205,null,2,1) asc,F0205||Decode(length(F0206),5,'0','')||F0206 asc,F0202,F0203 asc"
        If rsTmp1.State = adStateOpen Then rsTmp1.Close
        rsTmp1.CursorLocation = adUseClient
        rsTmp1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If rsTmp1.RecordCount > 0 Then
           Set GRD1.Recordset = rsTmp1
           SetGrd
        End If
        Command2(2).Enabled = False
        CmdAddInfo.Enabled = False: CmdAddInfo.Caption = "補件完成"
        txtNote.Locked = True
        stCPM35 = PUB_GetCPM35(txtF0301, cp(1))
        Select Case stCPM35
             Case "1" '先補文件再呈分案主管
             Case "2" '程序承不需經主管分案
                 CmdAddInfo.Caption = "呈分案主管"
             Case "3" '程序或工程師承辦
                 CmdAddInfo.Caption = "呈分案主管"
        End Select
        '北所分案人員已同意,才可按 確定 鈕
        If ChkConultRecFlow002(Me.Name, txtF0301, "A6", IsEConsultRec, stF0207_A6) = True Then
             Command2(2).Enabled = True
        End If
        arrTmp = Split(GetFlow003Data(txtF0301, , "F0308||';'||Nvl(F0309,'NULL')||';'||F0307"), ";")
        stF0309_Now = arrTmp(1)
        stF0307_Now = arrTmp(2)
        'Add by Amy 2022/11/15 +表單狀態/急件
        'Modify by Sindy 2022/11/22
        txtF0309 = PUB_GetCP157forF0309(cp(9)) '表單狀態
        Label24.Visible = True: txtF0309.Visible = True
        Check11.Visible = True '急件
        If cp(122) = "Y" Then Check11.Value = 1
        'end 2022/11/15
        '下一處理人員是A7(多筆案件性質已處理一筆)且目前表單狀態不是已分案
        If arrTmp(0) = "A7" And stF0309_Now <> "17" Then
            CmdAddInfo.Enabled = True
            txtNote.Locked = False
        End If
        '接洽單電子收文顯示簽核頁籤
        SSTab1.TabVisible(2) = True
        'Add by Amy 2022/11/15 狀態為 程序補件 時,切至 簽核 頁籤
        If stF0309_Now = "20" Then SSTab1.Tab = 2
        
        'Add By Sindy 2022/11/23
        If Text1(15).Text = "" And cp(157) = "" Then '與國內案號相同欄位空白才預帶
            If InStr(CaseMapIn, Text1(1)) > 0 And pa(1) = "P" Then
               strExc(10) = Pub_GetCRLCaseMap(txtF0301, "0", "P", pa(1), pa(2), pa(3), pa(4))
               If strExc(10) <> "" Then
                  strExc(6) = SystemNumber(strExc(10), 1)
                  strExc(7) = SystemNumber(strExc(10), 2)
                  strExc(8) = SystemNumber(strExc(10), 3)
                  strExc(9) = SystemNumber(strExc(10), 4)
                  Text1(15).Text = strExc(6) & strExc(7) & strExc(8) & strExc(9)
               End If
            End If
        End If
        'Add by Amy 2022/12/23 直接開啟接洽單-玲玲
        '從frm050701(基本檔維護)過來已開接洽單不需再開
        If intOpen090801 = 0 Then
            frm090801_Q.SetParent Me
            frm090801_Q.m_blnCallPrint = True
            frm090801_Q.Text5 = txtF0301
            Call frm090801_Q.cmdok_Click(4)
            frm090801_Q.Show
        ElseIf PUB_CheckFormExist("frm090801_Q") = True Then
            frm090801_Q.SetFocus
        End If
   End If

   'Add By Sindy 2024/1/30 各部門分案時，若本所期限與法定期限與接洽單的本所期限與法定期限不同時，要提醒
   Call PUB_ChkCRLdtCP06CP07(cp(9))
End Sub

'Added by Morgan 2012/7/20
Private Function GetCP147Default() As String
   '電子電機或生化醫學的申請程序預設 Y
   'Removed by Morgan 2016/5/18 取消 --郭雅娟
   'If (Left(Combo3, 1) = "2" Or Left(Combo3, 1) = "3") And (Text1(1) = "101" Or Text1(1) = "102" Or Text1(1) = "103") Then
   '   GetCP147Default = "Y"
   'End If
End Function

Private Sub GetCustom(ByVal iSitu As Integer)
 Dim strTmp As String, strTmp1 As String
   Select Case iSitu
      Case 8
         strExc(1) = pa(iSitu)
         If Not IsEmptyText(strExc(1)) Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetCustomer(strExc(1), strTmp) Then
            If ClsPDGetCustomer(strExc(1), strTmp) Then
               Label3(2) = strTmp
            Else
               Label3(2) = ""
            End If
         End If
      'Add by Morgan 2005/11/25
      Case 58, 59
         strExc(1) = pa(iSitu)
         If Not IsEmptyText(strExc(1)) Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetCustomer(strExc(1), strTmp) Then
            If ClsPDGetCustomer(strExc(1), strTmp) Then
               Label3(iSitu - 55) = strTmp
            Else
               Label3(iSitu - 55) = ""
            End If
         End If
      Case 26, 27, 28, 29, 30
         strExc(1) = pa(iSitu)
         If Not IsEmptyText(strExc(1)) Then
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.GetCustomer(strExc(1), strTmp) Then
            If ClsPDGetCustomer(strExc(1), strTmp) Then
               Label3(iSitu - 24) = strTmp
            Else
               Label3(iSitu - 24) = ""
            End If
         Else
            Label3(iSitu - 24) = ""
         End If
   End Select
End Sub
'Add by Morgan 2004/7/23
Private Sub SetAD(ByVal i As Integer)
   Dim strAD10 As String, strCU15 As String
   txtAD(i).Enabled = False
   txtAD(i).Tag = ""
   txtAD(i).Text = ""
   If pa(i + 25) <> "" And Text1(13).Text = "000" Then
      txtAD(i).Text = PUB_GetAD03(pa(i + 25), Text1(13).Text, strAD10, strCU15)
      txtAD(i).Tag = txtAD(i).Text
      '個人只可設定自然人(1)
      If strCU15 = "0" Then
         txtAD(i).Text = "1"
      'Added by Morgan 2014/7/15 學校也預設--玲玲
      ElseIf strCU15 = "2" Then
         txtAD(i).Text = "2"
      'end 2014/7/15
      '公司
      Else
         If txtAD(i).Text = "Y" Then
            txtAD(i).Text = strAD10
            txtAD(i).Tag = txtAD(i).Text
         End If
         txtAD(i).Enabled = True
      End If
   End If
End Sub
Private Sub GetFagent(ByVal iSitu As Integer)
 Dim strTmp As String
   strExc(1) = pa(iSitu)
   If Not IsEmptyText(strExc(1)) Then
      'Modify By Cheng 2002/07/08
      '若系統種類對照檔的SK03=0, 則代理人名稱抓中-->英-->日, 否則抓英-->中-->日
'      If objPublicData.GetAgent(strExc(1), strTmp) Then
      If PUB_GetAgentName(pa(1), strExc(1), strTmp) Then
         Label3(7) = strTmp
      Else
         Label3(7) = ""
      End If
   Else
      Label3(7) = ""
   End If
End Sub
'Add by Morgan 2004/7/20
'若沒有客戶減免身分需輸入則游標預設在承辦人
Private Sub Form_Activate()
   'Added by Morgan 2012/9/21
   '台灣申復修正逾法限不可分案
   If cp(14) = "" And pa(9) = "000" And (cp(10) = "204" Or cp(10) = "205") And Val(cp(7)) > 0 And DBDATE(cp(7)) < strSrvDate(1) Then
      MsgBox "本程序已逾法定期限，請交主管處理！", vbExclamation
      Command2_Click 3
      Exit Sub
   End If
   'end 2012/9/21
   
   'Added by Morgan 2012/10/8
   If pa(9) = "000" Then
      '衍生設計若母案已公告則不能申請
      If cp(10) = "125" Then
         strExc(0) = "select sqldatet(pa14) from patent where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='0' and pa04='" & pa(4) & "' and pa14>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "本衍生設計案之母案已於 " & RsTemp(0) & " 公告，不可分案！", vbCritical
            Command2_Click 3
            Exit Sub
         End If
      End If
      
      'Added by Morgan 2012/12/27
      If cp(10) = "601" Then
         If pa(20) <> "" Then
            strExc(1) = CompDate(2, 1, pa(20))
            strExc(1) = CompDate(1, 3, strExc(1))
            strExc(1) = CompDate(2, -1, strExc(1))
            If strExc(1) < "20130101" And strExc(1) < strSrvDate(1) Then
               MsgBox "原領證法限已逾期且早於 102/1/1，不可分案！"
               Command2_Click 3
               Exit Sub
            End If
         End If
      End If
      
      If cp(10) = "605" Then
         strExc(2) = Right(pa(72), 2)
         If Left(strExc(2), 1) = "," Then strExc(2) = Mid(strExc(2), 2)
         strExc(1) = CompDate(0, Val(strExc(2)), pa(14))
         strExc(1) = CompDate(2, -1, strExc(1))
         
         If strExc(1) < "20120701" And strExc(1) < strSrvDate(1) Then
            MsgBox "原年費法限已逾期且早於 101/7/1，不可分案！"
            Command2_Click 3
            Exit Sub
         End If
      End If
      'end 2012/12/27
      
   End If
   'end 2012/10/8
   
   'Add by Morgan 2010/1/28
   If m_bolCP98Check Then
      m_bolCP98Check = False
      strExc(0) = "select cp98 from caseprogress where cp09='" & cp(9) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '原值有變動
         If cp(98) <> "" & RsTemp("cp98") Then
            '若有改加乘註記時提醒
            If txtCP98 <> cp(98) Then
               MsgBox "原加乘註記值已變更(" & cp(98) & " -> " & RsTemp("cp98") & ")，請重新確認該值是否正確！"
               SSTab1.Tab = 1
               txtCP98.SetFocus
            End If
            cp(98) = "" & RsTemp("cp98")
            txtCP98 = cp(98)
         End If
      End If
   End If
   'end 2010/1/28
   
   If m_bolActive = True Then Exit Sub
   m_bolActive = True
   
   Dim i As Integer
   
   For i = 1 To 5
      If txtAD(i).Enabled = True And txtAD(i).Text = "" Then
         txtAD(i).SetFocus
         Exit Sub
      End If
   Next
   If Text1(0).Enabled = True Then Text1(0).SetFocus
   'Add by Morgan 2005/1/5 延緩公告若領證已發文時提醒
   If cp(10) = "412" And pa(9) = "000" And cp(27) = "" Then
      If PUB_ChkCPExist(pa, 領證及繳年費, 2) = True Then
         MsgBox "本案有【領證及繳年費】已發文！", vbExclamation, "延緩公告分案提醒"
      End If
   End If
   
   'Added by Morgan 2012/10/9
   If pa(9) = "000" And pa(8) = "3" And (cp(10) = "701" Or cp(10) = "708" Or cp(10) = "704" Or cp(10) = "705") And (Len(pa(11)) = 9 Or Mid(pa(11), 10, 1) = "D") Then
      strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)" & _
         " from patent,caseprogress where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03<>'" & pa(3) & "' and nvl(pa17,'Y')='Y'" & _
         " and nvl(substr(pa11,10,1),'D')='D' and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
         " and cp10(+)='" & cp(10) & "' and cp27(+) is null and cp57(+) is null and cp09 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         RsTemp.MoveFirst
         MsgBox "衍生設計案收文【 " & Label3(1) & " 】時母案及其他衍生設計案也需" & vbCrLf & "一併收文，但目前有未收文案號如下：" & vbCrLf & vbCrLf & RsTemp.GetString(adClipString, , , vbCrLf), vbExclamation
      End If
   End If
   'end 2012/10/9
   
'Remove by Morgan 2010/6/11 郭已回收請作單
'Add by Morgan 2010/6/9
'   If pa(57) <> "" And cp(27) = "" And cp(57) = "" Then
'      MsgBox "本案已結案閉卷，須與客戶再做進一步確認！"
'   End If

   'Add by Morgan 2010/6/23
   If Text1(0) <> "" And Text1(13) = "020" And InStr(CaseMapIn, Text1(1)) > 0 And Text1(15) = "" Then
      Text1(15).SetFocus
   End If
   'end 2010/6/23

   '2013/10/31 add by sonia 非台灣新申請案收費0,第一次分案時要提醒二案案件備註加註同時合併計算結餘" T-189182(T-188512)
   'Modified by Lydia 2022/03/09 'Modified by Lydia 2022/03/09 改判斷分案日;  ex.2021/06/18 CFT案件承辦人若空白時，預設為國家檔之CFT承辦人---統一判斷
   'If Text1(13) <> "000" And InStr(NewCasePtyList, Text1(1)) > 0 And Text1(0) = "" And Val(cp(16)) = 0 And Left(cp(12), 1) <> "F" Then
   If Text1(13) <> "000" And InStr(NewCasePtyList, Text1(1)) > 0 And Val(cp(149)) = 0 And Val(cp(16)) = 0 And Left(cp(12), 1) <> "F" Then
      MsgBox "此新申請案未收費, 若有前案則請至其他頁籤之案件備註欄加註與前案號合併計算結餘(前案之案件備註也要加註)!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1
      Text1_GotFocus (20)
      Text1(20).SetFocus
   End If
   '201/10/31 end
End Sub

Private Sub Form_Initialize()
   'Add by Morgan 2005/3/2 改用變數
   ReDim pa(TF_PA) As String
   ReDim cp(TF_CP) As String
End Sub

Private Sub Form_Load()
 Dim i As Integer
   MoveFormToCenter Me
   intWhere = 國內
   SSTab1.Tab = 0 'Add by Amy    2014/09/04
   With frm040101.MSHFlexGrid1
      IntTot = 0
'      If .Text = "" Then Exit Sub
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" Then
            'Modify By Cheng 2003/06/09
'            StrTot1(IntTot) = .TextMatrix(i, 7) '本所案號
            'Modify by Amy 2014/11/19 改為GetValue
            StrTot1(IntTot) = frm040101.GetValue(i, "CaseNo1") '.TextMatrix(i, 8) '本所案號
            StrTot2(IntTot) = frm040101.GetValue(i, "收文號") '.TextMatrix(i, 1)  '收文號
            IntTot = IntTot + 1
         End If
      Next
   End With
   'Added by Lydia 2020/05/20 法律所案源收文
   FraLOS.Visible = False
   FraLOS.BackColor = &H8000000F
   txtLOSagree.Text = ""
   FraLOS.Top = 2880
   'end 2020/05/20
   
   IntNow = 0
   GetData IntNow
   'Add by Amy 2014/09/19 承辦人期限隱藏
   Label1(23).Visible = False
   Text1(25).Enabled = False
   Text1(25).Visible = False
   'end 2014/09/19
End Sub

Private Function ChgType(iSitu As Integer) As Boolean
   Dim strTempName As String, bolChk As Boolean, i As Integer
   ChgType = True
   Select Case iSitu
      Case 0
         If Text1(iSitu) <> "" Then
            'Modify by Morgan 2011/10/13 已發文或承辦人沒改時不用考慮是否離職(補資料)--玲玲
            If cp(27) <> "" Or Text1(iSitu) = Text1(iSitu).Tag Then
               ChgType = ClsPDGetStaffN(Text1(iSitu), strTempName)
            Else
               ChgType = ClsPDGetStaff(Text1(iSitu), strTempName)
            End If
            Label3(0) = strTempName
         Else
            Label3(0) = ""
         End If
         
         'Modified by Morgan 2013/10/23 考慮程序新人
         'If Text1(iSitu) = "81002" Or Text1(iSitu) = "73017" Then
         'Modified by Morgan 2015/6/3 81002 權限已改但仍會承辦
         'Modified by Morgan 2022/10/26 已離職不可預設
         'If Text1(iSitu) = "81002" Or PUB_GetST05(Text1(iSitu)) = "75" Then
         'Modified by Lydia 2025/06/18 P及CFP分案時，預設N不計件管制：承辦人為內專程序P12都預設CP26=N
         'If Label3(0) <> "" And (Text1(iSitu) = "81002" Or PUB_GetST05(Text1(iSitu)) = "75") And cp(27) = "" Then
         If Label3(0) <> "" And PUB_GetST03(Text1(iSitu)) = "P12" And cp(27) = "" Then
         'end 2013/10/23
            Text1(10) = "N"
            Text1(25).Enabled = True 'Add by Morgan 2010/3/15 控制承辦期限欄位
            'Add by Morgan 2010/3/17 +收款後送件欄位
            'Modify by Morgan 2010/12/29 改用 Option
            'txtTakeCtrl.Locked = False
            Frame1.Enabled = True
         Else
            Text1(25).Enabled = False
            Text1(25).Text = TransDate(cp(48), 1)
            'Modify by Morgan 2010/12/29 改用 Option
            'txtTakeCtrl.Locked = True
            'txtTakeCtrl.Text = ""
            If Not m_bolFMP Then  'Added by Morgan 2023/8/31
               If strSrvDate(1) < 指定日期啟用日 Then 'Added by Morgan 2023/12/28
                  Frame1.Enabled = False
               End If
               'OptSendType(2).Value = False 'Removed by Morgan 2023/7/26 取消,接洽單已電子化--玲玲
            End If 'Added by Morgan 2023/8/31
         'end 2010/3/15
         End If

      Case 1
         If Text1(iSitu) <> "" Then
            'Modify by Morgan 2004/5/31
            '加 916-培訓費、917-超頁超項費
            '2005/7/15 MODIFY BY SONIA 加 其他 910
            'If Text1(iSitu) = 提早公開 Or Text1(iSitu) = 準備程序 Or Text1(iSitu) = 言詞辯論 Or Text1(iSitu) = 調卷 Or Text1(iSitu) = 補文件 Or Text1(iSitu) = 催審 Or Text1(iSitu) = 延緩公告 Or Text1(iSitu) = 後金 Or Text1(iSitu) = 補收款 Or Text1(iSitu) = 回覆代理人 Or Text1(iSitu) = 告知代理人 Then
            'Modify by Morgan 2007/4/27 加急件費(920)--玲玲
            'If Text1(iSitu) = 提早公開 Or Text1(iSitu) = 準備程序 Or Text1(iSitu) = 言詞辯論 Or Text1(iSitu) = 調卷 Or Text1(iSitu) = 補文件 Or Text1(iSitu) = 催審 Or Text1(iSitu) = 延緩公告 Or Text1(iSitu) = 後金 Or Text1(iSitu) = 補收款 Or Text1(iSitu) = 回覆代理人 Or Text1(iSitu) = 告知代理人 Or Text1(iSitu) = "916" Or Text1(iSitu) = "917" Or Text1(iSitu) = 其他 Then
            'Modify by Morgan 2010/1/6 +938,939
            'Modified by Lydia 2025/06/18 P及CFP分案時，預設N不計件管制：改用CasePropertyMap.CPM05控制
            'If Text1(iSitu) = 提早公開 Or Text1(iSitu) = 準備程序 Or Text1(iSitu) = 言詞辯論 Or Text1(iSitu) = 調卷 Or Text1(iSitu) = 補文件 Or Text1(iSitu) = 催審 Or Text1(iSitu) = 延緩公告 Or Text1(iSitu) = 後金 Or Text1(iSitu) = 補收款 Or Text1(iSitu) = 回覆代理人 Or Text1(iSitu) = 告知代理人 Or Text1(iSitu) = "916" Or Text1(iSitu) = "917" Or Text1(iSitu) = 其他 Or Text1(iSitu) = "920" Or Text1(iSitu) = "938" Or Text1(iSitu) = "939" Then
            '   Text1(10) = "N"
            'End If
            ''91.12.9 END
            ''94.1.31 add by sonia
            ''modify by sonia 2019/3/13 +123主張優惠期
            'If Text1(iSitu) = "106" Or Text1(iSitu) = "121" Or Text1(iSitu) = "123" Or Text1(iSitu) = "215" Or Text1(iSitu) = "401" Or Text1(iSitu) = "404" Or Text1(iSitu) = "405" Or Text1(iSitu) = "406" Or Text1(iSitu) = "413" Or Text1(iSitu) = "416" Or Text1(iSitu) = "420" Or Text1(iSitu) = "601" Or Text1(iSitu) = "602" Or Text1(iSitu) = "603" Or Text1(iSitu) = "604" Or Text1(iSitu) = "605" Or Text1(iSitu) = "606" Then
            '   Text1(10) = "N"
            'End If
            'If Text1(iSitu) = "608" Or Text1(iSitu) = "701" Or Text1(iSitu) = "702" Or Text1(iSitu) = "703" Or Text1(iSitu) = "705" Or Text1(iSitu) = "706" Or Text1(iSitu) = "707" Or Text1(iSitu) = "708" Or Text1(iSitu) = "905" Or Text1(iSitu) = "907" Or Text1(iSitu) = "908" Or Text1(iSitu) = "913" Or Text1(iSitu) = "915" Or Text1(iSitu) = "919" Or Text1(iSitu) = "920" Or Text1(iSitu) = "921" Then
            '   Text1(10) = "N"
            'End If
            ''94.1.31 END
            If PUB_GetCPMbyCP10(pa(1), Text1(iSitu), "cpm05") = "N" Then
               Text1(10) = "N"
            End If
            'end 2025/06/18
            
            'Modify by Morgan 2007/8/30 加第三人申請技術報告807
            'If Text1(iSitu) = 異議_專 Or Text1(iSitu) = 舉發 Or Text1(iSitu) = 鑑定報告 Then
            If Text1(iSitu) = 異議_專 Or Text1(iSitu) = 舉發 Or Text1(iSitu) = 鑑定報告 Or Text1(iSitu) = "807" Then
               Text1(7).Enabled = True
               Text1(17).Enabled = True
               If Text1(iSitu) = "807" Then
                  Text1(22).Enabled = True
               End If
            Else
               
               Text1(7).Enabled = False
               Text1(17).Enabled = False
               'Modify by Morgan 2007/8/31
               'Text1(7) = ""
               'Text1(17) = ""
               Text1(22).Enabled = False
               'end 2007/8/31
            End If
            If Text1(iSitu) = 異議_專 Then
               Text1(4).Enabled = False
               Text1(5).Enabled = False
            Else
               Text1(4).Enabled = True
               Text1(5).Enabled = True
            End If
            If Text1(13) = 台灣國家代號 Then
               bolChk = False
            Else
               bolChk = True
            End If
            'edit by nickc 2007/02/02 不用 dll 了
            'ChgType = objPublicData.GetCaseProperty(pA(1), Text1(iSitu), strTempName, bolChk)
            ChgType = ClsPDGetCaseProperty(pa(1), Text1(iSitu), strTempName, bolChk)
            Label3(1) = strTempName
            'Add by Morgan 2004/3/23
            '案件性質為分割時，顯示分割母案本所案號
            If Text1(1) = "307" Then
                DivVisibleSwitch True
            Else
                DivVisibleSwitch False
            End If
            'Add by Morgan 2004/6/3
            '延緩公告延緩月數輸入控制
            CP71Switch Text1(13).Text, Text1(1).Text
            
            'Add by Morgan 2010/3/16
            'Modify by Morgan 2011/8/17 +非FMP的領證601
            'Modify by Amy 2014/09/04 秀玲說拿掉m_bolFMP = False
            'If (m_bolFMP = False And Text1(iSitu) = "601") Or Text1(iSitu) = "605" Or Text1(iSitu) = "606" Or Text1(iSitu) = "607" Then
            If Text1(iSitu) = "601" Or Text1(iSitu) = "605" Or Text1(iSitu) = "606" Or Text1(iSitu) = "607" Then
               If Text1(iSitu) = "605" Or Text1(iSitu) = "601" Then
                  lblFeeYear.Caption = "繳費年度：第         -         年"
               Else
                  lblFeeYear.Caption = "繳費次數：第         -         次"
               End If
               lblFeeYear.Visible = True
               txtFeeYear(1).Visible = True
               txtFeeYear(2).Visible = True
            Else
               lblFeeYear.Visible = False
               txtFeeYear(1).Visible = False
               txtFeeYear(2).Visible = False
            End If
            
            'Add by Morgan 2010/3/17
            If Text1(iSitu) = "123" Then
               'lblFavDt.Visible = True
               txtFavDt.Visible = True
               CmdFav.Visible = True 'Add by Lydia 2015/02/02
            Else
              ' lblFavDt.Visible = False
               txtFavDt.Visible = False
               CmdFav.Visible = False 'Add by Lydia 2015/02/02
            End If
         Else
            MsgBox "案件性質不可空白 !", vbCritical
            ChgType = False
         End If

      Case 2 '專利種類
         If Me.Text1(2).Enabled Then
            Me.Label3(12).Caption = "" & PUB_GetPatentKindName(Me.Text1(2).Text, Me.Text1(13).Text)
            If Me.Label3(12).Caption = "" Then
               MsgBox "專利種類輸入錯誤!!!", vbExclamation + vbOKOnly
               ChgType = False
               Me.Text1(2).SetFocus
            Else
               If (Me.Text1(1).Text >= "101" And Me.Text1(1).Text <= "103") Or _
                  (Me.Text1(1).Text >= "301" And Me.Text1(1).Text <= "303") Then
                  If Mid(Me.Text1(1).Text, 3, 1) <> Me.Text1(2).Text Then
                     MsgBox "專利種類必須與案件性質的第三碼相同!!!", vbExclamation + vbOKOnly
                     ChgType = False
                     Me.Text1(2).SetFocus
                  End If
               End If
               'Added by Morgan 2012/12/27
               If Me.Text1(1).Text = "125" Or Me.Text1(1).Text = "308" Then
                  If Me.Text1(2).Text <> "3" Then
                     MsgBox Label3(1) & "專利種類必須為 3 設計!!!", vbExclamation + vbOKOnly
                     ChgType = False
                     Me.Text1(2).SetFocus
                  End If
               End If
               'end 2012/12/27
            End If
         End If
      
      Case 3
         If pa(1) = "P" Then
            If Text1(iSitu) = "" Then
               MsgBox "卷宗性質不可空白 !", vbCritical
               ChgType = False
            Else
               If Text1(1) = 異議_專 Then
                  If Text1(iSitu) <> "2" Then
                     MsgBox "案件性質為異議時，卷宗性質必須為 2 !", vbCritical
                     ChgType = False
                  End If
               'Modify by Morgan 2007/8/30 加第三人申請技術報告807
               'ElseIf Text1(1) = 舉發 Then
               ElseIf Text1(1) = 舉發 Or Text1(1) = "807" Then
                  If Text1(iSitu) <> "3" Then
                     MsgBox "案件性質為舉發或第三人申請技術報告時，卷宗性質必須為 3 !", vbCritical
                     ChgType = False
                  End If
               End If
            End If
         End If
      Case 4, 5 '本所期限, 法定期限
         If Text1(iSitu) <> "" Then
            If ChkDate(Text1(iSitu)) Then
                'Add By Cheng 2003/12/08
                '若本所期限非工作天則直接調整至最近的工作天
                If iSitu = 4 Then
                    Me.Text1(iSitu).Text = TransDate(PUB_GetWorkDay1(Me.Text1(iSitu).Text, True), 1)
                End If
                'End
               If iSitu = 5 Then
                  '2008/11/27 ADD BY SONIA
                  If Text1(1) = "226" Then
                     'Added by Lydia 2025/10/29
                     If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                        Text1(4).Text = TransDate(PUB_GetPOurDeadline(Text1(5), Text1(13)), 1)
                     Else
                     'end 2025/10/29
                        'Added by Morgan 2014/10/28
                        If Text1(13) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                           Text1(4) = TransDate(PUB_GetOurDeadline(Text1(5)), 1)
                        Else
                        'end 2014/10/28
                           Text1(4).Text = TransDate(PUB_GetWorkDay1(CompDate(2, -2, Text1(5).Text), True), 1)
                        End If
                     End If 'Added by Lydia 2025/10/29
                  End If
                  '2008/11/27 END

'Removed by Morgan 2021/4/8 使用在途期限已改需另外收文在途期限間(442)會於分案時自動發文並更新相關號期限
'                  'Added by Lydia 2015/05/13
'                  If Text1(6) <> cp(43) Then
'                    'Modified by Lydia 2015/08/07
'                    'If CheckCPExists("404", cp(9)) Then
'                    If CheckCPExists("404", cp(9), cp(43)) Then
'                       m_TogCP10 = "404": m_bolTogCP10 = True
'                    End If
'                  End If
'                  If Text1(13) <> 台灣國家代號 And InStr("205,204", Text1(1)) > 0 And m_bolTogCP10 = True Then
'                     '若收文(205)陳述意見、(204)補正,同時收文(404)延期,於分案時法定期限小於本所期限仍可進行分案
'
'                  Else
'                     m_TogCP10 = "": m_bolTogCP10 = False
                     
                     If Val(Text1(4)) > Val(Text1(5)) Then
                        MsgBox "本所期限必須小於法定期限 !", vbCritical
                        ChgType = False
                     End If
                     
'                  End If
'                  'end 2015/05/13
'end 2021/4/8
                  
               End If
            Else
               ChgType = False
            End If
         Else
            Select Case Text1(1)
               'Modify by Morgan 2004/10/13 加915退證註銷
               '2008/11/26 MODIFY BY SONIA 加226配合開庭 '2011/3/14取消配合開庭
               Case 答辯, 修正, 申復, 延期, 面詢, 閱卷, 訴願, 再訴願, 行政訴訟, 行政再審, 年費, 異議_專, 異議答辯, 舉發答辯, 915
                  MsgBox "此案件性質必須有期限 !", vbCritical
                  ChgType = False
               Case 領證及繳年費
                  If Text1(13) = 大陸國家代號 Then
                     MsgBox "此案件性質必須有期限 !", vbCritical
                     ChgType = False
                  End If

            End Select
         End If
         
         'Add By Cheng 2002/06/21 91.11.19 cancel 主張優先權106 by sonia,於存檔時自動計算
         '若案件性質為"主張優先權"(106), "再審申請"(107), "修正"(204), "申復"(205), "延期"(404), "面詢"(408), "閱卷"(410), _
         '"訴願"(501), "再訴願"(502), "行政訴訟"(503), "行政再審"(504), "領證及繳年費"(601)僅限申請國家非台灣, "年費"(605), "維持費"(606), "異議"(801), "異議答辯"(802), "舉發答辯"(804)
         '本所期限及法定期限不可空白
         If Me.Text1(1).Text = "107" Or _
            Me.Text1(1).Text = "204" Or Me.Text1(1).Text = "205" Or _
            Me.Text1(1).Text = "404" Or Me.Text1(1).Text = "408" Or _
            Me.Text1(1).Text = "410" Or Me.Text1(1).Text = "501" Or _
            Me.Text1(1).Text = "502" Or Me.Text1(1).Text = "503" Or _
            Me.Text1(1).Text = "504" Or (Me.Text1(1).Text = "601" And pa(9) <> 台灣國家代號) Or _
            Me.Text1(1).Text = "605" Or Me.Text1(1).Text = "606" Or _
            Me.Text1(1).Text = "801" Or Me.Text1(1).Text = "802" Or _
            Me.Text1(1).Text = "804" Then
            '檢查本所期限
            If Me.Text1(4).Text = "" And iSitu = 4 Then
               MsgBox "本所期限不可空白!!!", vbExclamation + vbOKOnly
               If Me.Text1(4).Enabled Then
                  Me.Text1(4).SetFocus
                  TextInverse Me.Text1(4)
               End If
               ChgType = False
            '檢查法定期限
            ElseIf Me.Text1(5).Text = "" And iSitu = 5 Then
               MsgBox "法定期限不可空白!!!", vbExclamation + vbOKOnly
               If Me.Text1(5).Enabled Then
                  Me.Text1(5).SetFocus
                  TextInverse Me.Text1(5)
               End If
               ChgType = False
            End If
         End If
         
         strDateTmp(iSitu - 3) = Text1(iSitu)
      Case 6
         If Text1(iSitu) <> "" Then
            intI = 1
            strExc(0) = "SELECT CP01||CP02||CP03||CP04,CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP09='" & Text1(6) & "'"
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Add by Morgan 2009/11/6
               '恢復權利的相關總收文號有變時重算期限,若期限有變時詢問是否更新
               Set414Date
               'end 2009/11/6
               If RsTemp.Fields(0) = pa(1) & pa(2) & pa(3) & pa(4) Then
                  If Text1(1) = 請求公告 Or Text1(1) = 延緩公告 Then
                     If Left(Text1(6), 1) <> "C" Then
                        MsgBox "案件性質為請求公告或延緩公告時，相關總收文號必須為 C 類之收文號 !", vbCritical
                        ChgType = False
                     End If
                  End If
               Else
                  'Modify by Morgan 2010/2/6 +香港大陸關聯
                  'Modified by Lydia 2015/09/09 +澳門發明案與大陸案之關聯
                  strExc(0) = "select 1 from casemap where cm10 in ('3','4','5') and cm01='" & RsTemp.Fields(1) & "' AND CM02='" & RsTemp.Fields(2) & "' AND CM03='" & RsTemp.Fields(3) & "' AND CM04='" & RsTemp.Fields(4) & "' AND CM05='" & pa(1) & "' AND CM06='" & pa(2) & "' AND CM07='" & pa(3) & "' AND CM08='" & pa(4) & "'"
                  strExc(0) = strExc(0) & " UNION select 1 from casemap where cm10 in ('3','4','5') and cm05='" & RsTemp.Fields(1) & "' AND CM06='" & RsTemp.Fields(2) & "' AND CM07='" & RsTemp.Fields(3) & "' AND CM08='" & RsTemp.Fields(4) & "' AND CM01='" & pa(1) & "' AND CM02='" & pa(2) & "' AND CM03='" & pa(3) & "' AND CM04='" & pa(4) & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  If intI <> 1 Then
                     MsgBox "必須為此本所案號之其他收文號 !", vbCritical
                     ChgType = False
                  End If
               End If
            Else
               MsgBox "必須為此本所案號之其他收文號 !", vbCritical
               ChgType = False
            End If
        End If
        
      Case 7
         If Text1(iSitu) = "" Then
            If Text1(13) = 台灣國家代號 Then
               'Modify by Morgan 2007/8/30
               'If Text1(1) = 異議_專 Then
               If Text1(1) = 異議_專 Or Text1(1) = "807" Then
                  MsgBox "案件性質為異議或第三人申請技術報告時公告日不可空白 !", vbCritical
                  ChgType = False
               End If
            End If
            
         Else
            If ChkDate(Text1(iSitu)) = False Then
               ChgType = False
            Else
               If Text1(1) = 異議_專 Then
                  Text1(5) = TransDate(CompDate(2, -1, CompDate(1, 3, TransDate(Text1(iSitu).Text, 2))), 1)
                  'Added by Lydia 2025/10/29
                  If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                     Text1(4).Text = TransDate(PUB_GetPOurDeadline(Text1(5), Text1(13)), 1)
                  Else
                  'end 2025/10/29
                     'Added by Morgan 2014/10/28
                     If Text1(13) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                        Text1(4) = TransDate(PUB_GetOurDeadline(Text1(5)), 1)
                     Else
                     'end 2014/10/28
                        Text1(4) = TransDate(CompDate(2, -2, TransDate(Text1(5).Text, 2)), 1)
                        '本所期限若非工作天則抓最近工作天
                        Text1(4).Text = TransDate(PUB_GetWorkDay1(Text1(4).Text, True), 1)
                     End If 'Added by Morgan 2014/10/28
                  End If 'Added by Lydia 2025/10/29
               End If
            End If
         End If
         
      Case 10
         'Modified by Morgan 2013/10/23 考慮程序新人
         'If Text1(0) = "81002" Or Text1(0) = "73017" Then
         '   If Text1(iSitu) <> "N" Then
         '      MsgBox "承辦人為 81002 或 73017 時，是否算案件數只可為 N !", vbInformation
         '
         'Modified by Morgan 2015/6/3 81002 權限已改但仍會承辦
         If Text1(0) = "81002" Or PUB_GetST05(Text1(0)) = "75" Then
            If Text1(iSitu) <> "N" Then
               MsgBox "承辦人為程序人員時，是否算案件數只可為 N !", vbInformation
         'end 2013/10/23
               Text1(iSitu) = "N"
            End If
         End If
      Case 11
         'Modified by Morgan 2024/11/20 有可能已閉卷且不可取消 Ex:一案兩請新型案的退費
         'If Text1(iSitu).Visible Then
         If Text1(iSitu).Visible And Text1(iSitu).Enabled Then
         'end 2024/11/20
            If Text1(iSitu) <> "Y" Then
               'Add by Morgan 2006/7/31 下列性質閉卷仍可收文
               If Text1(1) = "908" Or Text1(1) = "915" Or Text1(1) = "919" Then
                  If MsgBox("確定""不""取消閉卷 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                     ChgType = False
                  End If
               Else
               'end 2006/7/31
                  'Modified by Morgan 2019/7/8
                  'MsgBox "資料輸入錯誤 !", vbCritical
                  MsgBox "本案已閉卷 !", vbCritical
                  ChgType = False
               End If
            Else
               If MsgBox("是否確定取消閉卷 ?", vbQuestion + vbYesNo) = vbNo Then
                  Beep
                  ChgType = False
               End If
            End If
         End If
      Case 12
         If Text1(iSitu) <> "" Then
            ChgType = ChkDate(Text1(iSitu))
         Else
            MsgBox "收文日不可空白 !", vbCritical
            ChgType = False
         End If
      Case 13
         If Text1(iSitu) <> "" Then
            'edit by nickc 2007/02/02 不用 dll 了
            'ChgType = objPublicData.GetNation(Text1(13), strTempName)
            ChgType = ClsPDGetNation(Text1(13), strTempName)
            If ChgType Then pa(9) = Text1(13).Text
            Label3(11) = strTempName
         Else
            Label3(11) = ""
         End If
         'Add by Morgan 2004/6/7
         '延緩公告延緩月數輸入控制
         CP71Switch Text1(13).Text, Text1(1).Text

      'Add by Morgan 2006/5/23
      Case 14
         If Text1(iSitu) <> "" Then
            ChgType = ChkDate(Text1(iSitu))
         End If
         If ChgType = True Then
            SetPCTDate
         End If
         
      Case 15
         'Add By Sindy 2023/1/3
         If Text1(iSitu) <> "" Then
            If InStr(Text1(iSitu), "-") > 0 Then
               Text1(iSitu) = Replace(Text1(iSitu), "-", "")
            End If
         End If
         '2023/1/3 END
         
          'edit by nickc 2005/06/30 加入翻譯費
'         If (Text1(1) = "101" Or Text1(1) = "102" Or Text1(1) = "103" Or Text1(1) = "104" Or Text1(1) = "109" Or Text1(1) = "110" Or Text1(1) = "112") And _
'            Text1(13) <> 台灣國家代號 And pa(1) = "P" Then
         'Modify by Morgan 2007/10/11 不再限制非台灣
         'If (Text1(1) = "101" Or Text1(1) = "102" Or Text1(1) = "103" Or Text1(1) = "104" Or Text1(1) = "201" Or Text1(1) = "109" Or Text1(1) = "110" Or Text1(1) = "112") And _
         '   Text1(13) <> 台灣國家代號 And pa(1) = "P" Then
         'Modified by Morgan 2013/5/30 改用 CaseMapIn 常數判斷
         'If (Text1(1) = "101" Or Text1(1) = "102" Or Text1(1) = "103" Or Text1(1) = "104" Or Text1(1) = "201" Or Text1(1) = "109" Or Text1(1) = "110" Or Text1(1) = "112") And pa(1) = "P" Then
         If InStr(CaseMapIn, Text1(1)) > 0 And pa(1) = "P" Then
            If Text1(iSitu) <> "" Then
               intI = 1
               'Modify by Morgan 2005/2/24 檢查國內案申請國家須不同，但不管案件性質
               'strExc(0) = "SELECT CP27 FROM CASEPROGRESS WHERE " & ChgCaseprogress(Text1(iSitu)) & " AND CP10 IN ('101','102','103','104')"
               strExc(0) = "SELECT CP27,CP01 FROM CASEPROGRESS WHERE " & ChgCaseprogress(Text1(iSitu)) & _
                  " AND EXISTS( SELECT * FROM PATENT WHERE PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND PA09<>'" & Text1(13) & "') order by cp31"
   
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  'Add by Morgan 2005/4/29 檢查一件國內案不可關聯兩件相同國家之國外案
                  If PUB_CheckDaulCaseMap(pa(1) & pa(2) & pa(3) & pa(4), Text1(iSitu)) = True Then
                     ChgType = False
                  'Added by Morgan 2013/5/17
                  ElseIf CheckReverseCaseMap(pa(1) & pa(2) & pa(3) & pa(4), Text1(iSitu)) = True Then
                     ChgType = False
                  'end 2013/5/17
                  Else
                  '2005/4/29 end
                     If Not IsNull(RsTemp.Fields(0)) Or RsTemp.Fields(0) <> "" Then
                        MsgBox "國內案件已送件，" & Label3(11) & "案件可進行作業 !", vbExclamation
                                          
                     'Added by Morgan 2013/10/21
                     'Removed by Morgan 2015/5/26 取消--郭雅娟
                     'ElseIf RsTemp.Fields(1) = "CFP" Then
                     '   MsgBox "CFP案尚未發文不可為國內案!", vbCritical
                     '   ChgType = False
                     'end 2015/5/26
                     'end 2013/10/21
                     End If
                  End If
               Else
                  MsgBox "無此國內案號，請重新輸入 !", vbCritical
                  ChgType = False
               End If
               'Added by Morgan 2012/3/21 組別預設與國內案相同
               If ChgType = True And txtEngGroup.Visible = True Then
                  strExc(0) = "select pa150 from patent where " & ChgPatent(Text1(iSitu)) & " and pa150 is not null"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     txtEngGroup.Text = RsTemp.Fields(0)
                  End If
               End If
               'end 2012/3/21
            Else
               'MsgBox "無此國內案號，請重新輸入 !", vbCritical
               'ChgType = False
            End If
            
         ElseIf Text1(iSitu) <> "" Then
            'Modify by Morgan 2007/10/11
            'MsgBox "非台灣案才可輸入 !", vbCritical
            MsgBox "案件性質或系統別不符合，不可輸入 !", vbCritical
            'end 2007/10/11
            ChgType = False
         End If
         
      Case 16
         If Text1(iSitu) <> "" And Text1(iSitu) <> "0" Then
            If Val(cp(34)) >= Val(Text1(iSitu)) * 1000 Then
               If MsgBox("收文點數低於底價 ?", vbQuestion + vbYesNo) = vbNo Then
                  ChgType = False
               End If
            End If
         End If
         
      Case 17 '申請案號
         'Modify by Morgan 2007/8/30 加第三人申請技術報告
         'If Text1(1) = 異議_專 Or Text1(1) = 舉發 Then
         If Text1(1) = 異議_專 Or Text1(1) = 舉發 Or Text1(1) = "807" Then
            If Text1(iSitu) = "" Then
               MsgBox "案件性質為異議、舉發或第三人申請技術報告時，申請案號不可空白 !", vbCritical
               ChgType = False
            Else
               i = 2
               If Text1(13) = 台灣國家代號 Then
                  i = 0
               '92.1.11 MODIFY BY SONIA
               'ElseIf Text1(13) = 大陸國家代號 And Text1(14) <> "Y" Then
               'Modify by Morgan 2006/5/23
               'ElseIf Text1(13) = 大陸國家代號 And Text1(2) = "1" And Text1(14) <> "Y" Then
               ElseIf Text1(13) = 大陸國家代號 And Text1(2) = "1" And Text1(14) = "" Then
               '92.1.11 END
                  i = 1
               End If
               '2005/6/14 MODIFY BY SONIA
               'If i <> 2 Then ChgType = ChkAppNo(Text1(iSitu).Text, Val(Text1(2)), i)
               If i <> 2 Then ChgType = ChkAppNo(Text1(iSitu).Text, Val(Text1(2)), i, Val(Text1(3)))
               '2005/6/14 END
               
               'Add by Morgan 2007/8/30
               If ChgType = True Then
                  If Text1(1) = 舉發 Or Text1(1) = "807" Then
                     strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||pa04) CNo,pa23,pa26,cp10,pa01,pa02,pa03,pa04" & _
                        " from patent,caseprogress where pa01='P' and pa09='" & Text1(13) & "' and pa11='" & Text1(iSitu) & "'" & _
                        " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10(+)='807'" & _
                        " and not (pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "')"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        With RsTemp
                        '有其他申請案
                        If .Fields("pa23") = "1" Then
                           ChgType = False
                           MsgBox "申請案(" & .Fields("CNo") & ")與本案的申請案號相同，不可做雙方代理！"
                        '有其他爭議案
                        Else
                           'Modify by Morgan 2009/3/9 若是第2次提舉發時不要併案 Ex:P-90606 & P-66739
                           '同申請人
                           'If .Fields("pa26") = Left(pa(26) & "000", 9) Then
                              'ChgType = False
                              'If IsNull(.Fields("cp10")) Then
                                 
                              '   strExc(1) = "舉發案(" & .Fields("CNo") & ")與本案的申請案號及申請人皆相同，請轉案號合併之！"
                              'Else
                              '   strExc(1) = .Fields("CNo") & "案與本案的申請案號及申請人皆相同並已收文第三人申請技術報告，請轉案號合併之！"
                              'End If
                              'MsgBox strExc(1)
                           '關係企業
                           'ElseIf Left(RsTemp.Fields("pa26"), 6) = Left(pa(26), 6) Then
                              'If IsNull(RsTemp.Fields("cp10")) Then
                              '   strExc(1) = "舉發案(" & RsTemp.Fields("CNo") & ")與本案的申請案號相同且申請人為關係企業，是否仍要繼續？"""
                              'Else
                              '   strExc(1) = RsTemp.Fields("CNo") & "案與本案的申請案號相同且申請人為關係企業並已收文第三人申請技術報告，是否仍要繼續？"""
                              'End If
                              'If MsgBox(strExc(1), vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                              '   ChgType = False
                              'End If
                           'End If
                           
                           '已提第三人技術報告
                           strExc(1) = ""
                           strExc(2) = ""
                           strExc(3) = ""
                           If .Fields("pa26") = Left(pa(26) & "000", 9) Then
                              strExc(3) = "1"
                           ElseIf Left(RsTemp.Fields("pa26"), 6) = Left(pa(26), 6) Then
                              strExc(3) = "2"
                           End If
                           If Not IsNull(.Fields("cp10")) Then
                              If Text1(1) = "807" Then
                                 ChgType = False
                              '有收文技術報告且沒有爭議或救濟程序時才提醒併案
                              ElseIf Text1(1) = 舉發 Then
                                 strExc(0) = "select * from caseprogress where cp01='" & .Fields("pa01") & "' and cp02='" & .Fields("pa02") & "'" & _
                                    " and cp03='" & .Fields("pa03") & "' and cp04='" & .Fields("pa04") & "'" & _
                                    " and cp10<>'807' and substr(cp10,1,1) in ('5','8') and cp27>0"
                                 intI = 1
                                 Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                                 If intI = 0 Then
                                    ChgType = False
                                 End If
                              End If
                              If ChgType = False Then
                                 If strExc(3) = "1" Then
                                    strExc(1) = .Fields("CNo") & "案與本案的申請案號及申請人皆相同並已收文第三人申請技術報告，請轉案號合併之！"
                                 Else
                                    strExc(1) = .Fields("CNo") & "案與本案的申請案號相同且申請人為關係企業並已收文第三人申請技術報告，是否仍要繼續？"""
                                 End If
                              End If
                           
                           ElseIf Text1(1) = "807" Then
                              ChgType = False
                              If strExc(3) = "1" Then
                                 strExc(1) = "舉發案(" & .Fields("CNo") & ")與本案的申請案號及申請人皆相同，請轉案號合併之！"
                              Else
                                 strExc(1) = "舉發案(" & .Fields("CNo") & ")與本案的申請案號相同且申請人為關係企業，是否仍要繼續？"""
                              End If
                           End If
                           If ChgType = False Then
                              If strExc(3) = "1" Then
                                 MsgBox strExc(1)
                              Else
                                 If MsgBox(strExc(1), vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                                    ChgType = True
                                 End If
                              End If
                           End If
                           'end 2009/3/9
                        End If
                        End With
                     End If
                  End If
               End If
               'end 2007/8/30
            End If
         End If
      
      'Add by Morgan 2006/5/23
      Case 18 'PCT 優先權日
         If Text1(iSitu) <> "" Then
            ChgType = ChkDate(Text1(iSitu))
         End If
         If ChgType = True Then
            SetPCTDate
         End If
         
      'Add by Morgan 2007/8/30
      Case 22 '證書號
         If Text1(1) = "807" Then
            If Text1(iSitu) = "" Then
               MsgBox "案件性質為第三人申請技術報告時，證書號數不可空白 !", vbCritical
               ChgType = False
            Else
               If Text1(13) = 台灣國家代號 Then
                  If Left(Text1(22), 1) <> "M" Or Len(Text1(22)) <> 7 Then
                     MsgBox "輸入的專利號數錯誤"
                     ChgType = False
                  End If
               End If
            End If
         End If
         
      Case 24
         If Text1(iSitu) <> "" Then
            '92.6.28 modify by sonia
            'ChgType = objPublicData.GetStaff(Text1(iSitu), strTempName)
            'edit by nickc 2007/02/02 不用 dll 了
            'ChgType = objPublicData.GetStaffN(Text1(iSitu), strTempName)
            ChgType = ClsPDGetStaffN(Text1(iSitu), strTempName)
            '92.6.28 end
            Label3(10) = strTempName
         Else
            Label3(10) = ""
         End If
      'Add by Morgan 2010/3/15
      Case 25 '承辦期限
         If Text1(iSitu) <> "" Then
            If Not ChkDate(Text1(iSitu)) Then
               ChgType = False
            Else
               Text1(iSitu) = TransDate(PUB_GetWorkDay1(Me.Text1(iSitu).Text, True), 1)
            End If
         End If
   End Select
End Function

Private Function GetGrid(ByVal strRecive As String, ByVal intSitu As Integer) As Boolean
   Dim strKeyAll As String
   Dim strKey1 As String
   Dim StrKey2 As String
   Dim strKey3 As String
   Dim strKey4 As String
   strKey1 = textPA1
   StrKey2 = textPA2
   strKey3 = textPA3
   If IsEmptyText(strKey3) Then strKey3 = "0"
   strKey4 = textPA4
   If IsEmptyText(strKey4) Then strKey4 = "00"
   strKeyAll = strKey1 & StrKey2 & strKey3 & strKey4
   
   GetGrid = True
   If intSitu = 0 Then
      strExc(1) = Label3(9)
   Else
      'strExc(1) = Text1(2)
      strExc(1) = strKeyAll
   End If
   If pa(1) = "P" Then
      'Modify by Morgan 2009/12/23 延期加帶出AB類未發文未取消收文的程序,且下一程序要排除程序管制的案件性質
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      strExc(0) = "SELECT '',DECODE(PA09,'" & 台灣國家代號 & "',CPM03,CPM04)" & _
         ",SQLDateT(NP08),SQLDateT(NP09),NP13,NP14,SQLDateT(NP11)" & _
         ",NP01,NP07,NP22,NP15 FROM NEXTPROGRESS,CASEPROPERTYMAP,PATENT WHERE " & _
         ChgNextProgress(strExc(1)) & " AND (NP06<>'Y' OR NP06 IS NULL) AND " & _
         ChgPatent(strExc(1)) & " AND NP02=CPM01(+) AND NP07=CPM02(+)" & strNpSqlOfNoSalesDuty
      
      If cp(10) = "404" Then
         strExc(0) = strExc(0) & " union SELECT '',DECODE(PA09,'" & 台灣國家代號 & "',CPM03,CPM04)" & _
            ",SQLDateT(CP06),SQLDateT(CP07),CP08,NVL(CP40,NVL(CP41,CP42)),''" & _
            ",CP09,CP10,0,CP64 FROM CASEPROGRESS,CASEPROPERTYMAP,PATENT" & _
            " WHERE " & ChgCaseprogress(strExc(1)) & " AND CP09<'C' and cp10<>'404'  and cp07>0 AND CP27 IS NULL AND CP57 IS NULL" & _
            " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND pa01(+)=CP01 and pa02(+)=CP02 and pa03(+)=CP03 and pa04(+)=CP04 "
            
      'Added by Morgan 2020/2/13
      '442 在途期限抓未提申程序
      ElseIf cp(10) = "442" Then
         strExc(0) = strExc(0) & " union SELECT '',DECODE(PA09,'" & 台灣國家代號 & "',CPM03,CPM04)" & _
            ",SQLDateT(CP06),SQLDateT(CP07),CP08,NVL(CP40,NVL(CP41,CP42)),''" & _
            ",CP09,CP10,0,CP64 FROM CASEPROGRESS,CASEPROPERTYMAP,PATENT" & _
            " WHERE " & ChgCaseprogress(strExc(1)) & " AND CP09<'C' and cp10 not in ('442','404') and cp07>0 AND CP57 IS NULL AND CP47 IS NULL" & _
            " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND pa01(+)=CP01 and pa02(+)=CP02 and pa03(+)=CP03 and pa04(+)=CP04 "
      'end 2020/2/13
      End If
   Else
      'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
      strExc(0) = "SELECT '',DECODE(SP09,'" & 台灣國家代號 & "',CPM03,CPM04)," & _
         SQLDate("NP08") & "," & SQLDate("NP09") & ",NP13,NP14," & SQLDate("NP11") & _
         ",NP01,NP07,NP22,NP15 FROM NEXTPROGRESS,CASEPROPERTYMAP,SERVICEPRACTICE WHERE " & _
         ChgNextProgress(strExc(1)) & " AND (NP06<>'Y' OR NP06 IS NULL) AND " & _
         ChgService(strExc(1)) & " AND NP02=CPM01(+) AND NP07=CPM02(+)" & strNpSqlOfNoSalesDuty
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   If intSitu = 1 Then
      If pa(1) = "P" Then
         'strExc(0) = "SELECT count(*) FROM PATENT WHERE " & ChgPatent(Text1(2))
         strExc(0) = "SELECT count(*) FROM PATENT WHERE " & ChgPatent(strKeyAll)
      Else
         'strExc(0) = "SELECT count(*) FROM SERVICEPRACTICE WHERE " & ChgService(Text1(2))
         strExc(0) = "SELECT count(*) FROM SERVICEPRACTICE WHERE " & ChgService(strKeyAll)
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If RsTemp.Fields(0) = 0 Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetMaxNumber(pA(1), strExc(1)) Then
         If ClsPDGetMaxNumber(pa(1), strExc(1)) Then
            'If Text1(2) > pa(1) & String(6 - Len(strExc(1)), "0") & strExc(1) Then
            '2012/2/8 MODIFY BY SONIA 只判斷前二欄
            'If strKeyAll > pa(1) & String(6 - Len(strExc(1)), "0") & strExc(1) Then
            If strKey1 & StrKey2 > pa(1) & String(6 - Len(strExc(1)), "0") & strExc(1) Then
               MsgBox "新本所案號不可大於自動編號，請重新輸入 !", vbCritical
                'Modify By Cheng 2002/12/31
               GetGrid = False
            Else
               'If MsgBox("此本所案號不存在 ( " & Text1(2) & " ) ，請確認 ?", vbQuestion + vbYesNo) = vbNo Then
               If MsgBox("此本所案號不存在 ( " & strKeyAll & " ) ，請確認 ?", vbQuestion + vbYesNo) = vbNo Then
                  GetGrid = False
               End If
            End If
         End If
      End If
   End If
   GridHead
End Function

Private Sub MSHFlexGrid1_Click()
   Dim bolChecked As Boolean
   
   If MSHFlexGrid1.Rows < 2 Then Exit Sub
   
   'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
   If Pub_CheckNpTheSameShow(cp(1), Text1(1), Trim("" & MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 8))) = False Then
       Exit Sub
   End If
   'end 2021/08/31
   
   'Added by Morgan 2013/3/18
   '分析更新相關總收文號並檢查期限
   If Text1(1) = "941" Then
      Text1(6).Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 7)
      If Text1(4) <> "" Then
         If DBDATE(Text1(4)) > DBDATE(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 2)) Then
            MsgBox "本所期限已超過相關收文號的本所期限，將自動更新為相關收文號的本所期限！", vbExclamation
            Text1(4) = TransDate(DBDATE(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 2)), 1)
         End If
      End If
      Exit Sub
   End If
   'end 2013/3/18
   
   'Modify by Morgan 2009/12/23 延期只更新期限不可點選
   'Modified by Morgan 2020/2/7 +442 在途期限
   If Text1(1) = "404" Or Text1(1) = "442" Then
      bolChecked = True
   Else
      ' 91.01.22 modify by louis 可同時點選多筆
      'GridClick MSHFlexGrid1, intRow, 0
      GridClick MSHFlexGrid1, intRow, 0, 1
      If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 0) = "v" Then
         bolChecked = True
      End If
   End If
   
   If bolChecked Then
      '若已有本所期限
      If Len(Me.Text1(4).Text) > 0 Then
         'Modified by Morgan 2016/2/25 法限不同時也問,有可能接洽單只填所限收文時法限會輸一樣 P-108682
         'If Me.Text1(4).Text <> ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 2)) Then
         '      If MsgBox("是否要更新本所期限 ？", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
         If Me.Text1(4).Text <> ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 2)) Or Me.Text1(5).Text <> ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 3)) Then
               If MsgBox("是否要更新期限 ？", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
         'end 2016/2/25
                'Modify By Cheng 2003/12/08
'               Text1(4).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2))
               Text1(4).Text = TransDate(PUB_GetWorkDay1(ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 2)), True), 1)
               Text1(5).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 3))
               Text1(6).Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 7)
               ' 90.07.06 modify by louis
               Text1(19).Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 10)
               'Add by Morgan 2011/4/22
               If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 9) = "0" Then
                  m_CP30 = ""
               Else
                  m_CP30 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 9)
               End If
            End If
         End If
         
         'Added by Morgan 2013/4/10
         '若無相關總收文號時帶點選的 ex.P-91826 改請點選再審
         If Text1(6) = "" Then
            Text1(6).Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 7)
         End If
         'end 2013/4/10
         
      '若無本所期限
      Else
            'Modify By Cheng 2003/12/08
'         Text1(4).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2))
         Text1(4).Text = TransDate(PUB_GetWorkDay1(ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 2)), True), 1)
         Text1(5).Text = ChangeTDateStringToTString(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 3))
         Text1(6).Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 7)
         ' 90.07.06 modify by louis
         Text1(19).Text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 10)
         'Add by Morgan 2011/4/22
         If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 9) = "0" Then
            m_CP30 = ""
         Else
            m_CP30 = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.row, 9)
         End If
      End If
   Else
      Text1(4).Text = ""
      Text1(5).Text = ""
      Text1(6).Text = ""
      Text1(19).Text = ""
   End If

End Sub

'Add by Morgan 2010/12/29
Private Sub OptSendType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim oOpt As OptionButton
   If OptSendType(Index).Tag = "1" Then
      OptSendType(Index).Value = False
      OptSendType(Index).Tag = "0"
      If Index = 3 Then
         txtCP142.Text = ""
         txtCP142.Enabled = False
         'Added by Morgan 2023/8/29
         If Frame2.Visible = True Then
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False
            Option1(0).Enabled = False
            Option1(1).Enabled = False
            Option1(2).Enabled = False
         End If
         'end 2023/8/29
      End If
      
   Else
      For Each oOpt In OptSendType
         If oOpt.Index = Index Then
            oOpt.Tag = "1"
         Else
            oOpt.Tag = "0"
         End If
      Next
      'Modified by Morgan 2023/8/30
      'If Index = 3 Then
      If Index = 3 And OptSendType(Index).Value Then
      'end 2023/8/30
         txtCP142.Enabled = True
         txtCP142.SetFocus
         'Added by Morgan 2023/8/29
         If Frame2.Visible Then
            Option1(0).Enabled = True
            Option1(1).Enabled = True
            Option1(2).Enabled = True
         End If
         'end 2023/8/29
      Else
         txtCP142.Text = ""
         txtCP142.Enabled = False
         'Added by Morgan 2024/1/9
         If Frame2.Visible Then
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False
            Option1(0).Enabled = False
            Option1(1).Enabled = False
            Option1(2).Enabled = False
         End If
         'end 2024/1/9
      End If
   End If
End Sub

Private Sub Text1_Change(Index As Integer)
   Select Case Index
      Case 1 '案件性質
         Set434Date 'Added by Morgan 2012/3/30
         SetLOSagree 'Added by Lydia 2020/05/20 法律所案源收文
      Case 2 '專利種類
         If Me.Text1(Index).Enabled Then
            Me.Label3(12).Caption = "" & PUB_GetPatentKindName(Me.Text1(2).Text, Me.Text1(13).Text)
            If Me.Text1(Index).Text <> "" Then
               If Me.Label3(12).Caption = "" Then
                  MsgBox "專利種類輸入錯誤!!!", vbExclamation + vbOKOnly
                  Me.Text1(Index).SetFocus
                  TextInverse Me.Text1(Index)
                  Exit Sub
               End If
               'Add By Sindy 2014/7/14
               If Text1(Index) = "3" Or Text1(Index) = "4" Then
                  Combo3.Enabled = False
                  Combo3.Text = ""
               Else
                  Combo3.Enabled = True
               End If
               '2014/7/14 END
            End If
         End If
      Case 13  '申請國家
         If pa(1) = "P" Then
            SetAD 1
            SetAD 2
            SetAD 3
            SetAD 4
            SetAD 5
            Set434Date 'Added by Morgan 2012/3/30
            SetLOSagree 'Added by Lydia 2020/05/20 法律所案源收文
         End If
   End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 3
         KeyAscii = UpperCase(KeyAscii)
         If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 10
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 11
         KeyAscii = UpperCase(KeyAscii)
         'Modify By Cheng 2002/04/24
'         If KeyAscii <> 89 Then
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 0, 2, 6, 15, 22
         KeyAscii = UpperCase(KeyAscii)
         CloseIme
      
      'Modify by Morgan 2006/5/23 改輸入PCT申請日
'      Case 14
'         KeyAscii = UpperCase(KeyAscii)
'         If KeyAscii <> 89 And KeyAscii <> 8 Then
'            KeyAscii = 0
'            Beep
'         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   Dim strShow As String   'Add by Amy 2015/01/22
   
   Cancel = Not ChgType(Index)
   'Add by Amy 2015/01/22
   Select Case Index
        Case 0
            m_bolChkCP14OK = True
            If m_bolIsFirstKeyCP14 = True And Trim(Text1(Index)) <> "" And Text1(Index) <> Text1(Index).Tag Then
                'Add by Amy 2015/03/17 訊息依原承辦人業務區區分,增加顯示原承辦人員編-玲玲
                If Left(cp(12), 1) = "F" Then
                    strShow = "與外專"
                Else
                    strShow = "與分所"
                End If
                strShow = strShow & "輸入之承辦人 「" & GetStaffName(Text1(0).Tag) & "(" & Text1(Index).Tag & ") 」不同" & vbCrLf & "請再次輸入承辦人！"
                'end 2015/03/17
                If CheckReKey(Text1(Index), Label1(0), strShow) = False Then
                    Text1(Index).Text = cp(14)
                     ChgType (0)
                     Cancel = True
                     m_bolChkCP14OK = False
                End If
            End If
        'Add by Amy 2018/10/18 智權人員非國外部FXX且修改案件性質時,不可改為 902(回覆代理人)
        Case 1
            If Text1(1).Tag <> Text1(1) And Text1(1) = "902" Then
                If Left(PUB_GetStaffST15(Text1(24), 1), 1) <> "F" Then
                    Cancel = True
                    MsgBox "智權人員非國外部，案件性質不可改為902(回覆代理人)"
                    Text1(1).SetFocus
                End If
            End If
            OptSendType(1).Caption = PUB_GetCP114Opt1Desc(cp(1), Text1(1)) 'Added by Morgan 2024/1/19
        'Added by Lydia 2017/05/05 客戶案件案號長度控制
        Case 8
            'Modified by Lydia 2017/06/14 改常數
            'Cancel = Not CheckLengthIsOK(Text1(Index), 100)
            Cancel = Not CheckLengthIsOK(Text1(Index), 專利客戶案號max)
        'end 2017/05/05
        Case Else
   End Select
   'end 2015/01/22
   
   If Cancel Then
      TextInverse Text1(Index)
   'Added by Morgan 2022/12/28
   ElseIf Index = 1 Or Index = 13 Then
      SetPA178
   'end 2022/12/28
   End If
End Sub

Private Sub GridHead()
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1000: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 900: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 900: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1500: .Text = "機關文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1200: .Text = "解除期限日期"
      .col = 7: .ColWidth(7) = 0
      .col = 8: .ColWidth(8) = 0
      .col = 9: .ColWidth(9) = 0
      ' 90.07.06 modify by louis
      .col = 10: .ColWidth(10) = 2000: .Text = "備註"
      .CellAlignment = flexAlignCenterCenter
      '判斷是否有資料
      .Visible = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2007/3/23
   
   'Add by Amy 2023/01/05 frm880002從此支開啟改不為強制表單,故需判斷存在時要關
   strPrity1 = "": strPrity2 = "": strPrity3 = "": strPrity4 = "": strPrity5 = ""
   'Add by Sindy 2022/12/15 若接洽單已開需關閉
   If PUB_CheckFormExist("frm090801_Q") = True Then
      Unload frm090801_Q
   End If
   '2022/12/15 END
   'Set frm040101_1 = Nothing 'Removed by Morgan 2021/12/9 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim ii As Integer
Dim Cancel As Boolean
   
   CheckDataValid = False
    'Add By Cheng 2002/12/31
    Dim strKey1 As String
    Dim StrKey2 As String
    Dim strKey3 As String
    Dim strKey4 As String
   
    'Add By Cheng 2002/12/31
    If IsEmptyText(textPA1) = True And IsEmptyText(textPA2) = False Then
       strTit = "檢核資料"
       strMsg = "轉本所案號輸入不完整 !"
       nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
       textPA1.SetFocus
       textPA1_GotFocus
       Exit Function
    End If
    If IsEmptyText(textPA1) = False And IsEmptyText(textPA2) = True Then
       strTit = "檢核資料"
       strMsg = "轉本所案號輸入不完整 !"
       nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
       textPA1.SetFocus
       textPA1_GotFocus
       Exit Function
    End If
    If IsEmptyText(textPA1) = False And IsEmptyText(textPA2) = False Then
       strKey1 = textPA1
       StrKey2 = textPA2
       strKey3 = textPA3
       If IsEmptyText(strKey3) Then strKey3 = "0"
       strKey4 = textPA4
       If IsEmptyText(strKey4) Then strKey4 = "00"
       If GetGrid(strKey1 & StrKey2 & strKey3 & strKey4, 1) = False Then
             Me.textPA2.SetFocus
             textPA2_GotFocus
             Exit Function
       End If
       '2010/12/2 add by sonia
       If strKey1 = cp(1) And StrKey2 = cp(2) And strKey3 = cp(3) And strKey4 = cp(4) Then
         strTit = "檢核資料"
         strMsg = "轉本所案號不可與原本所案號相同 !"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         textPA2.SetFocus
         textPA2_GotFocus
         Exit Function
       End If
       '2010/12/2 end
    End If
    
    'Modify By Cheng 2002/11/22
    '若非執行轉本所案號功能
   If Me.textPA1.Text = "" Or Me.textPA2.Text = "" Then
        'Add By Sindy 2022/11/22
        'Modify By Sindy 2023/1/18 + And m_bolFMP = False : 排除FMP案件
        If strSrvDate(1) >= 接洽單電子收文啟用日 And m_bolFMP = False Then
            'Modify By Sindy 2023/4/12 + , , , , Label3(8)
            If PUB_CRLUseCP07CheckCP06(m_CP31, Text1(13), cp(1), Text1(1), Text1(4), Text1(5), , , , Label3(8)) = False Then
               Text1(4).SetFocus
               Exit Function
            End If
        End If
        '2022/11/22 END
      
        ' 檢查本所期限與法定期限
        Select Case m_CP10
           ' 異議及舉發不檢查本所期限及法定期限
           'Modify by Morgan 2007/8/30 加第三人申請技術報告
           'Case "803":
           Case "803", "807":
           Case "801":
              '93.4.7 ADD BY SONIA
              If DBDATE(Text1(5)) < strSrvDate(1) Then
                 strTit = "法定期限"
                 strMsg = "法定期限不可小於系統日期"
                 nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
                 Text1(7).SetFocus
                 Text1_GotFocus (7)
                 GoTo EXITSUB
              End If
              '93.4.7 END
           Case Else '
              If DBDATE(m_CP06) <> DBDATE(Text1(4)) Or DBDATE(m_CP07) <> DBDATE(Text1(5)) Then
                 strTit = "本所期限與法定期限"
                 strMsg = "本所期限與法定期限與原資料不同, 確定是否存檔?"
                 nResponse = MsgBox(strMsg, vbYesNo, strTit)
                 If nResponse = vbNo Then
                    GoTo EXITSUB:
                 End If
              End If
        End Select
        
        ' 檢查PCT案件
        '92.1.11 MODIFY BY SONIA
        'If Text1(13) = 大陸國家代號 And (cp(10) = "101" Or cp(10) = 102 Or cp(10) = 103 Or cp(10) = 104) Then
        If Text1(13) = 大陸國家代號 And Text1(1) = "101" Then
        '92.1.11 END
           If IsEmptyText(Text1(14)) = True Then
              strTit = "是否PCT案件"
              strMsg = "請確認是否不是PCT案件?"
              nResponse = MsgBox(strMsg, vbYesNo + vbDefaultButton1, strTit)
              If nResponse = vbNo Then
                 Text1(14).SetFocus
                 Text1_GotFocus (14)
                 GoTo EXITSUB
              End If
              'If nResponse = vbNo Then
              '   GoTo EXITSUB:
              'End If
           End If
        End If
        
        '91.11.3 add by sonia
        'Modify by Morgan 2004/10/14    加 121-主張國內優先權
        'If Me.Text1(1).Text = 主張優先權 And strPriority(2) = "" Then
        'Modified by Morgan 2012/7/2 +回復優先權主張124
        '20140211START Modify By eric 改為--詢問是否先輸入優先權資料,若是,則顯示優先權輸入畫面讓user輸入,否則繼續往下執行
        'Modify by Amy 2023/01/05 原:strPriority(1)
        If (Me.Text1(1).Text = 主張優先權 Or Text1(1).Text = "121" Or Text1(1).Text = "124") And strPrity1 = "" Then
           strTit = "檢核資料"
           If Me.Text1(1).Text = 主張優先權 Then
               strMsg = "案件性質為" & Label3(1) & ", 是否先輸入優先權資料? 若選擇「否」,系統於 7 天後會通知智權同仁補資料 "
               If MsgBox(strMsg, vbYesNo + vbQuestion, strTit) = vbYes Then
                  'Modify by Amy 2014/06/10 +strPriority(5)
                  'Modify by Amy 2023/01/05 strPriority原陣列,改變數,並加表單名
                  ModifyPriority strPrity1, strPrity2, strPrity3, pa(8), , pa(1) & pa(2) & pa(3) & pa(4), pa(9), , strPrity4, strPrity5, , Me
                  GoTo EXITSUB
                  
               'Added by Morgan 2022/9/14
               ElseIf Text1(4) = "" Then
                  MsgBox "案件性質為" & Label3(1) & ", 若不輸入優先權資料則本所期限不可為空白！", vbExclamation
                  Text1(4).SetFocus
                  GoTo EXITSUB
               'end 2022/9/14
               End If
           Else
              strMsg = "案件性質為" & Label3(1) & ", 請先輸入優先權資料 !"
              nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
              Text1(1).SetFocus
              Text1_GotFocus (1)
              GoTo EXITSUB
           End If
           
        End If
        'If (Me.Text1(1).Text = 主張優先權 Or Text1(1).Text = "121" Or Text1(1).Text = "124") And strPriority(1) = "" Then
        '   strTit = "檢核資料"
        '   strMsg = "案件性質為" & Label3(1) & ", 請先輸入優先權資料 !"
        '   nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
        '   Text1(1).SetFocus
        '   Text1_GotFocus (1)
        '   GoTo EXITSUB
        'End If
        '20140211END
        '91.11.3 end
        
        
        'Add By Cheng 2003/08/13
        '若案件性質為延期, 則不可點選本案期限
        If Me.Text1(1).Text = "404" Then
            For ii = 1 To Me.MSHFlexGrid1.Rows - 1
                If Me.MSHFlexGrid1.TextMatrix(ii, 0) <> "" Then
                    MsgBox "此案僅收文<延期>，不可點選下一程序期限資料，" & vbCrLf & "否則無法管制下一程序的期限!!!", vbExclamation + vbOKOnly
                    GoTo EXITSUB
                End If
            Next ii
        End If

    Else
        'Add By Cheng 2002/05/17
        '檢查轉本所案號系統類別
        If IsEmptyText(textPA1) = False Then
           If textPA1 <> pa(1) Then
              strTit = "檢核資料"
              strMsg = "轉本所案號必須與原本所案號之系統類別相同 !"
              nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
              Me.textPA1.SetFocus
              textPA1_GotFocus
              GoTo EXITSUB
           End If
        End If
        'Add By Cheng 2002/05/17
        '檢查轉本所案號輸入之完整性
        If (IsEmptyText(textPA1) = True And IsEmptyText(textPA2) = False) Or _
           (IsEmptyText(textPA1) = False And IsEmptyText(textPA2) = True) Then
           strTit = "檢核資料"
           strMsg = "轉本所案號輸入不完整 !"
           nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
           textPA1.SetFocus
           textPA1_GotFocus
           GoTo EXITSUB
        End If
    End If
    
   'Add By Sindy 2010/10/29
   'Modify by Morgn 2010/11/2 加控制 101,102 才要
   'If Val(DBDATE(Text1(12))) >= 20101102 And m_CP31 = "Y" And Left(PUB_GetStaffST15(Trim(Text1(24)), 1), 1) <> "F" Then
   If Val(DBDATE(Text1(12))) >= 20101102 And m_CP31 = "Y" And Left(cp(12), 1) <> "F" And (Text1(1) = "101" Or Text1(1) = "102") Then
      If Combo3 = "" Then
         strTit = "檢核資料"
         strMsg = "案件屬性不可空白 !"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
         Combo3.SetFocus
         GoTo EXITSUB
      End If
   End If
   'Modified by Morgan 2018/9/18
   'If Combo3 <> "" Then
   If Combo3 <> "" And Combo3.Enabled = True Then
   'end 2018/9/18
      'Added by Morgan 2015/10/2
      If Text1(13) <> "000" And Text1(2) = "3" Then
         MsgBox "非台灣設計案不可輸入案件屬性！", vbExclamation
         Combo3.SetFocus
         GoTo EXITSUB
      End If
      'end 2015/10/2
      Cancel = False
      Combo3_Validate Cancel
      If Cancel = True Then
         GoTo EXITSUB
      End If
      'Add by Morgan 2010/11/1
      If PUB_ChkRefCasePA158(cp(1), cp(2), cp(3), cp(4), Left(Combo3, 1)) = False Then
         GoTo EXITSUB
      End If
      
   End If
   '2010/10/29 End
    
   'Add by Amy 2014/09/05 若承辦加乘註記或承辦人基數有改則註記修改理由一定要輸
   If (Val(txtCP97) <> Val(cp(97)) Or Val(txtCP98) <> Val(cp(98))) And (txtCP99 = cp(99) Or txtCP99 = "") Then
        MsgBox "承辦加乘註記 或 承辦人基數有修改" & vbCrLf & "「註記修改理由」一定要輸入!", vbExclamation + vbOKOnly
        SSTab1.Tab = 1
        txtCP99.SetFocus
        txtCP99_GotFocus
        GoTo EXITSUB
   End If
   'end 2014/09/05
   
   'Add By Sindy 2014/9/17 承辦人有異動時檢查是否有設定核判表
   If Text1(0).Text <> "" And Text1(0).Text <> Text1(0).Tag Then
      'Add By Sindy 2014/9/17 檢查是否有設定核判表
      If InStr("P10,P11", GetST15(Text1(0).Text)) > 0 Then
         'Modify By Sindy 2024/6/26 +Text1(13)
         If PUB_ChkIsSetPromoterReader(Text1(0).Text, cp(1), cp(10), , , , Text1(13)) = False Then
            'Modified by Morgan 2025/2/20 游經理->李柏翰經理
            MsgBox "此承辦人該案件性質尚未設定核判表，請通知李柏翰經理轉電腦中心設定後再進行分案。", vbInformation
            SSTab1.Tab = 0
            Text1(0).SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   '2014/9/17 END
   
   CheckDataValid = True
   
EXITSUB:
End Function

Private Sub textPA1_GotFocus()
   'Add By Cheng 2002/05/17
   TextInverse Me.textPA1
   CloseIme
End Sub

Private Sub textPA1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPA1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textPA1) = False Then
      If textPA1 <> pa(1) Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "轉本所案號必須與原本所案號之系統類別相同 !"
         nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      End If
   End If
End Sub

Private Sub textPA2_GotFocus()
   'Add By Cheng 2002/05/17
   TextInverse Me.textPA2
End Sub

Private Sub textPA3_GotFocus()
   'Add By Cheng 2002/05/17
   TextInverse Me.textPA3
End Sub

Private Sub textPA3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textPA4_GotFocus()
   'Add By Cheng 2002/05/17
   TextInverse Me.textPA4
End Sub

Private Sub textPA4_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
    'Add By Cheng 2002/12/31
    Dim strKey1 As String
    Dim StrKey2 As String
    Dim strKey3 As String
    Dim strKey4 As String
   
   If IsEmptyText(textPA1) = True And IsEmptyText(textPA2) = False Then
      strTit = "檢核資料"
      strMsg = "轉本所案號輸入不完整 !"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      textPA1.SetFocus
      GoTo EXITSUB
   End If
   If IsEmptyText(textPA1) = False And IsEmptyText(textPA2) = True Then
      strTit = "檢核資料"
      strMsg = "轉本所案號輸入不完整 !"
      nResponse = MsgBox(strMsg, vbOKOnly + vbCritical, strTit)
      textPA1.SetFocus
      GoTo EXITSUB
   End If
    'Add By Cheng 2002/12/31
   If IsEmptyText(textPA1) = False And IsEmptyText(textPA2) = False Then
      strKey1 = textPA1
      StrKey2 = textPA2
      strKey3 = textPA3
      If IsEmptyText(strKey3) Then strKey3 = "0"
      strKey4 = textPA4
        'Modify By Cheng 2002/12/31
'      If IsEmptyText(strKey4) Then strKey3 = "00"
      If IsEmptyText(strKey4) Then strKey4 = "00"
      If GetGrid(strKey1 & StrKey2 & strKey3 & strKey4, 1) = False Then
            Me.textPA2.SetFocus
            Exit Sub
      End If
      Text1(6) = ""
   End If
   
   'Add By Cheng 2002/08/23
   If Me.textPA1.Text <> "" And Me.textPA2.Text <> "" Then
      MsgBox "不可同時做" & """轉本所案號""" & "及" & """修改其他欄位""" & "，否則資料可能有不一致的情形產生!!!", vbExclamation + vbOKOnly
   End If

EXITSUB:
End Sub

Private Sub textPA4_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
    'Modify By Cheng 2002/12/31
    '移至LostFocus做Check
'   If IsEmptyText(textPA1) = False And IsEmptyText(textPA2) = False Then
'      Dim strKey1 As String
'      Dim strKey2 As String
'      Dim strKey3 As String
'      Dim strKey4 As String
'      strKey1 = textPA1
'      strKey2 = textPA2
'      strKey3 = textPA3
'      If IsEmptyText(strKey3) Then strKey3 = "0"
'      strKey4 = textPA4
'        'Modify By Cheng 2002/12/31
''      If IsEmptyText(strKey4) Then strKey3 = "00"
'      If IsEmptyText(strKey4) Then strKey4 = "00"
'      If GetGrid(strKey1 & strKey2 & strKey3 & strKey4, 1) = False Then
'         Cancel = True
'      End If
'      Text1(6) = ""
'   End If
   'Add By Cheng 2002/09/09
'   'Add By Cheng 2002/08/23
'   If Cancel = False Then
'      If Me.textPA1.Text <> "" And Me.textPA2.Text <> "" Then
'         MsgBox "不可同時做" & """轉本所案號""" & "及" & """修改其他欄位""" & "，否則資料可能有不一致的情形產生!!!", vbExclamation + vbOKOnly
'      End If
'   End If

End Sub
'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim stInCNo(1 To 4) As String '國內案案號
Dim stCP98 As String 'Added by Morgan 2021/8/24

TxtValidate = False

'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

'Modify By Cheng 2002/11/22
'若非執行轉本所案號功能
If Me.textPA1.Text = "" Or Me.textPA2.Text = "" Then
   
   'Added by Morgan 2020/11/9
   If Text1(13) = "000" Then
      If ChkTwSupa() = False Then Exit Function
      
      'Added by Morgan 2025/3/13
      '台灣自請撤回相關收文號檢查--韻丞
      If Text1(1) = "413" And Text1(6) <> "" Then
         If PUB_ChkTW413(Text1(6)) = False Then
            Exit Function
         End If
      End If
      'end 2025/3/13
      
   End If
   'end 2020/11/9
   
    'Add by Morgan 2010/12/29
    If Frame1.Enabled = True And OptSendType(3).Value = True And txtCP142.Enabled = True Then
      If txtCP142.Text = "" Then
         MsgBox "送件方式選指定日期送件時，指定日期不可空白！", vbExclamation
         txtCP142.SetFocus
         Exit Function
      Else
         Cancel = False
         txtCP142_Validate Cancel
         If Cancel = True Then
            txtCP142.SetFocus
            Exit Function
         End If
         
         'Added by Morgan 2023/8/29
         If Frame2.Visible = True Then
            If Option1(0).Value = False And Option1(1).Value = False And Option1(2).Value = False Then
               MsgBox "有輸入指定日期，當天或之前或之後請擇一。", vbExclamation
               Exit Function
            End If
         End If
         'end 2023/8/29
      End If
    End If
    
    For Each objTxt In Text1
        If objTxt.Enabled = True Then
            Cancel = False
            'Add by Amy 2015/01/22 +if
            If objTxt.Index = 0 And m_bolChkCP14OK = True Then
                Cancel = Not ChgType(objTxt.Index)
            Else
                Text1_Validate objTxt.Index, Cancel
            End If
            If Cancel = True Then
                Me.Text1(objTxt.Index).SetFocus
                Text1_GotFocus objTxt.Index
                Exit Function
            End If
        End If
    Next
    
   'add by sonia 2019/5/17 改案件性質時(P108898申復改為改請)要提醒
   If Me.Text1(2).Enabled = False And _
      ((Me.Text1(1).Text >= "101" And Me.Text1(1).Text <= "103") Or _
      (Me.Text1(1).Text >= "301" And Me.Text1(1).Text <= "303")) Then
      If Mid(Me.Text1(1).Text, 3, 1) <> Me.Text1(2).Text Then
         MsgBox "專利種類必須與案件性質的第三碼相同, 請存檔後再分案一次修改專利種類!!!", vbExclamation + vbOKOnly
      End If
   End If
   'end 2019/5/17

   'Add by Morgan 2004/9/30
   If txtCP71.Visible = True And txtCP71.Enabled = True Then
      If txtCP71 = "" Then
         MsgBox "延緩月數/日期不可空白！"
         txtCP71_GotFocus
         Cancel = True
         Exit Function
      Else
         Cancel = False
         txtCP71_Validate Cancel
         If Cancel = True Then
            txtCP71.SetFocus
            txtCP71_GotFocus
            Exit Function
         End If
      End If
   End If
   
   'Add by Morgan 2010/3/16
   For ii = 1 To 2
      If txtFeeYear(ii).Visible And txtFeeYear(ii).Enabled = True Then
         Cancel = False
         txtFeeYear_Validate ii, Cancel
         If Cancel = True Then
            txtFeeYear(ii).SetFocus
            txtFeeYear_GotFocus ii
            Exit Function
         End If
      End If
   Next
   'end 2010/3/16
   
   'Add by Morgan 2006/5/24
    If cp(27) = "" Then
    
    
      '檢查PCT資料
      If Text1(14) <> "" Or Text1(18) <> "" Then
         If Text1(14) = "" Then
            MsgBox "有PCT優先權日時PCT申請日不可空白！", vbExclamation
            Text1(14).SetFocus
            Cancel = True
            Exit Function
         ElseIf Val(Text1(18)) > Val(Text1(14)) Then
            MsgBox "PCT優先權日不可晚於PCT申請日！", vbExclamation
            Text1(18).SetFocus
            Cancel = True
            Exit Function
         Else
            SetPCTDate
         End If
      End If
      
      'Add by Morgan 2009/11/30
      If Text1(13) = 大陸國家代號 And (Text1(1) = "101" Or Text1(1) = "102") Then
         If Text1(14) <> "" And Text1(23) = "" Then
            MsgBox "PCT案之PCT申請號不可空白！", vbExclamation
            Text1(23).SetFocus
            Exit Function
         End If
      End If
      
      If Text1(13) = 大陸國家代號 Then
         '大陸 PCT 案的發明申請必須要有期限
         'Modify by Morgan 2010/6/15 新型也要
         If Text1(14) <> "" And (Text1(1) = 發明申請 Or Text1(1) = 新型申請) Then
            If Text1(5) = "" Or Text1(4) = "" Then
               MsgBox "大陸 PCT 案必須要有期限 !", vbCritical
               Cancel = True
               Exit Function
            End If
         End If
         
'Remove by Morgan 2009/12/22 通知實審會更新,不必再控制
'         '2009/8/24 add by sonia大陸發明案主動補正,必須要有期限P-084560
'         If pa(8) = "1" And Text1(1) = 主動修正 Then
'            If Text1(5) = "" Or Text1(4) = "" Then
'               MsgBox "大陸發明案主動補正,必須要有期限 !", vbCritical
'               Cancel = True
'               Exit Function
'            End If
'         End If
'         '2009/8/24 end
         
         'Add by Morgan 2009/11/5
         If Text1(1) = "414" And Text1(6) = "" Then
            MsgBox "大陸恢復權利必須要輸相關總收文號 !", vbCritical
            Text1(6).SetFocus
            Exit Function
         End If
         'end 2009/11/5
         
         'Added by Morgan 2016/1/4
         '大陸案申請人國籍必須為臺灣才可主張臺灣優先權--郭
         'Modified by Morgan 2018/1/23 +大陸、香港、澳門國籍也可以可主張臺灣優先權--郭
         'Modify by Amy 2023/01/05 原:strPriority(1)
         If Text1(1) = "106" And InStr(strPrity1, "000") > 0 Then
            strExc(1) = ""
            For ii = 0 To 4
               If pa(26 + ii) <> "" Then
                  strExc(2) = Left(pa(26 + ii) & "000", 9)
                  strExc(0) = "select cu01||cu02 from customer where cu01='" & Left(strExc(2), 8) & "' and cu02='" & Mid(strExc(2), 9) & "' and cu10>'010' and cu10<>'020' and cu10<>'044' and cu10<>'013'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strExc(1) = "Y"
                     Exit For
                  End If
               End If
            Next
            If strExc(1) = "Y" Then
               MsgBox "大陸案申請人國籍必須為臺灣或大陸地區才可主張臺灣優先權！", vbCritical
               Exit Function
            End If
         End If
         'end 2016/1/4
      End If
   
      'Added by Lydia 2015/09/10 P案案件性質為自請撤回413時，相關總收文號不可空白
      If pa(1) = "P" And Text1(1).Text = "413" And Text1(6).Text = "" Then
         MsgBox "自請撤回必須要輸相關總收文號 !", vbCritical
         Text1(6).SetFocus
         Exit Function
      End If
   
      'Added by Lydia 2024/09/02 422加速審查一定要掛相關總收文號，以利後續抓資料能正確處理
      'Modified by Morgan 2024/11/18 +447再審查加速審查
      If pa(1) = "P" And Text1(13) = "000" And (Text1(1) = "422" Or Text1(1) = "447") And Text1(6) = "" Then
         MsgBox "必須輸入相關收文號！", vbExclamation
         Text1_GotFocus 6
         Text1(6).SetFocus
         Exit Function
      End If
      'end 2024/09/02
   
      'Added by Morgan 2013/3/18
      '分析更新相關總收文號並檢查期限
      If Text1(1) = "941" Then
         If Text1(6) = "" Then
            If MsgBox(Label3(1) & "是否要串相關總收文號？", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
               Exit Function
            End If
         Else
            For intI = 1 To MSHFlexGrid1.Rows - 1
               If Text1(6).Text = MSHFlexGrid1.TextMatrix(intI, 7) Then
                  If DBDATE(Text1(4)) > DBDATE(MSHFlexGrid1.TextMatrix(intI, 2)) Then
                     MsgBox "本所期限已超過相關收文號的本所期限，將自動更新為相關收文號的本所期限！", vbExclamation
                     Text1(4) = TransDate(DBDATE(MSHFlexGrid1.TextMatrix(intI, 2)), 1)
                  End If
                  Exit For
               End If
            Next
         End If
      End If
      'end 2013/3/18
      
      'Added by Morgan 2013/8/29
      If Text1(13) = 台灣國家代號 And Text1(1) = "705" Then
         If Text1(6) = "" Then
            MsgBox "終止授權必須選取相關總收文號!!", vbExclamation
            Exit Function
         Else
            strExc(0) = "select * from caseprogress where cp09='" & Text1(6) & "' and cp10 in ('704','709')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               MsgBox "相關總收文號案件性質必須為 [授權] 或 [專屬授權] !!", vbExclamation
               Exit Function
            End If
         End If
      End If
      'end 2013/8/29
   End If
   'end 2006/5/24
   
   'Add by Morgan 2010/2/5
   '若有國內案且為相同承辦之P案則加乘註記-0.5()
   'Modified by Morgan 2016/10/18 +109--郭雅娟
   If Text1(0) <> "" And Text1(15) <> "" And InStr("101,102,103,109", Text1(1)) > 0 Then
      If GetStaffDepartment(Text1(0)) <> "P12" Then 'Add by Morgan 2010/3/9 程序不必
         ChgCaseNo Text1(15), strExc
         '若加乘註記未修改
         'Modified by Morgan 2021/8/24 +最後註記是一案兩請(系統自動更新)要先還原加乘註記後計算 Ex:P127595--郭
         'If strExc(1) = "P" And txtCP98 = cp(98) And Val(cp(98)) >= 1 Then
         '   strExc(0) = "select * from flagstory where fs01='" & cp(9) & "' and fs04='1'"
         '   intI = 1
         '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '   If intI = 0 Then
         If strExc(1) = "P" And txtCP98 = cp(98) Then
            stCP98 = cp(98)
            strExc(0) = "select * from flagstory where fs01='" & cp(9) & "' and fs04='1' order by fs02 desc,fs03 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If InStr(RsTemp("fs07"), "@一案兩請") = 1 Then
                  stCP98 = RsTemp("fs05") '原加乘註記
                  intI = 0
               End If
            End If
            
            If Val(stCP98) >= 1 And intI = 0 Then
         'end 2021/8/24
               'Modify by Morgan 2010/5/11 設計不用(因為計件值較少=0.3)
               strExc(0) = "select cp14 from caseprogress" & _
                  " where " & ChgCaseprogress(Text1(15)) & " and cp10 in ('101','102') and cp57 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If IsNull(RsTemp("cp14")) Then
                     MsgBox "本案有國內案但尚未設定承辦人故無法完成加乘註記控管，請先完成該國內案的分案作業！"
                     Exit Function
                  ElseIf Text1(0) = RsTemp("cp14") Then
                     'Modify by Morgan 2010/5/11 改*0.5
                     'strExc(0) = Val(cp(98)) - 0.5
                     'Modified by Morgan 2014/3/7 4月起新規則(一案兩請不改)
                     'strExc(0) = Round(Val(cp(98)) * 0.5, 1)
                     If Val(DBDATE(cp(5))) >= 20140401 Then
                        strExc(0) = Round(Val(stCP98) * 0.3, 1)
                     Else
                        strExc(0) = Round(Val(stCP98) * 0.5, 1)
                     End If
                     'end 2014/3/7
                     If MsgBox("因本案有國內案且承辦人相同故加乘註記將調整為 " & strExc(0) & "，是否要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                        txtCP99 = "有國內案且承辦人相同(原加乘註記:" & stCP98 & ")"
                        txtCP98 = strExc(0)
                     Else
                        Exit Function
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   'end 2010/2/5
   
   If m_bolFMP Then
      'Add by Morgan 2007/8/9 若為外專收文時員工編號改外譯編號
      'Modify by Morgan 2008/8/20 翻譯才要
      'Modified by Lydia 2016/07/07 改成模組 'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
'      If Left(Text1(0), 1) <> "F" And Text1(1).Text = "201" Then
'         strExc(1) = PUB_GetMapID(Text1(0), 0)
'         If strExc(1) <> "" Then
'            Text1(0) = strExc(1)
'         End If
'      End If
'      'end 2007/8/9
       'Remvoe by Lydia 2021/06/30 已不符合現況; 取消內專分案在新案翻譯員工編號改外譯編號的設定, ex.P-126779的核稿人F5370
       'Call PUB_GetPfmpCP14(Text1(1).Text, Text1(0))
      
      'Added by Morgan 2012/3/9 等舊資料補齊後改強制
      'Modified by Morgan 2012/3/19 改強制要輸入
      If txtEngGroup.Visible = True Then
         'Added by Lydia 2023/04/26 預防從非新案的收文輸入組別; ex.P-131332從中說942先分案,造成命名作業未分案
         If txtEngGroup <> "" And txtEngGroup.Tag = "" And cp(31) <> "Y" Then
            MsgBox "國外部收文案件，請先從新案開始分案！", vbExclamation
            Exit Function
         End If
         'end 2023/04/26
         If txtEngGroup = "" Then
            MsgBox "國外部收文案件，工程師組別不可空白！", vbExclamation
            Exit Function
         End If
      End If
      
   End If
   
   'Add by Morgan 2010/3/15
   'Modified by Morgan 2022/9/28
   'If Text1(25).Enabled = True Then
   If Text1(25).Enabled And Text1(25).Visible Then
   'end 2022/9/28
      If Text1(4) <> "" And Text1(25) <> "" Then
         If Val(Text1(25)) > Val(Text1(4)) Then
            MsgBox "承辦期限不可晚於本所期限！"
            Text1(25).SetFocus
            Exit Function
         End If
      End If
   End If
   
   'Add by Morgan 2010/3/17
   If txtFavDt.Visible Then
      If txtFavDt.Text = "" Then
         MsgBox "新穎性優惠期日期不可空白！"
         txtFavDt.SetFocus
         Exit Function
         
      'Removed by Morgan 2018/3/16
      'Else
      '   Cancel = False
      '   txtFavDt_Validate Cancel
      '   If Cancel = True Then
      '      txtFavDt.SetFocus
      '      txtFavDt_GotFocus
      '      Exit Function
      '   End If
      'end 2018/3/16
      End If
   End If
   
   'Added by Morgan 2012/3/30 台灣高速審查一定要有期限
   If Text1(1) = "434" And Text1(13) = "000" Then
      'Modified by Morgan 2016/10/7 若美國案未提申時只檢查所限
      'If Text1(5) = "" Or Text1(4) = "" Then
      '   MsgBox Label3(11) & Label3(1) & "必須要有期限 !", vbCritical
      If (m_bolTw_SUPA_LawDateChk And Text1(5) = "") Or Text1(4) = "" Then
         MsgBox Label3(11) & Label3(1) & "必須要有" & IIf(m_bolTw_SUPA_LawDateChk, "期限 !", "所限 !"), vbCritical
      'end 2016/10/7
         Cancel = True
         Exit Function
      End If
   End If
   
   'Added by Morgan 2012/7/20
   If txtCP147 <> txtCP147.Tag Then
      If txtCP147 <> GetCP147Default() Then
         Me.SSTab1.Tab = 1
         txtCP147.SetFocus
         If MsgBox("是否為複雜或特殊案件已變更為 [ " & IIf(txtCP147 = "Y", "是", "否") & " ] 與預設值不同是否確定要繼續?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
            Exit Function
         End If
      End If
   End If
   'end 2012/7/20

   'Add By Sindy 2013/12/16
   '若未設定特殊出名公司則提醒
   '2014/1/10 MODIFY BY SONIA 加入NOT m_bolFMP 條件
   If Val(m_CP31isYGetCP05) >= Val(InvoiceStartDate) Then 'Add By Sindy 2014/1/29 +if
      '檢查是否有客戶不開發票
      If txtPA161.Visible = True And txtPA161 = "J" Then
         For ii = 1 To 5
            If Trim(pa(26 + ii - 1)) <> "" Then
               'Modify By Sindy 2024/9/25 增加傳入公司別做判斷
               If PUB_ChkCU144isN(Left(ChangeCustomerL(pa(26 + ii - 1)), 8), Right(ChangeCustomerL(pa(26 + ii - 1)), 1), "", txtPA161, False) = True Then
                  MsgBox Left(ChangeCustomerL(pa(26 + ii - 1)), 8) & Right(ChangeCustomerL(pa(26 + ii - 1)), 1) & "此客戶為不開發票，因此特殊出名公司不可選智權公司 !", vbCritical
                  Me.SSTab1.Tab = 1
                  txtPA161.SetFocus
                  Cancel = True
                  Exit Function
               End If
            End If
         Next ii
      End If
   End If
   '2013/12/16 END
   
   'Add By Sindy 2014/5/22 台灣申請案
   If pa(1) = "P" And Text1(13) = "000" And InStr(NewCasePtyList, Text1(1)) > 0 Then
      If Text1(15) <> "" Then
         ChgCaseNo Text1(15).Text, stInCNo
         '所鍵之關聯案為大陸案已發文,而未收保密審查者
         'Modify By Sindy 2014/6/3 +and pa57 is null and nvl(pa108,0)=0 未閉卷未銷卷
         strExc(0) = "select cp09 from caseprogress,patent where pa01='" & stInCNo(1) & "' and pa02='" & stInCNo(2) & "' and pa03='" & stInCNo(3) & "' and pa04='" & stInCNo(4) & "'" & _
                     " and pa09='020' and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+)" & _
                     " and instr('" & NewCasePtyList & "',cp10)>0 and nvl(cp27,0)>0" & _
                     " and pa57 is null and nvl(pa108,0)=0" & _
                     " and not exists (select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='430')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '檢查發明人國籍是否為中國大陸
            'Modify By Sindy 2014/11/6
            If strSrvDate(1) >= 專利發明人檔啟用日 Then
               strExc(0) = "select pi06 from patentInventor,inventor where pi01='" & stInCNo(1) & "' and pi02='" & stInCNo(2) & "' and pi03='" & stInCNo(3) & "' and pi04='" & stInCNo(4) & "' and substr(pi06,1,8)=in01(+) and substr(pi06,9,2)=in02(+) and in11='020'"
            Else
            '2014/11/6 END
               'Memo by Lydia 2021/08/17 刪除舊程式碼：專利發明人在專利基本檔60~69
            End If
            intI = 1
            Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If adoRecordset.RecordCount > 0 Then
                  MsgBox "請與智權同仁確認大陸案是否要保密審查！" & vbCrLf & stInCNo(1) & stInCNo(2) & "-" & stInCNo(3) & "-" & stInCNo(4), vbInformation
               End If
            End If
         End If
      End If
   End If
   '2014/5/22 END
   
   'Added by Morgan 2015/12/15
   '１.本所期限<=分案日+3個工作天時系統自動產生B類收文（４０４）延期，承辦人請掛處理分案的程序人員
   '２.可延期的案件性質如後：107再審,204修正,205申復,206補充說明,501訴願,804舉發答辯
   m_bolAuto404 = False
   If Text1(13) = "000" And m_bolIsFirstKeyCP14 And cp(157) = "" And Text1(4) <> "" Then
      If InStr("107,204,205,206,501,804", Text1(1)) > 0 Then
         If DBDATE(Text1(4)) <= CompWorkDay(3, CompDate(2, 1, strSrvDate(1))) Then
            strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='404' and cp43='" & cp(9) & "' and cp27||cp57 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               MsgBox Label3(1) & "本所期限<=分案日+3個工作天，系統將自動收文延期！", vbExclamation
               m_bolAuto404 = True
            End If
         End If
      End If
   End If
   'end 2015/12/15
   
   'Added by Morgan 2019/10/28
   '108.11.1新法 --玲玲
   If Text1(13) = 台灣國家代號 And strSrvDate(1) >= 20191101 And cp(27) = "" Then
      '舉發案-補充說明(相關收文號非來函且舉發未審定時，法限=舉發案發文日+3個月)
      If Text1(1) = "206" And Text1(3) = "3" And Left(Text1(6), 1) <> "C" Then
         strExc(0) = "select cp27 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='803' and cp27>0 and cp24||cp57 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Text1(5).Tag = Text1(5)
            strExc(1) = CompDate(1, 3, RsTemp("cp27")) '法限
            strExc(2) = PUB_GetOurDeadline(strExc(1)) '所限
            Text1(5) = TransDate(strExc(1), 1)
            Text1(4) = TransDate(strExc(2), 1)
            If Text1(5) <> Text1(5).Tag Then
               MsgBox "【" & Label3(1) & "】法定期限已設定為舉發案發文日+3個月！", vbExclamation
            End If
         End If
      End If
      
      '新型案-更正
      If Text1(1) = "402" And Text1(2) = "2" Then
         MsgBox "提【更正】必須是被舉發或申請技術報告審查中才能受理，須繳納規費2000!!", vbExclamation
      End If
   End If
   'end 2019/10/28
   
   'Added by Morgan 2020/2/6
   m_str442DeadLine = ""
   If pa(1) = "P" And Text1(1) = "442" And Text1(0) <> "" And cp(27) = "" Then
      If Text1(13) <> "020" Then
         MsgBox "必須為大陸案！", vbCritical
         Exit Function
      ElseIf Text1(6) = "" Then
         MsgBox "相關總收文號不可為空白！", vbCritical
         Exit Function
      Else
         'Added by Morgan 2021/10/5 FMP分割案除外--Sharon
         m_RefCP10 = ""
         strExc(0) = "select cp10 from caseprogress where cp09='" & Text1(6) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_RefCP10 = RsTemp("cp10")
         End If
         If Not (m_bolFMP And m_RefCP10 = "307") Then
         'end 2021/10/5
            If PUB_GetOADeadline(Text1(6), m_str442DeadLine, True) = False Then
               Exit Function
            End If
            
         End If 'Added by Morgan 2021/10/5
      End If
   End If
   'end 2020/2/6
   
   'Added by Morgan 2022/12/28
   If txtPA178.Visible Then
      If strSrvDate(1) > "20230000" Or DBDATE(txtCP142) > "20230000" Then
         If txtPA178 = "" Then
            MsgBox "請輸入證書形式！", vbExclamation
            SSTab1.Tab = 1
            txtPA178.SetFocus
            Exit Function
         End If
      End If
   End If
   'end 2022/12/28
   
   'Added by Morgan 2024/10/7
   '若原來有設案件屬性則不可清除(誤刪) Ex:P-134334
   If pa(158) <> "" And Combo3.Enabled Then
      If Combo3 = "" Then
         MsgBox "案件屬性不可清除！", vbExclamation
         Combo3.SetFocus
         Exit Function
      End If
   End If
   'end 2024/10/7
   
'若執行轉本所案號功能
Else
    If Me.textPA1.Enabled = True Then
       Cancel = False
       textPA1_Validate Cancel
       If Cancel = True Then
          Me.textPA1.SetFocus
          textPA1_GotFocus
          Exit Function
       End If
    End If
    
    If Me.textPA4.Enabled = True Then
       Cancel = False
       textPA4_Validate Cancel
       If Cancel = True Then
          Me.textPA4.SetFocus
          textPA4_GotFocus
          Exit Function
       End If
    End If
End If

   'Added by Morgan 2024/1/30
   If txtCP118.Enabled And Text1(13) = "020" And txtCP118 = "" And cp(9) < "C" Then
      If MsgBox("請確認本案是否為紙本送件？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         txtCP118.SetFocus
         Exit Function
      End If
   End If
   'end 2024/1/30
   
   'Added by Morgan 2025/9/11
   '台大主張優先權管控
   If Text1(13) = "020" And Text1(1) = "106" Then
      If InStr(strPrity1, "000") > 0 Then
         If PUB_ChkCNTWPriority(pa(), strPrity1, strPrity3) = False Then
            If PUB_CNPriorityMsg() = vbNo Then
               Exit Function
            End If
         End If
      End If
   End If
   'end 2025/9/11
   
   TxtValidate = True
End Function

'Add By Cheng 2002/06/10
'取得案件收費表的下次期限天數
Private Function GetCF12(strCF01 As String, strCF02 As String, strCF03 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetCF12 = "0"
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
'91.10.31 MODIFY BY SONIA
'strSQLA = "Select CF12,CF28 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "' AND CF12 IS NOT NULL"
StrSQLa = "Select CF12 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "'"
'91.10.31 END
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount <> 0 Then
   If Not IsNull(rsA.Fields(0).Value) Then
      GetCF12 = rsA.Fields(0).Value
   End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By SONIA 2002/10/31
'取得案件收費表的下次期限月份
Private Function GetCF28(strCF01 As String, strCF02 As String, strCF03 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetCF28 = "0"
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
StrSQLa = "Select CF28 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount <> 0 Then
   If Not IsNull(rsA.Fields(0).Value) Then
      GetCF28 = rsA.Fields(0).Value
   End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2002/06/10
'取得案件收費表的規費
Private Function GetCF08(strCF01 As String, strCF02 As String, strCF03 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

GetCF08 = "0"
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
StrSQLa = "Select CF08 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "' AND CF08 IS NOT NULL"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   GetCF08 = rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2003/07/22
'檢查本所期限是否為當日或假日期限
Private Function WorkDayCheck() As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

WorkDayCheck = False
If m_strCP06Update = "" Then Exit Function
If m_strCP06Update = strSrvDate(1) Then
    WorkDayCheck = True
    Exit Function
End If
StrSQLa = "Select * From Workday Where WD01>" & strSrvDate(1) & " Order By 1 "
rsA.CursorLocation = adUseClient
'Add by Morgan 2003/12/31
rsA.MaxRecords = 1

rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    If Val(m_strCP06Update) >= Val(strSrvDate(1)) And Val(m_strCP06Update) < Val("" & rsA.Fields(0).Value) Then
        WorkDayCheck = True
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function
'取得本所期限
Private Function GetCP06(strCP09 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select CP06 From Caseprogress Where CP09='" & strCP09 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCP06 = "" & rsA.Fields(0).Value
Else
    GetCP06 = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'取得法定期限 add by Toni
Private Function GetCP07(strCP09 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select CP07 From Caseprogress Where CP09='" & strCP09 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCP07 = "" & rsA.Fields(0).Value
Else
    GetCP07 = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function
'end 2008/10/24


'Add by Morgan 2004/3/29
'讀取分割案資料
Private Function GetDivCase() As Boolean

   Dim stSQL As String, rsQuery As String
   
On Error GoTo flgErr

   stSQL = "SELECT DC05, DC06, DC07, DC08 FROM DIVISIONCASE" & _
      " WHERE DC01='" & pa(1) & "' AND DC02='" & pa(2) & "' AND DC03='" & pa(3) & "' AND DC04='" & pa(4) & "'"
      
   intI = 1
   'edit by nickc 2007/02/05 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, stSQL)
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      With RsTemp
         If RsTemp.RecordCount > 0 Then
            'Add by Morgan 2004/3/17
            '分割母案本所案號
            txtDivCaseNo(1) = "" & .Fields("DC05").Value: txtDivCaseNo(1).Tag = txtDivCaseNo(1)
            txtDivCaseNo(2) = "" & .Fields("DC06").Value: txtDivCaseNo(2).Tag = txtDivCaseNo(2)
            txtDivCaseNo(3) = "" & .Fields("DC07").Value: txtDivCaseNo(3).Tag = txtDivCaseNo(3)
            txtDivCaseNo(4) = "" & .Fields("DC08").Value: txtDivCaseNo(4).Tag = txtDivCaseNo(4)
         End If
      End With
      GetDivCase = True
   End If
   If Text1(1) = "307" Then
       DivVisibleSwitch True
   Else
       DivVisibleSwitch False
   End If
         
flgErr:

   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   
End Function
'Add by Morgan 2004/3/29
'分割母案本所案號控制
Private Sub DivVisibleSwitch(Optional bolVisible As Boolean = True)
   Dim idx As Integer
   lblDivCase.Visible = bolVisible
   For idx = 1 To 4
      If bolVisible = False Then
         txtDivCaseNo(idx).Text = txtDivCaseNo(idx).Tag
      End If
      txtDivCaseNo(idx).Visible = bolVisible
   Next
End Sub
'Add by Morgan 2004/7/21
Private Sub txtAD_GotFocus(Index As Integer)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtAD(Index).IMEMode = 2
   CloseIme
   TextInverse txtAD(Index)
End Sub
'Add by Morgan 2004/7/21
'只有公司可輸入 2,3,N
Private Sub txtAD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Modified by Morgan 2014/7/15 學校改預設且不可改
   'If Not (KeyAscii = 8 Or KeyAscii = 50 Or KeyAscii = 51 Or KeyAscii = 78) Then
   If Not (KeyAscii = 8 Or KeyAscii = 51 Or KeyAscii = 78) Then
      KeyAscii = 0
   End If
End Sub

'2012/7/19 add by sonia
Private Sub txtCP118_GotFocus()
   TextInverse txtCP118
End Sub

Private Sub txtCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub
'2012/7/19 end

Private Sub txtCP142_GotFocus()
   TextInverse txtCP142
End Sub

Private Sub txtCP142_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCP142_Validate(Cancel As Boolean)
   If txtCP142 <> "" Then
      If ChkDate(txtCP142) = False Then
         Cancel = True
      'Modified by Morgan 2023/8/31 FMP指定日期可以超過所限
      'ElseIf Val(Text1(4)) > 0 And Val(txtCP142) > Val(Text1(4)) Then
      ElseIf Val(Text1(5)) > 0 And Val(txtCP142) > Val(Text1(5)) Then
         MsgBox "指定送件日期不可晚於法定期限！", vbExclamation
         Cancel = True
      'Modified by Morgan 2024/10/28 +指定日期當天或之後 ex:P-134236
      'ElseIf Not m_bolFMP And Val(Text1(4)) > 0 And Val(txtCP142) > Val(Text1(4)) Then
      ElseIf Not m_bolFMP And Val(Text1(4)) > 0 And Val(txtCP142) > Val(Text1(4)) And (Option1(0).Value Or Option1(2).Value) Then
      'end 2024/10/28
      'end2023/8/31
         MsgBox "指定送件日期不可晚於本所期限！", vbExclamation
         Cancel = True
      'end 2024/10/28
      ElseIf Not ChkWorkDay(DBDATE(txtCP142)) Then
         MsgBox "指定送件日期必須是工作天 !", vbExclamation
         Cancel = True
      End If
   End If
End Sub

Private Sub txtCP147_GotFocus()
   TextInverse txtCP147
End Sub

Private Sub txtCP147_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCP71_GotFocus()
   TextInverse txtCP71
End Sub
'只能key數字
Private Sub txtCP71_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And (KeyAscii < 48 Or KeyAscii > 57) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtCP71_Validate(Cancel As Boolean)
   If txtCP71 <> "" Then
      If Len(txtCP71) = 1 Then
         'Modified by Morgan 2016/3/11 105/3/9日起延緩公告最長改6個月(原3個月)
         If Val(txtCP71) < 1 Or Val(txtCP71) > 6 Then
            MsgBox "延緩公告月數只可輸入1~6！", vbExclamation
            txtCP71_GotFocus
            Cancel = True
         End If
         'end 2016/3/10
      Else
         If ChkDate(txtCP71) = False Then
            txtCP71_GotFocus
            Cancel = True
         'Add by Morgan 2005/3/16
         ElseIf Val(txtCP71) < Val(strSrvDate(2)) Then
            MsgBox "延緩日期不可小於系統日！", vbExclamation
            txtCP71_GotFocus
            Cancel = True
         '2005/3/16 end
         'Added by Morgan 2012/1/31 控制只能輸入 1,11,21--敏惠
         ElseIf Right(txtCP71, 2) <> "01" And Right(txtCP71, 2) <> "11" And Right(txtCP71, 2) <> "21" Then
            MsgBox "延緩日期必須為01,11,21!", vbExclamation
            txtCP71_GotFocus
            Cancel = True
         'end 2012/1/31
         End If
      End If
   End If
End Sub

'Add by amy 2014/09/05
Private Sub txtCP97_GotFocus()
    TextInverse txtCP97
End Sub

Private Sub txtCP97_Validate(Cancel As Boolean)
     If Not IsNumeric(txtCP97) Then
        MsgBox "承辦人計件值輸入錯誤！", vbExclamation
        Cancel = True
      End If
End Sub
'end 2014/09/05

Private Sub txtCP98_GotFocus()
   TextInverse txtCP98
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtCP98.IMEMode = 2
   CloseIme
End Sub

'Add by Morgan 2005/3/4
Private Sub txtCP98_Validate(Cancel As Boolean)
   Dim iMax As Integer
   iMax = 3
   
   'Add by Morgan 2010/9/27
   If bolNewPromoterRule Then
      iMax = 9
   End If
   'end 2010/9/27
   
   If Not IsNumeric(txtCP98) Then
      MsgBox "資料輸入錯誤！", vbExclamation
      Cancel = True
   ElseIf Val(txtCP98) > iMax Then
      MsgBox "資料輸入錯誤！", vbExclamation
      Cancel = True
   End If
End Sub

Private Sub txtCP99_GotFocus()
   TextInverse txtCP99
End Sub

'Add by Morgan 2004/3/29
Private Sub txtDivCaseNo_GotFocus(Index As Integer)
    TextInverse txtDivCaseNo(Index)
End Sub

Private Sub txtDivCaseNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDivCaseNo_Validate(Index As Integer, Cancel As Boolean)
   If Index = 4 Then
      If (txtDivCaseNo(1) <> "" Or txtDivCaseNo(2) <> "" Or txtDivCaseNo(3) <> "" Or txtDivCaseNo(4) <> "") Then
         'Modified by Morgan 2012/11/8 改呼叫公用函數檢查
         'Call CheckDivCase(m_stPA09)
         Call PUB_CheckDivCase(txtDivCaseNo, pa)
      End If
   End If
End Sub

'Removed by Morgan 2012/11/8 改寫為 PUB_CheckDivCase
''Add by Morgan 2004/3/30
''檢查母案是否存在
'Private Function CheckDivCase(ByRef stPA09 As String) As Boolean
'Dim dc(4) As String   '2009/8/24 add by sonia
'
'On Error GoTo flgErr
'
'   Dim stSQL As String, rsQuery As New ADODB.Recordset, stPA08 As String
'
'   If (txtDivCaseNo(1) = "" Or txtDivCaseNo(2) = "") Then
'      MsgBox "分割母案本所案號輸入錯誤！", vbExclamation
'      Exit Function
'   End If
'
'   txtDivCaseNo(1) = Trim(txtDivCaseNo(1))
'   txtDivCaseNo(2) = Right("00000" & txtDivCaseNo(2), 6)
'   txtDivCaseNo(3) = Right("0" & txtDivCaseNo(3), 1)
'   txtDivCaseNo(4) = Right("00" & txtDivCaseNo(4), 2)
'
'   'Add by Morgan 2004/4/29
'   If (txtDivCaseNo(1) = pa(1) And txtDivCaseNo(2) = pa(2) And txtDivCaseNo(3) = pa(3) And txtDivCaseNo(4) = pa(4)) Then
'      MsgBox "分割案不可為母案！", vbExclamation
'      Exit Function
'   End If
'
'   'Modified by Morgan 2012/8/24
'   'stSQL = "select PA08, PA09,PA16 from patent where pa01='" & ChgSQL(txtDivCaseNo(1)) & "' and pa02='" & ChgSQL(txtDivCaseNo(2)) & "' and  pa03='" & ChgSQL(txtDivCaseNo(3)) & "' and pa04='" & ChgSQL(txtDivCaseNo(4)) & "'"
'   stSQL = "select PA08, PA09,PA16,c1.cp25 RD1st,c2.cp25 RD2nd from patent,caseprogress c1,caseprogress c2 where pa01='" & ChgSQL(txtDivCaseNo(1)) & "' and pa02='" & ChgSQL(txtDivCaseNo(2)) & "' and  pa03='" & ChgSQL(txtDivCaseNo(3)) & "' and pa04='" & ChgSQL(txtDivCaseNo(4)) & "'" & _
'      " and c1.cp01(+)=pa01 and c1.cp02(+)=pa02 and c1.cp03(+)=pa03 and c1.cp04(+)=pa04 and c1.cp10(+)='101' and c1.cp24(+)='1'" & _
'      " and c2.cp01(+)=pa01 and c2.cp02(+)=pa02 and c2.cp03(+)=pa03 and c2.cp04(+)=pa04 and c2.cp10(+)='107' and c2.cp24(+) is not null"
'   rsQuery.CursorLocation = adUseClient
'   rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsQuery.RecordCount > 0 Then
'      stPA08 = "" & rsQuery.Fields(0)
'      stPA09 = "" & rsQuery.Fields(1)
'      'Add by Morgan 2004/4/29
'      '分割案與母案的申請國家和專利種類需相同
'      If stPA09 <> pa(9) Then
'         MsgBox "分割案與母案的申請國家需相同！", vbExclamation
'      ElseIf stPA08 <> Text1(2) Then
'         MsgBox "分割案與母案的專利種類需相同！", vbExclamation
'      'Added by Morgan 2012/8/14 102新法
'      ElseIf stPA09 = "000" And rsQuery("RD2nd") > 0 Then
'         MsgBox "台灣母案再審已有准駁不可分割！", vbExclamation
'      ElseIf stPA09 = "000" And rsQuery("RD1st") < 20121202 Then
'         MsgBox "台灣發明母案已初審核准不可分割！", vbExclamation
'      ElseIf stPA09 = "000" And stPA08 = "2" And Not IsNull(rsQuery.Fields("pa16")) Then
'         MsgBox "台灣新型母案已有准駁不可分割！", vbExclamation
'      'end 2012/8/14
'      Else
'         CheckDivCase = True
'         '2009/8/24 add by sonia 分割案未輸入優先權時帶母案優先權
'         '因發文及申請案號時都要再輸一次,故再考慮是否要做
'         'If Text1(1) = "307" And strPriority(1) = "" Then
'         '   dc(1) = txtDivCaseNo(1): dc(2) = txtDivCaseNo(2): dc(3) = txtDivCaseNo(3): dc(4) = txtDivCaseNo(4)
'         '   If ClsPDReadPriority(dc(), strPriority(1), strPriority(2), strPriority(3), strPriority(4)) = False Then GoTo flgErr
'         'End If
'         '2009/8/24 end
'      End If
'   Else
'      MsgBox "分割母案本所案號不存在！", vbExclamation
'   End If
'
'flgErr:
'   Set rsQuery = Nothing
'   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
'
'End Function


'Add by Morgan 2004/6/3
'延緩公告延緩月數輸入控制
Private Sub CP71Switch(ByVal stPA09 As String, ByVal stCP10 As String)
   Dim bolVisible As Boolean
   If stPA09 = "000" And stCP10 = "412" Then
      bolVisible = True
      'Add by Morgan 2004/8/16  一旦有值後就不可空白
      If txtCP71.Text = "" Then txtCP71.Text = cp(71)
   Else
      bolVisible = False
      txtCP71.Text = ""
   End If
   lblCP71.Visible = bolVisible
   txtCP71.Visible = bolVisible
End Sub
'Add by Morgan 2004/6/8
'檢查是否公告日於93.7.1以後且未收技術報告
Private Function Chk421Exist(ByRef p_PA() As String, p_PA23 As String) As Boolean
   Dim stCP10 As String
On Error GoTo ErrHnd
   'Add by Morgan 2007/8/31
   If p_PA23 = "3" Then
      stCP10 = "807"
   Else
      stCP10 = "421"
   End If
   'end 2007/8/31
   strSql = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & p_PA(1) & "' AND CP02='" & p_PA(2) & "' AND CP03='" & p_PA(3) & "' AND CP04='" & p_PA(4) & "' AND CP10='" & stCP10 & "' AND CP57 IS NULL"
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      Chk421Exist = True
   End If
   
ErrHnd:
   
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC
   
End Function

'Add by Morgan 2004/12/15 檢查國內外案件關聯是否存在
Private Function IfCaseMapExist(ByRef stCM() As String) As Boolean

On Error GoTo ErrHnd
   
   strSql = "select 1 from casemap" & _
      " where cm01='" & stCM(1) & "' and cm02='" & stCM(2) & "' and cm03='" & stCM(3) & "' and cm04='" & stCM(4) & "'" & _
      " and cm10='0'"
   
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If AdoRecordSet3.RecordCount > 0 Then
      IfCaseMapExist = True
   End If
   
ErrHnd:

   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Private Sub txtCP99_Validate(Cancel As Boolean)
   Cancel = Not CheckLengthIsOK(txtCP99, txtCP99.MaxLength)
End Sub

'Add by Morgan 2006/5/24
'計算PCT期限
Private Sub SetPCTDate(Optional ByVal p_Msg As Boolean = False)
   '大陸發明案且未發文才要算期限
   'Modify by Morgan 2010/6/15 +新型也要
   'If Text1(13) = 大陸國家代號 And Text1(1) = "101" And cp(27) = "" Then
   If Text1(13) = 大陸國家代號 And (Text1(1) = "101" Or Text1(1) = "102") And cp(27) = "" Then
      '優先權日
      If Text1(18) <> "" And Text1(14) <> "" Then
         strExc(0) = Text1(18)
      ElseIf Text1(14) <> "" Then
         strExc(0) = Text1(14)
      Else
         strExc(0) = ""
      End If
      If strExc(0) <> "" Then
         'Remove by Morgan 2010/1/8 改恢復專利權分案做控制就好
         ''Add by Morgan 2006/9/18 檢查若有收文恢復專利權414時再加2個月
         'If PUB_ChkCPExist(pa, "414") = True Then
         '   PUB_GetPCTLimit strExc(0), strExc(1), strExc(2), Text1(13), True
         'Else
         ''end 2006/9/18
        'Add by Lydia 2014/11/24 PCT進各個國家階段之期限設定
        '    PUB_GetPCTLimit strExc(0), strExc(1), strExc(2), Text1(13)
         PUB_GetPCTLimit strExc(0), strExc(1), strExc(2), pa(1), pa(2), pa(3), pa(4), Text1(13)
         'End If
         'end 2010/1/8
         
         'Add by Morgan 2009/11/4
         'FMP期限抓設定
         If m_bolFMP Then
            strExc(2) = PUB_GetDeadLine(DBDATE(cp(5)), DBDATE(strExc(1)), 2)
            If strExc(2) = "" Then strExc(2) = strExc(1) 'Add by Morgan 2010/2/3
            strExc(2) = TransDate(strExc(2), 1)
         End If
         'end 2009/11/4
         
         If (Text1(4) <> "" Or Text1(5) <> "") And (Text1(5) <> strExc(1) Or Text1(4) <> strExc(2)) Then
            If p_Msg = True Then
               If MsgBox("本案為PCT案，是否更新期限？(法限：" & strExc(1) & "；" & "所限：" & strExc(2) & ")", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                  Exit Sub
               End If
            End If
         End If
         Text1(5) = strExc(1)
         Text1(4) = strExc(2)
      End If
   End If
End Sub

'Add by Morgan 2006/9/8
'設定大陸PCT案恢復專利權的期限
Public Sub Set414Date()
   Dim stSQL As String, adoRst As ADODB.Recordset, adoRst1 As ADODB.Recordset, iR As Integer
   Dim bUpdate As Boolean, strCP133 As String
'Modify by Morgan 2009/11/5
'改抓相關總收文號的期限+2個月且不限制PCT案(領證也可恢復)
'   'Add by Morgan 2006/9/8 恢復權利期限為發明申請期限+2個月
'   If Text1(13) = "020" And Text1(14) <> "" And Text1(1) = "414" Then
'      If Text1(5) = "" Then
'         strPCTPriDate = PUB_GetPCTPriDate(pa(91))
'         If strPCTPriDate <> "" Then
'            strExc(0) = strPCTPriDate
'         Else
'            strExc(0) = pa(10)
'         End If
'         If strExc(0) <> "" Then
'            strExc(0) = CompDate(1, 2, strExc(0))
'            PUB_GetPCTLimit strExc(0), strExc(1), strExc(2), "020"
'            Text1(5) = TransDate(strExc(1), 1)
'            Text1(4) = TransDate(strExc(2), 1)
'         End If
'      End If
'   End If
'   'end 2006/9/8
   If Text1(13) = "020" And Text1(1) = "414" And Text1(6) <> "" Then
      stSQL = "select cp06,cp07,cp05,cp10,cp31 from caseprogress where cp09='" & Text1(6) & "'"
      iR = 1
      Set adoRst = ClsLawReadRstMsg(iR, stSQL)
      If iR = 1 Then
            strExc(1) = "": strExc(2) = ""
            '新案
            If adoRst("cp31") = "Y" Then
               If IsNull(adoRst("cp07")) Then
                  MsgBox "相關收文號尚無期限，請先做該程序分案作業！"
                  GoTo ExitPort
               Else
                'Add by Lydia 2014/11/24 PCT進各個國家階段之期限設定
'                  strExc(2) = CompDate(1, 2, adoRst("cp07")) '法限
'                  'FMP新案要用規則算所限
'                  If m_bolFMP And adoRst("cp31") = "Y" Then
'                     strExc(1) = PUB_GetDeadLine(adoRst("cp05"), strExc(2), 2)
'                  '所限=法限-7天
'                  Else
'                     strExc(1) = CompDate(2, -7, strExc(2))
'                  End If
               
                '(最初)優先權日=PCT優先權日>PCT申請日
                If Text1(18) <> "" And Text1(14) <> "" Then
                    strExc(0) = Text1(18)
                ElseIf Text1(14) <> "" Then
                    strExc(0) = Text1(14)
                End If
                '新案與414恢復的算法一致
                    PUB_GetPCTLimit strExc(0), strExc(1), strExc(2), pa(1), pa(2), pa(3), pa(4), Text1(13)
                'end 'Add by Lydia 2014/11/24
                  
                  'Added by Morgan 2018/3/6
                  'FMP新案要用規則算所限
                  If m_bolFMP Then
                     strExc(2) = PUB_GetDeadLine(adoRst("cp05"), DBDATE(strExc(1)), 2)
                     If strExc(2) = "" Then strExc(2) = strExc(1)
                     strExc(2) = TransDate(strExc(2), 1)
                  End If
                  'end 2018/3/6
               End If
            Else
               '有專利權消滅函
               'Modified by Morgan 2024/3/25 +1610視為撤回 --玲玲 Ex:P-123827
               strExc(0) = "select cp133 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in ('1604','1610') order by cp05 desc,cp09 desc"
               iR = 1
               Set adoRst1 = ClsLawReadRstMsg(iR, strExc(0))
               If iR = 1 Then
                  strCP133 = "" & adoRst1("CP133")
               End If
               '有官方發文日
               If strCP133 <> "" Then
                  '法限
                  strExc(1) = CompDate(1, 2, strCP133)
                  If adoRst("cp10") = "416" Or adoRst("cp10") = "605" Then
                     '所限=法限-10天
                     If m_bolFMP Then
                        strExc(2) = CompDate(2, -10, strExc(1))
                     '所限=法限-1月5天
                     Else
                        'Added by Lydia 2025/10/29
                        If strSrvDate(1) >= 內專本所約定期限啟用日 Then
                           strExc(2) = PUB_GetPOurDeadline(strExc(1), pa(9), , pa(1), "" & adoRst("cp10"))
                        Else
                        'end 2025/10/29
                           strExc(2) = CompDate(2, -5, CompDate(1, -1, strExc(1)))
                        End If 'Added by Lydia 2025/10/29
                     End If
                  Else
                     '所限=法限-7天
                     If m_bolFMP Then
                        strExc(2) = CompDate(2, -7, strExc(1))
                     '所限=法限-10天
                     Else
                        'Added by Lydia 2025/10/29
                        If strSrvDate(1) >= 內專本所約定期限啟用日 Then
                           strExc(2) = PUB_GetPOurDeadline(strExc(1), pa(9), , pa(1), "" & adoRst("cp10"))
                        Else
                        'end 2025/10/29
                           strExc(2) = CompDate(2, -10, strExc(1))
                        End If 'Added by Lydia 2025/10/29
                     End If
                  End If
               Else
                  strExc(2) = strSrvDate(1)
                  strExc(1) = strSrvDate(1)
               End If
            End If
            If DBDATE(strExc(2)) < strSrvDate(1) Then strExc(2) = strSrvDate(1) '所限不可小於系統日
            
            '轉民國年
            strExc(1) = TransDate(strExc(1), 1)
            strExc(2) = TransDate(PUB_GetWorkDay1(strExc(2), True), 1) '抓工作日
            
            If Text1(5) = "" And Text1(4) = "" Then
               bUpdate = True
            ElseIf strExc(2) <> Text1(4) Or strExc(1) <> Text1(5) Then
               strExc(0) = "期限將變更(所限 " & Text1(4) & "->" & strExc(2) & ",法限 " & Text1(5) & "->" & strExc(1) & "),是否要更新?"
               If MsgBox(strExc(0), vbYesNo, "") = vbYes Then
                  bUpdate = True
               End If
            End If
            
            If bUpdate Then
               Text1(4) = TransDate(strExc(2), 1) '所限
               Text1(5) = TransDate(strExc(1), 1) '法限
            End If
         End If
      
   End If
'end 2009/11/5

ExitPort:

   Set adoRst = Nothing
   Set adoRst1 = Nothing
End Sub

Private Sub Mail2Eng()
   Dim CNo
   Dim i As Integer
   CNo = Split(strMail2FEngCP09, ",")
   For i = LBound(CNo) To UBound(CNo)
      strExc(0) = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04 cno,cp14 from caseprogress where cp09='" & CNo(i) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Call PUB_SendMail(strUserNum, RsTemp.Fields("cp14"), CNo(i), RsTemp.Fields(0) & "(" & CNo(i) & ") 的國內案 " & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & " 已改號為 " & textPA1 & "-" & textPA2 & "-" & textPA3 & "-" & textPA4 & "，請修改接洽單上之關聯案號！")
      End If
   Next
End Sub

'Add by Morgan 2010/1/20
'最後繳費年度
Private Function GetLastYear() As String
   If Left(Right(pa(72), 2), 1) = "," Then
      GetLastYear = Right(pa(72), 1)
   Else
      GetLastYear = Right(pa(72), 2)
   End If
End Function
'輸入起迄年度並比較費用
Private Function InputYear(ByRef iStar As Integer, ByRef iEnd As Integer) As Boolean
   Dim stRtn As String
   Dim stFrom As String, stTO As String
   Dim iNext As Integer
   
ReInput:
   iNext = GetLastYear + 1
   stRtn = InputBox("本案尚有預繳年費未退，請輸入欲繳年費的起迄年度以便計算規費！" & vbCrLf & vbCrLf & "輸入格式:起年-迄年" & vbCrLf & "( Ex. 7 or 7-8 )", "預繳年費移作次年輸入")
   If stRtn = "" Then
      Exit Function
   ElseIf InStr(stRtn, "-") > 0 Then
      stFrom = Left(stRtn, InStr(stRtn, "-") - 1)
      stTO = Mid(stRtn, InStr(stRtn, "-") + 1)
      If Not IsNumeric(stFrom) Or Not IsNumeric(stTO) Then
         MsgBox "輸入格式錯誤！"
         GoTo ReInput
      Else
         iStar = Val(stFrom)
         iEnd = Val(stTO)
         If iStar <> iNext Then
            MsgBox "起年輸入錯誤！(是否應為" & iNext & "??)"
            GoTo ReInput
         ElseIf iEnd < iStar Then
            MsgBox "迄年輸入錯誤！"
            GoTo ReInput
         End If
      End If
   Else
      If Not IsNumeric(stRtn) Then
         MsgBox "輸入格式錯誤！"
         GoTo ReInput
      Else
         iStar = Val(stRtn)
         If iStar <> iNext Then
            MsgBox "年度輸入錯誤！(是否應為" & iNext & "??)"
            GoTo ReInput
         End If
         iEnd = iStar
      End If
   End If
   
   InputYear = True
End Function

'Add by Morgan 2010/1/21
'比較可退繳年費與已收文規費
Private Function CompRefund(ByVal lngRefund As Long, lngFee As Long, ByVal stCP10 As String, ByVal lngCP17 As Long) As Boolean
   Dim lngDiff As Long, strAddMsg As String
   If stCP10 = "908" Then
      '少
      If lngFee > lngRefund Then
         lngDiff = lngFee - lngRefund
         MsgBox "本案有可退預繳年費共 " & Format(lngRefund, DDollar) & " 元" & _
            "，現欲繳年費計 " & Format(lngFee, DDollar) & " 元" & _
            "，尚缺 " & Format(lngDiff, DDollar) & " 元，請智權人員改收文【年費】並向客戶補收差額！"
         Exit Function
      End If
   Else
      If lngCP17 > 0 Then strAddMsg = "並通知財務處暫緩開收據"
      '多
      If lngRefund > lngFee Then
         lngDiff = lngRefund - lngFee
         MsgBox "本案有可退預繳年費共 " & Format(lngRefund, DDollar) & " 元" & _
            "，現欲繳年費計 " & Format(lngFee, DDollar) & " 元" & _
            "，尚可退 " & Format(lngDiff, DDollar) & " 元，請智權人員改收文【退費】" & strAddMsg & "！"
         Exit Function
      '+規費還少
      ElseIf lngFee > lngRefund + lngCP17 Then
         lngDiff = lngFee - lngRefund - lngCP17
         
         'Add by Morgan 2011/8/9 若確認為拆收據時可繼續
         'Modified by Morgan 2012/3/19 改只提醒不確認 Ex.P-78797--玲玲
         'If MsgBox("本案收文規費不足，是否為拆收據案件？", vbYesNo + vbDefaultButton2) = vbNo Then
         
            MsgBox "本案有可退預繳年費共 " & Format(lngRefund, DDollar) & " 元" & _
               "，現欲繳年費計 " & Format(lngFee, DDollar) & " 元，收文規費 " & lngCP17 & " 元" & _
               "，尚缺 " & Format(lngDiff, DDollar) & " 元，請智權人員先向客戶補收差額" & strAddMsg & "！"
               
            'Exit Function
         'End If
      End If
   End If
   CompRefund = True
End Function

Private Sub txtEngGroup_Change()
   If txtEngGroup <> "" Then
      Label3(13) = PUB_GetFCPGrpName(txtEngGroup, True)
   Else
      Label3(13) = ""
   End If
End Sub

Private Sub txtEngGroup_GotFocus()
   TextInverse txtEngGroup
End Sub

Private Sub txtEngGroup_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (Chr(KeyAscii) >= "1" And Chr(KeyAscii) <= "4") Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtFavDt_GotFocus()
   TextInverse txtFavDt
   CloseIme
End Sub

'Removed by Morgan 2018/3/16 此處已改不可輸入,統一在frm880020控制以免規則不一致 Ex.P-119764
''Add by Morgan 2010/3/17
''法限=新穎性優惠期日期+6個月
''所限=法限-7天(台灣案-2天)
'Private Sub txtFavDt_Validate(Cancel As Boolean)
'Dim stDate As String, iMonth As Integer
'
'   If txtFavDt <> "" Then
'      If ChkDate(txtFavDt) Then
'         'Modified by Morgan 2017/7/6 台灣改12個月--潘韻丞(請作單)
'         'frm880020 也要同步修改
'         'stDate = TransDate(CompDate(1, 6, txtFavDt), 1)
'         If Text1(13) = 台灣國家代號 Then
'            'Modified by Morgan 2017/9/29 台灣設計6個月 -蕭茹曣
'            If Text1(2) <> "3" Then
'               iMonth = 6
'            Else
'               iMonth = 12
'            End If
'            'end 2017/9/29
'         Else
'            iMonth = 6
'         End If
'         stDate = TransDate(CompDate(1, iMonth, txtFavDt), 1)
'         'end 2017/7/6
'
'         'add by sonia 2014/3/28
'         If Val(stDate) < Val(strSrvDate(2)) Then
'            MsgBox "優惠期+" & iMonth & "個月不可小於系統日！"
'            txtFavDt.SetFocus
'            Cancel = True
'            Exit Sub
'         End If
'         '2014/3/28 end
'         If Text1(5) = "" Or Val(Text1(5)) > Val(stDate) Then
'            Text1(5) = stDate
'            'Added by Morgan 2012/9/10 --郭
'            If Text1(13) = "000" Then
'               'Added by Morgan 2014/10/28
'               If Text1(13) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
'                  Text1(4) = TransDate(PUB_GetOurDeadline(stDate), 1)
'               Else
'               'end 2014/10/28
'                  Text1(4) = TransDate(PUB_GetWorkDay1(CompDate(2, -2, stDate), True), 1)
'               End If 'Added by Morgan 2014/10/28
'            Else
'            'end 2012/9/10
'               Text1(4) = TransDate(PUB_GetWorkDay1(CompDate(2, -7, stDate), True), 1)
'            End If 'Added by Morgan 2012/9/10
'            If Val(Text1(4)) < Val(strSrvDate(2)) Then
'               Text1(4) = strSrvDate(2)
'            End If
'         End If
'      Else
'         Cancel = True
'      End If
'   End If
'   'add by sonia 2013/12/27
'   If Val(txtFavDt) >= Val(strSrvDate(2)) Then
'      MsgBox "優惠期事實發生日期不可大於等於系統日！"
'      txtFavDt.SetFocus
'      Cancel = True
'   End If
'   'end 2013/12/27
'End Sub
'end 2018/3/16

Private Sub txtFeeYear_GotFocus(Index As Integer)
   TextInverse txtFeeYear(Index)
   CloseIme
End Sub
'Add by Morgan 2010/3/16
Private Sub txtFeeYear_Validate(Index As Integer, Cancel As Boolean)
   Dim strNext As String, strYear As String
   If txtFeeYear(Index) = "" Then
      'Modify by Morgan 2011/8/17 +601
      If Text1(1) = "605" Or Text1(1) = "601" Then
         MsgBox "繳費年度不可空白！"
      Else
         MsgBox "繳費次數不可空白！"
      End If
      txtFeeYear(Index).SetFocus
      Cancel = True
      Exit Sub

   Else
      If Index = 1 Then
         'Add by Morgan 2011/8/17
         If Text1(1) = "601" Then
            If m_RefCP53 <> "" Then
               If Val(txtFeeYear(Index)) <> Val(m_RefCP53) Then
                  MsgBox "繳費(起)年度有誤，應為" & m_RefCP53 & "！"
                  txtFeeYear(Index).SetFocus
                  Cancel = True
                  Exit Sub
               End If
               
            ElseIf pa(9) = "000" Then
            
               If Val(txtFeeYear(Index)) <> 1 Then
                  MsgBox "繳費(起)年度有誤，應為 1！"
                  txtFeeYear(Index).SetFocus
                  Cancel = True
                  Exit Sub
               End If
               
            End If
            
            'Add by Morgan 2011/8/18
            If Val(txtFeeYear(1)) <> Val(cp(53)) Then
               MsgBox "繳費(迄)年度與櫃檯輸入[ " & cp(53) & " ]不同！"
            End If
            
         Else
         'end 2011/8/17
         
            strNext = PUB_Getnexttimes(pa(1), pa(2), pa(3), pa(4), strYear)
            If strNext <> "" Then
               
               If Text1(1) = "605" Then '繳費年度
                  If Val(txtFeeYear(Index)) <> Val(strYear) Then
                     MsgBox "繳費(起)年度有誤，應為" & strYear & "！"
                     txtFeeYear(Index).SetFocus
                     Cancel = True
                     Exit Sub
                  End If
               Else '繳費次數
                  If Val(txtFeeYear(Index)) <> Val(strNext) Then
                     MsgBox "繳費(起)次數有誤，應為" & strNext & "！"
                     txtFeeYear(Index).SetFocus
                     Cancel = True
                     Exit Sub
                  End If
               End If
            Else
               If Text1(1) = "605" Then  '繳費年度
                  MsgBox "無下次繳費年度！"
               Else '繳費次數
                  MsgBox "無下次繳費次數！"
               End If
               txtFeeYear(Index).SetFocus
               Cancel = True
               Exit Sub
            End If
            
         End If
         
      Else
         If Val(txtFeeYear(1)) > Val(txtFeeYear(2)) Then
            'Modify by Morgan 2011/8/17 +601
            If Text1(1) = "605" Or Text1(1) = "601" Then  '繳費年度
               MsgBox "繳費(迄)年度不可小於(起)年度！"
            Else '繳費次數
               MsgBox "繳費(迄)次數不可小於(起)次數！"
            End If
            txtFeeYear(Index).SetFocus
            Cancel = True
            Exit Sub
            
         'Add by Morgan 2011/8/18
         ElseIf Val(txtFeeYear(2)) <> Val(cp(54)) Then
            MsgBox "繳費(迄)年度與櫃檯輸入[ " & cp(54) & " ]不同！"
         
         End If
      End If
   End If
End Sub

'Added by Morgan 2012/3/29
'預設434(TW-SUPA)期限
Private Sub Set434Date()
   Dim stCP07 As String, stCP06 As String
   
   m_bolTw_SUPA_LawDateChk = False
   If Text1(1) = "434" And Text1(13) = "000" And cp(27) = "" Then
      strExc(0) = "select pa10 from pridate,patent" & _
         " where pd06='" & pa(11) & "' and pd07='000'" & _
         " and pa01(+)=pd01 and pa02(+)=pd02 and pa03(+)=pd03 and pa04(+)=pd04 and pa09='101' and pa08='1' and pa10>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_bolTw_SUPA_LawDateChk = True 'Added by Morgan 2016/10/7
         stCP07 = CompDate(1, 6, RsTemp.Fields(0))
         'modify by sonia 2023/5/19 若已有申請日則TW-SUPA本所期限改為收文日+10個工作天,法定維持CFP申請日+6個月
         ''Added by Morgan 2014/10/28
         'If Text1(13) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
         '   stCP06 = PUB_GetOurDeadline(stCP07)
         'Else
         ''end 2014/10/28
         '   stCP06 = PUB_GetWorkDay1(CompDate(2, -2, stCP07), True)
         'End If 'Added by Morgan 2014/10/28
         'Added by Lydia 2025/10/29
         If strSrvDate(1) >= 內專本所約定期限啟用日 Then
            stCP06 = PUB_GetPOurDeadline(Text1(12), Text1(13))
         Else
         'end 2025/10/29
            stCP06 = CompWorkDay(10, DBDATE(Text1(12)), 0)
         End If 'Added by Lydia 2025/10/29
         
         If stCP06 > stCP07 Then
            stCP06 = stCP07
         End If
         'end 2023/5/19
         If stCP06 < strSrvDate(1) Then
            stCP06 = strSrvDate(1)
         End If
         
         Text1(5) = TransDate(stCP07, 1)
         Text1(4) = TransDate(stCP06, 1)
         
      'Added by Morgan 2016/10/7
      '若美國案未提申時所限設為收文日+1個月,無需帶法限
      ElseIf cp(6) = "" Then
         stCP06 = PUB_GetOurDeadline(CompDate(1, 1, cp(5)))
         If stCP06 < strSrvDate(1) Then
            stCP06 = strSrvDate(1)
         End If
         Text1(4) = TransDate(stCP06, 1)
      'end 2016/10/7
      End If
   End If
End Sub

'Added by Morgan 2012/7/4
'主張優先權期限:pData=優先權資料,pType=主張優先權期限適用類別 1.發明或新型,2.設計
'Modify by Amy 2023/01/05 傳入的pData改變數
'Private Function GetAppDateLimit(pPA08 As String, pData() As String, ByRef pType As String) As String
Private Function GetAppDateLimit(pPA08 As String, pData2 As String, pData4 As String, ByRef pType As String) As String
   Dim stDate As String
   stDate = PUB_GetFirstPriDate2(pData2)
   If pPA08 = "3" Or InStr(pData4, "3") > 0 Then
'end 2023/01/05
      pType = "2"
   Else
      pType = "1"
   End If
   If pType = 1 Then
      stDate = CompDate(0, 1, stDate)
   Else
      stDate = CompDate(1, 6, stDate)
   End If
   GetAppDateLimit = stDate
End Function

Private Sub txtPA161_GotFocus()
   TextInverse txtPA161
   CloseIme
End Sub

Private Sub txtPA161_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Added by Lydia 2020/03/31
   If txtPA161.Enabled = False Then Exit Sub
   
   '事務所合併日起新案只能空白或J，不可輸T
   If Val(m_CP31isYGetCP05) >= 事務所合併日 Then
        If KeyAscii <> 8 And KeyAscii <> Asc("J") Then
            KeyAscii = 0
            Beep
        End If
        Exit Sub
   Else
   'end 2020/03/31
        'Modify By Sindy 2013/12/15
        'If strSrvDate(1) >= InvoiceStartDate Then
        'Modify by Amy 2016/08/12 +台灣判斷-秀玲
        If Val(m_CP31isYGetCP05) >= Val(InvoiceStartDate) Then
           'Modify by Amy 2017/07/13 服務業務 非台灣才可輸,只能輸J或空白
           If pa(1) = "PS" And pa(9) <> "000" And KeyAscii <> 8 And KeyAscii <> Asc("J") Then
                 KeyAscii = 0
                 Beep
           '專利
           ElseIf pa(9) <> "000" And KeyAscii <> 8 And KeyAscii <> Asc("T") And KeyAscii <> Asc("J") Then
              KeyAscii = 0
              Beep
           End If
           'end 2017/07/13
        Else
           'Modified by Morgan 2015/5/12
           'If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
           If KeyAscii <> 8 And KeyAscii <> Asc("T") Then
           'end 2015/5/12
              KeyAscii = 0
              Beep
           End If
        End If
        'end 2015/5/12
   End If 'Added by Lydia 2020/03/31

End Sub
'Added by Morgan 2012/10/4
'設定台灣改請期限:改請新型=核駁日+2個月,改請發明=核駁日+30天,新型改請設計=核駁日+30天,發明改請設計=核駁日+2個月
Private Sub Set30xDate()
   Dim strOldPA08 As String
   If Text1(13) = "000" And (Text1(1) = "301" Or Text1(1) = "302" Or Text1(1) = "303") Then
      If pa(16) = "2" And pa(20) <> "" Then
         '改請設計若專利種類已改則判斷申請程序
         If Text1(1) = "303" Then
            If pa(8) = "3" Then
               strExc(0) = "select cp10 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
                  " and  cp03='" & pa(3) & "' and  cp04='" & pa(4) & "' and cp10='101'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strOldPA08 = "1"
               Else
                  strOldPA08 = "2"
               End If
            Else
                strOldPA08 = pa(8)
            End If
         End If
         '改請新型,發明改請設計=核駁日+2個月
         If Text1(1) = "302" Or (Text1(1) = "303" And strOldPA08 = "1") Then
            strExc(1) = CompDate(1, 2, pa(20))
            'Added by Morgan 2014/10/28
            If Text1(13) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
               strExc(2) = PUB_GetOurDeadline(strExc(1))
            Else
            'end 2014/10/28
               strExc(2) = PUB_GetWorkDay1(CompDate(2, -4, strExc(1)), True)
            End If 'Added by Morgan 2014/10/28
            
            If strExc(2) < strSrvDate(1) Then
               strExc(2) = strSrvDate(1)
            End If
            Text1(5) = TransDate(strExc(1), 1)
            Text1(4) = TransDate(strExc(2), 1)
            
         '核駁日+30天
         Else
            strExc(1) = CompDate(2, 30, pa(20))
            'Added by Morgan 2014/10/28
            If Text1(13) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
               strExc(2) = PUB_GetOurDeadline(strExc(1))
            Else
            'end 2014/10/28
               strExc(2) = PUB_GetWorkDay1(CompDate(2, -2, strExc(1)), True)
            End If 'Added by Morgan 2014/10/28
            
            If strExc(2) < strSrvDate(1) Then
               strExc(2) = strSrvDate(1)
            End If
            Text1(5) = TransDate(strExc(1), 1)
            Text1(4) = TransDate(strExc(2), 1)
         End If
      End If
   End If
End Sub

Private Function CheckReverseCaseMap(ByRef p_OutCaseNo As String, ByRef p_InCaseNo As String, Optional ByVal p_bolMsg As Boolean = True) As Boolean

   Dim strCaseNoOut(1 To 4) As String
   Dim strCaseNoIn(1 To 4) As String
   Dim stSQL As String, iQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   ChgCaseNo p_OutCaseNo, strCaseNoOut
   ChgCaseNo p_InCaseNo, strCaseNoIn
   
On Error GoTo ErrHnd
   
   CheckReverseCaseMap = True
   
   stSQL = "SELECT * FROM CASEMAP" & _
      " WHERE CM10='0' AND CM01='" & strCaseNoIn(1) & "' AND CM02='" & strCaseNoIn(2) & "' AND CM03='" & strCaseNoIn(3) & "' AND CM04='" & strCaseNoIn(4) & "'" & _
      " AND CM05='" & strCaseNoOut(1) & "' AND CM06='" & strCaseNoOut(2) & "' AND CM07='" & strCaseNoOut(3) & "' AND CM08='" & strCaseNoOut(4) & "'"
      
   iQ = 1
   Set rsQuery = ClsLawReadRstMsg(iQ, stSQL)
   If iQ = 0 Then
      CheckReverseCaseMap = False
   ElseIf p_bolMsg = True Then
      MsgBox p_OutCaseNo & " 已建立為 " & p_InCaseNo & " 的國內案，不可再建立反向的關聯！", vbCritical
   End If
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   Set rsQuery = Nothing
End Function

'Add by Lydia 2015/02/02 輸入新穎性優惠期公開事實 (多筆)
Private Sub CmdFav_Click()
If cp(10) = "123" Then
   Set frm880020.m_PrevF = Me
   frm880020.m_dbCheck = False   'Modified by Lydia 2015/02/25  發文DoubleCheck
   frm880020.strFPD01 = pa(1):   frm880020.strFPD02 = pa(2)
   frm880020.strFPD03 = pa(3):   frm880020.strFPD04 = pa(4)
   frm880020.strLimit1 = Text1(4)
   frm880020.strLimit2 = Text1(5)
   frm880020.strNation = pa(9)
   frm880020.strPA08 = Text1(2) 'Added by Morgan 2018/3/16
   frm880020.strPA10 = pa(10) 'Added by Morgan 2022/8/19
   frm880020.strPA140 = IIf(txtFavDt.Text = "", IIf(pa(140) <> "", Val(pa(140)) - 19110000, ""), txtFavDt.Text)
   frm880020.Show
End If
End Sub

'Added by Lydia 2015/05/13
'查同時收文的案件進度
'Modified by Lydia 2015/08/07 +判斷未收文(掛下一程序)
'Private Function CheckCPExists(ByVal strX1 As String, ByVal strX2 As String) As Boolean
Private Function CheckCPExists(ByVal strX1 As String, ByVal strX2 As String, ByVal strX3 As String) As Boolean
  Dim sqlP As String, Id As Integer
  Dim rsP As New ADODB.Recordset
     
    CheckCPExists = False
    'Modified by Lydia 2015/08/07 +未發文
   ' sqlP = "select CP09,CP43 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
           " and cp10='" & strX1 & "' and cp27 is null and cp57 is null"
    'Modified by Lydia 2019/01/16 可能有多筆延期 (ex.P111996)
    'sqlP = "select CP27,CP43 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
           " and cp10='" & strX1 & "' and cp57 is null"
    sqlP = "select CP27,CP43 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
           " and cp10='" & strX1 & "' and cp57 is null and (cp27||cp43 is null " & IIf(strX2 <> "", "or cp43='" & strX2 & "' ", "") & IIf(strX3 <> "", "or cp43='" & strX3 & "' ", "") & ")"
    Id = 1
    Set rsP = ClsLawReadRstMsg(Id, sqlP)
    If Id = 1 Then
       '2015/08/07 +判斷未收文
       'If IsNull(rsP(1)) Or rsP(1) = strX2 Then '相關收文號是空,或是有串
       If (IsNull(rsP(1)) And IsNull(rsP(0))) Or rsP(1) = strX2 Or rsP(1) = strX3 Then
          CheckCPExists = True
       End If
    End If
    Set rsP = Nothing
End Function
'end 2015/05/13

'Added by Lydia 2020/05/20 法律所案源收文：案件性質=>案源案件類型
Private Sub SetLOSagree()
Dim m_LOSkind As String

    If strSrvDate(1) >= 法律所案源收文啟用日 And pa(1) = "P" Then
        'Modified by Lydia 2020/06/29 直接用案源檔的類型
        'm_LOSkind = PUB_GetLOSkind(pa(1), Text1(1), Text1(13))
        m_LOSkind = m_LOS02
        txtLOSagree.Text = ""
        FraLOS.Visible = False
        If Text1(13) = "000" Then
            If Left(m_LOSkind, 1) = "C" And m_LOS01 = "" Then 'C類-未分案通知
                 FraLOS.Visible = True
                 txtLOSagree.Text = "Y"
            End If
        End If
    End If
End Sub

Private Sub txtLOSagree_GotFocus()
   TextInverse txtLOSagree 'Added by Lydia 2020/05/29
End Sub

Private Sub txtLOSagree_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 89 And KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
End Sub

'Added by Lydia 2020/05/20 法律所案源收文：讀取法務案源檔
Private Sub ReadLOS()
Dim stSQL As String, intQ As Integer
Dim RsQ As ADODB.Recordset
   
   m_LOS01 = ""
   m_LOS07 = ""
   m_LOS15 = ""
   Text1(1).Locked = False
   Text1(13).Locked = False
   If strSrvDate(1) >= 法律所案源收文啟用日 And pa(1) = "P" Then
        stSQL = "select X.* from CaseProgress, LawOfficeSource X where CP09='" & Label3(8) & "'  and CP162=LOS15(+) and cp162 is not null "
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
        If intQ = 1 Then
           '案源總收文號
           m_LOS01 = "" & RsQ.Fields("los01")
           'Added by Lydia 2020/06/09 案源案件類型
           m_LOS02 = "" & RsQ.Fields("los02")
           '放棄日期
           m_LOS07 = "" & RsQ.Fields("los07")
           '案源單號
           m_LOS15 = "" & RsQ.Fields("los15")
           '已分案通知: 不可變更案件性質和申請國家
           'If m_LOS01 <> "" Then 'Mark by Lydia 2020/07/14 都不可以變更
               Text1(1).Locked = True
               Text1(13).Locked = True
           'End If
        End If
        Set RsQ = Nothing
   End If
End Sub

'Added by Morgan 2020/11/9
'TW-SUPA僅適用於沒有主張優先權的台灣案
Private Function ChkTwSupa() As Boolean
   Dim stCP10_2 As String
   
   ChkTwSupa = True
   If Text1(1) = "434" Then
      stCP10_2 = "106"
   ElseIf Text1(1) = "106" Then
      stCP10_2 = "434"
   End If
   If stCP10_2 <> "" Then
      If PUB_ChkCPExist(cp(), stCP10_2) = True Then
         ChkTwSupa = False
         If stCP10_2 = "434" Then
            MsgBox "此案有申請TW-SUPA，不應主張國際優先權，請確認!!", vbCritical
         Else
            MsgBox "此案有主張國際優先權，不得申請TW-SUPA!!", vbCritical
         End If
      End If
   End If
End Function

'Add by Amy 2022/10/17
Private Sub SetGrd()
    Dim arrGridHeadText, arrGridHeadWidth
    Dim iRow As Integer
    
    arrGridHeadText = Array("簽核人員", "身份", "日期", "時間", "簽核結果", "B1104")
    arrGridHeadWidth = Array(1050, 600, 800, 800, 800, 0)
    GRD1.Visible = False
    GRD1.Cols = UBound(arrGridHeadText) + 1
    For iRow = 0 To GRD1.Cols - 1
        GRD1.row = 0
        GRD1.col = iRow
        GRD1.Text = arrGridHeadText(iRow)
        GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
        GRD1.CellAlignment = flexAlignCenterCenter
    Next
    GRD1.Visible = True
End Sub

Private Sub txtPA178_GotFocus()
   TextInverse txtPA178
End Sub

Private Sub txtPA178_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Chr(KeyAscii) <> "1" And Chr(KeyAscii) <> "2" Then
      Beep
      KeyAscii = 0
   End If
End Sub
'Added by Morgan 2022/12/28
'台灣112年以後領證繳年費需輸入形式
Private Sub SetPA178()
   lblPA178.Visible = False
   txtPA178.Visible = False
   If Text1(13) = "000" And Val(cp(27)) = 0 Then
      If PUB_TWCertPty(cp(1), Text1(1), cp(2), cp(3), cp(4)) = True Then
         lblPA178.Visible = True
         txtPA178.Visible = True
      End If
   End If
End Sub

