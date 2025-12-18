VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050101_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "分案"
   ClientHeight    =   5772
   ClientLeft      =   156
   ClientTop       =   996
   ClientWidth     =   9180
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   9180
   Begin VB.TextBox txtF0309 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   7455
      Locked          =   -1  'True
      TabIndex        =   147
      Top             =   390
      Width           =   1665
   End
   Begin VB.CommandButton cmdCPP 
      Caption         =   "卷宗區"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   0
      TabIndex        =   54
      Top             =   0
      Width           =   705
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "檢視接洽單"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   706
      TabIndex        =   55
      Top             =   0
      Width           =   1065
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   4
      Left            =   5355
      MaxLength       =   1
      TabIndex        =   37
      Top             =   1170
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   2
      Left            =   5355
      MaxLength       =   1
      TabIndex        =   35
      Top             =   915
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   5
      Left            =   900
      MaxLength       =   1
      TabIndex        =   38
      Top             =   1410
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   3
      Left            =   900
      MaxLength       =   1
      TabIndex        =   36
      Top             =   1170
      Width           =   240
   End
   Begin VB.TextBox txtAD 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   270
      Index           =   1
      Left            =   900
      MaxLength       =   1
      TabIndex        =   34
      Top             =   915
      Width           =   240
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "一案兩請資料(&D)"
      Height          =   345
      Index           =   8
      Left            =   7470
      TabIndex        =   121
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "美國IDS維護"
      Height          =   345
      Index           =   7
      Left            =   1772
      TabIndex        =   56
      Top             =   0
      Width           =   1155
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4080
      Left            =   48
      TabIndex        =   40
      Top             =   1680
      Width           =   9048
      _ExtentX        =   15939
      _ExtentY        =   7197
      _Version        =   393216
      TabsPerRow      =   8
      TabHeight       =   420
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm050101_2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label29"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label17(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label28"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label13(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label15(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblPromoter"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblCaseProperty"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblNation"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblPoints"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label3(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label16"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label27(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblFee"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label17(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label24"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label3(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label40"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label3(11)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(12)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label35"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label18(0)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label20(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblCancelReceiveDate"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblDivCase"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label6(3)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label3(2)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label1(22)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblFeeYear"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label17(1)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label1(168)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtCaseField(10)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtCaseField(9)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtCaseField(7)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtCaseField(0)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtCaseField(15)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtCaseField(1)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtCaseField(2)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtCaseField(3)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtCaseField(14)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtCaseField(5)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtCaseField(6)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtCaseField(8)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtCaseField(16)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtCaseField(17)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtCaseField(4)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "optChoose(0)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtCode(0)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "txtCode(3)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "txtCode(2)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "txtCode(1)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "cmdCountry"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "grdDataList"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "Text1(21)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtCode(4)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "txtCode(5)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "txtCode(6)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "txtCode(7)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "txtDivCaseNo(3)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "txtDivCaseNo(2)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "txtDivCaseNo(1)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "txtDivCaseNo(4)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "optChoose(1)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "Text1(23)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "txtFeeYear(2)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "txtFeeYear(1)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "txtFavDt"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "Combo3"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "optChoose(2)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "Check11"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "CmdFav"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).ControlCount=   74
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm050101_2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPA161"
      Tab(1).Control(1)=   "txtPA61"
      Tab(1).Control(2)=   "txtCP97"
      Tab(1).Control(3)=   "txtCP147"
      Tab(1).Control(4)=   "txtEngGroup"
      Tab(1).Control(5)=   "txtCP98"
      Tab(1).Control(6)=   "cmdPriority"
      Tab(1).Control(7)=   "Frame2"
      Tab(1).Control(8)=   "Frame1"
      Tab(1).Control(9)=   "Label1(121)"
      Tab(1).Control(10)=   "LblPA61"
      Tab(1).Control(11)=   "txtCP99"
      Tab(1).Control(12)=   "txtCaseField(11)"
      Tab(1).Control(13)=   "txtCaseField(12)"
      Tab(1).Control(14)=   "txtCaseField(13)"
      Tab(1).Control(15)=   "LblCP97"
      Tab(1).Control(16)=   "lblPA161"
      Tab(1).Control(17)=   "Label1(172)"
      Tab(1).Control(18)=   "Label5(1)"
      Tab(1).Control(19)=   "Label1(24)"
      Tab(1).Control(20)=   "Label3(13)"
      Tab(1).Control(21)=   "Label1(17)"
      Tab(1).Control(22)=   "Label1(19)"
      Tab(1).Control(23)=   "Label37"
      Tab(1).Control(24)=   "Label26(0)"
      Tab(1).Control(25)=   "Label27(2)"
      Tab(1).Control(26)=   "Label27(1)"
      Tab(1).ControlCount=   27
      TabCaption(2)   =   "簽核"
      TabPicture(2)   =   "frm050101_2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdAddInfo"
      Tab(2).Control(1)=   "txtF0301"
      Tab(2).Control(2)=   "GRD1"
      Tab(2).Control(3)=   "Label68"
      Tab(2).Control(4)=   "txtF0407"
      Tab(2).Control(5)=   "Label66"
      Tab(2).Control(6)=   "txtNote"
      Tab(2).Control(7)=   "Label67"
      Tab(2).ControlCount=   8
      Begin VB.CommandButton CmdFav 
         Caption         =   "優惠期實際發生日期"
         Height          =   270
         Left            =   1820
         TabIndex        =   139
         Top             =   1579
         Width           =   1815
      End
      Begin VB.TextBox txtPA161 
         Height          =   300
         Left            =   -73572
         MaxLength       =   1
         TabIndex        =   50
         Top             =   3420
         Width           =   255
      End
      Begin VB.TextBox txtPA61 
         Height          =   300
         Left            =   -68472
         MaxLength       =   1
         TabIndex        =   51
         Top             =   3420
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check11 
         Caption         =   "急件"
         ForeColor       =   &H00000000&
         Height          =   200
         Left            =   3660
         TabIndex        =   146
         Top             =   300
         Width           =   765
      End
      Begin VB.CommandButton CmdAddInfo 
         Caption         =   "補件完成"
         CausesValidation=   0   'False
         Height          =   400
         Left            =   -67380
         TabIndex        =   53
         Top             =   360
         Width           =   1200
      End
      Begin VB.TextBox txtF0301 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -73830
         Locked          =   -1  'True
         TabIndex        =   140
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox txtCP97 
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   -69390
         MaxLength       =   4
         TabIndex        =   49
         Top             =   3120
         Width           =   732
      End
      Begin VB.OptionButton optChoose 
         Caption         =   "微個體"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   7305
         TabIndex        =   137
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtCP147 
         Height          =   300
         Left            =   -72990
         MaxLength       =   1
         TabIndex        =   48
         Top             =   3120
         Width           =   255
      End
      Begin VB.TextBox txtEngGroup 
         Height          =   300
         Left            =   -67755
         MaxLength       =   1
         TabIndex        =   42
         Top             =   330
         Width           =   285
      End
      Begin VB.ComboBox Combo3 
         Height          =   276
         ItemData        =   "frm050101_2.frx":0054
         Left            =   3180
         List            =   "frm050101_2.frx":0061
         TabIndex        =   29
         Top             =   2430
         Width           =   1395
      End
      Begin VB.TextBox txtFavDt 
         Height          =   300
         Left            =   3630
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   18
         Top             =   1584
         Width           =   885
      End
      Begin VB.TextBox txtFeeYear 
         Height          =   300
         Index           =   1
         Left            =   7965
         TabIndex        =   21
         Top             =   1590
         Width           =   285
      End
      Begin VB.TextBox txtFeeYear 
         Height          =   300
         Index           =   2
         Left            =   8460
         TabIndex        =   22
         Top             =   1590
         Width           =   285
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   23
         Left            =   1320
         TabIndex        =   23
         Top             =   1830
         Width           =   1710
      End
      Begin VB.TextBox txtCP98 
         Height          =   300
         Left            =   -73590
         MaxLength       =   12
         TabIndex        =   46
         Top             =   2076
         Width           =   675
      End
      Begin VB.OptionButton optChoose 
         Caption         =   "小個體"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   6000
         TabIndex        =   123
         Top             =   2760
         Width           =   1260
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   300
         Index           =   4
         Left            =   7635
         MaxLength       =   2
         TabIndex        =   33
         Top             =   2427
         Width           =   390
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   300
         Index           =   1
         Left            =   6180
         MaxLength       =   3
         TabIndex        =   30
         Top             =   2427
         Width           =   390
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   300
         Index           =   2
         Left            =   6570
         MaxLength       =   6
         TabIndex        =   31
         Top             =   2427
         Width           =   705
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   300
         Index           =   3
         Left            =   7275
         MaxLength       =   1
         TabIndex        =   32
         Top             =   2427
         Width           =   390
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   7
         Left            =   7215
         MaxLength       =   2
         TabIndex        =   12
         Top             =   786
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   6
         Left            =   6975
         MaxLength       =   1
         TabIndex        =   11
         Top             =   786
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   5
         Left            =   6135
         MaxLength       =   6
         TabIndex        =   10
         Top             =   786
         Width           =   852
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   4
         Left            =   5655
         MaxLength       =   3
         TabIndex        =   9
         Top             =   786
         Width           =   492
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   300
         Index           =   21
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   28
         Top             =   2427
         Width           =   375
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
         Height          =   1008
         Left            =   1080
         TabIndex        =   39
         Top             =   2976
         Width           =   7836
         _ExtentX        =   13822
         _ExtentY        =   1778
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         HighLight       =   0
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
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton cmdPriority 
         Caption         =   "輸入(&I)"
         Height          =   300
         Left            =   -73590
         TabIndex        =   41
         Top             =   300
         Width           =   972
      End
      Begin VB.CommandButton cmdCountry 
         Caption         =   "指定國家"
         Height          =   300
         Left            =   2580
         TabIndex        =   8
         Top             =   720
         Width           =   885
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   1
         Left            =   6135
         MaxLength       =   6
         TabIndex        =   4
         Top             =   534
         Width           =   852
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   2
         Left            =   6975
         MaxLength       =   1
         TabIndex        =   5
         Top             =   534
         Width           =   252
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   3
         Left            =   7215
         MaxLength       =   2
         TabIndex        =   6
         Top             =   534
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   0
         Left            =   5655
         MaxLength       =   3
         TabIndex        =   3
         Top             =   534
         Width           =   492
      End
      Begin VB.OptionButton optChoose 
         Caption         =   "大個體"
         Enabled         =   0   'False
         Height          =   252
         Index           =   0
         Left            =   5010
         TabIndex        =   122
         Top             =   2724
         Width           =   1035
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm050101_2.frx":0085
         Height          =   1795
         Left            =   -70410
         TabIndex        =   141
         Top             =   2130
         Width           =   4215
         _ExtentX        =   7451
         _ExtentY        =   3175
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
      Begin VB.Frame Frame2 
         Height          =   396
         Left            =   -70032
         TabIndex        =   156
         Top             =   3648
         Visible         =   0   'False
         Width           =   2220
         Begin VB.OptionButton Option1 
            Appearance      =   0  '平面
            BackColor       =   &H80000005&
            Caption         =   "之後"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   1476
            TabIndex        =   159
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
            TabIndex        =   158
            Top             =   144
            Width           =   660
         End
         Begin VB.OptionButton Option1 
            Appearance      =   0  '平面
            BackColor       =   &H80000005&
            Caption         =   "當天"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   24
            TabIndex        =   157
            Top             =   144
            Width           =   660
         End
      End
      Begin VB.Frame Frame1 
         Height          =   396
         Left            =   -73560
         TabIndex        =   150
         Top             =   3648
         Visible         =   0   'False
         Width           =   3636
         Begin VB.OptionButton OptSendType 
            Caption         =   "不限制"
            Height          =   180
            Index           =   1
            Left            =   24
            TabIndex        =   154
            Top             =   132
            Width           =   850
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "收款後"
            Height          =   180
            Index           =   2
            Left            =   864
            TabIndex        =   153
            Top             =   132
            Width           =   850
         End
         Begin VB.OptionButton OptSendType 
            Caption         =   "指定日期"
            Height          =   180
            Index           =   3
            Left            =   1710
            TabIndex        =   152
            Top             =   132
            Width           =   1035
         End
         Begin VB.TextBox txtCP142 
            Height          =   270
            Left            =   2745
            MaxLength       =   7
            TabIndex        =   151
            Top             =   90
            Width           =   825
         End
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   4
         Left            =   1320
         TabIndex        =   16
         Top             =   1560
         Width           =   885
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "送件方式"
         Height          =   252
         Index           =   121
         Left            =   -74880
         TabIndex        =   155
         Top             =   3768
         Visible         =   0   'False
         Width           =   912
      End
      Begin VB.Label LblPA61 
         AutoSize        =   -1  'True
         Caption         =   "CFP有無關聯P案：         (  N：無)"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   -70008
         TabIndex        =   149
         Top             =   3468
         Visible         =   0   'False
         Width           =   2568
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "內容："
         Height          =   300
         Left            =   -74910
         TabIndex        =   145
         Top             =   2130
         Width           =   540
      End
      Begin MSForms.TextBox txtF0407 
         Height          =   1795
         Left            =   -74370
         TabIndex        =   144
         Top             =   2130
         Width           =   3885
         VariousPropertyBits=   -1466939361
         ScrollBars      =   3
         Size            =   "6853;3166"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "呈報內容："
         Height          =   180
         Left            =   -74910
         TabIndex        =   143
         Top             =   600
         Width           =   900
      End
      Begin MSForms.TextBox txtNote 
         Height          =   1200
         Left            =   -74910
         TabIndex        =   52
         Top             =   840
         Width           =   8745
         VariousPropertyBits=   -1466939365
         ScrollBars      =   3
         Size            =   "15425;2117"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "接洽單編號："
         Height          =   180
         Left            =   -74910
         TabIndex        =   142
         Top             =   330
         Width           =   1080
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   17
         Left            =   3630
         TabIndex        =   19
         Top             =   1035
         Width           =   885
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   16
         Left            =   8010
         TabIndex        =   25
         Top             =   1860
         Width           =   975
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCP99 
         Height          =   732
         Left            =   -73584
         TabIndex        =   47
         Top             =   2376
         Width           =   7272
         VariousPropertyBits=   -1467987941
         MaxLength       =   100
         ScrollBars      =   2
         Size            =   "12832;1296"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   8
         Left            =   5655
         TabIndex        =   14
         Top             =   1035
         Width           =   495
         VariousPropertyBits=   671107097
         MaxLength       =   1
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   6
         Left            =   5865
         TabIndex        =   17
         Top             =   1305
         Width           =   495
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   5
         Left            =   1320
         TabIndex        =   26
         Top             =   2130
         Width           =   1335
         VariousPropertyBits=   671107099
         MaxLength       =   9
         Size            =   "2355;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   14
         Left            =   1320
         TabIndex        =   15
         Top             =   1305
         Width           =   885
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1561;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   3
         Left            =   1320
         TabIndex        =   13
         Top             =   1035
         Width           =   495
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   2
         Left            =   1320
         TabIndex        =   7
         Top             =   780
         Width           =   495
         VariousPropertyBits=   671107099
         MaxLength       =   3
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   2
         Top             =   540
         Width           =   495
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "873;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   15
         Left            =   5655
         TabIndex        =   24
         Top             =   1860
         Width           =   975
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   11
         Left            =   -73590
         TabIndex        =   43
         Top             =   630
         Width           =   7260
         VariousPropertyBits=   671107099
         MaxLength       =   50
         Size            =   "12806;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   564
         Index           =   12
         Left            =   -73584
         TabIndex        =   44
         Top             =   948
         Width           =   7272
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "12827;995"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   564
         Index           =   13
         Left            =   -73584
         TabIndex        =   45
         Top             =   1500
         Width           =   7272
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "12827;995"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Top             =   255
         Width           =   975
         VariousPropertyBits=   671107099
         MaxLength       =   6
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   7
         Left            =   5655
         TabIndex        =   1
         Top             =   285
         Width           =   735
         VariousPropertyBits=   671107099
         MaxLength       =   4
         Size            =   "1296;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   9
         Left            =   5655
         TabIndex        =   20
         Top             =   1590
         Width           =   975
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseField 
         Height          =   300
         Index           =   10
         Left            =   5895
         TabIndex        =   27
         Top             =   2145
         Width           =   2880
         VariousPropertyBits=   671107099
         Size            =   "5080;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label LblCP97 
         AutoSize        =   -1  'True
         Caption         =   "承辦人基數"
         Height          =   180
         Left            =   -70452
         TabIndex        =   138
         Top             =   3168
         Width           =   900
      End
      Begin VB.Label lblPA161 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司            (T:專利商標 J:智權公司 空白:系統預設)"
         Height          =   180
         Left            =   -74880
         TabIndex        =   136
         Top             =   3468
         Width           =   4704
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "與他案合併計算結餘，請於案件備註欄註明""與某案號合併計算結餘""！"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   172
         Left            =   -72492
         TabIndex        =   135
         Top             =   2076
         Width           =   5556
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "是否為複雜或特殊案件         (Y:是)"
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   134
         Top             =   3168
         Width           =   2676
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工程師組別"
         Height          =   180
         Index           =   24
         Left            =   -68745
         TabIndex        =   133
         Top             =   375
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "lblEngGroup"
         Height          =   180
         Index           =   13
         Left            =   -67350
         TabIndex        =   132
         Top             =   375
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件屬性"
         Height          =   180
         Index           =   168
         Left            =   2430
         TabIndex        =   131
         Top             =   2475
         Width           =   720
      End
      Begin VB.Label Label17 
         Caption         =   "承辦期限："
         Height          =   180
         Index           =   1
         Left            =   2760
         TabIndex        =   130
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblFeeYear 
         AutoSize        =   -1  'True
         Caption         =   "繳費年度：第         -         年"
         Height          =   180
         Left            =   6840
         TabIndex        =   129
         Top             =   1590
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PCT申請號："
         Height          =   180
         Index           =   22
         Left            =   90
         TabIndex        =   128
         Top             =   1905
         Width           =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "PCT優先權日："
         Height          =   180
         Index           =   2
         Left            =   6795
         TabIndex        =   127
         Top             =   1905
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人加乘註記"
         Height          =   240
         Index           =   17
         Left            =   -74904
         TabIndex        =   126
         Top             =   2076
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "註記修改理由"
         Height          =   180
         Index           =   19
         Left            =   -74880
         TabIndex        =   125
         Top             =   2340
         Width           =   1080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "（美、加、法國案）"
         Height          =   180
         Index           =   3
         Left            =   3435
         TabIndex        =   124
         Top             =   2760
         Width           =   1620
      End
      Begin VB.Label lblDivCase 
         Caption         =   "分割母案本所案號："
         Height          =   180
         Left            =   4575
         TabIndex        =   120
         Top             =   2475
         Width           =   1620
      End
      Begin VB.Label lblCancelReceiveDate 
         BackColor       =   &H80000009&
         Height          =   255
         Left            =   3630
         TabIndex        =   111
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label Label20 
         Caption         =   "卷宗性質：                （1.申請  2.異議  3.舉發）"
         Height          =   180
         Index           =   0
         Left            =   4575
         TabIndex        =   88
         Top             =   1080
         Width           =   3645
      End
      Begin VB.Label Label18 
         Caption         =   "是否算案件數：         （N：不算）"
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   92
         Top             =   570
         Width           =   2820
      End
      Begin VB.Label Label35 
         Caption         =   "轉本所案號："
         Height          =   180
         Left            =   4575
         TabIndex        =   81
         Top             =   570
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利種類："
         Height          =   180
         Index           =   12
         Left            =   90
         TabIndex        =   119
         Top             =   2475
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   180
         Index           =   11
         Left            =   1755
         TabIndex        =   118
         Top             =   2475
         Width           =   465
      End
      Begin VB.Label Label40 
         Caption         =   "本案期限："
         Height          =   255
         Left            =   90
         TabIndex        =   98
         Top             =   2985
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "是否取消閉卷：             （Y：取消閉卷）"
         Height          =   180
         Index           =   0
         Left            =   4575
         TabIndex        =   99
         Top             =   1350
         Width           =   3225
      End
      Begin VB.Label Label24 
         Caption         =   "相關總收文號："
         Height          =   255
         Left            =   90
         TabIndex        =   85
         Top             =   2145
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "本所期限："
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   87
         Top             =   1635
         Width           =   900
      End
      Begin VB.Label lblFee 
         Height          =   255
         Left            =   780
         TabIndex        =   112
         Top             =   2730
         Width           =   960
      End
      Begin VB.Label Label27 
         Caption         =   "費用："
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   84
         Top             =   2730
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "是否為多國案：         （Y：是）"
         Height          =   180
         Left            =   90
         TabIndex        =   89
         Top             =   1080
         Width           =   2640
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PCT申請日："
         Height          =   180
         Index           =   1
         Left            =   4575
         TabIndex        =   116
         Top             =   1905
         Width           =   1035
      End
      Begin VB.Label Label37 
         Caption         =   "優先權資料："
         Height          =   252
         Left            =   -74880
         TabIndex        =   115
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label26 
         Caption         =   "分所案號："
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   114
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblPoints 
         Height          =   255
         Left            =   2580
         TabIndex        =   113
         Top             =   2730
         Width           =   825
      End
      Begin MSForms.Label lblNation 
         Height          =   255
         Left            =   1860
         TabIndex        =   103
         Top             =   795
         Width           =   1755
         VariousPropertyBits=   27
         Size            =   "3096;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCaseProperty 
         Height          =   255
         Left            =   6465
         TabIndex        =   102
         Top             =   285
         Width           =   2400
         VariousPropertyBits=   27
         Size            =   "4233;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label27 
         Caption         =   "進度備註："
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   97
         Top             =   945
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "案件備註："
         Height          =   252
         Index           =   1
         Left            =   -74880
         TabIndex        =   96
         Top             =   1500
         Width           =   972
      End
      Begin VB.Label lblPromoter 
         Height          =   255
         Left            =   2400
         TabIndex        =   95
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "承辦人："
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   94
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label13 
         Caption         =   "案件性質："
         Height          =   180
         Index           =   0
         Left            =   4575
         TabIndex        =   93
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label14 
         Caption         =   "申請國家："
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   91
         Top             =   825
         Width           =   900
      End
      Begin VB.Label Label28 
         Caption         =   "與國內                                                        案號相同"
         Height          =   180
         Left            =   4575
         TabIndex        =   90
         Top             =   825
         Width           =   3780
      End
      Begin VB.Label Label17 
         Caption         =   "法定期限："
         Height          =   180
         Index           =   2
         Left            =   4575
         TabIndex        =   86
         Top             =   1635
         Width           =   900
      End
      Begin VB.Label Label29 
         Caption         =   "點數："
         Height          =   255
         Left            =   1845
         TabIndex        =   83
         Top             =   2730
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "客戶案件案號："
         Height          =   180
         Index           =   1
         Left            =   4575
         TabIndex        =   82
         Top             =   2190
         Width           =   1260
      End
      Begin VB.Label Label2 
         Caption         =   "收文日："
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   80
         Top             =   1350
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "取消收文日："
         Height          =   180
         Index           =   2
         Left            =   2565
         TabIndex        =   79
         Top             =   1350
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   1
      Left            =   7648
      TabIndex        =   62
      Top             =   0
      Width           =   870
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   6997
      TabIndex        =   61
      Top             =   0
      Width           =   650
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   2
      Left            =   8520
      TabIndex        =   63
      Top             =   0
      Width           =   650
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   3
      Left            =   6291
      TabIndex        =   60
      Top             =   0
      Width           =   705
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "國內外案件維護"
      Height          =   345
      Index           =   4
      Left            =   2928
      TabIndex        =   57
      Top             =   0
      Width           =   1425
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "多國案卷號"
      Height          =   345
      Index           =   5
      Left            =   4354
      TabIndex        =   58
      Top             =   0
      Width           =   1050
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "案件進度"
      Height          =   345
      Index           =   6
      Left            =   5405
      TabIndex        =   59
      Top             =   0
      Width           =   885
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "目前表單狀態："
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   6210
      TabIndex        =   148
      Top             =   390
      Width           =   1230
   End
   Begin MSForms.ComboBox cboPatentName 
      CausesValidation=   0   'False
      Height          =   300
      Left            =   1110
      TabIndex        =   64
      Top             =   600
      Width           =   7935
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "13996;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblClose 
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
      Height          =   255
      Left            =   4830
      TabIndex        =   117
      Top             =   390
      Width           =   1245
   End
   Begin VB.Label lblCaseCode 
      Height          =   255
      Left            =   3270
      TabIndex        =   110
      Top             =   390
      Width           =   1500
   End
   Begin MSForms.Label lblSalesName 
      Height          =   225
      Left            =   5535
      TabIndex        =   109
      Top             =   1440
      Width           =   1530
      VariousPropertyBits=   27
      Size            =   "2699;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   225
      Index           =   3
      Left            =   6615
      TabIndex        =   108
      Top             =   1185
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   225
      Index           =   1
      Left            =   6615
      TabIndex        =   107
      Top             =   930
      Width           =   2415
      VariousPropertyBits=   27
      Size            =   "4260;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   225
      Index           =   4
      Left            =   2145
      TabIndex        =   106
      Top             =   1425
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   225
      Index           =   2
      Left            =   2145
      TabIndex        =   105
      Top             =   1185
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblPetitionName 
      Height          =   225
      Index           =   0
      Left            =   2145
      TabIndex        =   104
      Top             =   930
      Width           =   2295
      VariousPropertyBits=   27
      Size            =   "4048;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblPetition 
      Height          =   225
      Index           =   1
      Left            =   5655
      TabIndex        =   101
      Top             =   930
      Width           =   855
   End
   Begin VB.Label lblSales 
      Height          =   225
      Left            =   7650
      TabIndex        =   100
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "申請人："
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   78
      Top             =   1410
      Width           =   735
   End
   Begin VB.Label lblPetition 
      Height          =   225
      Index           =   4
      Left            =   1185
      TabIndex        =   77
      Top             =   1425
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   76
      Top             =   390
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   2310
      TabIndex        =   75
      Top             =   390
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "智權人員："
      Height          =   225
      Index           =   0
      Left            =   4545
      TabIndex        =   74
      Top             =   1425
      Width           =   945
   End
   Begin VB.Label Label6 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   73
      Top             =   645
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "申請人："
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   72
      Top             =   930
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "申請人："
      Height          =   225
      Index           =   0
      Left            =   4545
      TabIndex        =   71
      Top             =   930
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "申請人："
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   70
      Top             =   1185
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "申請人："
      Height          =   225
      Index           =   0
      Left            =   4545
      TabIndex        =   69
      Top             =   1185
      Width           =   735
   End
   Begin VB.Label lblPetition 
      Height          =   225
      Index           =   0
      Left            =   1185
      TabIndex        =   68
      Top             =   930
      Width           =   855
   End
   Begin VB.Label lblPetition 
      Height          =   225
      Index           =   2
      Left            =   1185
      TabIndex        =   67
      Top             =   1185
      Width           =   855
   End
   Begin VB.Label lblPetition 
      Height          =   225
      Index           =   3
      Left            =   5655
      TabIndex        =   66
      Top             =   1185
      Width           =   855
   End
   Begin VB.Label lblReceiveCode 
      Height          =   255
      Left            =   990
      TabIndex        =   65
      Top             =   390
      Width           =   1300
   End
End
Attribute VB_Name = "frm050101_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Morgan 2021/12/4 改成Form2.0 (grdDataList,txtCaseField,lblPetitionName...)
'Modified by Morgan 2021/12/8 符號"ˇ"改為"V"(因Ext-B字會變小)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'bolLeave判斷離開時，是否要彈出詢問視窗
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim bolLeave As Boolean, intLeaveKind As Integer
'strReceiveCode上一畫面frm050101_1勾選的收文號
'intTotalReceive上一畫面frm050101_1勾選的收文號總數
'intNowReceive現在Query的收文號Index
Dim strReceiveCode() As String, intTotalReceive As Integer, intNowReceive As Integer
'控制Form_Activate()不要再執行
Dim bolIsRun As Boolean
'此本所案號之系統類別，在ReadAllData中傳回真正的系統類別
Dim intCaseKind As Integer
'cp()存放CaseProgress,field()存放基本資料檔
Dim cp() As String, field() As String
'strCountry存放指定國家
Dim strCountry As String
Dim strCountryOld As String 'Add by Morgan 2007/12/24
'strPriority存放優先權
Public strPriority1 As String, strPriority2 As String, strPriority3 As String, strPriority4 As String, strPriority5 As String 'Modify by Amy 2023/01/06 原Dim
'strCode存放國內外關聯 0-3此案之本所案件 4-7原國內本所案號 8-11修改後國內本所案號
'12-15轉本所案號
'edit by nickc 2007/02/02
'Dim strCode(0 To T_PA) As String
Dim strCode() As String
'是否轉新本所案號
Dim bolNew As Boolean
'Add By Cheng 2002/06/11
Dim m_strCP27 As String '發文日
Dim strFirstPriDate As String  '最早的優先權日期
'控制是否讀過
Dim Nick920224Bol As Boolean
Dim m_strCP06Update As String '更新後的本所期限
'Add by Morgan 2004/2/18
'若承辦人是王協理且未發文則要發EMail通知
Dim stCP09 As String, stCP14 As String, stCP27 As String
'Add by Morgan 2004/6/7
'一案兩申請
Dim m_PA2 As String
Dim m_stCP98 As String 'Add by Morgan 2005/3/4
Dim m_EP06 As String 'Add by Morgan 2005/3/29
'Add by Morgan 2005/6/10
Dim m_bolIsMutiNation As Boolean '是否為多國案
Dim m_bolIsInsCR As Boolean '是否有新增多案相關
Dim m_bolCheckCM As Boolean '是否檢查國內外案資料
Dim m_bolCheckCP21 As Boolean '是否檢查多國
'Add by Morgan 2006/1/19
Dim old_Entity As String   '原大小個體
Dim new_Entity As String   '新大小個體
'Add by Morgan 2006/5/23
Dim iPos1 As Integer, iPos2 As Integer, strPCTPriDate As String
Dim m_bolUpdCP27 As Boolean 'Add by Morgan 2006/12/29是否上發文日
Dim m_strCP44 As String 'Add by Morgan 2010/2/3 自動上發文日的發文代理人
'Add by Morgan 2007/8/24
Dim m_bolAnnuityAlert As Boolean, m_strAlertMsg As String '是否提醒年費未收文
Dim m_CP14ST06 As String '2010/3/3 add by sonia 承辦人所別
'Add by Morgan 2010/6/2
Dim m_bol106Chk As Boolean '優先權是否需直譯本
Dim m_bolMail927Inform As Boolean '是否通知智權人員收文其他翻譯
Dim m_bolActive As Boolean
Dim m_CP13 As String, m_CP31 As String 'Add By Sindy 2010/10/29
Dim m_strOldCP10 As String, m_strOldPA09 As String 'Add by Morgan 2010/10/29
Dim m_CP30 As String 'Add by Morgan 2011/4/22
Dim m_CP31isYGetCP05 As String 'Add By Sindy 2014/1/29
Dim m_field46 As String 'Add by Lydia 2014/11/24
'Add by Amy 2015/01/22
Dim m_bolIsFirstKeyCP14 As Boolean '北所第一次輸承辦人
Dim m_strOldCP14 As String '原承辦人
Dim bolCP14Mail As Boolean '承辦人員修改後是否發mail
Dim m_bolAskEngInCase As Boolean 'Added by Morgan 2017/9/12 Email承辦工程師確認是否有國內案
'Added by Morgan 2019/11/27
Dim m_bolXCACase As Boolean '母案非本所之接續案
Dim m_bolPCTbyPass As Boolean 'PCT接續案
'Add by Amy 2022/10/21
Dim stCPM35 As String, stF0207_A6 As String, m_F0308 As String, m_F0309 As String, strUpdDate As String, strUpdTime As String, IsEConsultRec As Boolean
Dim stF0307_Now As String, stF0309_Now As String '登入時F030X值
Dim stCP141 As String, stCP142 As String, stCP164 As String 'Added by Morgan 2024/2/6
Dim m_IDSCP44 As String, m_IDSCP45 As String 'Added by Morgan 2024/3/8
Dim m_B209Msg As String 'Added by Morgan 2025/6/17
Dim strMsgCloseCancel As String 'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之年費605、維持費606、延展費607，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。

Private Sub ReadAllData()
 Dim i As Integer, varSaveCursor
 Dim nRow As Integer
 Dim strSql As String
 Dim rsTmp As New ADODB.Recordset
 Dim arrTmp 'Add by Amy 2022/10/21
 
On Error GoTo ErrHnd
varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass

'Add By Cheng 2002/06/11
Me.Text1(21).Enabled = False
Me.Text1(21).Text = ""
Me.Label3(11).Caption = ""
m_strCP27 = ""
'Added by Morgan 2017/12/28 修正EPC子案重複問題(連續分案發明申請及指定費，變數未清除造成)
Nick920224Bol = False
strCountryOld = strCountry
'end 2017/12/28

For i = 0 To 7
   txtCode(i) = Empty
Next

'Modify by Morgan 2005/3/4 不再Call dll
'If objPublicData.ReadAllData(strReceiveCode(intNowReceive), cp(), field(), intCaseKind, 0) Then
ReDim cp(TF_CP) As String
cp(9) = strReceiveCode(intNowReceive)
If PUB_ReadAllData(cp(), field(), intCaseKind, 0) Then
'2005/3/4 end
   lblReceiveCode = cp(9)
   'Modified by Lydia 2019/10/17 去掉空白; 為了統一PointReAssignInform的email主旨顯示
   'lblCaseCode = cp(1) + " - " + cp(2) + IIf(cp(4) = "00" And cp(3) = "0", "", " - " + cp(3)) + IIf(cp(4) = "00", "", " - " + cp(4))
   lblCaseCode = cp(1) + "-" + cp(2) + IIf(cp(4) = "00" And cp(3) = "0", "", "-" + cp(3)) + IIf(cp(4) = "00", "", "-" + cp(4))
   lblSales = cp(13)
   
  'Modify by Amy 2015/01/22 操作人員為北所且北所分案日期有值才顯示辦人
   'Add By Cheng 2003/10/29 記錄原承辦人
   Me.txtCaseField(0).Tag = cp(14)
   m_strOldCP14 = cp(14) '因再確認輸入後會更新 .Tag資料,故+m_strOldCP14變數
   'End 2003/10/29
   If pub_strUserOffice = "1" And cp(157) = "" Then
        m_bolIsFirstKeyCP14 = True
        'Modify by Amy 2022/11/03 接洽單電子收文上線後直接顯示(cp14=cra09),不需再輸
        If strSrvDate(1) >= 接洽單電子收文啟用日 Then
            Me.txtCaseField(0).Tag = "" 'Add By Sindy 2022/12/27 因為電子收文上線畫面上承辦人會顯示出來,反過來Tag=空白,後續程式在比對承辦人才會是不一致的
        Else
            cp(14) = ""
        End If
   End If
   txtCaseField(0) = cp(14)
   'end 2015/01/22
   
   'Added by Morgan 2025/5/23
   'P/CFP 設定為程序承辦且不需專業部主管分案的性質，承辦人都自動預設為程序人員，若有需要，分案人員再自行修改，但CFP實體審查除外--郭
   If cp(14) = "" And cp(10) <> "416" And cp(157) = "" Then
      'Modified by Lydia 2025/06/19 改模組名稱
      'If PUB_GetCPM35byCP10(cp(1), cp(10)) = "2" Then
      If PUB_GetCPMbyCP10(cp(1), cp(10), "cpm35") = "2" Then
         txtCaseField(0) = PUB_GetCFPHandler(cp(1) & cp(2) & cp(3) & cp(4))
      End If
   End If
   'end 2025/5/23
      
   txtCaseField(1) = cp(26)
   txtCaseField(3) = cp(21)
   txtCaseField(4) = cp(6)
   
    'Add By Cheng 2003/12/08
    '本所期限若非工作天則抓最近工作天
    'Me.txtCaseField(4).Text = TransDate(PUB_GetWorkDay1(Me.txtCaseField(4).Text, True), 1) 'Modify by Amy 2013/09/03 往下搬
    
   txtCaseField(5) = cp(43)
   txtCaseField(5).Tag = txtCaseField(5) 'Add by Morgan 2010/6/30
   txtCaseField(7) = cp(10)
   txtCaseField(7).Tag = txtCaseField(7) 'Add by Amy 2018/10/18
   txtCaseField_LostFocus 7 'Added by Morgan 2013/3/14 觸發是否計件的預設(Ex.延期不可改原程式就不會觸發)
   
   m_strOldCP10 = cp(10) 'Add by Morgan 2010/10/27
   txtCaseField(9) = cp(7)
   
   'Add by Amy 2013/09/03 +if 印度商業使用聲明分案時，系統自動帶未過期期限
   If txtCaseField(4) = "" And txtCaseField(9) = "" And field(8) = "1" And field(9) = "040" And txtCaseField(7) = "930" Then
        'Modified by Morgan 2020/12/4 2020/10/20印度新法:法限改為9/30,所限=法限-1個月--禧佩
        'If Val(Right(strSrvDate(1), 4)) <= 331 Then
        '    '系統日 <=3/31 本限=當年1/31(系統日 > 1/31 本限設系統日)，法限=當年3/31
        '    '系統日 <3/31   本限=隔年1/31，法限=隔年3/31
        '    If Val(Right(strSrvDate(1), 4)) > 131 Then
        '        txtCaseField(4) = TransDate(PUB_GetWorkDay1(strSrvDate(2), True), 1)
        '    Else
        '         txtCaseField(4) = TransDate(PUB_GetWorkDay1(Left(strSrvDate(1), 4) - 1911 & "0131", True), 1)
        '    End If
        '    txtCaseField(9) = Left(strSrvDate(1), 4) - 1911 & "0331"
        'Else
        '    txtCaseField(4) = TransDate(PUB_GetWorkDay1(Left(strSrvDate(1), 4) - 1910 & "0131", True), 1)
        '    txtCaseField(9) = Left(strSrvDate(1), 4) - 1910 & "0331"
        'End If
        If Val(Right(strSrvDate(1), 4)) <= 930 Then
            txtCaseField(9) = Left(strSrvDate(1), 4) - 1911 & "0930"
        Else
            txtCaseField(9) = Left(strSrvDate(1), 4) - 1910 & "0930"
        End If
        strExc(1) = CompDate(1, -1, txtCaseField(9))
        If strExc(1) < strSrvDate(1) Then strExc(1) = strSrvDate(1)
        txtCaseField(4) = TransDate(PUB_GetWorkDay1(strExc(1), True), 1)
        'end 2020/12/4
        
   'Added by Morgan 2020/3/4
   '美國IDS若已輸入申請案號則設定所限為系統日+1週,法限不設
   'Removed by Morgan 2022/2/21 取消,已有其他管控--郭
   'Modified by Lydia 2024/06/28 恢復此本所期限之設定(拿掉txtCaseField(4))--郭
   'ElseIf txtCaseField(4) = "" And txtCaseField(9) = "" And field(9) = "101" And txtCaseField(7) = "214" Then
   'Modified by Morgan 2024/7/9 不管有無法限都適用此規則--郭
   'ElseIf txtCaseField(9) = "" And field(9) = "101" And txtCaseField(7) = "214" Then
   ElseIf field(9) = "101" And txtCaseField(7) = "214" Then
   'end 2024/7/9
      If field(11) <> "" Then
         'Modified by Lydia 2024/06/28 原IDS的本限早於此預設本限，則請以原有本限為主
         'txtCaseField(4) = TransDate(PUB_GetWorkDay1(CompDate(2, 7, strSrvDate(1)), True), 1)
         strExc(1) = PUB_GetWorkDay1(CompDate(2, 7, strSrvDate(1)), True)
         If txtCaseField(4) = "" Or DBDATE(txtCaseField(4)) > strExc(1) Then
            txtCaseField(4) = TransDate(strExc(1), 1)
         End If
         'end 2024/06/28
      End If
   'end 2024/06/28
   'end 2022/2/21
   'end 2020/3/4
   Else
        Me.txtCaseField(4).Text = TransDate(PUB_GetWorkDay1(Me.txtCaseField(4).Text, True), 1)
    End If
    'end 2013/09/03
   txtCaseField(12) = cp(64)
   txtCaseField(14) = cp(5)
   
   txtCaseField(17) = TransDate(cp(48), 1) 'Add by Morgan 2010/3/19
   
   'Add by Morgan 2020/3/9 --郭
   '分析預設本所期限=承辦期限=收文日+5個工作天(不含收文日)
   If cp(10) = "941" Then
      If cp(14) = "" And cp(27) = "" Then
         strExc(2) = CompWorkDay(5, CompDate(2, 1, cp(5)))
         If Val(strExc(2)) < Val(strSrvDate(1)) Then
            strExc(2) = strSrvDate(1)
         End If
         strExc(2) = TransDate(strExc(2), 1)
         txtCaseField(4) = strExc(2)
         If PUB_IfSetCP48(cp(9)) Then
            If txtCaseField(17) = "" Then
               txtCaseField(17) = strExc(2)
               cp(48) = DBDATE(strExc(2))
            End If
         End If
      End If
   End If
   'end 2020/3/9
   
   'Add by Morgan 2010/3/17
   If cp(10) = "123" Then
     ' lblFavDt.Visible = True
      txtCaseField(4).Left = 960
      txtFavDt.Visible = True
      'Modified by Morgan 2012/3/22 改抓 PA140
      'txtFavDt.Text = TransDate(PUB_GetFavorDate(field(91)), 1)
      txtFavDt.Text = TransDate(field(140), 1)
      CmdFav.Visible = True 'Add by Lydia 2015/02/02
   Else
     ' lblFavDt.Visible = False
      txtCaseField(4).Left = 1320
      txtFavDt.Visible = False
      CmdFav.Visible = False 'Add by Lydia 2015/02/02
   End If
   
   m_CP30 = cp(30) 'Add by Morgan 2011/4/22
   m_CP31 = cp(31) 'Add By Sindy 2010/10/29
   m_CP13 = cp(13) 'Add By Sindy 2010/10/29
   lblCancelReceiveDate = TransDate(cp(57), 1)
   lblFee = cp(16)
   lblPoints = cp(18)
   SetNameToCombo cboPatentName, field(5), field(6), field(7)
   txtCaseField(8) = field(23)
   'Add By Cheng 2002/06/11
   m_strCP27 = "" & cp(27)
   
   
   'Add by Morgan 2007/9/17 已發文不可改指定國家
   'Modify by Morgan 2009/6/10 改不限制未發文,因為存檔時有控制只能新增了
   'If m_strCP27 <> "" Then
   '   cmdCountry.Enabled = False
   'End If
   'end 2009/6/10
   'end 2007/9/17
   
   'Add by Morgan 2010/3/16
   txtFeeYear(1) = cp(53)
   txtFeeYear(2) = cp(54)
   'end 2010/3/16
   
   If intCaseKind = 專利 Then
      'Added by Morgan 2013/3/20
      'Removed by Morgan 2023/3/23 改在下面設定
      'If field(9) = "101" Then
      '   optChoose(2).Enabled = True
      'Else
      '   optChoose(2).Enabled = False
      'End If
      'end 2013/3/20
      
      For i = 0 To 4
             lblPetition(i) = field(i + 26)
      Next
      txtCaseField(2) = field(9)
      'Modify By Cheng 2002/04/24
      '是否取消閉卷欄不要顯示
'      txtCaseField(6) = field(57)
      'Add By Cheng 2002/04/23
      If field(57) = "Y" Then
         Me.Label3(0).Visible = True
         Me.txtCaseField(6).Visible = True
         Me.lblClose.Caption = "已閉卷"
      Else
         Me.Label3(0).Visible = False
         Me.txtCaseField(6).Visible = False
         Me.lblClose.Caption = ""
      End If
      
      'Add By Sindy 2010/10/29
      If field(158) = "" Then
         Combo3 = ""
      Else
         Combo3 = field(158) + "." + PUB_GetCaseAttributeName(field(158))
      End If
      '2010/10/29 End
      
      txtCaseField(8) = field(23)
      txtCaseField(10) = field(48)
      txtCaseField(13) = field(91)
      txtCaseField(11) = field(47)
      'Add By Cheng 2002/03/06
      
      m_field46 = ""
      If field(1) = "CFP" Then
         'Modify by Morgan 2006/5/24
         'Me.txtCaseField(15).Text = "" & field(46)
         'Text1(14) = pA(46)
         If field(9) <> "056" Then
            If field(46) = "Y" Then
               'Add by Lydia 2014/11/24
               m_field46 = field(46)
               'PCT申請日
               txtCaseField(15) = TransDate(field(10), 2)
               'PCT優先權日
               txtCaseField(16) = PUB_GetPCTPriDate(field(91))
               'Added by Morgan 2014/10/27
               'PCT申請號
               Text1(23) = PUB_GetPCTPriNo(field(91))
               'end 2014/10/27
            End If
         End If
         'end 2006/5/24
      End If
      'Add By Cheng 2002/06/11
      Me.Text1(21).Enabled = True
      Me.Text1(21).Text = "" & field(8)
      
      'Add by Morgan 2006/1/19
      'Modify by Morgan 2006/9/20 加法國
      'Modified by Morgan 2015/11/23 +印度040,菲律賓030--禧佩
      'Modified by Morgan 2023/3/24 國家改抓常數，收文有設定的也顯示
      'If field(9) = "101" Or field(9) = "102" Or txtCaseField(2) = "203" Or txtCaseField(2) = "040" Or txtCaseField(2) = "030" Then
      'If field(9) = "101" Or field(9) = "102" Or txtCaseField(2) = "203" Or txtCaseField(2) = "040" Or txtCaseField(2) = "030" Then
      PUB_SetEntityOpt field(1), field(9), field(8), OptChoose
      If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
         If strSrvDate(1) >= PA179啟用日 Then
            If field(179) = "1" Then
               OptChoose(0).Value = True
               old_Entity = OptChoose(0).Caption
            ElseIf field(179) = "2" Then
               OptChoose(1).Value = True
               old_Entity = OptChoose(1).Caption
            ElseIf field(179) = "3" Then
               OptChoose(2).Value = True
               old_Entity = OptChoose(2).Caption
            Else
               old_Entity = ""
            End If
         Else
      'end 2023/3/24
            If InStr(1, field(91), "大個體", 1) > 0 Then
               OptChoose(0).Value = True
               old_Entity = "大個體"
            ElseIf InStr(1, field(91), "小個體", 1) > 0 Then
               OptChoose(1).Value = True
               old_Entity = "小個體"
            'Added by Morgan 2013/3/20
            ElseIf InStr(1, field(91), "微個體", 1) > 0 Then
               OptChoose(2).Value = True
               old_Entity = "微個體"
            'end 2013/3/20
            Else
               old_Entity = ""
            End If
            
         End If 'Added by Morgan 2023/3/24
      End If
      '2006/1/19 end
      txtEngGroup = field(150)  'Added by Morgan 2012/3/12
   Else
      lblPetition(0) = field(8)
      txtCaseField(2) = field(9)
      'Modify By Cheng 2002/04/24
      '是否取消閉卷欄不要顯示
'      txtCaseField(6) = field(15)
      'Add By Cheng 2002/04/23
      If field(15) = "Y" Then
         Me.Label3(0).Visible = True
         Me.txtCaseField(6).Visible = True
         Me.lblClose.Caption = "已閉卷"
      Else
         Me.Label3(0).Visible = False
         Me.txtCaseField(6).Visible = False
         Me.lblClose.Caption = ""
      End If
      
      txtCaseField(10) = field(29)
      txtCaseField(13) = field(18)
      txtCaseField(11) = field(28)
      txtEngGroup = field(79)  'Added by Morgan 2012/3/12
   End If
   
   
   stCP141 = cp(141): stCP142 = cp(142): stCP164 = cp(164) 'Added by Morgan 2024/2/6
   
   OptSendType(1).Caption = PUB_GetCP114Opt1Desc(cp(1), cp(10))   'Added by Morgan 2024/1/22
   
   'Added by Morgan 2024/1/5
   txtCP142 = ""
   OptSendType(1).Value = False
   OptSendType(2).Value = False
   OptSendType(3).Value = False
   Option1(0).Value = False
   Option1(1).Value = False
   Option1(2).Value = False
   If strSrvDate(1) >= 指定日期啟用日 Then
      Label1(121).Visible = True
      Frame1.Visible = True
      Frame2.Visible = True
      Select Case cp(141)
         Case "1"
            OptSendType(1).Value = True
         Case "2"
            OptSendType(2).Value = True
         Case "3"
            OptSendType(3).Value = True
            txtCP142.Text = TransDate(cp(142), 1)
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
      End Select
      If cp(27) = "" Then
         Frame1.Enabled = True
         Frame2.Enabled = True
      Else
         Frame1.Enabled = False
         Frame2.Enabled = False
      End If
   Else
      Label1(121).Visible = False
      Frame1.Visible = False
      Frame2.Visible = False
   End If
   'end 2024/1/5

   
   'Added by Morgan 2012/3/9
   Label3(13) = ""
   Label3(13).Visible = False
   txtEngGroup.Visible = False
   Label1(24).Visible = False
   If Left(cp(12), 1) = "F" Then
      Label3(13).Visible = True
      txtEngGroup.Visible = True
      Label1(24).Visible = True
   End If
   'end 2012/3/9
   
   'Add by Morgan 2010/10/29
   m_strOldCP10 = txtCaseField(7)
   m_strOldPA09 = txtCaseField(10)
   
   'Remove by Morgan 2007/12/24 點指定國按鈕的時候會做
   'If cmdCountry.Visible Then
   '   'edit by nickc 2007/02/02 不用 dll 了
   '   'If objPublicData.ReadCountry(intCaseKind, cp(), strCountry) = False Then GoTo Err
   '   If ClsPDReadCountry(intCaseKind, cp(), strCountry) = False Then GoTo Err
   'Else
   '   strCountry = ""
   'End If
   
   '2007/8/13 ADD BY SONIA銷卷提醒
   CheckCaseDestroy cp(1), cp(2), cp(3), cp(4)
   '2007/8/13 END
   
   Dim strTemp As String
   strTemp = txtCaseField(8)
   CheckKeyIn 0
   CheckKeyIn 2
   CheckKeyIn 7
   txtCaseField(8) = strTemp
   '2008/10/23 MODIFY BY SONIA
   'GetCaseDeadLineData grdDataList, intLastRow, cp(1), cp(2), cp(3), cp(4)
   GetGrid grdDataList, intLastRow, cp(1), cp(2), cp(3), cp(4)
   '2008/10/23 END
   For i = 0 To 15
          strCode(i) = ""
   Next
   For i = 0 To 3
      strCode(i) = cp(i + 1)
   Next
   
   'Modify by Morgan 2005/8/24 改共用
   Call ReadCM
   '2005/8/24 end
   
   '91.11.1 MODIFY BY SONIA
   'If cp(10) <> 發明申請 And cp(10) <> 設計申請 And cp(10) <> 新型申請 And cp(10) <> 聯合申請 Then
   '91.11.19 modify by sonia
   'If cp(10) <> 發明申請 And cp(10) <> 設計申請 And cp(10) <> 新型申請 And cp(10) <> 聯合申請 And cp(10) <> 主張優先權 Then
   'Modify by Morgan 2004/10/14 加 121
   'If cp(10) <> 主張優先權  Then
   If cp(10) <> 主張優先權 And cp(10) <> "121" Then
   '91.11.19 end
   '91.11.3 END
      cmdPriority.Visible = False
      strPriority1 = ""
      strPriority2 = ""
      strPriority3 = ""
   Else
      'Modify by Amy 2014/04/14 加 strPriority5
      'Modify by Morgan 2007/4/25 加 strPriority4
      If ClsPDReadPriority(cp(), strPriority1, strPriority2, strPriority3, strPriority4, strPriority5) = False Then GoTo ErrHnd
      cmdPriority.Visible = True
   End If
   SSTab1.Tab = 0
   '2005/5/6 CANCEL BY SONIA
   'If txtCaseField(5) = "" Then
   '   If intLastRow > 0 Then
   '      grdDataList.Row = 1
   '      grdDataList.Col = 7
   '      txtCaseField(5) = grdDataList.Text
   '   End If
   'End If
   '2005/5/6 END
   txtCaseField(0).SetFocus

   ' 90.07.06 modify by louis
   grdDataList.ColWidth(8) = 0
   grdDataList.ColWidth(9) = 2000
   grdDataList.ColWidth(10) = 0
   
   'Added by Morgan 2020/12/23
   '美國IDS要看備註以便點選相關案號
   If field(9) = "101" And cp(10) = "214" Then
      grdDataList.ColWidth(1) = 900
      grdDataList.ColWidth(4) = 0
      grdDataList.ColWidth(5) = 0
      grdDataList.ColWidth(6) = 1200
      grdDataList.ColWidth(7) = 0
      grdDataList.ColWidth(9) = 3300
   End If
   'end 2020/12/23
      
   'Add by Morgan 2003/12/07
   Call PUB_CheckSales(cp(1), cp(2), cp(3), cp(4), cp(5), cp(13), lblSalesName)
   'End 2003/12/07
   
   'Add by Morgan 2004/3/29
   Call GetDivCase
   
   'Added by Lydia 2016/10/13 分割案若有主張優先權於分案時，不必輸入優先權資料，請直接帶母案的優先權資料P-114427(母案P-104573)
   If strPriority1 = "" And txtDivCaseNo(1) & txtDivCaseNo(2) <> "" And Me.txtCaseField(7) = 主張優先權 Then
        strExc(1) = txtDivCaseNo(1)
        strExc(2) = txtDivCaseNo(2)
        strExc(3) = txtDivCaseNo(3)
        strExc(4) = txtDivCaseNo(4)
        If Not ClsPDReadPriority(strExc, strPriority1, strPriority2, strPriority3, strPriority4, strPriority5) Then
        End If
   End If
   'end 2016/10/13
   
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
   
   'Modify By Sindy 2014/1/29
   m_CP31isYGetCP05 = GetCP31isY_CP05(cp(1), cp(2), cp(3), cp(4)) '取得本所案號新案件的收文日
   
   'Added by Lydia 2020/03/31 事務所合併日起新案只能空白或J，不可輸T
   If Val(m_CP31isYGetCP05) >= 事務所合併日 Then
       lblPA161.Caption = "特殊出名公司                     (J:智權公司 空白:系統預設)"
   'end 2020/03/31
   
   'Add By Sindy 2013/12/16
   'If strSrvDate(1) < InvoiceStartDate Then
   'Modified by Lydia 2020/03/31 +ElseIf
   ElseIf Val(m_CP31isYGetCP05) < Val(InvoiceStartDate) Then
      lblPA161.Caption = "是否以專利商標出名         (Y:是)"
   ElseIf field(1) = "PS" And field(9) <> "000" Then
      lblPA161.Caption = "特殊出名公司                     (J:智權公司 空白:系統預設)"
   End If
   '2013/12/16 END
Else

   frm050101_1.ReChoose intNowReceive, strReceiveCode()
   bolLeave = True
   intLeaveKind = 1
   Unload Me
End If

'Add by Morgan 2009/12/23 延期不可改案件性質
grdDataList.Enabled = True
If cp(10) = "404" Then
   txtCaseField(7).Enabled = False
   If cp(27) <> "" Then
      grdDataList.Enabled = False '已發文不可再點選
   End If
Else
   txtCaseField(7).Enabled = True
End If
   
   'Added by Morgan 2012/6/18
   '若為併號請以連絡單通知電腦中心處理
   If cp(27) <> "" Or cp(9) > "C" Then
      txtCode(0).Enabled = False
      txtCode(1).Enabled = False
      txtCode(2).Enabled = False
      txtCode(3).Enabled = False
   Else
      txtCode(0).Enabled = True
      txtCode(1).Enabled = True
      txtCode(2).Enabled = True
      txtCode(3).Enabled = True
   End If
   'end 2012/6/18
   
   'Added by Morgan 2012/7/20
   '是否為複雜或特殊案件
   txtCP147 = cp(147)
   txtCP147.Tag = txtCP147
   If cp(14) = "" And txtCP147 = "" Then txtCP147 = GetCP147Default()
   'end 2012/7/20
   
   'Modify by Amy 2017/07/13 開放服務業務也顯示顯示專利商標出名欄位 ex:CPS-000096
   If cp(1) = "CFP" Then
     txtPA161 = field(161)
     'Add By Sindy 2023/3/31
     LblPA61.Visible = True
     txtPA61.Visible = True
     txtPA61 = field(61)
     txtPA61.Tag = txtPA61
     '2023/3/31 END
   Else
     txtPA161 = field(85)
   End If
   'Modify by Amy 2016/08/29
   txtPA161.Tag = txtPA161
   
   'Added by Morgan 2012/9/5
   '新案可設定是否以專利商標出名欄位
   'Modify by Amy 2017/07/13 +服務業務 新案且非台灣 就顯示專利商標出名欄位-秀玲
   If cp(31) = "Y" And (cp(1) = "CFP" Or (cp(1) = "CPS" And field(9) <> "000")) Then
      lblPA161.Visible = True
      txtPA161.Visible = True
      'Add by Amy 2016/08/12 +客戶檔收據公司別
      'Mark by Amy 2018/07/03 個案為空白會預設申請人出名公司,若個案改為空仍會一直預帶(CFP-029915)-秀玲:拿掉(2017/11/24 Mark掉的不見了)
'      If txtPA161 = MsgText(601) Then
'        If cp(1) = "CFP" Then
'            txtPA161 = GetReceiptCmp(Left(GetNewFagent(field(26)), 8), Mid(GetNewFagent(field(26)), 9, 1), cp(1), field(9))
'        ElseIf field(8) <> MsgText(601) Then
'            txtPA161 = GetReceiptCmp(Left(GetNewFagent(field(8)), 8), Mid(GetNewFagent(field(8)), 9, 1), cp(1), field(9))
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
   'end 2012/9/5
   
   '2013/10/31 add by sonia 非台灣新申請案收費0,第一次分案時要提醒二案案件備註加註同時合併計算結餘" T-189182(T-188512)
   'Modified by Lydia 2022/03/09 'Modified by Lydia 2022/03/09 改判斷分案日;  ex.2021/06/18 CFT案件承辦人若空白時，預設為國家檔之CFT承辦人---統一判斷
   'If txtCaseField(2) <> "000" And InStr(NewCasePtyList, txtCaseField(7)) > 0 And txtCaseField(0) = "" And Val(cp(16)) = 0 And Left(cp(12), 1) <> "F" Then
   If txtCaseField(2) <> "000" And InStr(NewCasePtyList, txtCaseField(7)) > 0 And Val(cp(149)) = 0 And Val(cp(16)) = 0 And Left(cp(12), 1) <> "F" Then
      MsgBox "此新申請案未收費, 若有前案則請至第二頁頁籤之案件備註欄加註與前案號合併計算結餘(前案之案件備註也要加註)!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1
      txtCaseField_GotFocus (13)
      txtCaseField(13).SetFocus
   End If
   '201/10/31 end
   
   txtCP97 = cp(97) 'Add by Amy 2014/09/05 增加承辦人計件值欄位讓user修改-玲玲
   
   'Added by Lydia 2021/02/19 (有承辦人)讀取齊備日
   m_EP06 = ""
   If txtCaseField(0) <> "" Then
       strExc(0) = "select ep06 from EngineerProgress Where EP02='" & cp(9) & "' "
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
       If intI = 1 Then
           m_EP06 = "" & RsTemp.Fields("ep06")
       End If
   End If
   'end 2021/02/19
   
   'Add by Amy 2022/10/21 +簽核頁籤,接洽單電子收文才顯示「檢視接洽單」鈕
   cmdFile.Visible = False
   Me.cmdCPP.Visible = False 'Add By Sindy 2022/11/1
   SSTab1.TabVisible(2) = False
   'Modify by Sindy 2022/11/23
   'txtF0301 = cp(140)
   txtF0301 = Pub_GetIsFlowCP140(cp(9))
   '2022/11/23 END
   'Add by Amy 2022/11/16
   Label11.Visible = False: txtF0309.Visible = False '目前狀態
   Check11.Visible = False '急件
   Check11.Value = 0 'Add By Sindy 2023/1/10 要先清欄位值,再後續判斷是否急件
   'end 2022/11/16
   'Modify by Amy 2023/01/03 +Len(txtF0301) = 10,8碼(結案單)不可開接洽單會錯
   If strSrvDate(1) >= 接洽單電子收文啟用日 And txtF0301 <> MsgText(601) And Len(txtF0301) = 10 Then
        cmdFile.Visible = True
        Me.cmdCPP.Visible = True 'Add By Sindy 2022/11/1
        '補件完成 欄-案件表單流程備註檔屬於分案作業相關資訊
        SetFlow004TextBox txtF0407, txtF0301, " And F0408 in('A5','A6','A7') And F0409 in('A5','A6','A7') "
        '案件表單簽核檔
        strSql = "SELECT ST02||nvl(F0208,'') 簽核人員,decode(F0202," & ShowFlow簽核人員身份 & ") 身份,sqldateT(F0205) 日期,sqltime6(F0206) 時間,decode(F0207," & ShowFlow簽核結果 & ") 簽核結果,F0204 FROM FLOW002,Staff WHERE F0201='" & txtF0301 & "' and F0204=ST01(+) order by decode(F0205,null,2,1) asc,F0205||Decode(length(F0206),5,'0','')||F0206 asc,F0202,F0203 asc"
        If rsTmp.State = adStateOpen Then rsTmp.Close
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If rsTmp.RecordCount > 0 Then
           Set GRD1.Recordset = rsTmp
           SetGrd
        End If
        cmdOK(0).Enabled = False
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
             cmdOK(0).Enabled = True
        End If
        arrTmp = Split(GetFlow003Data(txtF0301, , "F0308||';'||Nvl(F0309,'NULL')||';'||F0307"), ";")
        stF0309_Now = arrTmp(1)
        stF0307_Now = arrTmp(2)
        'Add by Amy 2022/11/16 +表單狀態/急件
        'Modify by Sindy 2022/11/22
        txtF0309 = PUB_GetCP157forF0309(cp(9)) '表單狀態
        Label11.Visible = True: txtF0309.Visible = True
        Check11.Visible = True '急件
        If cp(122) = "Y" Then Check11.Value = 1
        'end 2022/11/16
        
        '下一處理人員是A7(多筆案件性質已處理一筆)且目前表單狀態不是已分案
        If arrTmp(0) = "A7" And stF0309_Now <> "17" Then
            CmdAddInfo.Enabled = True
            txtNote.Locked = False
        End If
        '接洽單電子收文顯示簽核頁籤
        SSTab1.TabVisible(2) = True
        'Add by Amy 2022/11/16 狀態為 程序補件 時,切至 簽核 頁籤
        If stF0309_Now = "20" Then SSTab1.Tab = 2
        'Add by Amy 2022/12/26 直接開啟接洽單-玲玲
        frm090801_Q.SetParent Me
        frm090801_Q.m_blnCallPrint = True
        frm090801_Q.Text5 = txtF0301
        Call frm090801_Q.cmdok_Click(4)
        frm090801_Q.Show
        'end 2022/12/26
   End If
   
   'Add By Sindy 2024/1/30 各部門分案時，若本所期限與法定期限與接洽單的本所期限與法定期限不同時，要提醒
   Call PUB_ChkCRLdtCP06CP07(cp(9))
   
ErrHnd:
ErrorMsg
Screen.MousePointer = varSaveCursor
End Sub

'Added by Morgan 2012/7/20
Private Function GetCP147Default() As String
   '電子電機或生化醫學的申請程序預設 Y
   'Removed by Morgan 2016/5/18 取消 --郭雅娟
   'If (Left(Combo3, 1) = "2" Or Left(Combo3, 1) = "3") And (txtCaseField(7) = "101" Or txtCaseField(7) = "102" Or txtCaseField(7) = "103") Then
   '   GetCP147Default = "Y"
   'End If
End Function

'2008/10/23 ADD BY SONIA
'取得本案期限
Public Sub GetGrid(ByRef grdTemp As MSHFlexGrid, ByRef intLastRow As Integer, ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String, Optional strCode5 As String)
Dim varSaveCursor, varGridWidth() As Variant
Dim strSql As String

   varSaveCursor = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   varGridWidth = Array(300, 1350, 900, 900, 1200, 1200, 1200, 1000, 0, 0, 0, 0, 0)
   SetGridDataListWidth grdTemp, varGridWidth()
   
   'Modify by Morgan 2009/12/23 延期加帶出AB類未發文未取消收文的程序,且下一程序要排除程序管制的案件性質
   If strCode1 = "CFP" Then
      strSql = "select '' V,decode(pa09," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,"
   Else
      strSql = "select '' V,decode(sp09," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,"
   End If
   strSql = strSql + SQLDate("np08") & " 本所期限," & SQLDate("np09") & " 法定期限,np13 機關文號,np14 相關人," & SQLDate("np11") & " 解除期限日期, np01 總收文號, dbms_rowid.rowid_to_restricted(nextprogress.rowid,0), NP15 備註, NP22 序號, NP07,NP08 "
   If strCode1 = "CFP" Then
      strSql = strSql + "from nextprogress,patent,casepropertymap where pa01(+)=np02 and pa02(+)=np03 and pa03(+)=np04 and pa04(+)=np05"
   Else
      strSql = strSql + "from nextprogress,servicepractice,casepropertymap where sp01(+)=np02 and sp02(+)=np03 and sp03(+)=np04 and sp04(+)=np05"
   End If
   strSql = strSql + " and np02=cpm01(+) and np07=cpm02(+) and (np06<>'Y' or np06 is null)"
   'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
   strSql = strSql + " and np02=" + CNULL(strCode1) + " and np03=" + CNULL(strCode2) + " and np04=" + CNULL(strCode3) + " and np05=" + CNULL(strCode4) & strNpSqlOfNoSalesDuty
   
   'Modified by Morgan 2023/10/27
   'If cp(10) = "404" Then
   If txtCaseField(7) = "404" Then
   'end 2023/10/27
      strSql = strSql & " union SELECT '',DECODE(PA09,'" & 台灣國家代號 & "',CPM03,CPM04)" & _
         ",SQLDateT(CP06),SQLDateT(CP07),CP08,NVL(CP40,NVL(CP41,CP42)),''" & _
         ",CP09,DBMS_ROWID.ROWID_TO_RESTRICTED(CASEPROGRESS.RowID,0),CP64,0,CP10,CP06 FROM CASEPROGRESS,CASEPROPERTYMAP,PATENT" & _
         " WHERE " & ChgCaseprogress(strCode1 & strCode2 & strCode3 & strCode4) & " AND CP09<'C' and cp10<>'404' and cp07>0 AND CP27 IS NULL AND CP57 IS NULL" & _
         " AND CPM01(+)=CP01 AND CPM02(+)=CP10 AND pa01(+)=CP01 and pa02(+)=CP02 and pa03(+)=CP03 and pa04(+)=CP04"
   End If
   
   Set grdTemp.Recordset = ClsPDReadRst(strSql)
   
   SetDataListVision grdTemp, True
   intLastRow = 0
   If grdTemp.Rows > 1 Then
      ShowBar grdTemp, intLastRow, grdTemp.Cols - 1
   End If
   Screen.MousePointer = varSaveCursor
End Sub

'Add by Amy 2022/10/21 補件完成
Private Sub CmdAddInfo_Click()
    Dim oTopForm As Form
    Dim i As Integer
        
On Error GoTo ErrHand
    
    If txtNote = MsgText(601) Then
        MsgBox "呈報內容不可為空！", vbExclamation
        Exit Sub
    End If
    
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
    
    intNowReceive = intNowReceive + 1
    If intNowReceive = intTotalReceive Then
        frm050101_1.ReChoose intNowReceive, strReceiveCode()
        Screen.MousePointer = vbHourglass
        '重新搜尋資料
        frm050101_1.RefreshData
        ' 設定滑鼠游標為預設
        Screen.MousePointer = vbDefault
        bolLeave = True
        Unload Me
    Else
        If intNowReceive = intTotalReceive Then
            cmdOK(3).Visible = False
        End If
        ReadAllData
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
    
    
ErrHand:
    Screen.MousePointer = vbDefault
    cnnConnection.RollbackTrans
    MsgBox "補件失敗！" & vbCrLf & Err.Description
End Sub

'2008/10/23 END
Private Sub cmdCountry_Click()
    '920224 nick 新增
    If Nick920224Bol = False Then
   Dim nick920224rs As New ADODB.Recordset
   Set nick920224rs = New ADODB.Recordset
   Dim nick920224str As String
   strCountry = ""
   nick920224str = "select pa09 from patent where pa01='" & cp(1) & "' and pa02='" & cp(2) & "' and pa03='" & cp(3) & "' and pa04>'00'"
   With nick920224rs
        .CursorLocation = adUseClient
        '.Open nick920224str, Connection, adOpenStatic, adLockReadOnly
        .Open nick920224str, cnnConnection, adOpenStatic, adLockReadOnly
        If Not .EOF And Not .BOF Then
            .MoveFirst
            Do While Not .EOF
                If Trim(strCountry) = "" Then
                    strCountry = strCountry & CheckStr(.Fields(0).Value)
                Else
                    strCountry = strCountry & "," & CheckStr(.Fields(0).Value)
                End If
                .MoveNext
            Loop
        End If
        
   End With
    Nick920224Bol = True
    strCountryOld = strCountry 'Add by Morgan 2007/12/24
   End If
   'Modify by Morgan 2005/9/12
   'ModifyAssignCountry strCountry
   ModifyAssignCountry strCountry, TransDate(field(10), 2)
End Sub

Private Sub Process()
   Dim i As Integer
   Dim strOfficeKind As String '所別
   '若承辦人是王協理且未發文則要發EMail通知
   Dim bolMail As Boolean
   Dim bol106 As Boolean '是否有收文主張國際優先權
   Dim bol106Mail As Boolean '是否通知智權人員收文主張國際優先權
   Dim m_605 As String   '檢查收文領證時逾期年費是否有消除期限 2008/10/23 add by sonia
   Dim strA0K11 As String 'Add By Sindy 2014/2/12
   'Added by Lydia 2016/01/28
   Dim strPD As String '判斷國際或國內優先權
   Dim tmpContent As String
   'Added by Lydia 2017/06/06
   Dim rsTmp As New ADODB.Recordset
   Dim strPA(1 To 4) As String 'Added by Lydia 2017/07/18
   Dim m_StrTo As String, m_StrSub As String, m_StrCont As String 'Added by Lydia 2022/05/30 整理frm880005改用寄信模組
   Dim strAnoPA() As String, strCaseNo As String 'Add by Amy 2022/12/09
   
   '若非執行轉本所案號功能

   If Me.txtCode(0).Text = "" Or Me.txtCode(1).Text = "" Then
   
      'Added by Morgan 2018/9/11 CFP電子化
      'Modified by Morgan 2018/10/2 +C類也要檢查接洽單
      'Modified by Morgan 2021/2/23 改國外部案件也要檢查
      If Pub_StrUserSt03 = "P12" And strSrvDate(1) >= CFP第一階段電子化啟用日 And InStr("A,C", Left(cp(9), 1)) > 0 Then
         'Modify By Sindy 2022/12/16 電子收文不用檢查
         If Not (txtF0301 <> "" And Len(txtF0301) = 10) Then
         '2022/12/16 END
            If PUB_CheckPDF2(cp(9), 0, True, , txtCaseField(7)) = False Then
                MsgBox "無接洽單電子檔(.ORDER.PDF),不可分案!", vbCritical
                Exit Sub
            End If
         End If
      End If
      'end 2018/9/11
      'Add by Amy 2022/12/09 電子收文需檢查一案兩請是否有資料
      'Modify by Amy 2022/12/22 +案件性質新型申請(102)才彈
      'Modify by Sindy 2023/3/9 +案件性質新型申請(101)也要檢查
      If txtF0301 <> MsgText(601) And txtCaseField(2) = "231" And (txtCaseField(7) = "102" Or txtCaseField(7) = "101") Then
            ReDim Preserve strAnoPA(1 To TF_PA) As String
            If Pub_GetField("ConsultRecordList", "CRL01='" & txtF0301 & "'", "CRL67") <> MsgText(601) Then
                If PUB_IsDualApplyCom(strPA, strAnoPA, strCaseNo) = False Then
                    MsgBox "一案兩請無資料,不可繼續!", vbCritical
                    Exit Sub
                End If
            End If
      End If
      'end 2022/12/09
      
      'Added by Lydia 2023/12/14 檢查智財協作在分案時若未建立相關案號(caserelation1)時則跳提醒程序人員，但可選擇輸或不輸 !
      'Modified by Lydia 2023/12/15 PS及CPS之智財協作967，TT及S之智財協作737，L之智財協作7601，(也可用案件性質中文判斷)在分案時若未建立相關案號且為ACS且為TIPS的案件時，提醒文字：「案件性質為智財協作，請先依接洽單輸入相關卷號資料」。
'      If field(1) = "CPS" And txtCaseField(7) = "967" Then
'         If PUB_IfCaseRelation1Exists(field(1), field(2), field(3), field(4)) = False Then
'            If MsgBox("案件性質為" & lblCaseProperty.Caption & "，請確認接洽單是否有相關案號，是否補輸入？", vbExclamation + vbYesNo + vbDefaultButton1) = vbYes Then
'               Exit Sub
'            End If
'         End If
'      End If
'      'end 2023/12/14
      If field(1) = "CPS" And InStr(lblCaseProperty.Caption, "智財協作") > 0 Then
         If PUB_ChkACSforTIPS(field(1) & field(2) & field(3) & field(4), , True) = False Then
            MsgBox "案件性質為" & lblCaseProperty.Caption & "，請先依接洽單輸入相關卷號資料", vbExclamation
            Exit Sub
         End If
      End If
      'end 2023/12/15
      
      For i = 0 To 14
         If txtCaseField(i).Enabled Then
            If CheckKeyIn(i) <> 1 Then
               txtCaseField(i).SetFocus
               txtCaseField_GotFocus (i)
               Exit Sub
            End If
         End If
      Next
      If txtCaseField(2) = 221 And (txtCaseField(7) = "101" Or txtCaseField(7) = "102" Or txtCaseField(7) = "103") And strCountry = "" Then
         Screen.MousePointer = vbDefault
         ShowMsg MsgText(9180)
         Exit Sub
      End If
      'Add by Lydia 2014/11/24 PCT進各個國家階段之期限設定
         If m_field46 = "Y" And txtCaseField(7) = 申請復活 Then  '414
          If txtCaseField(5) = "" Then
            MsgBox "申請復活必須要輸相關總收文號 !", vbCritical
            txtCaseField(5).SetFocus
            Exit Sub
          ElseIf txtCaseField(4) = "" Or txtCaseField(9) = "" Then
                 MsgBox "無法計算期限，請先分案發明/新型申請！"
                 Exit Sub
              Else
                  MsgBox "本程序期限資料將會回寫到申請程序！"
          End If
         End If
      
      'Added by Morgan 2020/12/23
      '美國IDS要檢查是否有NP期限沒點選
      If txtCaseField(2) = "101" And txtCaseField(7) = "214" Then
         strExc(0) = ""
         For intI = 1 To grdDataList.Rows - 1
            If grdDataList.TextMatrix(intI, 11) = "214" Then
               If grdDataList.TextMatrix(intI, 0) <> "V" Then
                  strExc(0) = strExc(0) & vbCrLf & "　" & grdDataList.TextMatrix(intI, 9)
               End If
            End If
         Next
         
         If strExc(0) <> "" Then
            If MsgBox("尚有IDS期限未點選，請確認接洽單上所列案號是否皆已點選？" & vbCrLf & vbCrLf & "未點選之IDS期限：" & strExc(0), vbExclamation + vbYesNo + vbDefaultButton2, "美國IDS期限點選確認") = vbNo Then
               Exit Sub
            End If
         End If
      End If
      'end 2020/12/23
      
   '若執行轉本所案號功能
   Else
      i = 15
   End If
   If i = 15 Then
      If CheckKeyIn1(False) Then
         If CheckKeyIn2(False) Then
            '重新檢查欄位有效性
            'Add by Lydia 2014/11/24 TxtValidate內含txtcasefield檢查(checkkeyin)
            If TxtValidate = False Then Screen.MousePointer = vbDefault: Exit Sub
            
            If txtCode(0) <> "" And txtCode(1) <> "" Then
                  strExc(1) = txtCode(0)
                  strExc(2) = txtCode(1)
                  strExc(3) = txtCode(2)
                  If strExc(3) = "" Then strExc(3) = "0"
                  strExc(4) = txtCode(3)
                  If strExc(4) = "" Then strExc(4) = "00"
                  'Modified by Morgan 2023/10/27
                  'strExc(5) = cp(10) '案件性 質
                  strExc(5) = txtCaseField(7) '案件性 質
                  'end 2023/10/27
                  strExc(6) = lblCaseProperty '案件性質名稱
                  strExc(7) = cp(5) '收文日
                  strExc(8) = cp(9) '總收文號
                  strExc(9) = lblPetition(0)
                  'edit by nickc 2007/02/05 不用 dll 了
                  'If Not objLawDll.ChkSameCase(strExc) Then Exit Sub
                  If Not ClsLawChkSameCase(strExc) Then Exit Sub
                  'Added by Lydia 2020/08/18 更新相關卷號前,先檢查是否有重複
                  If m_CP31 = "Y" Then
                      If PUB_ChkUpdCR(field(1), field(2), field(3), field(4), strExc(1), strExc(2), strExc(3), strExc(4)) = False Then
                          Exit Sub
                      End If
                  End If
                  'end 2020/08/18
            End If
            
          'Add By Cheng 2002/08/23 關聯案提醒
          '執行轉本所案號
          If Me.txtCode(0).Text <> "" And Me.txtCode(1).Text <> "" Then
               MsgBox "本次只更新轉本所案號的資料，不更新其他欄位的修改!!!", vbCritical
               'Add by Morgnan 2006/9/19
               Me.Tag = ""
               If InStr(CaseMapOut, cp(10)) > 0 Then
                  Set frm1104_1.m_form = Me
                  frm1104_1.m_CP09 = cp(9)
                  frm1104_1.m_CaseNoBefore = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4)
                  frm1104_1.m_CaseNoAfter = txtCode(0) & "-" & txtCode(1) & "-" & Right("0" & txtCode(2), 1) & "-" & Right("00" & txtCode(3), 2)
                  If frm1104_1.GetRelation = True Then
                     frm1104_1.Show vbModal
                     If Me.Tag = "0" Then
                        Exit Sub
                     End If
                  End If
               End If
               'end 2006/9/19
          'Add By Cheng 2002/12/03
          '執行分案
          Else
               '2008/10/23 add by sonia 領證時若下一程序年費605~607期限已過期未續辦,一定要勾選
               If txtCaseField(7) = "601" Then
                  m_605 = "Y"
                  With grdDataList
                     If .Recordset.RecordCount > 0 Then
                        For intI = 1 To .Rows - 1
                           '若有逾期年費未勾選時跳離
                           If .TextMatrix(intI, 0) = "" And (.TextMatrix(intI, 11) = "605" Or .TextMatrix(intI, 11) = "606" Or .TextMatrix(intI, 11) = "607") And .TextMatrix(intI, 12) < strSrvDate(1) Then
                              m_605 = "N"
                              Exit For
                           End If
                        Next
                        '未勾選
                        If m_605 = "N" Then
                           MsgBox "下一程序之年費已過期，一定要勾選消除期限，並一併於領證時繳納 !"
                           Exit Sub
                        End If
                     End If
                  End With
               End If
               '2008/10/23 end
               'Add by Amy 2018/04/09 年費移作次年時若下一程序有605~607期限則提醒並不可存檔
               If cp(1) = "CFP" And txtCaseField(7) = "612" Then
                    With grdDataList
                        For intI = 1 To .Rows - 1
                            '因605~607 可能同時存在,當已分某其中一個再進入分案時仍會再檢查,故只先檢查605
                            If .TextMatrix(intI, 0) = "" And .TextMatrix(intI, 11) = "605" Then
                                MsgBox "有未收文之" & .TextMatrix(intI, 1) & "期限，不可存檔!"
                                Exit Sub
                            End If
                        Next intI
                    End With
               End If
               'end 2018/04/09
               
               'Add by Morgan 2008/10/22
               If txtCaseField(7) = "413" And txtCaseField(5) = "" Then
                  MsgBox "自請撤回的相關總收文號不可空白!!"
                  txtCaseField(5).SetFocus
                  Exit Sub
               End If
              'Add by Lydia 2014/11/24 PCT進各個國家階段之期限設定
               If m_field46 = "Y" And txtCaseField(7) = 申請復活 Then  '414
                If txtCaseField(5) = "" Then
                  MsgBox "申請復活必須要輸相關總收文號 !", vbCritical
                  txtCaseField(5).SetFocus
                  Exit Sub
                ElseIf txtCaseField(4) = "" Or txtCaseField(9) = "" Then
                       MsgBox "無法計算期限，請先分案發明/新型申請！"
                       Exit Sub
'                    Else
'                       MsgBox "本程序期限資料將會回寫到申請程序！" '第2次檢查不彈訊息
                End If
               End If
               
               'Add by Morgan 2007/3/14 多國案若有其他相同案已核准、發證、公開時不可分案、發文、提申
               If cp(27) = "" And (txtCaseField(7) = "101" Or txtCaseField(7) = "102" Or txtCaseField(7) = "103") Then
                  If PUB_SameCaseCheck(cp) = False Then
                     Exit Sub
                  End If
               End If
               'end 2007/3/14
               
               strExc(10) = "" 'Add By Sindy 2023/3/13
               '若本所期限小於收文日
               If Me.txtCaseField(4).Text <> "" And Me.txtCaseField(14).Text <> "" Then
                   If Val(DBDATE(txtCaseField(4))) < Val(DBDATE(txtCaseField(14))) Then
                       MsgBox "本所期限小於收文日, 將更新本所期限為系統日!!!", vbExclamation + vbOKOnly
                       strExc(10) = "不用再檢查期限" 'Add By Sindy 2023/3/13
                   End If
               End If
               If strExc(10) <> "不用再檢查期限" Then 'Add By Sindy 2023/3/13
                  'Add By Sindy 2022/11/22
                  If strSrvDate(1) >= 接洽單電子收文啟用日 Then
                     'Modify By Sindy 2023/4/12 + , , , , lblReceiveCode
                     If PUB_CRLUseCP07CheckCP06(m_CP31, txtCaseField(2), cp(1), txtCaseField(7), txtCaseField(4), txtCaseField(9), , , , lblReceiveCode) = False Then
                        txtCaseField(4).SetFocus
                        Exit Sub
                     End If
                  End If
                  '2022/11/22 END
               End If
               
               'Add by Morgan 2004/3/17
               '分割案提示
               If txtCaseField(7) = "307" Then
                  If txtDivCaseNo(1).Text = "" And txtDivCaseNo(2).Text = "" And txtDivCaseNo(3).Text = "" And txtDivCaseNo(4).Text = "" Then
                     If MsgBox("本進度的案件性質為分割，確定不輸入分割母案本所案號？", vbExclamation + vbYesNo) = vbNo Then
                        txtDivCaseNo_GotFocus 1
                        txtDivCaseNo(1).SetFocus
                        Exit Sub
                     End If
                  '檢查母案本所案號是否存在
                  ElseIf CheckDivCase() = False Then
                     txtDivCaseNo_GotFocus 1
                     txtDivCaseNo(1).SetFocus
                     Exit Sub
                  End If
                  
                  'Added by Morgan 2018/3/29
                  If cp(27) = "" Then
                    'Modified by Morgan 2020/6/29
                    'PUB_Get307CtrlDate txtDivCaseNo(1), txtDivCaseNo(2), txtDivCaseNo(3), txtDivCaseNo(4), strExc(1)
                    PUB_Get307CtrlDate txtDivCaseNo(1), txtDivCaseNo(2), txtDivCaseNo(3), txtDivCaseNo(4), , strExc(1)
                    'end 2020/6/29
                    If Val(strExc(1)) > 0 And strExc(1) < strSrvDate(1) Then
                       If MsgBox("分案期限(法限：" & ChangeWStringToTDateString(strExc(1)) & ")已過！是否確認要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                          Exit Sub
                       End If
                     'Added by Morgan 2024/6/18 若分割案沒有期限，也一併再跳提醒程序，請再次確認是否無期限--玫音
                     ElseIf Val(strExc(1)) = 0 And txtCaseField(4) = "" Then
                        If MsgBox("請再次確認是否無期限？", vbQuestion + vbExclamation + vbDefaultButton2) = vbNo Then
                           Exit Sub
                        End If
                     End If
                  End If
                  'end 2018/3/29
               End If
               'Add end---
               'Add by Morgan 2004/6/8
               'P、CFP一案二申請於分案時建立關聯；提醒條件：同一申請人同一天收文同一申請國家同一案件名稱但不同專利種類時，若未建立關聯則提醒使用者。
               m_PA2 = ""
               If PUB_DualCaseExist(cp, m_PA2) = True Then
                  If PUB_DualCaseRelationExist(cp) = False Then
                     If MsgBox("本案與 " & m_PA2 & " 案可能為一案兩申請且尚未建立關聯，確定要繼續？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then Exit Sub
                  End If
               End If
               
               'Add by Morgan 2005/5/3 檢查一件國內案不可關聯兩件相同國家之國外案
               If txtCode(5) <> "" Then
                  If PUB_CheckDaulCaseMap(field(1) & field(2) & field(3) & field(4), txtCode(4) & txtCode(5) & txtCode(6) & txtCode(7), , True) = True Then
                     Exit Sub
                  End If
               End If
               '2005/5/3 end
               
               'Add by Morgan 2006/12/14
               'Modified by Morgan 2012/4/18 +判斷CFP案才要
               'If txtCaseField(2) = "101" Or txtCaseField(2) = "102" Or txtCaseField(2) = "203" Then
               'Modified by Morgan 2015/11/23 +印度040,菲律賓030--禧佩
               'Modified by Lydia 2016/09/13 國別改用共用變數
               'If intCaseKind = 專利 And txtCaseField(2) = "101" Or txtCaseField(2) = "102" Or txtCaseField(2) = "203" Or txtCaseField(2) = "040" Or txtCaseField(2) = "030" Then
               If intCaseKind = 專利 And txtCaseField(2) <> "" And InStr(CFP_ChkEntity, txtCaseField(2)) > 0 Then
                  strExc(8) = ""
                  If OptChoose(0).Value = True Then
                     strExc(8) = OptChoose(0).Caption
                  ElseIf OptChoose(1).Value = True Then
                     strExc(8) = OptChoose(1).Caption
                  'Added by Morgan 2013/3/20
                  ElseIf OptChoose(2).Value = True Then
                     strExc(8) = OptChoose(2).Caption
                  'end 2013/3/20
                  End If
                  
                  'Modified by Morgan 2024/12/10 個體別順序會因國家有所不同,且客戶設定目前只設定是否可減免,故只能用大個體中文判斷(非大個體都是可減免)
                  'strExc(9) = OptChoose(1).Caption
                  'For intI = 1 To 5
                  '   If txtAD(intI).Text = "N" Then
                  '      strExc(9) = OptChoose(0).Caption
                  '      Exit For
                  '   End If
                  'Next
                  'If strExc(8) = OptChoose(0).Caption Or strExc(9) = OptChoose(0).Caption Then 'Added by Morgan 2013/5/8 微個體另外有檢查
                  '   If strExc(9) <> strExc(8) Then
                  '      If MsgBox("本案客戶減免設定為【" & strExc(9) & "】與基本檔不同，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                  
                  'Modified by Morgan 2025/5/6 法人也不可減免 Ex:CFP-035133--玫音
                  strExc(9) = ""
                  For intI = 1 To 5
                     If txtAD(intI).Text = "N" Then
                        'strExc(9) = "大個體"
                        strExc(9) = "大個體,法人"
                        Exit For
                     End If
                  Next
                  'If (strExc(9) = "大個體" Or strExc(8) = "大個體") Then
                  '   If strExc(9) <> strExc(8) Then
                  If (InStr(strExc(9), "大個體") > 0 Or strExc(8) = "大個體" Or strExc(8) = "法人") Then
                     If InStr(strExc(9), strExc(8)) = 0 Then
                        If MsgBox("申請人減免身份與案件個體別不一致，是否要繼續？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
                  'end 2024/12/10
                           Exit Sub
                        End If
                     End If
                  End If
                  
               End If
               
               'Add by Morgan 2006/12/29 判斷是否上發文日
               m_bolUpdCP27 = False
               m_strCP44 = ""
               If txtCaseField(0).Text <> "" And cp(27) = "" Then
                  '超頁、超項費(917)若新案已發文則詢問是否上發文日(預設是)
                  '2010/1/6 modify by sonia 加938超頁費,939超項費
                  If txtCaseField(7) = "917" Or txtCaseField(7) = "938" Or txtCaseField(7) = "939" Then
                      strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' AND CP10 IN (" & CaseMapOut & ") and cp27>0"
                      intI = 1
                      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                      If intI = 1 Then
                         If MsgBox("新申請案已發文，請問本程序是否要上發文日？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
                            m_bolUpdCP27 = True
                            m_strCP44 = "" & RsTemp.Fields("CP44") 'Add by Morgan 2010/2/3
                         End If
                     End If
                  End If
               End If
               'end 2006/12/29
               
               'Add by Morgan 2007/4/20
               'PCT案
               '1.若已收文主張國際優先權則PCT優先權日一定要輸
               '2.沒收文主張國際優先權且沒輸PCT優先權日時要提醒確認
               '3.沒收文主張國際優先權但有輸PCT優先權日時發Mail通知智權人員補收文
               bol106 = False: bol106Mail = False
               strPD = "": tmpContent = ""  'Added by Lydia 2016/01/28
               'Modified by Morgan 2019/11/27 排除接續案
               'If txtCaseField(15) <> "" Then
               If txtCaseField(15) <> "" And m_bolXCACase = False Then
               'end 2019/11/27
                  Select Case txtCaseField(7)
                     Case "101", "102" '發明申請,新型申請
                        'Modified by Lydia 2016/01/28 以優先權檔判斷是否要主張國際優先權(106)或國內優先權(121)
                        'strExc(0) = "SELECT * FROM CASEPROGRESS WHERE CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND CP10='106' AND CP57 IS NULL"
                        'intI = 1
                        'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        ''有收106
                        'If intI = 1 Then bol106 = True
                        'Modified by Lydia 2016/02/18 改成共用模組,另外增加106,121的判斷
                        If ClsPDReadPriority(cp(), strExc(5), strExc(6), strExc(7)) = False Then
                           strExc(5) = ""
                        End If
                        bol106 = PUB_CheckPDMsg(cp(1), cp(2), cp(3), cp(4), txtCaseField(7), field(9), tmpContent, strExc(5))
                        
                        '有輸PCT優先權日
                        If txtCaseField(16) <> "" Then
                           strFirstPriDate = PUB_GetFirstPriDate(cp)
                           If strFirstPriDate <> "" Then
                              If strFirstPriDate <> txtCaseField(16) Then
                                 If MsgBox("PCT優先權日與最早優先權日不符，請確認是否無誤！", vbYesNo + vbDefaultButton2) = vbNo Then
                                    Exit Sub
                                 End If
                              End If
                           End If
                           If bol106 = False Then bol106Mail = True
                        '沒輸PCT優先權日
                        Else
                           If bol106 = True Then
                              'Modified by Lydia 2016/01/28
                              'MsgBox "本案為PCT案且已收文主張國際優先權，PCT優先權日不可空白！"
                              MsgBox "本案為PCT案且已收文" & IIf(tmpContent = "", "主張國際優先權", tmpContent) & "，PCT優先權日不可空白！"
                              Exit Sub
                           Else
                              'Modified by Lydia 2016/01/28
                              'If MsgBox("本案是否有主張國際優先權？", vbYesNo + vbDefaultButton1) = vbYes Then
                              If MsgBox("本案是否有" & IIf(tmpContent = "", "主張國際優先權", tmpContent) & "？", vbYesNo + vbDefaultButton1) = vbYes Then
                                 Exit Sub
                              End If
                           End If
                        End If
                     'Modified by Lydia 2016/02/18 + 121
                     Case "106", "121" '主張國際優先權,主張國內優先權
                        'Added by Lydia 2016/02/18 判斷優先權資料是否存在
                        bol106 = PUB_CheckPDMsg(cp(1), cp(2), cp(3), cp(4), txtCaseField(7), field(9), tmpContent, strPriority1)
                        If bol106 = False Then
                           MsgBox tmpContent & "尚未輸入優先權資料！"
                           Exit Sub
                        End If
                        'end 2016/02/18
                        strPCTPriDate = PUB_GetPCTPriDate(field(91))
                        strFirstPriDate = PUB_GetFirstPriDate2(strPriority2)
                        If strPCTPriDate = "" Then
                           MsgBox "本案未記錄PCT優先權日，請檢查申請案期限是否正確！"
                        ElseIf strFirstPriDate <> strPCTPriDate Then
                           If MsgBox("PCT優先權日與最早優先權日不符，請確認是否無誤！", vbYesNo + vbDefaultButton2) = vbNo Then
                              Exit Sub
                           End If
                        End If
                     
                  End Select
                  
               End If
               'end 2007/4/20

          End If
                    
         'Add by Morgan 2005/6/10 多國案檢查
         m_bolIsMutiNation = False
         m_bolIsInsCR = False
         If txtCode(4) <> "" And txtCode(0) = "" And (txtCaseField(7) = "101" Or txtCaseField(7) = "102" Or txtCaseField(7) = "103") Then
            m_bolIsMutiNation = CheckMutiNation
            If m_bolIsMutiNation = True Then
               'Add by Morgan 2006/6/28 判斷多國主案
               i = InStr(MultiCountryPriority, txtCaseField(2))
               If i > 0 Then
                  '為主案第一順位國家(美國)
                  If i = 1 Then
                     If txtCaseField(3) = "Y" Then
                        'Modified by Morgan 2017/5/12 改系統提醒請照跳,但可讓程序人員輸入--慧汶 Ex.CFP-029474
                        'MsgBox "本案應為多國案主案，是否為多國案不可上[Y]！", vbInformation
                        If MsgBox("本案應為多國案主案，是否為多國案不應上[Y]！是否確認要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                        'end 2017/5/12
                           txtCaseField(3).SetFocus
                           Exit Sub
                        End If 'Added by Morgan 2017/5/12
                     End If
                  '檢查是否有其他主案優先順位國家
                  Else
                     intI = 1
                     strExc(0) = "select 1 from (select cr05,cr06,cr07,cr08 from caserelation" & _
                        " where cr01='" & cp(1) & "' and cr02='" & cp(2) & "' and cr03='" & cp(3) & "' and cr04='" & cp(4) & "' and cr05='CFP'" & _
                        " union select cm01,cm02,cm03,cm04 from casemap" & _
                        " where cm05='" & txtCode(4) & "' and cm06='" & txtCode(5) & "' and cm07='" & txtCode(6) & "' and cm08='" & txtCode(7) & "' and cm10='0' and cm01='CFP'" & _
                        "),patent where pa01(+)=cr05 and pa02(+)=cr06 and pa03(+)=cr07 and pa04(+)=cr08" & _
                        " and instr('" & Left(MultiCountryPriority, i - 1) & "',pa09)>0 and pa57 is null"
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     '有
                     If intI = 1 Then
                        'Modified by Morgan 2015/5/18 多國案可於建立關聯時設多個主案
                        'If txtCaseField(3) <> "Y" Then
                        '   MsgBox "本案不可為多國案主案，請檢查其他多國案！", vbInformation
                        If txtCaseField(3) <> "Y" And txtCaseField(3).Text <> cp(21) Then
                           If MsgBox("本多國案有其他主案優先順位國家，本案是否確定設為主案？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        'end 2015/5/18
                              txtCaseField(3).SetFocus
                              Exit Sub
                           End If
                        End If
                     '無
                     Else
                        If txtCaseField(3) = "Y" Then
                           'Modify by Morgan 2006/9/7 因為不一定會先分主案，改確定後可繼續
                           'MsgBox "本案應為多國案主案，是否為多國案不可上[Y]！", vbInformation
                           If MsgBox("本案應為多國案主案，是否多國案確定要上[Y]？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                              txtCaseField(3).SetFocus
                              Exit Sub
                           End If
                        End If
                     End If
                  End If
               Else
               'end 2006/6/28
                  If txtCaseField(3) <> "Y" Then
                     If MsgBox("本案應為多國案，確定是否為多國案不上[Y]？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                        txtCaseField(3).SetFocus
                        Exit Sub
                     End If
                  End If
               End If
            End If
         End If
            
            'Add By Sindy 2014/2/12 若財務處已開立收據,且收據的公司別與案件的特殊出名公司不符時,
            '顯示訊息,讓使用者可選擇是否修改,預設在"是"
            strSql = "select cp60,a0k01,a0k11 from caseprogress,acc0k0" & _
                     " where cp09='" & cp(9) & "' and cp60 is not null and cp60=a0k01(+)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If "" & RsTemp.Fields("a0k01") <> "" Then
                  strA0K11 = "" & RsTemp.Fields("a0k11")
                  If strA0K11 = "1" Then strA0K11 = "T"
                  If (txtPA161 <> "" Or strA0K11 = "T" Or strA0K11 = "J") And _
                     txtPA161 <> strA0K11 Then
                     If MsgBox("財務處開立的收據公司與分案之特殊出名公司不符,是否修改特殊出名公司？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbYes Then
                        Exit Sub
                     End If
                  End If
               End If
            End If
            '2014/2/12 END
            
            'Add By Sindy 2014/9/17 承辦人有異動時檢查是否有設定核判表
            If txtCaseField(0).Text <> "" And txtCaseField(0).Text <> txtCaseField(0).Tag Then
               'Add By Sindy 2014/9/17 檢查是否有設定核判表
               If InStr("P10,P11", GetST15(txtCaseField(0).Text)) > 0 Then
                  'Modified by Morgan 2023/10/27
                  'If PUB_ChkIsSetPromoterReader(txtCaseField(0).Text, cp(1), cp(10)) = False Then
                  'Modify By Sindy 2024/6/26 +txtCaseField(2)
                  If PUB_ChkIsSetPromoterReader(txtCaseField(0).Text, cp(1), txtCaseField(7), , , , txtCaseField(2)) = False Then
                  'end 2023/10/27
                     'Modified by Morgan 2025/2/20 游經理->李柏翰經理
                     MsgBox "此承辦人該案件性質尚未設定核判表，請通知李柏翰經理轉電腦中心設定後再進行分案。", vbInformation
                     txtCaseField(0).SetFocus
                     Exit Sub
                  End If
               End If
            End If
            '2014/9/17 END
            'Added by Lydia 2020/01/21 所有系統別分案，若該收文號存在於下一程序(NP24)，若修改案件性質則彈訊息
            If txtCaseField(7).Tag <> txtCaseField(7).Text Then
                If Pub_CheckNP24Exists(lblReceiveCode.Caption) = True Then
                End If
            End If
            'end 2020/01/21
            
            'Add by Amy 2023/01/03 從Command2搬過來,接洽單電子收文後玲玲反應,太早關
            If PUB_CheckFormExist("frm090801_Q") = True Then
                 Unload frm090801_Q
            End If
    
            ' 設定滑鼠游標為等待狀態
            Screen.MousePointer = vbHourglass
            
            'If SaveDatabase Then
            If FormSave Then
               If m_B209Msg <> "" Then MsgBox m_B209Msg, vbInformation, "內部收文「檢視中說」分案提醒" 'Added by Morgan 2025/6/17
               'Add by Morgan 2008/1/11
               '當電腦中心人員做新案的分案時需詢問是否發Mail通知工程師
               If Pub_StrUserSt03 = "M51" And txtCaseField(0) <> "" And cp(31) = "Y" Then
                  PUB_M51Mail cp(1) & cp(2) & cp(3) & cp(4), txtCaseField(0).Text
               End If
               'end 2008/1/11
               
               'Add by Morgan 2005/6/10
               If m_bolIsInsCR = True Then
                  MsgBox "本案已自動建立多案相關！", vbInformation
               End If
               
               '2009/10/23 add by sonia 新加坡發明實審若未提檢索報告則提醒操作者,仍可分案但不可發文
               'Modified by Morgan 2012/11/13 +非PCT案才要--慧汶
               If txtCaseField(2) = "014" And Text1(21) = "1" And txtCaseField(7) = "416" And field(46) = "" Then
                  CheckOC3
                  strSql = "Select cp09,cp57 From caseProgress WHERE " & ChgCaseprogress(field(1) & field(2) & field(3) & field(4)) & " and cp10='421' "
                  AdoRecordSet3.CursorLocation = adUseClient
                  AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If AdoRecordSet3.RecordCount > 0 Then
                     If AdoRecordSet3.Fields("cp57") <> "" Then MsgBox "請提醒智權同仁：新加坡發明必須先提檢索報告才可提實審！", vbInformation
                  Else
                     MsgBox "請提醒智權同仁：新加坡發明必須先提檢索報告才可提實審！", vbInformation
                  End If
               End If
               '2009/10/23 end
               
               'Added by Lydia 2017/06/06 韓國案件若主張台灣案優先權，於分案及韓國案主張優先權106發文時檢查台灣案是否收文"優先權電子交換"(437)，若未收文請e-mail提醒業務收文。
               'Modified by Lydia 2021/09/13 排除設計案 + And Text1(21) <> "3"  ; ex. CFP-32691是韓國案，主張台灣案P-126542優先權
               If txtCaseField(2) = "012" And txtCaseField(7) = "106" And cp(13) <> "" And Text1(21) <> "3" Then
                  Dim strDesc As String 'Added by Lydia 2017/08/28
                  'Modified by Lydia 2017/08/28 被主張之優先權案若尚未發文,優先權為案號; 若中途轉本所案
                  'strExc(0) = "SELECT PA01,PA02,PA03,PA04 From PRIDATE, PATENT WHERE PD01='" & field(1) & "' AND PD02='" & field(2) & "' AND PD03='" & field(3) & "' AND PD04='" & field(4) & "' AND PD07='000' AND PD06=PA11(+) AND PD05=PA10(+) AND PD07=PA09(+) "
                  strExc(0) = "SELECT NVL(P1.PA01,P2.PA01) PA01,NVL(P1.PA02,P2.PA02) PA02,NVL(P1.PA03,P2.PA03) PA03,NVL(P1.PA04,P2.PA04) PA04,PD06 From PRIDATE, PATENT P1,PATENT P2 " & _
                              "WHERE PD01='" & field(1) & "' AND PD02='" & field(2) & "' AND PD03='" & field(3) & "' AND PD04='" & field(4) & "' AND PD07='000' " & _
                              "AND PD06=P1.PA11(+) AND PD05=P1.PA10(+) AND PD07=P1.PA09(+) " & _
                              "AND SUBSTR(PD06,1,1)=P2.PA01(+) AND SUBSTR(PD06,-9,6)=P2.PA02(+) AND SUBSTR(PD06,-3,1)=P2.PA03(+) AND SUBSTR(PD06,-2)=P2.PA04(+) "
                  intI = 1
                  Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strExc(1) = ""
                     rsTmp.MoveFirst
                     Do While Not rsTmp.EOF
                        'Modified by Lydia 2017/07/18 debug:要傳台灣案號
                        'If PUB_ChkCPExist(field, "437") = False Then
                        
                        'Added by Lydia 2017/08/28 中途轉本所案,一律發信
                        If Trim("" & rsTmp.Fields("PD06")) <> "" And Trim("" & rsTmp.Fields("PA01")) = "" Then
                            strExc(1) = strExc(1) & IIf(strExc(1) <> "", "、", "") & rsTmp.Fields("PD06")
                        Else
                        'end 2017/08/28
                            strPA(1) = rsTmp.Fields("PA01"): strPA(2) = rsTmp.Fields("PA02")
                            strPA(3) = rsTmp.Fields("PA03"): strPA(4) = rsTmp.Fields("PA04")
                            If PUB_ChkCPExist(strPA, "437") = False Then
                            'end 2017/07/18
                               strExc(1) = strExc(1) & IIf(strExc(1) <> "", "、", "") & IIf(rsTmp.Fields("PA03") & rsTmp.Fields("PA04") = "000", rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02"), rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02") & "-" & rsTmp.Fields("PA03") & "-" & rsTmp.Fields("PA04"))
                            End If
                        End If 'end 2017/08/28
                        rsTmp.MoveNext
                     Loop
                     If strExc(1) <> "" Then
                         strExc(1) = IIf(field(3) & field(4) = "000", field(1) & "-" & field(2), field(1) & "-" & field(2) & "-" & field(3) & "-" & field(4)) & "主張國際優先權,被主張之台灣案" & strExc(1) & "尚未收文優先權電子交換 !"
                         Call PUB_SendMail("", cp(13), cp(9), strExc(1))
                     End If
                  End If
                  Set rsTmp = Nothing
               End If
               'end 2017/06/06
               
                  'Add By Cheng 2003/08/12
                  '執行分案
                  If Me.txtCode(0).Text = "" Or Me.txtCode(1).Text = "" Then
                      '若本所期限為當日或假日期限, 則發E-Mail給承辦人
                      If Me.txtCaseField(0).Text <> "" Then
                          '取得更新後的本所期限
                          m_strCP06Update = GetCP06(Me.lblReceiveCode.Caption)
                          If WorkDayCheck = True Then
                              'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
                              'strOfficeKind = PUB_GetST06(strUserNum)
                              'Load frm880005
                              'strOfficeKind = PUB_GetST06(strUserNum)
                              ''若使用者為北所人員, 則E-Mail後面不加@taie.com.tw
                              'If strOfficeKind = "1" Then
                              '    'IAIN傳給韓聖文 93.9.30 91028->79075  '2005/6/22 79075->94003  '2006/4/3->95008  '2006/7/10->79075
   '                           '             frm880005.txtEmail(0).Text = Me.txtCaseField(0).Text
                               '   '88024已離職,故此句無妨
                               '   frm880005.txtEmail(0).Text = IIf(Me.txtCaseField(0).Text = "88024", "79075", Me.txtCaseField(0).Text)
                              ''若使用者非北所人員, 則E-Mail後面加@taie.com.tw
                              'Else
                              '    'Modify By Cheng 2003/08/27
                              '    '分所照寄
   '                           '         frm880005.txtEmail(0).Text = ReGetStaffCode(Me.txtCaseField(0).Text)
                               '   'IAIN傳給韓聖文 93.9.30 91028->79075  '2005/6/22 79075->94003  '2006/4/3->95008  '2006/7/10->79075
   '                            '            frm880005.txtEmail(0).Text = Me.txtCaseField(0).Text & "@taie.com.tw"
                               '   frm880005.txtEmail(0).Text = IIf(Me.txtCaseField(0).Text = "88024", "79075", Me.txtCaseField(0).Text) & "@taie.com.tw"
                              'End If
                              'frm880005.txtEmail(1).Text = "本所期限到期通知"
                              'frm880005.txtEmail(2).Text = "收文號：" & Me.lblReceiveCode.Caption & vbCrLf & _
                                                                          "本所案號：" & Me.lblCaseCode.Caption & vbCrLf & _
                                                                          "案件名稱" & Me.cboPatentName.Text & vbCrLf & _
                                                                          "案件性質：" & Me.txtCaseField(7).Text & " " & Me.lblCaseProperty.Caption & vbCrLf & _
                                                                          "收文日：" & DBYEAR(Me.txtCaseField(14).Text) - 1911 & " 年 " & DBMONTH(Me.txtCaseField(14).Text) & " 月 " & DBDAY(Me.txtCaseField(14).Text) & " 日 " & vbCrLf & _
                                                                          "承辦人：" & Me.txtCaseField(0).Text & " " & Me.lblPromoter.Caption & vbCrLf & _
                                                                          "本所期限：" & DBYEAR(m_strCP06Update) - 1911 & " 年 " & DBMONTH(m_strCP06Update) & " 月 " & DBDAY(m_strCP06Update) & " 日 " & vbCrLf & vbCrLf & _
                                                                          "※本所期限為當日期限或假日期限!!!"
                              'frm880005.Form_Activate: DoEvents
                              'frm880005.cmdOK_Click 0: DoEvents
                              m_StrTo = Me.txtCaseField(0).Text
                              m_StrSub = "本所期限到期通知"
                              m_StrCont = "本所案號：" & Me.lblCaseCode.Caption & vbCrLf & _
                                                                          "案件名稱" & Me.cboPatentName.Text & vbCrLf & _
                                                                          "案件性質：" & Me.txtCaseField(7).Text & " " & Me.lblCaseProperty.Caption & vbCrLf & _
                                                                          "收文日：" & DBYEAR(Me.txtCaseField(14).Text) - 1911 & " 年 " & DBMONTH(Me.txtCaseField(14).Text) & " 月 " & DBDAY(Me.txtCaseField(14).Text) & " 日 " & vbCrLf & _
                                                                          "承辦人：" & Me.txtCaseField(0).Text & " " & Me.lblPromoter.Caption & vbCrLf & _
                                                                          "本所期限：" & DBYEAR(m_strCP06Update) - 1911 & " 年 " & DBMONTH(m_strCP06Update) & " 月 " & DBDAY(m_strCP06Update) & " 日 " & vbCrLf & vbCrLf & _
                                                                          "※本所期限為當日期限或假日期限!!!"
                              PUB_SendMail strUserNum, m_StrTo, cp(9), m_StrSub, m_StrCont
                              'end 2022/05/30
                              'Add by Morgan 2004/2/18
                              '若已發過EMail則控制不再發送
                              bolMail = True
                          End If
                      End If
                  End If
                  
                  'Add by Morgan 2007/4/20
                  If bol106Mail = True Then
                    'Modified by Lydia 2016/01/28
                    'Call PUB_SendMail(strUserNum, cp(13), cp(9), cp(1) & cp(2) & cp(3) & cp(4) & " 案為PCT發明案且有輸[PCT優先權日]但未收文[主張國際優先權]，請補收文該程序！", " ")
                    Call PUB_SendMail(strUserNum, cp(13), cp(9), cp(1) & cp(2) & cp(3) & cp(4) & " 案為PCT發明案且有輸[PCT優先權日]但未收文[" & IIf(tmpContent = "", "主張國際優先權", tmpContent) & "]，請補收文該程序！", " ")
                  End If
                  'end 2007/4/20
                  
                  'Add by Morgan 2007/8/27
                  If m_bolAnnuityAlert = True Then
                     Call PUB_SendMail(strUserNum, "79017", cp(9), cp(1) & cp(2) & cp(3) & cp(4) & " 案為PCT發明案" & m_strAlertMsg & "！", " ")
                  End If
                  
                  'Add by Morgan 2004/2/18
                  '若承辦人是王協理且未發文則要發EMail通知
                  stCP14 = Me.txtCaseField(0).Text
                  'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
                  If bolMail = False And stCP14 = "99050" Then
                      stCP09 = Me.lblReceiveCode.Caption
                      Call PUB_SendMail(strUserNum, "99050", stCP09, "分案通知")
                  End If
          
                  'Modify by Morgan 2005/4/29 有承辦人才發
                  If txtCaseField(0).Text <> "" Then
                     'Add by Morgan 2005/4/15
                     'Mark by Lydia 2021/02/19 改在ReadAllData
                     'CheckOC3
                     'strSql = "Select EP06 From EngineerProgress Where EP02='" & cp(9) & "'"
                     'AdoRecordSet3.CursorLocation = adUseClient
                     'AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                     'If AdoRecordSet3.RecordCount > 0 Then
                     '  m_EP06 = "" & AdoRecordSet3.Fields("EP06")
                     'Else
                     '  m_EP06 = ""
                     'End If
                     ''2005/4/15 end
                     'end 2021/02/19
                     
                     'Add by Morgan 2005/3/29 若承辦人變更時通知智權人員，並註明文件齊備日
                     'Modify by Morgan 2008/7/31 B類收文不用通知智權人員
                     'If txtCaseField(0).Tag <> "" And txtCaseField(0).Text <> txtCaseField(0).Tag Then
                     If txtCaseField(0).Tag <> "" And txtCaseField(0).Text <> txtCaseField(0).Tag And Left(cp(9), 1) <> "B" Then
                       strExc(0) = "原承辦人：" & txtCaseField(0).Tag & " " & GetStaffName(txtCaseField(0).Tag) & vbCrLf
                       'Added by Lydia 2017/05/25 若承辦人之部門為程序(P12),不要帶文件齊備日的訊息(如最後一行),因為程序一定不會輸文件齊備日欄,程序承辦的案件性質也不一定有文件的問題.
                       If PUB_GetStaffST15(txtCaseField(0).Tag, "1") <> "P12" Then
                          If m_EP06 <> "" Then
                             strExc(0) = strExc(0) & "文件齊備日：" & ChangeWStringToTDateString(m_EP06) & vbCrLf
                          'Remove by Morgan 2008/8/12 未齊備不必註明--禧佩(智權人員反應)
                          'Else
                          '   strExc(0) = strExc(0) & "文件齊備日：未齊備" & vbCrLf
                          End If
                       End If 'end 2017/05/22
                                              
                        'Added by Morgan 2023/8/14
                        PUB_SetEngInform cp(9), txtCaseField(0).Tag, False, strExc(1)
                        If strExc(1) <> "" Then strExc(0) = strExc(0) & vbCrLf & strExc(1)
                        'end 2023/8/14
           
                       'Modified by Lydia 2021/04/19 收件人為智權人員+承辦人，另外CC給原承辦人
                       'Call PUB_SendMail(strUserNum, cp(13), cp(9), "承辦人變更通知", "", strExc(0))
                       'Modified by Lydia 2021/11/02 收件人和CC排除舜禹F5588
                       'Call PUB_SendMail(strUserNum, cp(13) & ";" & txtCaseField(0), cp(9), "承辦人變更通知", "", strExc(0), , , , , txtCaseField(0).Tag)
                       'Modified by Lydia 2025/03/13 改用模組取得
                       'Call PUB_SendMail(strUserNum, cp(13) & IIf(txtCaseField(0) <> "F5588", ";" & txtCaseField(0), ""), cp(9), "承辦人變更通知", "", strExc(0), , , , , IIf(txtCaseField(0).Tag <> "F5588", txtCaseField(0).Tag, ""))
                       Call PUB_SendMail(strUserNum, cp(13) & IIf(InStr(Pub_SetF51Order("F", ""), txtCaseField(0)) = 0, ";" & txtCaseField(0), ""), cp(9), "承辦人變更通知", "", strExc(0), , , , , IIf(InStr(Pub_SetF51Order("F", ""), txtCaseField(0).Tag) = 0, txtCaseField(0).Tag, ""))
                     End If
                     
                     'Add By Sindy 2022/12/27 承辦人為P12.專利處程序時,要發分案通知 (And PUB_GetST03(txtCaseField(0).Text) = "P12")
                     If strSrvDate(1) >= 接洽單電子收文啟用日 Then
                        If txtCaseField(0).Text <> txtCaseField(0).Tag Then
                           'Modify By Sindy 2023/1/5 CFP案改為不分承辦人身分都要發通知給程序人員，仍依照業務區寄給負責的程序人員
                           '接洽單第一筆案件性質,才發
                           'If Pub_ConIsFirstCRC(txtF0301, cp(9)) = True Then
                              strExc(10) = Left(Trim(GetSignOffEmp("NP", CStr(field(1)), CStr(field(2)), txtCaseField(2), CStr(field(1)) & CStr(field(2)) & CStr(field(3)) & CStr(field(4)))), 5)
                              'Add By Sindy 2024/11/14 CFP的分案通知，
                              '   調整新申請案的案件性質若承辦人為'F'字頭時，請加CC給系統特殊設定人員「H」，目前是陳品薇98012
                              If InStr(NewCasePtyList, txtCaseField(7)) > 0 And Left(txtCaseField(0), 1) = "F" Then
                                 strExc(10) = strExc(10) & ";" & Pub_GetSpecMan("H")
                              End If
                              '2024/11/14 END
                              Call PUB_SendMail(strUserNum, strExc(10), cp(9), "分案通知（本所案號：" & field(1) & "-" & field(2) & IIf(field(3) & field(4) = "000", "", "-" & field(3) & "-" & field(4)) & "）" & lblCaseProperty.Caption, "", IIf(Check11.Value = 1, "注意：急件！", ""))
                           'End If
                        End If
                     End If
                     '2022/12/27 END
                  End If
                  
               ' 設定滑鼠游標為預設
               Screen.MousePointer = vbDefault
               intLeaveKind = 1
               If IsEmptyText(txtCode(0)) = False And IsEmptyText(txtCode(2)) = False Then
                  strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(field(1) & field(2) & field(3) & field(4))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If RsTemp.Fields(0) < 1 Then
                     MsgBox "原本所案號 " & field(1) & field(2) & field(3) & field(4) & "已無案件進度資料，請通知收文人員刪號！", vbInformation
                  Else
                     MsgBox "原本所案號為 " & field(1) & field(2) & field(3) & field(4) & "，請自行更新原本所案號之下一程序資料 !", vbInformation
                  End If
               End If
               'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之年費605、維持費606、延展費607，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。
               If strMsgCloseCancel <> "" Then
                  MsgBox "已還原「" & strMsgCloseCancel & "」期限", vbInformation, "取消閉卷"
               End If
               If txtCaseField(6) = "Y" And field(9) = "221" Then  '申請國EPC
                  MsgBox "請自行取消有效子案之閉卷欄位！", vbInformation, "取消閉卷"
               End If
               'end 2025/06/30
               
               intNowReceive = intNowReceive + 1
               If intNowReceive = intTotalReceive Then
                  frm050101_1.ReChoose intNowReceive, strReceiveCode()
                  ' 90.07.06 modify by louis
                  ' 設定滑鼠游標為等待狀態
                  Screen.MousePointer = vbHourglass
                  ' 90.07.06 modify by louis (重新搜尋資料)
                  frm050101_1.RefreshData
                  ' 設定滑鼠游標為預設
                  Screen.MousePointer = vbDefault
                  bolLeave = True
                  Unload Me
               Else
                  If intNowReceive = intTotalReceive Then
                     cmdOK(3).Visible = False
                  End If
                  ReadAllData
               End If
            '911202 nick
            Else
                MsgBox "存檔失敗, 請洽電腦中心人員!!!", vbExclamation + vbOKOnly
            End If
            ' 設定滑鼠游標為預設
            Screen.MousePointer = vbDefault
         End If
       End If
   End If
End Sub

'Add By Sindy 2022/10/31 該文號的卷宗區
Private Sub cmdCPP_Click()
    'Mark by Amy 2022/12/22 卷宗區與接洽單可同時開
'   'Add by Amy 2022/11/16
'   If PUB_CheckFormExist("frm090801_Q") = True Then
'        Unload frm090801_Q
'   End If
   Screen.MousePointer = vbHourglass
   frm100101_L.m_CP09 = lblReceiveCode
   frm100101_L.m_strKey = lblReceiveCode
   frm100101_L.SetParent Me
   If frm100101_L.QueryData = True Then
      frm100101_L.Show
      Me.Hide
   Else
      Unload frm100101_L
   End If
   Screen.MousePointer = vbDefault
End Sub

'Add by Amy 2022/10/21 檢視接洽單
Private Sub cmdFile_Click()
    frm090801_Q.SetParent Me
    frm090801_Q.m_blnCallPrint = True
    frm090801_Q.Text5 = txtF0301
    Call frm090801_Q.cmdok_Click(4)
    frm090801_Q.Show 'Add by Amy 2022/11/16 改獨立視窗
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim varSaveCursor
   'Add by Amy 2022/11/16 按 確定/回前畫面/下一筆/結束,若接洽單已開需關閉
   'Modify by Amy 2023/01/03  原Index >= 0 (確定鈕)改至檢查完存檔前關
   If Index >= 1 And Index <= 3 Then
        If PUB_CheckFormExist("frm090801_Q") = True Then
             Unload frm090801_Q
        End If
   End If

   Select Case Index
      Case 0 '確定
         varSaveCursor = Screen.MousePointer
         Screen.MousePointer = vbHourglass
         Process
         Screen.MousePointer = varSaveCursor
      Case 1, 2
         If Index = 2 Then
            intLeaveKind = 0
         Else
            intLeaveKind = 1
            frm050101_1.ReChoose intNowReceive, strReceiveCode()
         End If
         bolLeave = False
         Unload Me
      Case 3
         If MsgBox("你並未存檔，確定到下一筆嗎?", vbYesNo + vbCritical) = vbYes Then
            intNowReceive = intNowReceive + 1
            If intNowReceive = intTotalReceive - 1 Then
               cmdOK(3).Visible = False
            End If
            ReadAllData
            'Add by Morgan 2004/10/20 設定減免身分
            SetAD 1
            SetAD 2
            SetAD 3
            SetAD 4
            SetAD 5
         End If
      Case 4
         Me.Hide
         frm050106_1.intWhereToGo = 1
         frm050106_1.Show
         frm050106_1.txtCode(4) = txtCode(4)
         frm050106_1.txtCode(5) = txtCode(5)
         frm050106_1.txtCode(6) = txtCode(6)
         frm050106_1.txtCode(7) = txtCode(7)
         frm050106_1.txtCode(0) = cp(1)
         frm050106_1.txtCode(1) = cp(2)
         frm050106_1.txtCode(2) = cp(3)
         frm050106_1.txtCode(3) = cp(4)
         m_bolCheckCM = True
      Case 5
         'Modify by Morgan 2004/9/7 改多案相關卷號
         'Where1103ComeFrom Me, cp(1), cp(2), cp(3), cp(4)
         frm1104.intWhereComeFrom = 1
         Set frm1104.m_form = Me
         frm1104.m_CRL01 = txtF0301.Text 'Add By Sindy 2022/11/23
         frm1104.Show
         frm1104.txtSystem = cp(1)
         frm1104.txtCode(0) = cp(2)
         frm1104.txtCode(1) = cp(3)
         frm1104.txtCode(2) = cp(4)
         frm1104.GetRelation
         Me.Hide
         m_bolCheckCM = True
         m_bolCheckCP21 = True
      Case 6
         frm040101_2.iGo = 5
         frm040101_2.Show
         Me.Hide
      Case 7
         Me.Hide
         frm050107_1.intWhereToGo = 1
         frm050107_1.Show
         frm050107_1.txtCode(4) = cp(1)
         frm050107_1.txtCode(5) = cp(2)
         frm050107_1.txtCode(6) = cp(3)
         frm050107_1.txtCode(7) = cp(4)
      'Add by Morgan 2004/6/14   一案兩請資料
      Case 8
         Set frm040109_1.frmParent = Me
         Me.Hide
         frm040109_1.Show
         frm040109_1.txtCode(0) = cp(1)
         frm040109_1.txtCode(1) = cp(2)
         frm040109_1.txtCode(2) = cp(3)
         frm040109_1.txtCode(3) = cp(4)
         'Add By Sindy 2022/11/23
         If txtF0301 <> MsgText(601) Then
            strExc(10) = Pub_GetCRLCaseMap(txtF0301, "3", "CFP", cp(1), cp(2), cp(3), cp(4))
            If strExc(10) <> "" Then
               frm040109_1.txtCode(4) = SystemNumber(strExc(10), 1)
               frm040109_1.txtCode(5) = SystemNumber(strExc(10), 2)
               frm040109_1.txtCode(6) = SystemNumber(strExc(10), 3)
               frm040109_1.txtCode(7) = SystemNumber(strExc(10), 4)
            End If
         End If
         '2022/11/23 END
   End Select
End Sub

Private Function FormSave() As Boolean

   Dim i As Integer, intSaveMode As Integer
   Dim strTxt(1 To 50) As String, iStep As Integer, strTmp As String
   Dim StrSQLa As String, WorkDate1 As String, WorkDate2 As String
   Dim rsA As New ADODB.Recordset
   'edit by nickc 2007/02/02
   'Dim sPA(1 To T_PA) As String
  ' Dim sSP(1 To T_SP) As String
   Dim sPA() As String
   Dim sSP() As String
   'add by nickc 2007/02/02
   ReDim sPA(1 To TF_PA) As String
   ReDim sSP(1 To tf_SP) As String
   
   Dim StrSqlB As String
   Dim rsB As New ADODB.Recordset
   Dim strPromoteDate  As String '承辦期限
   Dim stCP31 As String
   Dim stDate(0 To 10) As String
   Dim strNP07 As String, strNP08 As String, strNP09 As String '年費期限
   Dim strNP22 As String
   Dim bolAddRec As Boolean 'Add by Morgan 2010/3/18 是否新增未收款無法發文紀錄
   Dim m_list() As String '相關案 Added by Morgan 2012/3/26
   Dim strSqlUpdatEP As String 'Added by Morgan 2012/8/3
   Dim st307Msg As String '分割案提醒訊息 Added by Morgan 2018/3/28
   Dim stCP122 As String 'Add by Amy 2022/11/16
   Dim douStPrice As Double, douLowPrice As Double
   
   m_bolAnnuityAlert = False
   
   FormSave = False
 
 On Error GoTo CheckingErr
 
   cnnConnection.BeginTrans

   iStep = 1
   '執行轉本所案號功能
   If Me.txtCode(0).Text <> "" And Me.txtCode(1).Text <> "" Then
      txtCode(2).Text = Right("0" & Me.txtCode(2).Text, 1)
      txtCode(3).Text = Right("00" & Me.txtCode(3).Text, 2)
      '判斷是否新增專利或服務業務基本案
      'Modified by Morgan 2018/3/9 加判斷沒有進度檔時也要設新案
      Select Case field(1)
         Case "P", "CFP", "FCP"
            strExc(0) = "SELECT * FROM PATENT,caseprogress WHERE " & ChgPatent(txtCode(0) & txtCode(1) & txtCode(2) & txtCode(3)) & " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and rownum<2"
         Case Else
            strExc(0) = "SELECT * FROM SERVICEPRACTICE,caseprogress WHERE " & ChgService(txtCode(0) & txtCode(1) & txtCode(2) & txtCode(3)) & " and cp01(+)=sp01 and cp02(+)=sp02 and cp03(+)=sp03 and cp04(+)=sp04 and rownum<2"
      End Select
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      '若無基本檔資料, 則複製原案號資料
      If intI = 1 Then
         stCP31 = ""
         If IsNull(RsTemp("CP09")) Then stCP31 = "Y" 'Added by Morgan
      Else
         stCP31 = "Y"
         Select Case field(1)
            Case "P", "CFP", "FCP":
               If PUB_ReadPatentData(sPA(), field(1), field(2), field(3), field(4)) Then
                  sPA(1) = Me.txtCode(0).Text
                  sPA(2) = Me.txtCode(1).Text
                  sPA(3) = Me.txtCode(2).Text
                  sPA(4) = Me.txtCode(3).Text
                  If Not PUB_AddNewPatent(sPA()) Then
                     GoTo CheckingErr
                  End If
               End If
               
            Case Else:
               If PUB_ReadServicePracticeData(sSP(), field(1), field(2), field(3), field(4)) Then
                  sSP(1) = Me.txtCode(0).Text
                  sSP(2) = Me.txtCode(1).Text
                  sSP(3) = Me.txtCode(2).Text
                  sSP(4) = Me.txtCode(3).Text
                  If Not PUB_AddNewServicePractice(sSP()) Then
                     GoTo CheckingErr
                  End If
               End If
         End Select
      End If
      
      
'cancel by sonia 2024/11/26 已不立卷不必再通知分所收文人員
'      'Add by Morgan 2010/8/6 若為分所收文案件則發Mail通知收文人員
'      strExc(0) = PUB_GetST06(cp(65))
'      If strExc(0) > "1" Then
'         strExc(1) = "原本所案號 " & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
'         strExc(1) = strExc(1) & " 已更改為 " & txtCode(0) & "-" & txtCode(1) & IIf(txtCode(2) & txtCode(3) = "000", "", "-" & txtCode(2) & "-" & txtCode(3)) & " 。"
'         '2010/12/2 modify by sonia
'         'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'            " values ('" & strUserNum & "','" & cp(65) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'            ",'" & ChgSQL(strExc(1)) & "','如旨' )"
'         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'            " values ('" & strUserNum & "','" & cp(65) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'            ",'" & ChgSQL(strExc(1)) & "','總收文號：" & lblReceiveCode & " 改本所案號如主旨')"
'         cnnConnection.Execute strSql, intI
'      End If
'      'end 2010/8/6
'end 2024/11/26
      
      'Modify by Morgan 2004/9/7 若該案號無基本檔則 CP31='Y' 否則 NULL
      strTxt(iStep) = "UPDATE CASEPROGRESS SET CP01='" & txtCode(0) & "' ,CP02='" & txtCode(1) & "'" & _
         " ,CP03='" & txtCode(2) & "' ,CP04='" & Me.txtCode(3).Text & "' ,CP43='', CP31=" & CNULL(stCP31) & _
         " WHERE CP09 = '" & Me.lblReceiveCode.Caption & "'"

      cnnConnection.Execute strTxt(iStep), intI
      iStep = iStep + 1
      
      'Add by Morgan 2006/9/7
      '更正財務相關資料
      PUB_UpdateAccData cp(9), cp(1) & cp(2) & cp(3) & cp(4)
      'Add by Morgan 2006/9/19
      '要轉關聯
      If Me.Tag = "1" Then
         strExc(1) = txtCode(0)
         strExc(2) = txtCode(1)
         strExc(3) = Right("0" & txtCode(2), 1)
         strExc(4) = Right("00" & txtCode(3), 2)
         PUB_UpdateCaseRelation cp, strExc
      End If
      
      
   '分案(非執行轉本所案號功能)
   Else
      'Add by Morgan 2004/9/23
      '設定客戶減免身分
      For i = 1 To 5
         If txtAD(i).Enabled = True Then
            '身分有變更才要做
            If txtAD(i).Tag <> txtAD(i).Text Then
               strSql = PUB_GetADSQL(field(25 + i), txtCaseField(2).Text, txtAD(i).Text)
               cnnConnection.Execute strSql
            End If
         End If
      Next
   
      'Add by Morgan 2004/3/22
      '若有輸入分割母案本所案號則更新 DIVISIONCASE
      If txtCaseField(7) = "307" Then
         If (txtDivCaseNo(1) <> txtDivCaseNo(1).Tag Or txtDivCaseNo(2) <> txtDivCaseNo(2).Tag Or txtDivCaseNo(3) <> txtDivCaseNo(3).Tag Or txtDivCaseNo(4) <> txtDivCaseNo(4).Tag) Then
            '若原先有建立關聯則更新，否則新增
            If txtDivCaseNo(1).Tag <> "" Then
               strTxt(iStep) = " UPDATE DIVISIONCASE SET DC05='" & txtDivCaseNo(1) & "', DC06='" & txtDivCaseNo(2) & "', DC07='" & txtDivCaseNo(3) & "', DC08='" & txtDivCaseNo(4) & "'" & _
                  " , DC12='" & strUserNum & "', DC13=TO_CHAR(SYSDATE,'YYYYMMDD'), DC14=TO_CHAR(SYSDATE,'HHMISS')" & _
                  " WHERE DC01='" & field(1) & "' AND DC02='" & field(2) & "' AND DC03='" & field(3) & "' AND DC04='" & field(4) & "'"
            Else
               strTxt(iStep) = " INSERT INTO DIVISIONCASE (DC01, DC02, DC03, DC04, DC05, DC06, DC07, DC08, DC09, DC10, DC11 )" & _
                  " VALUES('" & field(1) & "' ,'" & field(2) & "','" & field(3) & "','" & field(4) & "'" & _
                  " ,'" & txtDivCaseNo(1) & "', '" & txtDivCaseNo(2) & "', '" & txtDivCaseNo(3) & "', '" & txtDivCaseNo(4) & "'" & _
                  " ,'" & strUserNum & "', TO_CHAR(SYSDATE,'YYYYMMDD'), TO_CHAR(SYSDATE,'HHMISS'))"
            End If
            cnnConnection.Execute strTxt(iStep)
            iStep = iStep + 1
         End If
      End If
      'Add end 2004/3/22
      
      cp(14) = txtCaseField(0)
      cp(26) = txtCaseField(1)
      cp(21) = txtCaseField(3)
      cp(6) = txtCaseField(4)
      cp(43) = txtCaseField(5)
      cp(10) = txtCaseField(7)
      If txtCaseField(0).Text <> "" And m_strCP27 = "" Then
         '後金自動上發文
         'Modify by Morgan 2006/12/29
         'If cp(10) = "909" Then
         If cp(10) = "909" Or m_bolUpdCP27 = True Then
             cp(27) = strSrvDate(1)
             If m_strCP44 <> "" Then cp(44) = m_strCP44 'Add by Morgan 2010/2/3 代理人也要上--甄妮
         End If
      End If
      
      'Added by Morgan 2023/10/27
      '美國IDS代理人預設Y20825000--郭
      If field(9) = "101" And txtCaseField(7) = "214" And cp(27) = "" And cp(57) = "" Then
         'Added by Morgan 2025/5/20
         '同時有RCE(424)和IDS(214)收文未發文時，IDS不要帶固定的代理人Bacon
         strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='424' and cp158=0 and cp159=0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            cp(44) = ""
            cp(45) = ""
         'end 2025/5/20
         ElseIf cp(44) = "" Then
            cp(44) = "Y20825000"
            'Added by Morgan 2024/4/10 若已有IDS發文給該代理人時要將彼號也帶入--慧汶
            'Modified by Morgan 2024/4/19 不必限定IDS發文，因有可能是案件代理人 Ex:CFP-033687
            strExc(0) = "select cp44,cp45 from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp44='" & cp(44) & "' and cp27>0 order by cp27 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               cp(45) = "" & RsTemp("cp45")
            End If
            'end 2024/4/10
         End If 'Added by Morgan 2025/5/20
      End If
      'end 2023/10/27
      
      'Added by Morgan 2025/5/20
      '同時有RCE(424)和IDS(214)收文未發文時，IDS不要帶固定的代理人Bacon
      If field(9) = "101" And txtCaseField(7) = "424" And cp(27) = "" And cp(57) = "" Then
         strSql = "update caseprogress set cp44='',cp45='' where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='214' and cp158=0 and cp159=0 and cp44='Y20825000'"
         cnnConnection.Execute strSql, intI
      End If
      'end 2025/5/20
      
      
      'Added by Morgan 2024/3/8
      If m_IDSCP44 <> "" Then
         cp(44) = m_IDSCP44
         cp(45) = m_IDSCP45
      End If
      'end 2024/3/8
      
      cp(7) = txtCaseField(9)
      cp(64) = txtCaseField(12)
      cp(5) = txtCaseField(14)
      '若本所期限 < 收文日, 則本所期限 = 系統日
      If cp(6) <> "" And cp(5) <> "" Then
         If Val(cp(6)) < Val(cp(5)) Then cp(6) = strSrvDate(2)
      End If
      
      'Add By Sindy 2022/12/15 有修改案件性質
      If txtCaseField(7).Tag <> txtCaseField(7).Text Then
         'Modify By Sindy 2023/9/22 + , txtCaseField(7).Tag
         If PUB_ModCrLCRCData(cp(9), txtF0301, txtCaseField(7).Text, txtCaseField(7).Tag, field(9), txtCaseField(12)) = False Then
            GoTo CheckingErr
         End If
      End If
      '2022/12/15 END
      
      'Modify by Morgan 2010/10/29 改申請國家或案件性質時重抓(若改標準價原收文資料要維持不變)--秀玲
      If m_strOldCP10 <> txtCaseField(7) Or m_strOldPA09 <> txtCaseField(2) Then
         cp(33) = ""
         cp(34) = ""
         'Modify By Sindy 2022/12/15
'         strExc(0) = "select cf13,cf14 from casefee where cf01=" + CNULL(field(1)) + " and cf02=" + CNULL(txtCaseField(2)) + _
'            " and cf03=" + CNULL(cp(10)) + ""
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If Not IsNull(RsTemp.Fields("cf13")) Then cp(33) = RsTemp.Fields("cf13")
'            If Not IsNull(RsTemp.Fields("cf14")) Then cp(34) = RsTemp.Fields("cf14")
'         End If
         If ClsPDGetCaseLowPrice(field(1), txtCaseField(2), txtCaseField(7), douStPrice, douLowPrice, Text1(21), Left(Combo3, 1), cp(140)) = 1 Then
            cp(33) = douStPrice
            cp(34) = douLowPrice
         End If
         '2022/12/15 END
      End If
      
      'Add by Morgan 2010/3/16
      If txtFeeYear(1).Visible Then
         cp(53) = txtFeeYear(1)
         cp(54) = txtFeeYear(2)
      End If
         
      'Add by Morgan 2010/3/15
      If txtCaseField(17).Enabled Then
         cp(48) = txtCaseField(17)
      End If
      
      'Add by Morgan 2010/6/3
      If cp(10) = "106" Then
         If m_bol106Chk Then
            cp(71) = "Y"
         Else
            cp(71) = ""
         End If
      End If
      'end 2010/6/3
      
      'Add by Morgan 2011/4/22
      '延期要回寫 CP30
      If txtCaseField(7) = "404" Then
         cp(30) = m_CP30
      End If
      
      cp(147) = txtCP147 'Added by Morgan 2012/7/20
      
      'Add by Amy 2015/01/22 北所第一次輸入承辦人與原先不同更新承辦人,若北所分案日為null則更新為系統日
      bolCP14Mail = False
      If txtCaseField(0).Tag = "" Then
        If txtCaseField(0).Tag <> txtCaseField(0) Then
            cp(14) = txtCaseField(0) '原承辦人沒值,且有修改
         Else
            cp(14) = txtCaseField(0).Tag
        End If
        If txtCaseField(0) <> "" And cp(157) = "" Then cp(157) = strSrvDate(1)
      Else
        If Trim(txtCaseField(0)) = "" Then
            '原承辦人有值,但畫面上為空
            strExc(1) = ClsPDGetStaff(txtCaseField(0).Tag, strExc(0))
            strExc(0) = "此程序原承辦人為 " & strExc(0) & " 是否取消原承辦人?" & vbCrLf & vbCrLf & _
                            "是:取消原承辦人 / 否:保留原承辦人"
            
            If MsgBox(strExc(0), vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
                cp(14) = "": cp(157) = ""
                bolCP14Mail = True
            Else
                cp(14) = txtCaseField(0).Tag
                If m_bolIsFirstKeyCP14 And txtCaseField(0) <> "" And cp(157) = "" Then cp(157) = strSrvDate(1)
            End If
        ElseIf txtCaseField(0).Tag <> txtCaseField(0).Text Then
            '原承辦人有值,且有修改
            cp(14) = txtCaseField(0)
            If m_bolIsFirstKeyCP14 And txtCaseField(0) <> "" And cp(157) = "" Then cp(157) = strSrvDate(1)
            bolCP14Mail = True
        Else
            cp(14) = txtCaseField(0).Tag
            If m_bolIsFirstKeyCP14 And txtCaseField(0) <> "" And cp(157) = "" Then cp(157) = strSrvDate(1)
        End If
      End If
      'end 2015/01/22
            
      'Added by Morgan 2024/1/5
      If cp(27) = "" And Frame1.Visible = True And Frame1.Enabled Then
'Removed by Morgan 2024/2/22 已收款通知已改為有設收款後送件時通知管制人
'         If OptSendType(2).Value = True Then
'            If txtCaseField(0) <> "" Then
'               'Modified by Morgan 2024/1/22 不限制承辦人為程序改預設通知管制人(收款會判斷承辦人為程序時優先通知其次才是管制人)
'               'If GetStaffDepartment(txtCaseField(0)) = "P12" Then
'                  'strSql = "update UndeliveredRec set UD04='" & txtCaseField(0) & "' where UD01='" & cp(9) & "' and UD02=" & strSrvDate(1)
'                  strExc(1) = PUB_GetCFPHandler(cp(1) & cp(2) & cp(3) & cp(4))
'                  strSql = "update UndeliveredRec set UD04='" & strExc(1) & "' where UD01='" & cp(9) & "' and UD02=" & strSrvDate(1)
'                  cnnConnection.Execute strSql, intI
'                  If intI = 0 Then
'                     'strSql = "insert into UndeliveredRec(UD01,UD02,UD03,UD04) VALUES('" & cp(9) & "'," & strSrvDate(1) & ",'1','" & txtCaseField(0) & "')"
'                     strSql = "insert into UndeliveredRec(UD01,UD02,UD03,UD04) VALUES('" & cp(9) & "'," & strSrvDate(1) & ",'1','" & strExc(1) & "')"
'                     cnnConnection.Execute strSql, intI
'                  End If
'               'End If
'               'end 2024/1/22
'            End If
'         Else
'            strSql = "delete UndeliveredRec where UD01='" & cp(9) & "'"
'            cnnConnection.Execute strSql, intI
'         End If
'end 2024/2/22

         intI = Abs(OptSendType(1).Value * 1) + Abs(OptSendType(2).Value * 2) + Abs(OptSendType(3).Value * 3)
         If intI > 0 Then
            cp(141) = intI
         Else
            cp(141) = ""
         End If
         If txtCP142.Text <> "" Then
            cp(142) = DBDATE(txtCP142.Text)
         Else
            cp(142) = ""
         End If
         If Frame2.Visible = True Then
            cp(164) = IIf(Option1(0).Value = True, "1", IIf(Option1(1).Value = True, "2", IIf(Option1(2).Value = True, "3", "")))
         End If
                  
         If txtCaseField(0) <> "" Then
            If OptSendType(3).Value And (Option1(0).Value Or Option1(2).Value) Then
               If GetStaffDepartment(txtCaseField(0)) <> "P12" Then 'Added by Morgan 2024/1/22 程序不必通知--郭
                  strExc(0) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & _
                     "分案提醒 「本案需於指定日" & ChangeTStringToTDateString(txtCP142) & IIf(Option1(2).Value, "之後", "") & "方可送件」，請留意承辦時間！"
                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc13)" & _
                     " values('" & strUserNum & "','" & txtCaseField(0) & "',to_char(sysdate,'yyyymmdd')" & _
                     ",to_char(sysdate,'hh24miss'),'" & strExc(0) & "','" & cp(9) & "')"
                  cnnConnection.Execute strSql, intI
               End If
            End If
         End If
      End If
      'end 2024/1/5
      
      strTxt(1) = SubGetCPSQL(cp())
      
      'Add by Morgan 2005/3/29 若承辦人變更時紀錄異動人員日期時間
      If txtCaseField(0).Tag <> "" And txtCaseField(0).Text <> txtCaseField(0).Tag Then
         'Add By Sindy 2013/11/28
         Call PUB_ChgEmpUpdEEP05(cp(9), txtCaseField(0).Tag, txtCaseField(0).Text, "1")
         '2013/11/28 END
         
         'Modified by Morgan 2012/4/30 要寫 Log 並改觸發 Trigger 更新修改人員日期時間
         'strSql = "UPDATE CASEPROGRESS SET CP68='" & strUserNum & "',CP69=to_number(to_char(sysdate,'YYYYMMDD')),CP70=to_number(to_char(sysdate,'HH24MI')) WHERE CP09='" & cp(9) & "'"
         'cnnConnection.Execute strSql
         Pub_SeekTbLog strTxt(1)
         strTxt(1) = "begin user_data.user_enabled:=1; " & strTxt(1) & "; end;"
         'end 2012/4/30
         
      'Added by Morgan 2024/2/6 送件方式修改也要紀錄
      ElseIf cp(141) <> stCP141 Or cp(142) <> stCP142 Or cp(164) <> stCP164 Then
         Pub_SeekTbLog strTxt(1)
         strTxt(1) = "begin user_data.user_enabled:=1; " & strTxt(1) & "; end;"
      End If
      
      cnnConnection.Execute strTxt(1)
      
      'Added by Morgan 2023/8/11
      '若更換工程師或分案予不同的工程師之系統通知--杜燕文
      If txtCaseField(0).Text <> "" Then
         If InStr("P10,P11", GetST15(txtCaseField(0).Text)) > 0 Then
            '分案/改承辦人
            If m_bolIsFirstKeyCP14 Or txtCaseField(0).Text <> txtCaseField(0).Tag Then
               '排除已有發Mail的條件
               If Not (txtCaseField(0).Tag <> "" And txtCaseField(0).Text <> txtCaseField(0).Tag And Left(cp(9), 1) <> "B") Then
                  PUB_SetEngInform cp(9), txtCaseField(0).Tag, True
               End If
            End If
         End If
      End If
      'end 2023/8/11
   
      'Modified by Morgan 2012/8/3 更新齊備日必須在更新CP之後否則Trigger計算的承辦期限會被蓋掉
      '2010/3/3 add by sonia C類來函分所承辦人齊備日為收文日的下一工作天,北所為當日
      'Modified by Lydia 2021/02/19 有齊備日不用再變更 ; ex.CFP-31446 核駁分析CB0003093於1/28修改承辦人[89026=>A6022]，變更到齊備日
      'If lblReceiveCode > "C" And Me.txtCaseField(0).Text <> Me.txtCaseField(0).Tag And m_strCP27 = "" Then
      If lblReceiveCode > "C" And m_EP06 = "" And Me.txtCaseField(0).Text <> Me.txtCaseField(0).Tag And m_strCP27 = "" Then
         If m_CP14ST06 <> "1" Then
            strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & CompWorkDay(2, TransDate(txtCaseField(14), 2), 0) & " WHERE EP02='" & lblReceiveCode & "'"
            'Modify by Morgan 2010/10/1
            'cp(48) = TransDate(Pub_GetHandleDay(field(1), field(9), txtCaseField(7), CompWorkDay(2, TransDate(txtCaseField(14), 2), 0), IIf(txtCaseField(4) = "", "", TransDate(txtCaseField(4), 2))), 1)
            If PUB_IfSetCP48() Then
               cp(48) = TransDate(Pub_GetHandleDay(field(1), field(9), txtCaseField(7), CompWorkDay(2, TransDate(txtCaseField(14), 2), 0), IIf(txtCaseField(4) = "", "", TransDate(txtCaseField(4), 2))), 1)
               'Added by Morgan 2012/8/3
               strSql = "Update caseprogress set cp48=" & CNULL(cp(48), True) & " where cp09='" & cp(9) & "'"
               cnnConnection.Execute strSql, intI
               'end 2012/8/3
            'Removed by Morgan 2012/8/3 承辦期限改由 Trigger 觸發
            'Else
            '   cp(48) = ""
            'end 2012/8/3
            End If
            'end 2010/10/1
         Else
            strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & TransDate(txtCaseField(14), 2) & " WHERE EP02='" & lblReceiveCode & "'"
            'Modify by Morgan 2010/10/1
            'cp(48) = TransDate(Pub_GetHandleDay(field(1), field(9), txtCaseField(7), TransDate(txtCaseField(14), 2), IIf(txtCaseField(4) = "", "", TransDate(txtCaseField(4), 2))), 1)
            If PUB_IfSetCP48() Then
               cp(48) = TransDate(Pub_GetHandleDay(field(1), field(9), txtCaseField(7), TransDate(txtCaseField(14), 2), IIf(txtCaseField(4) = "", "", TransDate(txtCaseField(4), 2))), 1)
               'Added by Morgan 2012/8/3 更新承辦期限改單獨執行,
               strSql = "Update caseprogress set cp48=" & CNULL(cp(48), True) & " where cp09='" & cp(9) & "'"
               cnnConnection.Execute strSql, intI
               'end 2012/8/3
            'Removed by Morgan 2012/8/3 承辦期限改由 Trigger 觸發
            'Else
            '   cp(48) = ""
            'end 2012/8/3
            End If
            'end 2010/10/1
         End If
         cnnConnection.Execute strSql, intI
         
         'Added by Morgan 2018/10/2
         '分案通知
         If CFP第一階段電子化啟用日 <= Val(strSrvDate(1)) And txtCaseField(0) <> "" Then
            strExc(3) = Trim(field(5))
            If strExc(3) = "" Then strExc(3) = Trim(field(6))
            If strExc(3) = "" Then strExc(3) = Trim(field(7))
            strExc(0) = cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & "「" & strExc(3) & "」-->" & lblCaseProperty
            strExc(1) = "本所案號：" & cp(1) & "-" & cp(2) & "-" & cp(3) & "-" & cp(4) & vbCrLf & _
                       "案件名稱：" & strExc(3) & vbCrLf & _
                       "案件性質：" & lblCaseProperty & vbCrLf & _
                       "申請人　：" & GetCustomerName(field(26)) & vbCrLf & _
                       "本所期限：" & ChangeTStringToTDateString(txtCaseField(4)) & vbCrLf & _
                       "來函內容：請至卷宗區參看來函"
             strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " values ('" & strUserNum & "','" & txtCaseField(0) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & ChgSQL(strExc(0)) & "','" & ChgSQL(strExc(1)) & "')"
            cnnConnection.Execute strSql, intI
         End If
         'end 2018/10/2
         
      '2012/3/13 add by sonia 案件性質 941分析, 分案時自動上齊備日
      'modify by sonia 2018/9/12 +IDS214
      ElseIf txtCaseField(7) = "941" Or txtCaseField(7) = "214" Then
         strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & lblReceiveCode & "' AND EP06 IS NULL"
         cnnConnection.Execute strSql, intI
      End If
      '2010/3/3 END
      
      'Add by Morgan 2010/6/30
      '異議答辯、舉發答辯更新對造號數名稱為被異議(舉發)之C類來函資料
      If txtCaseField(7) = "802" Or txtCaseField(7) = "804" Then
         If txtCaseField(5) <> "" And txtCaseField(5) <> txtCaseField(5).Tag Then
            strSql = "update caseprogress a set (cp36,cp37,cp38,cp39,cp40,cp41,cp42)=(select b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42 from caseprogress b where b.cp09=a.cp43) where CP09='" & cp(9) & "'  and exists(select * from caseprogress b where b.cp09=a.cp43 and b.cp10 in ('1801','1802'))"
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2010/6/30
   
      'Add by Lydia 2014/11/24 PCT進各個國家階段之期限設定
      'PCT案恢復專利權(414)期限要都回寫相關總收文號
      If m_field46 = "Y" And txtCaseField(7) = 申請復活 And txtCaseField(5) <> "" And txtCaseField(4) <> "" And txtCaseField(9) <> "" Then
         strSql = "Update Caseprogress Set CP06=" & CNULL(TransDate(txtCaseField(4), 2)) & ",CP07=" & CNULL(TransDate(txtCaseField(9), 2)) & _
            " where cp09='" & txtCaseField(5) & "' and cp27 is null"
         cnnConnection.Execute strSql, intI
      End If
      
      '更新基本檔
      '專利
      If intCaseKind = 專利 Then
         field(150) = txtEngGroup 'Added by Morgan 2012/3/12
         field(91) = txtCaseField(13)
         field(9) = txtCaseField(2)
         
         'Add by Morgan 2006/1/19 大小個體
         'Modify by Morgan 2006/9/21 加法國且申請日>=20050901
         'Modified by Morgan 2015/11/23 +印度040,菲律賓030--禧佩
         'Modified by Morgan 2023/3/24 條件同 ReadAllData
         'If field(9) = "101" Or field(9) = "102" Or (field(9) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901")) Or txtCaseField(2) = "040" Or txtCaseField(2) = "030" Then
         If InStr(CFP_ChkEntity, field(9)) > 0 Or field(179) <> "" Then
            If strSrvDate(1) >= PA179啟用日 Then
               If OptChoose(0).Value = True Then
                  field(179) = "1"
               ElseIf OptChoose(1).Value = True Then
                  field(179) = "2"
               ElseIf OptChoose(2).Value = True Then
                  field(179) = "3"
               End If
            Else
         'end 2023/3/24
               If OptChoose(0).Value = True Then
                  new_Entity = "大個體"
               ElseIf OptChoose(1).Value = True Then
                  new_Entity = "小個體"
               'Added by Morgan 2013/3/20
               ElseIf OptChoose(2).Value = True Then
                  new_Entity = "微個體"
               'end 2013/3/20
               End If
               If InStr(1, field(91), old_Entity, 1) > 0 Then
                  'Modified by Morgan 2020/7/29
                  'field(91) = Replace(field(91), old_Entity, new_Entity, InStr(1, field(91), old_Entity, 1), , 1)
                  field(91) = Replace(field(91), old_Entity, new_Entity, , , 1)
               ElseIf new_Entity <> "" Then
                  If field(91) = "" Then
                     field(91) = new_Entity
                  Else
                     field(91) = new_Entity & "，" & field(91)
                  End If
               End If
               
            End If 'Added by Morgan 2023/3/24
         End If
         '2006/1/19 end
         
         'Add by Morgan 2010/3/17
         If txtFavDt.Visible = True Then
            'Modified by Morgan 2012/3/22 改放 PA140
            'strExc(0) = PUB_GetFavorDate(field(91))
            'If strExc(0) <> "" Then
            '   field(91) = Replace(field(91), "新穎性優惠期日期" & strExc(0) & ";", "")
            'End If
            'field(91) = "新穎性優惠期日期" & txtFavDt.Text & ";" & field(91)
            field(140) = DBDATE(txtFavDt)
            'end 2012/3/22
         End If
         
         '取消閉卷
         If Me.txtCaseField(6).Text = "Y" Then
            field(57) = Empty
            field(58) = Empty
            field(59) = Empty
         End If
         field(23) = txtCaseField(8)
         field(48) = txtCaseField(10)
         field(47) = txtCaseField(11)
         
         Select Case cp(10)
            Case 發明申請, 追加申請
               field(8) = "1"
            Case 新型申請
               field(8) = "2"
            Case 設計申請, 聯合申請
               field(8) = "3"
            '其他案件性質則依所輸入的專利種類更新
            Case Else
               field(8) = Me.Text1(21).Text
         End Select
         
         Select Case Mid(cp(10), 1, 1)
            Case "5"
               field(18) = "Y"
            Case "8"
               field(19) = "Y"
         End Select
      
         If txtCaseField(15).Visible Then
            'Modify by Morgan 2006/5/24
'            If txtCaseField(15) = "" Then
'               field(46) = ""
'            Else
'               field(46) = txtCaseField(15).Text
'            End If
            'Modify by Morgan 2006/7/19 PCT進國家階段可請發明或新型
            'If txtCaseField(7) = 發明申請 And txtCaseField(15) <> "" Then
            If (txtCaseField(7) = 發明申請 Or txtCaseField(7) = 新型申請) And txtCaseField(15) <> "" Then
               'Modify by Morgan 2009/12/24 +PCT申請號
               field(91) = PUB_GetNewCaseMemo(field(91), txtCaseField(16), Text1(23))
               field(46) = "Y"
               field(10) = TransDate(txtCaseField(15), 2)
               'Modify by Amy 2013/12/11 Mark 2013/08/28程式 改由AutoBatchDay做
'               'Add by Amy 2013/08/28 PCT進入國家階段之發明或新型需將PCT申請案下一程序「進入國家階段」期線上「Y」若未算過結餘自動上結餘
'               If Text1(23) <> "" Then
'                    strExc(0) = "Select pa01||pa02||pa03||pa04,pa01,pa02,pa03,pa04 From Patent Where pa11='" & Text1(23) & "' and pa09='056' "
'                    intI = 1
'                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                    If intI > 0 Then
'                        strExc(0) = "Update NextProgress set NP06='Y' Where np07='119' and " & ChgNextProgress(RsTemp.Fields(0))
'                        cnnConnection.Execute strExc(0), intI
'                        bolEndModCash = True
'                        Pub_UpdateEndModCash RsTemp.Fields("pa01"), RsTemp.Fields("pa02"), RsTemp.Fields("pa03"), RsTemp.Fields("pa04")
'                    End If
'                End If
'               'end 2013/08/28
            
            'Added by Morgan 2019/11/27 接續案
            ElseIf m_bolXCACase = True Then
               field(91) = PUB_GetNewCaseMemo(field(91), txtCaseField(16), Text1(23))
               field(91) = Trim(Replace(field(91), "PCT案 by pass;", ""))
               If m_bolPCTbyPass Then
                  field(91) = "PCT案 by pass; " & field(91)
               End If
               field(10) = TransDate(txtCaseField(15), 2)
            'end 2019/11/27
            End If
            'end 2006/5/24
         End If
         
         field(158) = Left(Combo3, 1) 'Add By Sindy 2010/10/29
         
         'Added by Morgan 2012/9/6
         If txtPA161.Visible = True And txtPA161.Tag <> txtPA161 Then
            field(161) = txtPA161
            'Add by Amy 2022/06/17 CFP-29915 EPC案 母案改出名公司別,子案也要更新 (補 2017/11/24 改的不見了)
            If field(1) = "CFP" And field(4) = "00" And field(9) = "221" Then
                strExc(1) = "Update Patent Set PA161='" & txtPA161 & "' " & _
                                    "Where PA01='" & field(1) & "' And PA02='" & field(2) & "' And PA04<>'00' "
                cnnConnection.Execute strExc(1)
            End If
         End If
         'end 2012/9/6
         
         'Add By Sindy 2023/3/31 回存基本檔
         If txtPA61.Visible = True And txtPA61.Tag <> txtPA61 Then
            field(61) = txtPA61
         End If
         '2023/3/31 END
         
         strTxt(2) = GetPASQL(field())
         
      '服務
      Else
         field(79) = txtEngGroup 'Added by Morgan 2012/3/12
         field(9) = txtCaseField(2)
         '取消閉卷
         If Me.txtCaseField(6).Text = "Y" Then
            field(15) = Empty
            field(16) = Empty
            field(17) = Empty
         End If
         field(29) = txtCaseField(10)
         field(18) = txtCaseField(13)
         field(28) = txtCaseField(11)
         
         'Add by Amy 2017/07/13 回存基本檔
         If txtPA161.Visible = True And txtPA161.Tag <> txtPA161 Then
            field(85) = txtPA161
         End If
         'end 2017/07/13
         
         strTxt(2) = GetSPSQL(field())
      End If
      
      Pub_SeekTbLog strTxt(2) 'Added by Morgan 2023/12/18
      cnnConnection.Execute strTxt(2)
   
      'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之年費605、維持費606、延展費607，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。
      If txtCaseField(6) = "Y" Then
         strMsgCloseCancel = PUB_GetCaseCloseCancel(field(1), field(2), field(3), field(4), field(9))
      End If
      
      If bolNew Then
         intSaveMode = 1
      ElseIf cmdCountry.Visible Then
         intSaveMode = 2
      Else
         intSaveMode = 0
      End If
      
      'Remove by Morgan 2006/9/14 國內案號改由多國卷號維護
      
      For i = 1 To grdDataList.Rows - 1
         If grdDataList.TextMatrix(i, 0) = "V" Then
            'Modify by Morgan 2006/1/24 加NP01
            'Modified by Morgan 2020/12/23 +更新NP24
            strTxt(iStep) = "UPDATE NEXTPROGRESS SET NP06 = 'Y',NP24='" & cp(9) & "' WHERE NP22 = " & grdDataList.TextMatrix(i, 10) & " and np01='" & grdDataList.TextMatrix(i, 7) & "'"
            Pub_SeekTbLog strTxt(iStep) 'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業，若畫面勾選下一程序期限且存檔有上續辦Y的都寫Log以便事後能追蹤
            cnnConnection.Execute strTxt(iStep)
            iStep = iStep + 1
         End If
      Next
         
      '若有更改承辦人, 則更新ENG的核稿人
      If Me.txtCaseField(0).Text <> Me.txtCaseField(0).Tag Then
         'Modify By Sindy 2016/5/24
         '無完稿日時,清掉核稿人及判發人
         strExc(0) = "SELECT ep02 FROM EngineerProgress WHERE ep02='" & lblReceiveCode & "' and ep09 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "update EngineerProgress set ep04=null,ep40=null WHERE ep02='" & lblReceiveCode & "'"
            cnnConnection.Execute strSql
         Else
            '檢查是否有送判或判發
            If PUB_ChkEmpFlowExists(lblReceiveCode, EMP_送判) = False And _
               PUB_ChkEmpFlowExists(lblReceiveCode, EMP_判發) = False Then
               strSql = "update EngineerProgress set ep40=null WHERE ep02='" & lblReceiveCode & "'"
               cnnConnection.Execute strSql
            End If
         End If
'         'edit by nickc 2007/08/16 修正更新欄位
'         'strTxt(iStep) = "UPDATE ENGINEERPROGRESS SET EP03=(" & _
'            "SELECT PP04 FROM CASEPROGRESS,PROMOTERPROOFREADER WHERE CP09='" & lblReceiveCode & _
'            "' AND CP01=PP01(+) AND '" & Me.txtCaseField(0).Text & "'=PP02(+) AND CP10=PP03(+)) WHERE EP02='" & lblReceiveCode & "'"
'         strTxt(iStep) = "UPDATE ENGINEERPROGRESS SET EP04=(" & _
'            "SELECT PP04 FROM CASEPROGRESS,PROMOTERPROOFREADER WHERE CP09='" & lblReceiveCode & _
'            "' AND CP01=PP01(+) AND '" & Me.txtCaseField(0).Text & "'=PP02(+) AND CP10=PP03(+)) WHERE EP02='" & lblReceiveCode & "'"
'         cnnConnection.Execute strTxt(iStep)
         '2016/5/24 END
         
         'Add by Morgan 2009/7/30 依規則更新齊備日,承辦期限
         'UpdEp06BySameCase cp() 'Removed by Morgan 2021/2/5 移到下面(日本案若改要計件也能要更新齊備日)
         
         'Added by Morgan 2023/8/2
         '承辦人是外翻人員,分案時,若P案已完稿,則請系統自動發MAIL通知承辦人(系統會透過ST14轉發給品薇)
         If Left(txtCaseField(0).Text, 2) = "F5" And InStr(CaseMapOut, cp(10)) > 0 Then
            If txtCode(5) <> "" Then
               CFPMail2F5xx cp(1), cp(2), cp(3), cp(4), txtCaseField(0).Text
            'Added by Morgan 2023/9/19 無P案關聯，則分案時,系統自動發MAIL通知品薇,讓其知道有外翻案件要處理--郭
            Else
               strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & "無Ｐ案關聯，請處理外翻事宜"
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                  " VALUES ( '" & strUserNum & "','" & txtCaseField(0).Text & "',to_char(sysdate,'yyyymmdd')" & _
                  ",to_char(sysdate,'hh24miss'),'" & strExc(1) & "','如旨')"
               cnnConnection.Execute strSql, intI
            'end 2023/9/19
            End If
         End If
         'end 2023/8/2
      End If
      
      UpdEp06BySameCase cp() 'Added by Morgan 2021/2/5
      
      '有指定國家
      'Modify by Morgan 2007/12/24
      'If intSaveMode = 2 And strCountry <> "" Then
      If intSaveMode = 2 And strCountry <> strCountryOld Then
      'end 2007/12/24
         'Add by Morgan 2008/1/3 指定費若有改國家則只能新增
         If txtCaseField(7) = "215" Then
            PUB_UpdateCountry intCaseKind, field, strCountry, strCountryOld
         Else
         'end 2008/1/3
            'edit by nickc 2007/02/02 不用 dll 了
            'If objPublicData.DeleteCountry(0, intCaseKind, field()) Then
            If ClsPDDeleteCountry(0, intCaseKind, field()) Then
               'Modify by Morgan 2006/12/25
               'Call objPublicData.SaveCountry(0, intCaseKind, field(1) & field(2) & field(3) & field(4), cp(9), strCountry)
               Call PUB_SaveCountry(0, intCaseKind, field(1) & field(2) & field(3) & field(4), cp(9), strCountry)
               'end 2006/12/25
            End If
         End If
      End If
      
      '91.11.3 ADD BY SONIA 更新主張優先權期限
      'Modify by Morgan 2004/9/22 加 主張國內優先權 121
      If txtCaseField(7).Text = 主張優先權 Or txtCaseField(7).Text = "121" Then
         'Modify by Amyn 2014/04/15 加 strPriority5
         'Modify by Morgan 2007/4/25 加 strPriority4
         If ClsPDSavePriority(field(), strPriority1, strPriority2, strPriority3, strPriority4, strPriority5) Then
            strFirstPriDate = PUB_GetFirstPriDate(cp)
         End If
         'Modify by Morgan 2007/4/25 判斷非PCT案才要
         'If strFirstPriDate <> "" Then
         'Modified by Morgan 2020/7/29 PCT案 by Pass 除外 --玫音
         If strFirstPriDate <> "" And field(46) <> "Y" And PUB_IsPCTByPass(field(91)) = False Then
         
            'Modify by Morgan 2008/9/3 改呼叫共用函數
            '更新CFP優先權日相關期限
            PUB_UpdCfpDate1 field(1), field(2), field(3), field(4), False
         
'            'Modify by Moragn 2007/4/25
'            '主張一個以上優先權的固定為最早優先權日+6個月,另外主張或被主張的若為設計時也是+6個月,剩下的則為+12個月
'            ''發明或新型之主張優先權法定期限為最早優先權日+12月, 設計為最早優先權日+6月
'            'Modify by Morgan 2007/5/9 加控制主張多個優先權且有設計時才為早優先權日+6個月
'            'Add by Morgan 2006/4/18 後面國家本所=法定-14天，韓國012,法國203,瑞士205,比利時209,保加利亞226,芬蘭217,西班牙211,義大利204,波蘭222,冰島218,泰國019
'            'Add by Morgan 2007/4/3 日本新型
'            '92.5.27 ADD BY SONIA 同時更新新案之期限

         End If
      End If
      '91.11.3 END
      
      'Add by Morgan 2010/3/17
      '新穎性優惠期要同時更新優先權及申請程序的期限
      If txtCaseField(7) = "123" And txtCaseField(9) <> "" Then
         strExc(1) = DBDATE(txtCaseField(9))
         strExc(2) = PUB_GetWorkDay1(txtCaseField(4), True)
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
      
      'Remove by Morgan 2011/3/31
      '100/4/1 以後多國案草墨圖改要計件(草齊日=墨齊日)
      ''Add by Morgan 2004/12/7 若為多國案時草圖是否計件上'N' 2004/12/13墨圖是否計件上'N'
      'If cp(21) = "Y" And cp(107) = "" Then
      '   strSql = "Update EngineerProgress SET EP20='N',EP29='N' Where EP02='" & cp(9) & "'"
      '   cnnConnection.Execute strSql
      'Else
      '   'Add by Morgan 2004/12/3 若有收文日較早(<=)的國內案時草圖是否計件上'N'
      '   '2008/12/4 modify by sonia 加入繪圖未分案確認條件 CFP-021680
      '   'If PUB_IfCaseMapExist(cp) = True Then
      '   If PUB_IfCaseMapExist(cp) = True And cp(107) = "" Then
      '      strSql = "Update EngineerProgress SET EP20='N' Where EP02='" & cp(9) & "'"
      '      cnnConnection.Execute strSql
      '   End If
      '   '2004/12/3 end
      'End If
      'end 2011/3/31
      
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
      
      'Add by Morgan 2006/10/4 EPC回覆檢索報告分案時若有收文實體審查且承辦人為程序則改掛工程師
      If txtCaseField(7) = "218" And txtCaseField(2) = "221" And cp(27) = "" And txtCaseField(0) <> "" Then
         strExc(0) = "select cp09 from caseprogress,staff where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='416' and cp27 is null and st01(+)=cp14 and st03='P12'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSql = "update caseprogress set cp14='" & txtCaseField(0) & "' where cp09='" & RsTemp.Fields(0) & "'"
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2006/10/4
      
      'Add by Morgan 2007/5/23 加判斷新申請案或主張優先權且未發文
      If cp(27) = "" And (txtCaseField(7).Text = 發明申請 Or txtCaseField(7).Text = 新型申請 Or txtCaseField(7).Text = 設計申請 Or txtCaseField(7).Text = 主張優先權 Or txtCaseField(7).Text = "121") Then
         'Add by Morgan 2007/4/23 PCT案或有主張優先權時需掛實審期限
         If txtCaseField(15).Text <> "" Or Trim(strPriority2) <> "" Then
            '最早優先權日
            If Trim(strPriority2) <> "" Then
               stDate(10) = PUB_GetFirstPriDate2(strPriority2)
            'PCT優先權日
            ElseIf txtCaseField(16) <> "" Then
               stDate(10) = txtCaseField(16)
            'PCT申請日
            Else
               stDate(10) = txtCaseField(15)
            End If
            
            'Modify by Morgan 2008/9/3 改呼叫共用函數
            PUB_UpdCfpDate2 field(1), field(2), field(3), field(4), stDate(10), cp(9)
            
         End If
         'end 2007/4/23
         'Add by Morgan 2007/8/27  PCT進國家階段若有年費未繳且未收文時發Mail給慧汶並新增下一程序年費期限
         If txtCaseField(15) <> "" Then
            If PUB_CheckAnnuity(field(8), field(9), DBDATE(txtCaseField(15)), strNP07, strNP09) = True Then
               strExc(0) = "select * from caseprogress where cp01='" & field(1) & "' and cp02='" & field(2) & "'" & _
                  " and cp03='" & field(3) & "' and cp04='" & field(4) & "' and cp10='" & strNP07 & "' and cp57 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  m_bolAnnuityAlert = True
                  m_strAlertMsg = "且需繳交年費而未收文"
                  stDate(1) = field(1)
                  stDate(2) = field(9)
                  stDate(3) = strNP09
                  GetCtrlDT stDate
                  strNP08 = PUB_GetWorkDay1(stDate(0), True)
                  strSql = "update nextprogress set np08=" & strNP08 & ",np09=" & strNP09 & _
                     " where np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "'" & _
                     " and np06 is null and np07='" & strNP07 & "'"
                  cnnConnection.Execute strSql, intI
                  If intI = 0 Then
                     strNP22 = GetNextProgressNo()
                     strSql = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                        "VALUES ('" & cp(9) & "','" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strNP07 & "," & strNP08 & "," & strNP09 & ",'" & cp(13) & "'," & strNP22 & ") "
                     cnnConnection.Execute strSql, intI
                  End If
               End If
            End If
            'Added by Morgan 2012/8/22
            'PCT進EPC還要檢查實審及指定費是否已收文
            If txtCaseField(2) = "221" Then
               If PUB_ChkCPExist(cp, "416") = False Then
                  m_bolAnnuityAlert = True
                  m_strAlertMsg = m_strAlertMsg & "且實審未收文"
               End If
               If PUB_ChkCPExist(cp, "215") = False Then
                  m_bolAnnuityAlert = True
                  m_strAlertMsg = m_strAlertMsg & "且指定費未收文"
               End If
            End If
            'end 2012/8/22
         End If
         'end 2007/8/27
      End If
      'end 2007/5/23
      
      '2009/4/21 ADD BY SONIA 美專CIP,CA或分割或CPA(但限設計)案抓母案之領證或答辯期限
      '母案可能閉卷或收文後延期故只抓下一程序領證或答辯期限最大者,不管是否續辦
      'Modified by Morgan 2024/6/19 分割改下面共用函直接更新
      'If field(9) = 美國國家代號 And (txtCaseField(7).Text = "113" Or txtCaseField(7).Text = "122" Or txtCaseField(7).Text = "307" Or (txtCaseField(7) = "114" And Text1(21) = "3")) Then
      If field(9) = 美國國家代號 And (txtCaseField(7).Text = "113" Or txtCaseField(7).Text = "122" Or (txtCaseField(7) = "114" And Text1(21) = "3")) Then
      'end 2024/6/19
         'Modified by Morgan 2016/3/3 +126 期末拋棄
         'strSql = "SELECT MAX(NP08||NP09) FROM (" & _
               "SELECT NP08,NP09 FROM NEXTPROGRESS WHERE NP02='" & cp(1) & "' and NP03='" & cp(2) & "' and NP04='0' " & _
               "AND NP07 IN ('601','107') UNION " & _
               "SELECT NP08,NP09 FROM NEXTPROGRESS,DivisionCase WHERE DC01='" & cp(1) & "' and DC02='" & cp(2) & "' and DC03='" & cp(3) & "' and DC04='" & cp(4) & "' " & _
               "AND DC05=NP02 AND DC06=NP03 AND DC07=NP04 AND DC08=NP05 AND NP07 IN ('601','107','126')) "
         'Modified by Lydia 2016/08/26 +438 再考量試行計畫(AFCP 2.0)
         'Modified by Morgan 2022/6/16 +424 請求繼續審查 --玫音
         strSql = "SELECT MAX(NP08||NP09) FROM (" & _
               "SELECT NP08,NP09 FROM NEXTPROGRESS WHERE NP02='" & cp(1) & "' and NP03='" & cp(2) & "' and NP04='0' " & _
               "AND NP07 IN ('601','107','126','438','424') UNION " & _
               "SELECT NP08,NP09 FROM NEXTPROGRESS,DivisionCase WHERE DC01='" & cp(1) & "' and DC02='" & cp(2) & "' and DC03='" & cp(3) & "' and DC04='" & cp(4) & "' " & _
               "AND DC05=NP02 AND DC06=NP03 AND DC07=NP04 AND DC08=NP05 AND NP07 IN ('601','107','126','438','424')) "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If Not IsNull(RsTemp.Fields(0)) Then
               '接續案或分割案無期限直接更新
               If txtCaseField(4) = "" And txtCaseField(9) = "" Then
                  strSql = "UPDATE CASEPROGRESS SET CP06=" & TransDate(Mid(RsTemp.Fields(0), 1, 8), 2) & ", CP07=" & TransDate(Mid(RsTemp.Fields(0), 9, 8), 2) & " WHERE CP09='" & cp(9) & "'"
                  cnnConnection.Execute strSql
               '期限不同則先詢問是否更新
               ElseIf TransDate(Mid(RsTemp.Fields(0), 1, 8), 2) <> TransDate(txtCaseField(4), 2) Or TransDate(Mid(RsTemp.Fields(0), 9, 8), 2) <> TransDate(txtCaseField(9), 2) Then
                  If MsgBox("母案之本所期限為" & TransDate(Mid(RsTemp.Fields(0), 1, 8), 1) & "，法定期限為" & TransDate(Mid(RsTemp.Fields(0), 9, 8), 1) & "，是否要更新？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                     strSql = "UPDATE CASEPROGRESS SET CP06=" & TransDate(Mid(RsTemp.Fields(0), 1, 8), 2) & ", CP07=" & TransDate(Mid(RsTemp.Fields(0), 9, 8), 2) & " WHERE CP09='" & cp(9) & "'"
                     cnnConnection.Execute strSql
                  End If
               End If
            End If
         End If
      End If
      '2009/4/21 END
      
      '2011/3/28 ADD BY SONIA EPC之分割案抓母案之答辯期限
      '母案可能閉卷或收文後延期故只抓下一程序答辯期限最大者,不管是否續辦
      'Removed by Morgan 2024/6/19 改下面共用函直接更新
      'If field(9) = "221" And txtCaseField(7).Text = "307" Then
      '   strSql = "SELECT MAX(NP08||NP09) FROM NEXTPROGRESS,DivisionCase WHERE DC01='" & cp(1) & "' and DC02='" & cp(2) & "' and DC03='" & cp(3) & "' and DC04='" & cp(4) & "' " & _
      '            "AND DC05=NP02 AND DC06=NP03 AND DC07=NP04 AND DC08=NP05 AND NP07='107' "
      '   intI = 1
      '   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      '   If intI = 1 Then
      '      If Not IsNull(RsTemp.Fields(0)) Then
      '         '接續案或分割案無期限直接更新
      '         If txtCaseField(4) = "" And txtCaseField(9) = "" Then
      '            strSql = "UPDATE CASEPROGRESS SET CP06=" & TransDate(Mid(RsTemp.Fields(0), 1, 8), 2) & ", CP07=" & TransDate(Mid(RsTemp.Fields(0), 9, 8), 2) & " WHERE CP09='" & cp(9) & "'"
      '            cnnConnection.Execute strSql
      '         '期限不同則先詢問是否更新
      '         ElseIf TransDate(Mid(RsTemp.Fields(0), 1, 8), 2) <> TransDate(txtCaseField(4), 2) Or TransDate(Mid(RsTemp.Fields(0), 9, 8), 2) <> TransDate(txtCaseField(9), 2) Then
      '            If MsgBox("母案之本所期限為" & TransDate(Mid(RsTemp.Fields(0), 1, 8), 1) & "，法定期限為" & TransDate(Mid(RsTemp.Fields(0), 9, 8), 1) & "，是否要更新？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
      '               strSql = "UPDATE CASEPROGRESS SET CP06=" & TransDate(Mid(RsTemp.Fields(0), 1, 8), 2) & ", CP07=" & TransDate(Mid(RsTemp.Fields(0), 9, 8), 2) & " WHERE CP09='" & cp(9) & "'"
      '               cnnConnection.Execute strSql
      '            End If
      '         End If
      '      End If
      '   End If
      'End If
      'end 2024/6/19
      '2011/3/28 END
      
      'Added by Morgan 2018/3/28
      If txtCaseField(7).Text = "307" And cp(27) = "" Then
         st307Msg = PUB_Update307Ref(cp(9))
      End If
      'end 2018/3/28
     
'Removed by Morgan 2024/2/22 已經沒有限制一定要收款後送件，改以收款後送件的選項控制
'      'Add by Morgan 2010/3/17
'      '承辦為程序且該程序有未收款且新申請案已發文時,新增未收款無法發文紀錄
'      If Val(cp(16)) > 0 And Val(cp(79)) > 0 And txtCaseField(0) <> "" And Frame1.Visible = True Then
'         If GetStaffDepartment(txtCaseField(0)) = "P12" Then
'            strExc(0) = "select cp27 from caseprogress where cp01='" & cp(1) & "'" & _
'               " and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'" & _
'               " and cp10 in (" & NewCasePtyList & ")"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               If RsTemp(0) > 0 Then
'                  bolAddRec = True
'               End If
'            '中間接來的案子也要新增
'            Else
'               bolAddRec = True
'            End If
'
'            If bolAddRec Then
'               strSql = "update UndeliveredRec set UD04='" & txtCaseField(0) & "' where UD01='" & cp(9) & "' and UD02=" & strSrvDate(1)
'               cnnConnection.Execute strSql, intI
'               If intI = 0 Then
'                  strSql = "insert into UndeliveredRec(UD01,UD02,UD03,UD04) VALUES('" & cp(9) & "'," & strSrvDate(1) & ",'1','" & txtCaseField(0) & "')"
'                  cnnConnection.Execute strSql, intI
'               End If
'            End If
'         End If
'      End If
'end 2024/2/22
      
      'Added by Morgan 2012/9/12
      '新案若設公司別與已開收據不同時發Mail通知財務處及智權人員
      'If txtPA161.Visible = True And txtPA161.Tag <> txtPA161 And Left(cp(60), 1) = "E" Then
      If txtPA161.Visible = True And Left(cp(60), 1) = "E" Then
         'Modify By Sindy 2013/12/15
         'If strSrvDate(1) >= InvoiceStartDate Then
         If Val(m_CP31isYGetCP05) >= Val(InvoiceStartDate) Then
            strExc(0) = "select a0k01||DECODE(A0K32,'N','(暫不印)','Y','(待列印)') from acc0j0,acc0k0 where a0j01='" & cp(9) & "' and a0k01(+)=a0j13 and a0k11<>'" & IIf(txtPA161 = "T", "1", IIf(txtPA161 = "J", "J", "2")) & "'"
         Else
         '2013/12/15 END
            strExc(0) = "select a0k01||DECODE(A0K32,'N','(暫不印)','Y','(待列印)') from acc0j0,acc0k0 where a0j01='" & cp(9) & "' and a0k01(+)=a0j13 and a0k11<>'" & IIf(txtPA161 = "Y", "1", "2") & "'"
         End If
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
            'Modified by Lydia 2020/03/31 智慧所更名日
            'If Val(m_CP31isYGetCP05) >= Val(InvoiceStartDate) Then
            '   strExc(1) = "專利案 " & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " 設定為以" & IIf(txtPA161 = "T", "專利商標", IIf(txtPA161 = "J", "台一智權", "專利法律")) & "出名與收據 " & strExc(1) & " 的公司別不同，請更正！"
            'Else
            ''2013/12/15 END
             '  strExc(1) = "專利案 " & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " 設定為以專利" & IIf(txtPA161 = "Y", "商標", "法律") & "出名與收據 " & strExc(1) & " 的公司別不同，請更正！"
            'End If
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
      
      'Add by Morgan 2010/6/3
      If m_bolMail927Inform Then
         strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
         strExc(2) = strExc(1) & "(" & lblNation & ")...煩請補收文""其他翻譯""..."
         strExc(3) = strExc(1) & "(" & lblNation & "),此國主張優先權時要求檢附優先權基礎案直譯本,煩請補收文""其他翻譯"",費用請參考價目表,謝謝!"
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " VALUES ( '" & strUserNum & "','" & cp(13) & "',to_char(sysdate,'yyyymmdd')" & _
            ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','" & strExc(3) & "')"
         cnnConnection.Execute strSql, intI
      End If
      
      'Add by Morgan 2010/6/17
      '若已開請款單則換承辦人或核稿人時發Mail通知靜芳
      If cp(60) > "X" Then
         PUB_PointReAssignInform lblCaseCode, cp(60), txtCaseField(0).Tag, txtCaseField(0)
      End If
      
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
           'strSql = "UPDATE CaseProgress Set CP97=" & Val(txtCP97) & " Where CP09='" & lblReceiveCode.Caption & "' "
           If ExistCheck("EXVALUE", "EV01", lblReceiveCode.Caption, strExc(0), False) = True Then
               strSql = "UPDATE EXVALUE Set EV02=" & Val(txtCP97) & " Where EV01='" & lblReceiveCode.Caption & "' "
           Else
               strSql = "INSERT INTO EXVALUE (EV01,EV02) VALUES ('" & lblReceiveCode.Caption & "'," & Val(txtCP97) & ") "
           End If
           'end 2014/09/09
           Pub_SeekTbLog strSql '記錄修改log
           cnnConnection.Execute strSql
      End If
      'end 2014/09/05
      
      'Modify By Sindy 2016/4/13 抽出來變共用Func
      Call PUB_UpdRelationCaseFixEP(cp(1), cp(2), cp(3), cp(4), txtCaseField(7), lblCaseProperty)
      '2016/4/13 END
         
      'Added by Morgan 2017/9/12
      'EMail承辦工程師確認是否有關聯P案 --甄妮
      'Modify By Sindy 2023/3/31 若CFP有無關聯P案=N時,不用再發Mail給工程師 => + And txtPA61 = ""
      If m_bolAskEngInCase And txtCaseField(0) <> "" And txtPA61 = "" Then
         strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
         strExc(2) = "敬請回覆確認 " & strExc(1) & "(" & lblNation & ") 案是否有相同P案需建關聯？"
         strExc(3) = "收文號：" & cp(9) & vbCrLf & _
            "本所案號：" & strExc(1) & vbCrLf & _
            "案件名稱：" & IIf(field(5) = "", field(6), field(5)) & vbCrLf & _
            "案件性質：" & lblCaseProperty & vbCrLf & _
            "收文日：" & ChangeTStringToTDateString(txtCaseField(14)) & vbCrLf & _
            "智權人員：" & cp(13) & " " & lblSalesName & vbCrLf & _
            "承辦人：" & txtCaseField(0) & " " & lblPromoter
            
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
            " VALUES ( '" & strUserNum & "','" & txtCaseField(0) & "',to_char(sysdate,'yyyymmdd')" & _
            ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','" & strExc(3) & "')"
         cnnConnection.Execute strSql, intI
      End If
      'end 2017/9/12
      
      'Added by Morgan 2023/3/21 未發文美國新案IDS檢查(相關案是否已有OA)
      If cp(27) = "" And txtCaseField(0) <> "" And txtCaseField(2) = 美國國家代號 And Text1(21) = "1" And InStr("'101','113','114','122','307'", "'" & txtCaseField(7) & "'") > 0 Then
         PUB_NewUsCaseIdsChk cp(1), cp(2), cp(3), cp(4)
      End If
      'end 2023/3/21

      'Added by Morgan 2025/4/14
      '日本/德國的發明/新型且承辦人為F編號時自動收文檢視中說並預設承辦人為國內案工程師(若無則設為李柏翰經理)
      If (txtCaseField(2) = "011" Or txtCaseField(2) = "231") And cp(27) = "" Then
         If (txtCaseField(7) = "101" Or txtCaseField(7) = "102") And Left(txtCaseField(0), 1) = "F" Then
            strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='209'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               
               strExc(1) = "" '承辦人
               strExc(2) = CompWorkDay(3, strSrvDate(1)) '本所期限=系統日+2工作天
               strExc(4) = ""
               'Modified by Morgan 2025/6/17 +無國內案或國內案承辦人已離職則不掛承辦人，並發email提醒分案人員：「原承辦人已離職，請重新分案。」--玫音
               strExc(0) = "select cp14,cp27,pa09,cp10,st04 from casemap,caseprogress,patent,staff" & _
                  " where cm01='" & cp(1) & "' and cm02='" & cp(2) & "' and cm03='" & cp(3) & "' and cm04='" & cp(4) & "' and cm10='0'" & _
                  " and cp01(+)=cm05 and cp02(+)=cm06 and cp03(+)=cm07 and cp04(+)=cm08 and cp10 in (" & CaseMapIn & ") and st01(+)=cp14" & _
                  " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 order by pa09,cp10"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp("st04") = "1" Then
                     strExc(1) = "" & RsTemp("cp14") '國內案承辦人
                  Else
                     m_B209Msg = "國內案承辦人已離職，請重新分案。"
                  End If
                  'Modified by Morgan 2025/6/9 修正已發文判斷
                  'If IsNull(RsTemp("cp27")) > 0 Then '若國內案未發文則於發文時上齊備及所限
                  If IsNull(RsTemp("cp27")) Then
                     strExc(2) = ""
                  End If
                  'end 2025/6/9
               Else
                  m_B209Msg = "本案無國內案，請依接洽單註記分案。"
               End If
               
               'Removed by Morgan 2025/6/17 改彈訊息提醒--玫音
               'If strExc(1) = "" Then strExc(1) = "99050" '無國內案
               
               strExc(3) = AutoNo("B", 6)
               strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06" & _
                  ",CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP32) " & _
                  " SELECT cp01,cp02,cp03,cp04," & strSrvDate(1) & "," & CNULL(strExc(2), True) & _
                  ",'" & strExc(3) & "','209','90',cp12,cp13,'" & strExc(1) & "','N','N'" & _
                  " FROM caseprogress WHERE cp09='" & cp(9) & "'"
                  cnnConnection.Execute strSql, intI
               
               '上齊備日(國內案先上齊備管控，再視需要人工調整)
               If strExc(2) <> "" Then
                  strSql = "update engineerprogress set ep06=" & strSrvDate(1) & " where ep02='" & strExc(3) & "'"
                  cnnConnection.Execute strSql, intI
               End If
            End If
         End If
      End If
      'end 2025/4/14
   
   End If
   
   'Add by Amy 2022/10/17 +接洽單電子化
   'Modify by Amy 2022/11/16 +急件
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
   'end 2022/11/16
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
   
   cnnConnection.CommitTrans
   
   If st307Msg <> "" Then MsgBox st307Msg, vbInformation 'Added by Morgan 2018/3/28
   
   FormSave = True
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   If Err.NUMBER <> 0 Then
      MsgBox Err.Description
   End If
End Function

'Update:True 更新 False:新增
Private Function SubGetCPSQL(cp() As String, Optional UPDATE As Boolean = True) As String
 Dim i As Integer, strTmp As String, strTmp1 As String
   cp(5) = TransDate(cp(5), 2)
   cp(6) = TransDate(cp(6), 2)
   cp(7) = TransDate(cp(7), 2)
   
   If Not UPDATE Then cp(9) = AutoNo(Left(cp(9), 1), 6)
   
   cp(25) = TransDate(cp(25), 2)
   cp(27) = TransDate(cp(27), 2)
   ''910917 nick 單引號修正
   cp(38) = ChgSQL(cp(38))
   cp(41) = ChgSQL(cp(41))
   
   cp(46) = TransDate(cp(46), 2)
   cp(47) = TransDate(cp(47), 2)
   cp(48) = TransDate(cp(48), 2)
   '2010/2/12 MODIFY BY SONIA
   'cp(53) = TransDate(cp(53), 2)
   'cp(54) = TransDate(cp(54), 2)
   '2012/11/6 modify by sonia 加入601領證
   'Modify by Amy 2018/04/10 +612年費移作次年
   If (cp(1) = "P" Or cp(1) = "CFP") And _
      (cp(10) = "601" Or cp(10) = "605" Or cp(10) = "606" Or cp(10) = "607" Or cp(10) = "612" Or cp(10) = "908") Then
   Else
      cp(53) = TransDate(cp(53), 2)
      cp(54) = TransDate(cp(54), 2)
   End If
   '2010/2/12 END
   cp(57) = TransDate(cp(57), 2)
   
   cp(44) = ChangeCustomerL(cp(44))
   cp(55) = ChangeCustomerL(cp(55))
   cp(56) = ChangeCustomerL(cp(56))
   'Added by Morgan 2023/12/1
   cp(89) = ChangeCustomerL(cp(89)) '移轉申請人(讓與申請人)2
   cp(90) = ChangeCustomerL(cp(90)) '移轉申請人(讓與申請人)3
   cp(91) = ChangeCustomerL(cp(91)) '移轉申請人(讓與申請人)4
   cp(92) = ChangeCustomerL(cp(92)) '移轉申請人(讓與申請人)5
   cp(93) = ChangeCustomerL(cp(93)) '移轉人(讓與人)2
   cp(94) = ChangeCustomerL(cp(94)) '移轉人(讓與人)3
   cp(95) = ChangeCustomerL(cp(95)) '移轉人(讓與人)4
   cp(96) = ChangeCustomerL(cp(96)) '移轉人(讓與人)5
   'end 2023/12/1
   SubGetCPSQL = ""
      
   'Modify by Morgan 2005/11/22
   'For i = 1 To T_CP
   For i = 1 To UBound(cp)
      Select Case i
         '2005/9/15 MODIBY SONIA 加 CP61,CP62,CP63,CP87,CP88
         'Modify by Morgan 2005/6/29 加cp60不可回寫
         'Case 65, 66, 67, 68, 69, 70
         'Modified by Morgan 2013/5/17 +CP151
         'Modify By Sindy 2023/4/17 +, 16, 17, 18
         Case 60, 65, 66, 67, 68, 69, 70, 61, 62, 63, 87, 88, 151, 16, 17, 18
         '2011/8/25 ADD by sonia 加cp16~18,cp73~79(2011/8/17加在上面會造成CFP-024078自動發證未開收據前也不會回寫)
         'Modify By Sindy 2023/4/17 秀玲:應該是開著分案的畫面，同時開進度檔改規費；回分案畫面存檔時又存回原金額。
         '                          看一下CFP及P的分案，應該都不可以回寫CP16、CP17、CP18、CP60
         Case 73, 74, 75, 76, 77, 78, 79 'Mark:16, 17, 18,
            If cp(60) = "" Then
               SubGetCPSQL = SubGetCPSQL + "cp" + Format(i, "00") + "=" + CNULL(ChgSQL(cp(i))) + ","
               If Not UPDATE Then
                  strTmp = strTmp & "CP" & Format(i, "00") & ","
                  strTmp1 = strTmp1 & CNULL(ChgSQL(cp(i))) & ","
               End If
            End If
         '2011/8/25 END
         Case Else
            SubGetCPSQL = SubGetCPSQL + "cp" + Format(i, "00") + "=" + CNULL(ChgSQL(cp(i))) + ","
            If Not UPDATE Then
               strTmp = strTmp & "CP" & Format(i, "00") & ","
               strTmp1 = strTmp1 & CNULL(ChgSQL(cp(i))) & ","
            End If
      End Select
   Next
   
   If UPDATE Then
      SubGetCPSQL = Left(SubGetCPSQL, Len(SubGetCPSQL) - 1)
      SubGetCPSQL = "update caseprogress set " & SubGetCPSQL & " where cp09=" + CNULL(cp(9))
   Else
   
      strTmp = Left(strTmp, Len(strTmp) - 1)
      strTmp1 = Left(strTmp1, Len(strTmp1) - 1)
      SubGetCPSQL = "insert into caseprogress (" & strTmp & ") values (" & strTmp1 & ")"
   End If
End Function

Private Sub cmdPriority_Click()
    'Add by Amy 2023/01/06 此支進優先權表單改不是強制表單,故進入時畫面鎖住
    mdiMain.Enabled = False
    Me.Enabled = False
    'Modify by Amy 2014/04/14 加 strPriority5
    'Modify by Amy 2023/01/06 +表單名
    ModifyPriority strPriority1, strPriority2, strPriority3, field(8), , field(1) & field(2) & field(3) & field(4), field(9), , strPriority4, strPriority5, , Me
End Sub
'Modify by Morgan 2006/9/18
Private Sub ReadCM()
   Dim i As Integer
   
   For i = 4 To 7
      txtCode(i) = ""
      strCode(i) = ""
   Next
   'edit by nickc 2007/02/05 不用 dll 了
   'If obj003.GetCaseMap(strCode()) Then
   If Cls003GetCaseMap(strCode()) Then
      'Modified by Morgan 2023/10/27
      'If InStr(CaseMapOut, cp(10)) > 0 Then
      'Modified by Morgan 2025/3/14 +cp(31)="Y"
      If InStr(CaseMapOut, txtCaseField(7)) > 0 Or cp(31) = "Y" Then
      'end 2023/10/27
            For i = 4 To 7
               txtCode(i) = strCode(i)
            Next
      End If
   End If
End Sub
'Add by Morgan 2005/8/24 重讀期限
'Modify by Morgan 2009/5/11 +CP64
Private Sub ReadCP0607()
   Dim bolMsg As Boolean 'Added by Morgan 2012/3/30
   strSql = "select cp06,cp07,CP64 from caseprogress where cp09='" & cp(9) & "' and cp31='Y' and cp06>0"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         If txtCaseField(12) <> "" & .Fields("cp64") Then
            txtCaseField(12) = "" & .Fields("cp64")
            bolMsg = True
         End If
         If txtCaseField(4) <> TransDate(.Fields(0), 1) Then
            txtCaseField(4) = TransDate(.Fields(0), 1)
            bolMsg = True
         End If
         
         If "" & .Fields(1) <> "" Then
            If txtCaseField(9) <> TransDate(.Fields(1), 1) Then
               txtCaseField(9) = TransDate(.Fields(1), 1)
               bolMsg = True
               'Add by Morgan 2011/6/29
               If .Fields(1) < Val(strSrvDate(1)) Then
                  MsgBox "本案法限已過期請確認!!", vbExclamation
               End If
            End If
         End If
         
         If bolMsg = True Then
            MsgBox "期限及備註資料已更新！", vbExclamation
         End If
      End If
   End With
End Sub
'Add by Morgan 2006/8/25 '檢查多國案不可輸入國內案號，控制統一從多國案維護輸入
Private Sub InnerCaseControl()

   If cp(1) <> "" Then
      strExc(0) = "select 1 from caserelation where cr01='" & cp(1) & "' and cr02='" & cp(2) & "' and cr03='" & cp(3) & "' and cr04='" & cp(4) & "' and rownum<2"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtCode(4).Locked = True
         txtCode(5).Locked = True
         txtCode(6).Locked = True
         txtCode(7).Locked = True
      Else
         txtCode(4).Locked = False
         txtCode(5).Locked = False
         txtCode(6).Locked = False
         txtCode(7).Locked = False
      End If
   End If
End Sub
'Added by Morgan 2013/3/28
Private Sub SettxtCP147()
   '若案件屬性更改時重新預設複雜或特殊案件
   If Combo3.Tag <> "" Or (Combo3.Tag = "" And field(158) <> Left(Combo3, 1)) Then
      If Combo3.Tag <> Left(Combo3, 1) Then
         txtCP147 = GetCP147Default()
      End If
   End If
   Combo3.Tag = Left(Combo3, 1)
End Sub

'Added by Morgan 2013/3/29
Private Sub Combo3_Click()
   SettxtCP147
End Sub

'Added by Morgan 2013/3/29
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
   SettxtCP147 'Added by Morgan 2013/3/29
End Sub
'2010/10/29 End

Private Sub Form_Activate()
   'Add by Morgan 2005/8/24 重新讀取國內外案資料
   If m_bolCheckCM = True Then
      Call ReadCM
      'If txtCaseField(4) = "" Then 'Removed by Morgan 2012/3/30 期限若有更新還是重新讀取並提醒,否則會有覆蓋之虞
         Call ReadCP0607
      'End If
      m_bolCheckCM = False
   End If
   '2005/8/24 end
   
   If m_bolCheckCP21 = True Then
      'Modify by Morgan 2006/11/15 若有進"多國卷案號"則相關欄位都要重讀
      'strExc(0) = "select cp21 from caseprogress where cp09='" & cp(9) & "'"
      strExc(0) = "select cp21,cp29,cp48 from caseprogress where cp09='" & cp(9) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         cp(21) = "" & RsTemp.Fields("cp21")
         txtCaseField(3).Text = cp(21)
         'add by Morgan 2006/11/15
         cp(29) = "" & RsTemp.Fields("cp29")
         cp(48) = "" & RsTemp.Fields("cp48")
         'end 2006/11/15
      End If
      m_bolCheckCP21 = False
   End If
   
   'Remove by Morgan 2006/9/19 改由多國案卷號輸入
   'InnerCaseControl 'Add by Morgan 2006/8/25
   
   '只跑一次的程式段
   If Not bolIsRun Then
      bolIsRun = True
      If intNowReceive = intTotalReceive - 1 Then
         cmdOK(3).Visible = False
      End If
      ReadAllData
      blnOKtoShow = True
      
'Remove by Morgan 2010/6/11 郭已回收請作單
'      'Add by Morgan 2010/6/9
'      If field(57) <> "" And cp(27) = "" And cp(57) = "" Then
'         MsgBox "本案已結案閉卷，須與客戶再做進一步確認！"
'      End If

      'Added by Morgan 2019/11/28
      SetPCTbyPass
      If m_bolXCACase = True Then
         txtCaseField(15) = TransDate(field(10), 2)
         If m_bolPCTbyPass Then
            txtCaseField(16) = PUB_GetPCTPriDate(field(91))
            Text1(23) = PUB_GetPCTPriNo(field(91))
         Else
            txtCaseField(16) = ""
            Text1(23) = ""
         End If
      End If
      'end 2019/11/28
            
   End If

End Sub

'Added by Morgan 2019/11/27
Private Sub SetPCTbyPass()
   m_bolXCACase = False
   m_bolPCTbyPass = False
   If cp(3) = "0" And (txtCaseField(7) = "113" Or txtCaseField(7) = "122") Then
      m_bolXCACase = True
      If MsgBox("是否為 PCT 案 by pass？", vbYesNo + vbQuestion) = vbYes Then
         m_bolPCTbyPass = True
      End If
   End If
   SetPCTVisible
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim strCode(0 To TF_PA) As String
End Sub

Private Sub Form_Load()
Dim i As Integer, j As Integer
   
   MoveFormToCenter Me
   For i = 1 To frm050101_1.grdDataList.Rows - 1
          If frm050101_1.grdDataList.TextMatrix(i, 0) <> "" Then
             ReDim Preserve strReceiveCode(j)
             'Modified by Morgan 2018/11/6 加數量欄後移一欄
             strReceiveCode(j) = frm050101_1.grdDataList.TextMatrix(i, 3)
             j = j + 1
          End If
   Next
   intTotalReceive = j
   intNowReceive = 0
   bolIsRun = False
   bolLeave = False
   intLeaveKind = 1
   '920224 nick
   Nick920224Bol = False
   
   'Add by Morgan 2006/9/14
   '國內案改由多國案卷號維護
   txtCode(4).Locked = True
   txtCode(5).Locked = True
   txtCode(6).Locked = True
   txtCode(7).Locked = True
   cmdOK(4).Visible = False
   'end 2006/9/14
   'Add by Amy 2014/09/19 承辦人期限隱藏
   Label17(1).Visible = False
   txtCaseField(17).Enabled = False
   txtCaseField(17).Visible = False
   'end 2014/09/19
   lblCancelReceiveDate.BackColor = Me.BackColor
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If bolLeave = False Then
   If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
      Cancel = 1
   End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
PUB_SendMailCache 'Add by Morgan 2007/3/23
'Add by Amy 2023/01/06 frm880002從此支開啟改不為強制表單,故需判斷存在時要關
strPriority1 = "": strPriority2 = "": strPriority3 = "": strPriority4 = "": strPriority5 = ""
If intLeaveKind = 1 Then
   frm050101_1.Show
Else
   Unload frm050101_1
End If
'Add By Cheng 2002/07/18
'Set frm050101_2 = Nothing 'Removed by Morgan 2021/12/8 form2.0會有問題，先取消
End Sub

Private Sub lblPetition_Change(Index As Integer)
Dim strTemp As String

If lblPetition(Index) = "" Then lblPetitionName(Index) = "": Exit Sub 'Add by Morgan 2004/10/20

'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetCustomer(lblPetition(Index), strTemp) Then
If ClsPDGetCustomer(lblPetition(Index), strTemp) Then
   lblPetitionName(Index) = strTemp
Else
   lblPetitionName(Index) = ""
End If
End Sub
Private Sub lblSales_Change()
Dim strTemp As String

'edit by nickc 2007/02/02 不用 dll 了
'If objPublicData.GetStaff(lblSales, strTemp) Then
If ClsPDGetStaff(lblSales, strTemp) Then
   lblSalesName = strTemp
Else
   lblSalesName = ""
End If
End Sub



Private Sub Text1_Change(Index As Integer)
Select Case Index
'Add By Cheng 2002/06/11
Case 21 '專利種類
   If Me.Text1(Index).Enabled Then
      Me.Label3(11).Caption = "" & PUB_GetPatentKindName(Me.Text1(21).Text, 台灣國家代號)
      If Me.Text1(Index).Text <> "" Then
         If Me.Label3(11).Caption = "" Then
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
End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
TextInverse Me.Text1(Index)
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
   'Add By Cheng 2002/06/11
   Case 21 '專利種類
      If Me.Text1(Index).Enabled Then
         Me.Label3(11).Caption = "" & PUB_GetPatentKindName(Me.Text1(21).Text, 台灣國家代號)
         If Me.Label3(11).Caption = "" Then
            MsgBox "專利種類輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            Me.Text1(Index).SetFocus
         Else
            If (Me.txtCaseField(7).Text >= "101" And Me.txtCaseField(7).Text <= "103") Or _
               (Me.txtCaseField(7).Text >= "301" And Me.txtCaseField(7).Text <= "303") Then
               If Mid(Me.txtCaseField(7).Text, 3, 1) <> Me.Text1(Index).Text Then
                  MsgBox "專利種類必須與案件性質的第三碼相同!!!", vbExclamation + vbOKOnly
                  Cancel = True
                  Me.Text1(Index).SetFocus
               End If
            End If
         End If
      End If
End Select
End Sub

'Add by Morgan 2004/9/23
Private Sub SetEntity()
   Dim i As Integer
   OptChoose(0).Value = False: OptChoose(1).Value = False: OptChoose(2).Value = False
   
   'Removed by Morgan 2025/5/23 個體別順序會因國家有所不同,取消預設
   'For i = 1 To 5
   '   If txtAD(i).Text = "N" Then
   '      OptChoose(0).Value = True
   '      Exit For
   '   '只要有未設定減免身分的公司申請人則不預設大小個體
   '   ElseIf txtAD(i).Enabled = True And txtAD(i).Text = "" Then
   '      Exit For
   '   End If
   'Next
   ''若五個申請人檢查完都不是大個體則為小個體
   'If OptChoose(2).Enabled = False Then 'Added by Morgan 2013/3/20 不可選微個體時才預設
   '   If OptChoose(0).Value = False And i = 6 Then OptChoose(1).Value = True
   'End If
   'end 2025/5/23
   
End Sub
'Add by Morgan 2004/9/23
Private Sub txtAD_Change(Index As Integer)
   SetEntity
End Sub

'Add by Morgan 2004/9/22
Private Sub txtAD_GotFocus(Index As Integer)
   'edit by nickc 2007/06/06 切換輸入法改用API
   'txtAD(Index).IMEMode = 2
   CloseIme
   TextInverse txtAD(Index)
End Sub
'Add by Morgan 2004/9/22
'只有公司可輸入 Y,N
Private Sub txtAD_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If Not (KeyAscii = 8 Or KeyAscii = 89 Or KeyAscii = 78) Then
      KeyAscii = 0
   End If
End Sub
'Add by Morgan 2004/9/23
Private Sub txtAD_Validate(Index As Integer, Cancel As Boolean)
   If txtAD(Index) = "" Then
      MsgBox "請設定減免身分(Y/N)！"
      Cancel = True
   End If
End Sub
'Add by Morgan 2004/9/23
Private Sub SetAD(ByVal i As Integer)
   txtAD(i).Enabled = False
   txtAD(i).Tag = ""
   txtAD(i).Text = ""
   'Modify by Morgan 2005/4/13 控制專利才要
   If intCaseKind = 專利 Then
      'Modify by Morgan 2006/9/20 加法國且申請日>=2005/9/1
      'Modify by Morgan 2006/12/19 法國不管申請日都顯示
      'If field(i + 25) <> "" And (txtCaseField(2) = "101" Or txtCaseField(2) = "102" Or (txtCaseField(2) = "203" And (field(10) = "" Or DBDATE(field(10)) >= "20050901"))) Then
      'Modified by Morgan 2015/11/23 +印度040,菲律賓030--禧佩
      'Modified by Morgan 2023/3/24
      'If field(i + 25) <> "" And (txtCaseField(2) = "101" Or txtCaseField(2) = "102" Or txtCaseField(2) = "203" Or txtCaseField(2) = "040" Or txtCaseField(2) = "030") Then
      If field(i + 25) <> "" And InStr(CFP_ChkEntity, txtCaseField(2)) > 0 Then
      'end 2023/3/24
         txtAD(i).Text = PUB_GetAD03(field(i + 25), txtCaseField(2).Text)
         txtAD(i).Tag = txtAD(i).Text
         txtAD(i).Enabled = True
      End If
   End If
End Sub
'92.5.9 ADD BY SONIA
Private Sub txtCaseField_LostFocus(Index As Integer)
   Select Case Index
      Case 0
         strExc(0) = "SELECT ST03 FROM STAFF WHERE ST01='" & txtCaseField(Index) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Not IsNull(RsTemp.Fields("ST03")) And RsTemp.Fields("ST03") = "P12" Then
               txtCaseField(1) = "N"
            End If
         End If
      '92.10.21 add by sonia
      Case 7
         'Modify by Morgan 2004/7/28
         '提供前案資料(207)改要計件
         'Modify by Morgan 2004/10/14 加 121
         '94.1.31 MODIFY BY SONIA 郭說再加 604,702,704,705,909,920, 但取消 請求面詢
         '2010/1/6 modify by sonia 加938超頁費,939超項費
         'modify by sonia 2019/3/13 +123主張優惠期
         'Modified by Morgan 2024/10/24 取消 專利調查--郭
         'Modified by Lydia 2025/06/18 P及CFP分案時，預設N不計件管制：改用CasePropertyMap.CPM05控制
         'If txtCaseField(7) = 主張優先權 Or txtCaseField(7) = "121" Or txtCaseField(7) = "123" Or txtCaseField(7) = 補文件 _
         'Or txtCaseField(7) = "215" Or txtCaseField(7) = "216" Or txtCaseField(7) = 公開費 Or txtCaseField(7) = 變更 _
         'Or txtCaseField(7) = 延期 Or txtCaseField(7) = 申請優先權證明 Or txtCaseField(7) = 實體審查 Or txtCaseField(7) = "427" _
         'Or txtCaseField(7) = 領證及繳年費 Or txtCaseField(7) = 年費 Or txtCaseField(7) = 維持費 Or txtCaseField(7) = 延展費 _
         'Or txtCaseField(7) = 讓與 Or txtCaseField(7) = 回覆代理人 Or txtCaseField(7) = 調卷 _
         'Or txtCaseField(7) = 不續辦 Or txtCaseField(7) = 補收款 Or txtCaseField(7) = "913" Or txtCaseField(7) = "917" Or txtCaseField(7) = "938" Or txtCaseField(7) = "939" _
         'Or txtCaseField(7) = 核准 Or txtCaseField(7) = 核駁 Or txtCaseField(7) = 通知補文件 Or txtCaseField(7) = "1005" _
         'Or txtCaseField(7) = 通知申請案號 Or txtCaseField(7) = 通知修正 Or txtCaseField(7) = 通知補充說明 Or txtCaseField(7) = 通知提供前案 _
         'Or txtCaseField(7) = 通知要求選取 Or txtCaseField(7) = 通知公開 Or txtCaseField(7) = 通知公告 Or txtCaseField(7) = 檢索報告 _
         'Or txtCaseField(7) = 通知證書號數 Or txtCaseField(7) = 專利證書 Or txtCaseField(7) = 其他來函 Or txtCaseField(7) = "1908" _
         'Or txtCaseField(7) = "604" Or txtCaseField(7) = "702" Or txtCaseField(7) = "704" Or txtCaseField(7) = "705" Or txtCaseField(7) = "909" Or txtCaseField(7) = "920" _
         'Then
         If PUB_GetCPMbyCP10(field(1), txtCaseField(7), "cpm05") = "N" Then
         'end 2025/06/18
            txtCaseField(1) = "N"
         Else
            'Add by Morgan 2011/3/30
            If txtCaseField(1) = "N" Then
               If txtCaseField(1).Tag <> "N" Then 'Added by Morgan 2018/9/13 回答一次就好--禧佩
                  If MsgBox("此程序是否改為要計件？", vbYesNo + vbDefaultButton2) = vbYes Then
                     txtCaseField(1) = ""
                  Else
                     txtCaseField(1).Tag = "N" 'Added by Morgan 2018/9/13
                  End If
               End If
            Else
            'end 2011/3/30
               txtCaseField(1) = ""
            End If
         End If
      '92.10.21 end
      
   End Select
End Sub
'92.5.9 END

Private Sub txtCode_GotFocus(Index As Integer)
   TextInverse txtCode(Index)
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtCode_LostFocus(Index As Integer)
   Select Case Index
      Case 3 '轉本所案號
         If Not CheckKeyIn1 Then
            txtCode(0).SetFocus
         Else
            If (Me.txtCode(0).Text = "" And Me.txtCode(1) <> "") Or (Me.txtCode(0).Text <> "" And Len(txtCode(1)) < 6) Then
                MsgBox "轉本所案號輸入不完整!!!", vbExclamation + vbOKOnly
                Me.txtCode(0).SetFocus
                txtCode_GotFocus 0
                Exit Sub
            End If
            If Me.txtCode(0).Text <> "" And Me.txtCode(1).Text <> "" Then
               MsgBox "不可同時做" & """轉本所案號""" & "及" & """修改其他欄位""" & "，否則資料可能有不一致的情形產生!!!", vbExclamation + vbOKOnly
            End If
         End If
      Case 7 '與國內案號相同
         If Not CheckKeyIn2 Then txtCode(4).SetFocus
   End Select
End Sub

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If IsEmptyText(txtCode(0)) = False Then
            If txtCode(0) <> cp(1) Then
               Cancel = True
               MsgBox "轉本所案號必須與原本所案號之系統別相同 !", vbCritical
               txtCode_GotFocus 0
            End If
         End If
      
      Case 4, 5, 6, 7
'Modify by Morgan 2006/4/14 案件性質改用常數控制
'         Select Case txtCaseField(7)
'            'Modify by Morgan 2006/4/14 加113(CIP申請)
'            Case 發明申請, 設計申請, 新型申請, 聯合申請, 翻譯, "113"
         'Modified by Morgan 2025/3/13 +cp(31)="Y"
         If InStr(CaseMapOut, txtCaseField(7)) > 0 Or cp(31) = "Y" Then
               If Index = 4 Then
                  If txtCode(Index) <> "" Then
                     'Modify by Morgan 2007/11/20 不再限制P案
                     'If txtCode(Index) <> "P" Then
                     '   MsgBox "系統只可為 P !", vbCritical
                     '   Cancel = True
                     'End If
                  End If
               End If
'            Case Else
         Else
               If txtCode(Index) <> "" Then
                  MsgBox "此案件性質此欄不可輸入 !", vbCritical
                  txtCode(Index) = ""
               End If
'         End Select
         End If
'2006/4/14 end
   End Select
End Sub
Private Function CheckKeyIn1(Optional bolShowMsg As Boolean = True) As Boolean
Dim i As Integer, strAutoNumber As String

bolNew = False
If txtCode(0) <> "" And txtCode(1) <> "" Then
   CheckKeyIn1 = True
   '2008/10/23 modify by sonia
   'If bolShowMsg Then GetCaseDeadLineData grdDataList, intLastRow, txtCode(0), txtCode(1), txtCode(2), txtCode(3)
   If bolShowMsg Then GetGrid grdDataList, intLastRow, txtCode(0), txtCode(1), txtCode(2), txtCode(3)
   '2008/10/23 END
   CheckKeyIn 2
End If
If i = 4 Then
   CheckKeyIn1 = True
   '2008/10/23 modify by sonia
   'If bolShowMsg Then GetCaseDeadLineData grdDataList, intLastRow, cp(1), cp(2), cp(3), cp(4)
   If bolShowMsg Then GetGrid grdDataList, intLastRow, cp(1), cp(2), cp(3), cp(4)
   '2008/10/23 END
   CheckKeyIn 2
   Exit Function
End If
If txtCode(0) = cp(1) Then
   If txtCode(1) = cp(2) And IIf(txtCode(2) = "", "0", txtCode(2)) = cp(3) And IIf(txtCode(3) = "", "00", txtCode(3)) = cp(4) Then
      ShowMsg MsgText(9181)
      For i = 0 To 3
             txtCode(i) = ""
      Next
      txtCode_GotFocus 0
   Else
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.CheckCaseCodeIsExist(txtCode(0), txtCode(1), _
           IIf(txtCode(2) = "", "0", txtCode(2)), IIf(txtCode(3) = "", "00", txtCode(3)), , , , , , , , False) = False Then
      If ClsPDCheckCaseCodeIsExist(txtCode(0), txtCode(1), _
           IIf(txtCode(2) = "", "0", txtCode(2)), IIf(txtCode(3) = "", "00", txtCode(3)), , , , , , , , False) = False Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetMaxNumber(txtCode(0), strAutoNumber) = False Then
         If ClsPDGetMaxNumber(txtCode(0), strAutoNumber) = False Then
            Exit Function
         Else
            If Val(strAutoNumber) < Val(txtCode(1)) Then
               ShowMsg MsgText(9182)
               
'               For i = 0 To 3
'                      txtCode(i) = ""
'               Next
               CheckKeyIn1 = False
               Exit Function
            End If
         End If
         ShowMsg MsgText(9183) + vbLf + vbLf + txtCode(0) + "-" + txtCode(1) + IIf(txtCode(2) = "", "", "-" + txtCode(2)) + IIf(txtCode(3) = "", "", "-" + txtCode(3))
         bolNew = True
      End If
      If cmdCountry.Visible And bolShowMsg Then
         If MsgBox("因為轉本所案號所以將清除指定國家", vbInformation + vbOKCancel) = vbOK Then
            '2008/10/23 MODIFY BY SONIA
            'If bolShowMsg Then GetCaseDeadLineData grdDataList, intLastRow, txtCode(0), txtCode(1), txtCode(2), txtCode(3)
            If bolShowMsg Then GetGrid grdDataList, intLastRow, txtCode(0), txtCode(1), txtCode(2), txtCode(3)
            '2008/10/23 END
            cmdCountry.Visible = False
         Else
            For i = 0 To 3
                   txtCode(i) = ""
            Next
         End If
      End If
      CheckKeyIn1 = True
   End If
Else
   If txtCode(0) = "" Then
      CheckKeyIn1 = True
   Else
      ShowMsg MsgText(9173)
      txtCode(0).SetFocus
   End If
End If
End Function

Private Function CheckKeyIn2(Optional bolShowMsg As Boolean = True)

   'Add By Cheng 2003/09/10
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
      
   If txtCaseField(7) <> cp(10) Then Call ReadCM 'Added by Morgan 2023/10/27 若案件性質改要重新讀取國內案號
   
   'Modify by Morgan 2006/9/6  改判斷新案案件性質
'   Select Case txtCaseField(7)
'      Case 發明申請, 設計申請, 新型申請, 聯合申請
'         If txtCode(4) = "" Or txtCode(5) = "" Then
'            If MsgBox("與國內案號相同之欄位未輸入，是否繼續存檔", vbInformation + vbOKCancel) = vbCancel Then
'               CheckKeyIn2 = False
'               txtCode(4) = ""
'               txtCode(5) = ""
'               Exit Function
'            End If
'         End If
'   End Select
   m_bolAskEngInCase = False
   'Modified by Morgan 2017/12/20 分割(307),CIP(113)只會輸入母案 --甄妮
   'Modified by Morgan 2018/9/10 改分割案仍要鍵關聯--郭
   'If InStr(CaseMapOut, txtCaseField(7)) > 0 And txtCaseField(7) <> "307" And txtCaseField(7) <> "113" Then
   'Modified by Morgan 2020/10/12 排除已設定無關聯P案者 Ex:CFP-31981 (改承辦人時) --玫音
   'If InStr(CaseMapOut, txtCaseField(7)) > 0 And txtCaseField(7) <> "113" Then
   'Modified by Morgan 2025/2/25 +新案都要(中間轉進本所案件建關聯的提醒)--玫音
   If (InStr(CaseMapOut, txtCaseField(7)) > 0 Or cp(31) = "Y") And txtCaseField(7) <> "113" And field(61) = "" Then
   'end 2020/10/12
   'end 2018/9/10
      If txtCode(4) = "" Or txtCode(5) = "" Then
         If MsgBox("未與國內案建關聯，是否繼續存檔", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Function
         'Added by Morgan2017/9/12
         'Modified by Morgan 2017/12/20 北所分案才要--甄妮
         ElseIf pub_strUserOffice = "1" Then
            m_bolAskEngInCase = True
         End If
      End If
   End If
   'end 2006/9/6
   
   If txtCode(4) & txtCode(5) <> "" Then
'Remove by Morgan 2006/9/14 改在多國卷號維護控制
'      If txtCode(6) = "" Then txtCode(6) = "0"
'      If txtCode(7) = "" Then txtCode(7) = "00"
'      If cp(1) = "CFP" Then
'        strSQLA = "Select PA05, PA06, PA07 From Patent Where " & ChgPatent(Me.txtCode(4).Text & Me.txtCode(5).Text & Me.txtCode(6).Text & Me.txtCode(7).Text)
'        rsA.CursorLocation = adUseClient
'        rsA.Open strSQLA, cnnConnection, adOpenStatic, adLockReadOnly
'        If rsA.RecordCount > 0 Then
'            If MsgBox("國內相同案號案件名稱：" & vbCrLf & "(中)" & rsA.Fields(0).Value & vbCrLf & "(英)" & rsA.Fields(1).Value & vbCrLf & "(日)" & rsA.Fields(2).Value, vbExclamation + vbOKCancel) = vbCancel Then
'                CheckKeyIn2 = False
'                If rsA.State <> adStateClosed Then rsA.Close
'                Set rsA = Nothing
'                Exit Function
'            Else
'                CheckKeyIn2 = True
'            End If
'        Else
'            MsgBox "查無此本所案號資料!!!", vbExclamation + vbOKOnly
'            CheckKeyIn2 = False
'            If rsA.State <> adStateClosed Then rsA.Close
'            Set rsA = Nothing
'            Exit Function
'        End If
'        If rsA.State <> adStateClosed Then rsA.Close
'        Set rsA = Nothing
'      End If
   Else
      CheckKeyIn2 = True
      Exit Function
   End If
   
   intI = 1
   'Moidfy by Morgan 2004/12/6 加案件性質 113,114,201,307
   'Modify by Morgan 2005/12/14 加 109
   'Modify by Morgan 2006/4/14 案件性質改用常數控制
   'Modify by Morgan 2006/9/14 所有國內案都發文才提醒
   'strExc(0) = "SELECT CP27 FROM CASEPROGRESS WHERE CP01='" & txtCode(4) & "' AND CP02='" & txtCode(5) & "' AND CP03='" & txtCode(6) & "' AND CP04='" & txtCode(7) & "' AND CP10 IN (" & CaseMapIn & ")"
   strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN (SELECT CM05,CM06,CM07,CM08 FROM CASEMAP WHERE CM10='0' AND CM01='" & cp(1) & "' AND CM02='" & cp(2) & "' AND CM03='" & cp(3) & "' AND CM04='" & cp(4) & "') AND CP10 IN (" & CaseMapIn & ") AND CP27 IS NULL AND CP57 IS NULL"
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      CheckKeyIn2 = True
      'If Not IsNull(rsTemp.Fields(0)) Or rsTemp.Fields(0) <> "" Then
      If RsTemp.Fields(0) = 0 Then
         'Added by Morgan 2017/3/27 國內有新案發文才算 Ex.CFP-29342,P117009(只有收撰稿,申請人自己送件沒有收新案)
         strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE (CP01,CP02,CP03,CP04) IN (SELECT CM05,CM06,CM07,CM08 FROM CASEMAP WHERE CM10='0' AND CM01='" & cp(1) & "' AND CM02='" & cp(2) & "' AND CM03='" & cp(3) & "' AND CM04='" & cp(4) & "') AND CP10 IN (" & CaseMapIn & ") AND CP27>0 AND CP57 IS NULL"
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) > 0 Then
            'end 2017/3/27
            
               MsgBox "國內案件已送件，此案件可進行作業 !", vbExclamation
               
            'Added by Morgan 2017/3/27
            End If
         End If
         'end 2017/3/27
      End If
   Else
      MsgBox "無此國內案號，請重新輸入 !", vbCritical
      CheckKeyIn2 = False
   End If
   
End Function

Private Sub txtCaseField_Change(Index As Integer)
Select Case Index
             Case 0
                        lblPromoter = ""
             Case 2
                        'Modify by Morgan 2007/12/24
                        'If txtCaseField(Index) = EPC指定國家 Then
                        '   If txtCode(0) = "" Then
                        '      cmdCountry.Visible = True
                        '   End If
                        'Else
                        '   cmdCountry.Visible = False
                        'End If
                        SetCountryButton
                        'end 2007/12/24
                        
                        lblNation = ""
                        'Add by Morgan 2004/9/22
                        SetAD 1
                        SetAD 2
                        SetAD 3
                        SetAD 4
                        SetAD 5
                        '2004/9/22 end
                        SetPCTVisible 'Add by Morgan 2006/7/19
             Case 7
                        lblCaseProperty = ""
'                        txtCaseField(8) = ""
                        'Add by Morgan 2007/12/24
                        SetCountryButton
                        SetPCTVisible 'Added by Morgan 2020/7/29
End Select

End Sub

'Add by Morgan 2007/12/24
'EPC的發明申請且未發文才可以點指定國家按鈕
Private Sub SetCountryButton()
   'Modify by Morgan 2009/6/10 改不限制未發文,因為存檔時有控制只能新增了
   'If cp(27) = "" And txtCaseField(2) = EPC指定國家 And (txtCaseField(7) = 發明申請 Or txtCaseField(7) = "215") Then
   If txtCaseField(2) = EPC指定國家 And (txtCaseField(7) = 發明申請 Or txtCaseField(7) = "215") Then
      cmdCountry.Visible = True
   Else
      cmdCountry.Visible = False
   End If
End Sub

Private Sub txtCaseField_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
             Case 2
                        If CheckKeyIn(Index) <> -1 Then
                           CheckKeyIn 7
                        Else
                           Cancel = True
                        End If
             'Add by Amy 2018/10/18 智權人員非國外部FXX且修改案件性質時,不可改為 902(回覆代理人)
'Removed by Morgan 2020/3/11 移到 CheckKeyIn 否則原檢查會沒跑到
'             Case 7
'                If txtCaseField(7).Tag <> txtCaseField(7) And txtCaseField(7) = "902" Then
'                    If Left(PUB_GetStaffST15(lblSales, 1), 1) <> "F" Then
'                        Cancel = True
'                        MsgBox "智權人員非國外部，案件性質不可改為902(回覆代理人)"
'                        txtCaseField(7).SetFocus
'                    End If
'                End If
'end 2020/3/11
                
             'add by sonia 2019/8/20 CFP-029973
             Case 8
               If field(1) = "CFP" Then
                  If txtCaseField(Index) = "" Then
                     MsgBox "卷宗性質不可空白 !"
                     Cancel = True
                     txtCaseField(8).SetFocus
                  Else
                     If txtCaseField(7) = 異議_專 Then
                        If txtCaseField(Index) <> "2" Then
                           MsgBox "案件性質為異議時，卷宗性質必須為 2 !"
                           Cancel = True
                           txtCaseField(8).SetFocus
                        End If
                     ElseIf txtCaseField(7) = 舉發 Then
                        If txtCaseField(Index) <> "3" Then
                           MsgBox "案件性質為舉發時，卷宗性質必須為 3 !"
                           Cancel = True
                           txtCaseField(8).SetFocus
                        End If
                     End If
                  End If
               End If
             'end 2019/8/20
             'Added by Lydia 2017/05/05 客戶案件案號長度控制
             Case 10
                  'Modified by Lydia 2017/06/14 改常數
                  'Cancel = Not CheckLengthIsOK(txtCaseField(Index), 100)
                  Cancel = Not CheckLengthIsOK(txtCaseField(Index), 專利客戶案號max)
             'end 2017/05/05
             
             'add by nickc 2005/10/06
             Case 11
                  Cancel = Not CheckLengthIsOK(txtCaseField(11), txtCaseField(11).MaxLength)
             Case 12, 13
                        cmdOK(0).Default = True
                        cmdOK(0).CausesValidation = True
             Case Else
                        Dim nRet As Integer
                        nRet = CheckKeyIn(Index)
                        If nRet = -1 Then
                            Cancel = True
                        End If
                        ' 90.10.09 modify by louis
                        'Modify by Amy 2015/01/22 +index=0
                        If Index = 4 Or Index = 9 Or Index = 0 Then
                           If nRet = 0 Then
                              Cancel = True
                           End If
                        End If
End Select
If Cancel Then txtCaseField_GotFocus (Index)
End Sub

Private Function CheckKeyIn(intIndex As Integer) As Integer
Dim strTemp As String, strTemp1 As String, bolIsChina As Boolean, strCusTemp As String
Dim strOldValue As String
Dim strShow As String 'Add by Amy 2015/01/22

CheckKeyIn = -1
Select Case intIndex
             Case 0
                        m_CP14ST06 = "1" '2010/3/3 add by sonia
                        If txtCaseField(intIndex) <> "" Then
                           'edit by nickc 2007/02/02 不用 dll 了
                           'If objPublicData.GetStaff(txtCaseField(intIndex).Text, strTemp, strTemp1) Then
                           If ClsPDGetStaff(txtCaseField(intIndex).Text, strTemp, strTemp1) Then
                              lblPromoter = strTemp
                              CheckKeyIn = 1
                              m_CP14ST06 = PUB_GetST06(txtCaseField(intIndex))  '2010/3/3 add by sonia
                           End If
                           'Add by Morgan 2010/3/19
                           If PUB_GetST03(txtCaseField(intIndex)) = "P12" Then
                              txtCaseField(17).Enabled = True
                           Else
                              txtCaseField(17).Text = TransDate(cp(48), 1)
                              txtCaseField(17).Enabled = False
                           End If
                           'Add by Amy 2015/01/22
                           If txtCaseField(intIndex) <> m_strOldCP14 Then
                                If m_bolIsFirstKeyCP14 = True Then
                                     strShow = "與分所輸入之承辦人 " & GetStaffName(txtCaseField(intIndex).Tag) & " 不同，請再次輸入承辦人！"
                                     If CheckReKey(txtCaseField(intIndex), Label15(0), strShow) Then
                                         CheckKeyIn = 1
                                     Else
                                         CheckKeyIn = 0
                                         txtCaseField(intIndex) = cp(14)
                                     End If
                                 Else
                                     CheckKeyIn = 1
                                End If
                           Else
                                CheckKeyIn = 1
                           End If
                           'end 2015/01/22
                        Else
                           CheckKeyIn = 1
                        End If
             Case 1
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "N" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(1038)
                        End If
             Case 2
                        If txtCode(0) <> "" And cmdCountry.Visible Then
                           ShowMsg MsgText(9186)
                        'edit by nickc 2007/02/02 不用 dll 了
                        'ElseIf objPublicData.GetNation(txtCaseField(intIndex).Text, strTemp) Then
                        ElseIf ClsPDGetNation(txtCaseField(intIndex).Text, strTemp) Then
                           lblNation.Caption = strTemp
                           CheckKeyIn = 1
                        End If
             Case 3
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "Y" Then
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9174)
                        End If
             Case 4 '本所期限
                        If txtCaseField(intIndex) = "" Then
                           'Add/Modify By Cheng 2002/06/24
                           '若案件性質為"答辯"(107),"選取"(208),"公開費"(217),"延期"(404),"請求繼續審查"(424),"訴願"(501),"領證"(601),"年費"(605),"維持費"(606)
                           'Modify by Morgan 2011/2/16 +修正(204)
                           'Modified by Morgan 2016/3/3 +126 期末拋棄
                           'Modified by Lydia 2016/08/26 +438 再考量試行計畫(AFCP 2.0)
                           'modify by sonia 2020/2/5 +607延展費
                           'modify by sonia 2024/10/8 +408面詢
                           If Me.txtCaseField(7).Text = "107" Or _
                              Me.txtCaseField(7).Text = "126" Or _
                              Me.txtCaseField(7).Text = "204" Or _
                              Me.txtCaseField(7).Text = "208" Or _
                              Me.txtCaseField(7).Text = "217" Or _
                              Me.txtCaseField(7).Text = "404" Or _
                              Me.txtCaseField(7).Text = "424" Or _
                              Me.txtCaseField(7).Text = "501" Or _
                              Me.txtCaseField(7).Text = "601" Or _
                              Me.txtCaseField(7).Text = "605" Or _
                              Me.txtCaseField(7).Text = "606" Or _
                              Me.txtCaseField(7).Text = "607" Or _
                              Me.txtCaseField(7).Text = "408" Or _
                              Me.txtCaseField(7).Text = "438" Then
                              '檢查本所期限
                              If Me.txtCaseField(4).Text = "" Then
                                 MsgBox "本所期限不可空白!!!", vbExclamation + vbOKOnly
                                 If Me.txtCaseField(4).Enabled Then
                                    Me.txtCaseField(4).SetFocus
                                    TextInverse Me.txtCaseField(4)
                                 End If
                                 Exit Function
                              End If
                           Else
                              CheckKeyIn = 1
                           End If
                        Else
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                             If CheckReKey(txtCaseField(intIndex)) Then
                                '91.11.19 CANCEL BY SONIA
                                'If txtCaseField(intIndex) < strSrvDate(2) Then
                                '   ShowMsg MsgText(1032)
                                'Else
                                 CheckKeyIn = 1
                                'End If
                                '91.11.19 END
                             Else
                                CheckKeyIn = 0
                             End If
                            'Add By Cheng 2003/12/08
                            '若本所期限非工作天則直接調整至最近的工作天
                            If CheckKeyIn = 1 Then
                                Me.txtCaseField(intIndex).Text = TransDate(PUB_GetWorkDay1(Me.txtCaseField(intIndex).Text, True), 1)
                            End If
                           End If
                        End If
                        If txtCaseField(7) = 年費 Or txtCaseField(7) = 延期 Then
                           If txtCaseField(intIndex) = "" Then
                              ShowMsg Label17(0) & MsgText(9015)
                              CheckKeyIn = 2
                           End If
                        End If
             Case 5 '相關總收文
                        If txtCaseField(intIndex) <> "" Then
                           'Added by Morgan 2021/1/20 IDS的期限會來自其他案件
                           'Modified by Morgan 2023/3/23 EPC子案期限會來自母案 Ex:CFP-028997-0-40 商業使用聲明
                           If txtCaseField(7) = "214" Or cp(4) <> "00" Then
                              strExc(0) = "select np01 from nextprogress where np01='" & txtCaseField(intIndex) & "' and np02='" & cp(1) & "' and np03='" & cp(2) & "' and np04='" & cp(3) & "' and np05='" & cp(4) & "'"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 CheckKeyIn = 1
                              End If
                           End If
                           If CheckKeyIn <> 1 Then
                           'end 2021/1/20
                           
                              'edit by nickc 2007/02/02 不用 dll 了
                              'If objPublicData.GetRelationalReceiveCode(txtCaseField(intIndex), cp(1), cp(2), cp(3), cp(4)) Then
                              If ClsPDGetRelationalReceiveCode(txtCaseField(intIndex), cp(1), cp(2), cp(3), cp(4)) Then
                               'Add by Lydia 2014/11/24 PCT進各個國家階段之期限設定
                               'field(15)PCT申請日、field(16)PCT優先權日在txtValidate判斷
                                 Set414Date
                                 CheckKeyIn = 1
                              End If
                              
                           End If 'Added by Morgan 2021/1/20
                        Else
                           CheckKeyIn = 1
                        End If
    
             Case 6
                        If txtCaseField(intIndex) = "" Or txtCaseField(intIndex) = "Y" Then
                           If txtCaseField(intIndex) = "Y" Then CheckReKey txtCaseField(intIndex)
                           CheckKeyIn = 1
                        Else
                           ShowMsg MsgText(9174)
                        End If
             Case 7
                        'Added by Morgan 2020/3/11 自Validate移來
                        If txtCaseField(7).Tag <> txtCaseField(7) And txtCaseField(7) = "902" Then
                            If Left(PUB_GetStaffST15(lblSales, 1), 1) <> "F" Then
                                MsgBox "智權人員非國外部，案件性質不可改為902(回覆代理人)"
                                txtCaseField(7).SetFocus
                                Exit Function
                            End If
                        End If
                        'end 2020/3/11
                
                       If txtCaseField(2) = 大陸國家代號 Then bolIsChina = True Else bolIsChina = False
                       If ClsPDGetCaseProperty(cp(1), txtCaseField(intIndex), strTemp, bolIsChina) Then
                           lblCaseProperty.Caption = strTemp
                           If intCaseKind = 專利 Then
                              'cancel by sonia 2019/8/20 CFP-029973
                              'Dim strTmp As String
                              'Cls001SetPAFileProperty txtCaseField(7), strTmp, field(46)
                              'txtCaseField(8) = strTmp
                              'end 2019/8/20
                              If field(23) = "1" Then
                                 txtCaseField(8).Enabled = True
                              Else
                                 txtCaseField(8).Enabled = False
                              End If
                           Else
                              txtCaseField(8).Enabled = False
                           End If
                           CheckKeyIn = 1
                        End If
                        If txtCaseField(7) = "307" Then
                            DivVisibleSwitch True
                        Else
                            DivVisibleSwitch False
                        End If
                        SetPCTVisible 'Add by Morgan 2006/7/19
                        
                        'Add by Morgan 2010/3/16
                        If txtCaseField(intIndex) = "605" Or txtCaseField(intIndex) = "606" Or txtCaseField(intIndex) = "607" Then
                           If txtCaseField(intIndex) = "605" Then
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
                        
                        If SSTab1.Tab = 0 Then 'Added by Morgan 2018/12/6 在所在頁籤設定文字框座標時才檢查否則文字框也會顯示在目前頁籤
                           'Add by Morgan 2010/3/17
                           If txtCaseField(intIndex) = "123" Then
                            '  lblFavDt.Visible = True
                              txtCaseField(4).Left = 960
                              txtFavDt.Visible = True
                              CmdFav.Visible = True 'Add by Lydia 2015/02/02
                           Else
                            '  lblFavDt.Visible = False
                              txtCaseField(4).Left = 1320
                              txtFavDt.Visible = False
                              CmdFav.Visible = False 'Add by Lydia 2015/02/02
                           End If
                        End If 'Added by Morgan 2018/12/6

                        OptSendType(1).Caption = PUB_GetCP114Opt1Desc(cp(1), txtCaseField(intIndex))   'Added by Morgan 2024/1/22
                  
             Case 9 '法定期限
                        If txtCaseField(intIndex) <> "" Then
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              If Val(txtCaseField(4)) <= Val(txtCaseField(9)) Then
                                 If CheckReKey(txtCaseField(intIndex)) Then
                                    CheckKeyIn = 1
                                 Else
                                    CheckKeyIn = 0
                                 End If
                              Else
                                 ShowMsg MsgText(1033)
                              End If
                           End If
                        
                        'Modified by Morgan 2016/3/3 若期限為國內核准日計算來的除外
                        'ElseIf txtCaseField(4) <> "" Then
                        'Removed by Morgan 2020/3/23 有所限不一定要有法限 Ex:IDS
                        'ElseIf txtCaseField(4) <> "" And InStr(txtCaseField(12), "所限=核准日") = 0 Then
                        '   ShowMsg MsgText(1033)
                        '   CheckKeyIn = 0
                        'end 2020/3/23
                        Else
                           'Add/Modify By Cheng 2002/06/21
                           '若案件性質為"答辨"(107),"選取"(208),"公開費"(217),"延期"(404),"請求繼續審查"(424),"訴願"(501),"領證"(601),"年費"(605),"維持費"(606)
                           'Modified by Lydia 2016/08/26 +126 期末拋棄, +438 再考量試行計畫(AFCP 2.0)
                           'modify by sonia 2024/10/9 +408面詢
                           If Me.txtCaseField(7).Text = "107" Or _
                              Me.txtCaseField(7).Text = "208" Or _
                              Me.txtCaseField(7).Text = "217" Or _
                              Me.txtCaseField(7).Text = "404" Or _
                              Me.txtCaseField(7).Text = "424" Or _
                              Me.txtCaseField(7).Text = "501" Or _
                              Me.txtCaseField(7).Text = "601" Or _
                              Me.txtCaseField(7).Text = "605" Or _
                              Me.txtCaseField(7).Text = "606" Or _
                              Me.txtCaseField(7).Text = "126" Or _
                              Me.txtCaseField(7).Text = "408" Or _
                              Me.txtCaseField(7).Text = "438" Then
                              '檢查法定期限
                              If Me.txtCaseField(9).Text = "" Then
                                 MsgBox "法定期限不可空白!!!", vbExclamation + vbOKOnly
                                 If Me.txtCaseField(9).Enabled Then
                                    Me.txtCaseField(9).SetFocus
                                    TextInverse Me.txtCaseField(9)
                                 End If
                                 Exit Function
                              End If
                           Else
                              CheckKeyIn = 1
                              If txtCaseField(7) = 年費 Or txtCaseField(7) = 延期 Then
                                 If txtCaseField(intIndex) = "" Then
                                    ShowMsg Label17(2) & MsgText(9015)
                                    CheckKeyIn = 2
                                 End If
                              End If
                           End If
                        End If
             Case 14
                        'If cp(1) = "CFP" Then  91.7.24 CFP系統不管系統類別, 日期都是同一曆制
                           If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                              CheckKeyIn = 1
                           End If
                        'Else
                        '   If CheckIsTaiwanDate(txtCaseField(intIndex).Text) Then
                        '      CheckKeyIn = 1
                        '   End If
                        'End If
             'Modify by Morgan 2006/5/24
             '改輸PCT申請日&優先權日
             'Case 15 '是否為PCT案
             Case 15, 16
               If txtCaseField(intIndex).Visible = True Then
                  If txtCaseField(intIndex) = "" Then
                     CheckKeyIn = 1
                  ElseIf ChkDate(txtCaseField(intIndex)) = True Then
                     If CheckReKey(txtCaseField(intIndex)) Then
                        CheckKeyIn = 1
                     End If
                  End If
                  If CheckKeyIn = 1 Then
                     SetPCTDate
                  End If
               Else
                  CheckKeyIn = 1
               End If
               
            'Add by Morgan 2010/3/19
            Case 17 '承辦期限
               If txtCaseField(intIndex) <> "" Then
                  If ChkDate(txtCaseField(intIndex)) Then
                     txtCaseField(intIndex) = TransDate(PUB_GetWorkDay1(txtCaseField(intIndex), True), 1)
                      CheckKeyIn = 1
                  End If
               Else
                  CheckKeyIn = 1
               End If
         
             Case Else
                        CheckKeyIn = 1
End Select
End Function
Private Sub txtCaseField_GotFocus(Index As Integer)
txtCaseField(Index).SelStart = 0
txtCaseField(Index).SelLength = Len(txtCaseField(Index).Text)
'Remove by Morgan 2005/4/29 移到下面，改控制特定欄位才紀錄
'儲存未修改前之值至Tag中,供再確認時使用
'txtCaseField(Index).Tag = txtCaseField(Index)

Select Case Index
   'Add  by Amy 2015/01/22 +index 0
   Case 0
    m_strOldCP14 = txtCaseField(Index)
    
   'Add by Morgan 2005/4/29
   Case 4, 6, 9, 15
      txtCaseField(Index).Tag = txtCaseField(Index)
      
   Case 12, 13
      cmdOK(0).Default = False
      cmdOK(0).CausesValidation = False
End Select

End Sub
Private Sub txtCaseField_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
Select Case Index
            'Modify By Cheng 2002/03/06
'             Case 0, 1, 2, 3, 5, 6
             Case 0, 1, 2, 3, 5, 6, 15
                        KeyAscii = UpperCase(KeyAscii)
                        'Add By Cheng 2002/04/24
                        If Index = 6 Then
                           If KeyAscii <> 89 And KeyAscii <> 8 Then
                              KeyAscii = 0
                           End If
                        End If
End Select
End Sub
Private Sub grdDataList_GotFocus()
GridGotFocus grdDataList
End Sub
Private Sub grdDataList_LostFocus()
GridLostFocus grdDataList
End Sub
Private Sub grdDataList_KeyPress(KeyAscii As Integer)
If KeyAscii = 32 Then GrdDataList_Click
End Sub
Private Sub GrdDataList_Click()
   If grdDataList.Rows < 2 Then Exit Sub 'Added by Morgan 2016/2/3
   
   'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
   If Pub_CheckNpTheSameShow(cp(1), txtCaseField(7), Trim("" & grdDataList.TextMatrix(grdDataList.row, 11))) = False Then
       Exit Sub
   End If
   'end 2021/08/31
   
   'Add by Morgan 2009/2/26 領證分案點選過期年費時只勾選不更新資料
   If (txtCaseField(7) = "601" And grdDataList.TextMatrix(grdDataList.row, 11) = "605" And Val(Replace(grdDataList.TextMatrix(grdDataList.row, 3), "/", "")) < Val(strSrvDate(2))) Then
      If grdDataList.TextMatrix(grdDataList.row, 0) = "V" Then
         grdDataList.TextMatrix(grdDataList.row, 0) = ""
      Else
         grdDataList.TextMatrix(grdDataList.row, 0) = "V"
      End If
      Exit Sub
   End If
   'end 2009/2/26
   
   'Added by Morgan 2020/12/23
   '美國IDS點選後NP備註都要帶到CP備註(內容為IDS相關案號及國家)
   If txtCaseField(2) = "101" And txtCaseField(7) = "214" Then
      strExc(0) = txtCaseField(12)
      strExc(0) = Replace(strExc(0), grdDataList.TextMatrix(grdDataList.row, 9), "")
      If grdDataList.TextMatrix(grdDataList.row, 0) = "V" Then
         grdDataList.TextMatrix(grdDataList.row, 0) = ""

      Else
         grdDataList.TextMatrix(grdDataList.row, 0) = "V"
         strExc(0) = strExc(0) & " " & grdDataList.TextMatrix(grdDataList.row, 9)
         
         'Added by Morgan 2021/4/22 若原來無法限或晚於點選的期限時更新 Ex:CFP-31943
         strExc(1) = Replace(grdDataList.TextMatrix(grdDataList.row, 3), "/", "")
         If txtCaseField(9) = "" Or Val(strExc(1)) < Val(txtCaseField(9)) Then
            txtCaseField(9) = strExc(1)
            txtCaseField(4) = Replace(grdDataList.TextMatrix(grdDataList.row, 2), "/", "")
         End If
         'end 2021/4/22
      End If

      txtCaseField(12) = Trim(strExc(0))
      Exit Sub
   End If
   'end 2020/12/23
         
   If grdDataList.TextMatrix(grdDataList.row, 0) = "V" Then
      grdDataList.TextMatrix(grdDataList.row, 0) = ""
      txtCaseField(4) = ""
      txtCaseField(9) = ""
      txtCaseField(5) = ""
      txtCaseField(12) = ""
      m_CP30 = ""
   Else
      'Modify by Morgan 2009/12/23 延期只更新期限不可點選
      'grdDataList.TextMatrix(grdDataList.row, 0) = "V"
      If txtCaseField(7) <> "404" Then
         grdDataList.TextMatrix(grdDataList.row, 0) = "V"
      End If
      'end 2009/12/23
      
      txtCaseField(4) = Replace(grdDataList.TextMatrix(grdDataList.row, 2), "/", "")
      txtCaseField(9) = Replace(grdDataList.TextMatrix(grdDataList.row, 3), "/", "")
      txtCaseField(5) = grdDataList.TextMatrix(grdDataList.row, 7)
      txtCaseField(12) = grdDataList.TextMatrix(grdDataList.row, 9)
      'Add by Morgan 2011/4/22
      If grdDataList.TextMatrix(grdDataList.row, 10) = "0" Then
         m_CP30 = ""
      Else
         m_CP30 = grdDataList.TextMatrix(grdDataList.row, 10)
      End If
   End If

End Sub
Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      'ShowBar grdDataList, intLastRow, 6
      ShowBar grdDataList, intLastRow, 9
      blnOKtoShow = True
   End If
End If
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strTit As String    '20140211ADD By eric
Dim strMsg As String    '20140211ADD By eric
Dim rsTmp As ADODB.Recordset 'Add By Sindy 2014/5/22
Dim stInCNo(1 To 4) As String '國內案案號
   
   TxtValidate = False
   
   'Added by Morgan 2021/12/6 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   
   'Add by Morgan 2010/6/3
   m_bol106Chk = False
   m_bolMail927Inform = False
   'end 2010/6/3
   
   If (Me.txtCode(0).Text = "" And Me.txtCode(1) <> "") Or (Me.txtCode(0).Text <> "" And Me.txtCode(1).Text = "") Then
       MsgBox "轉本所案號輸入不完整!!!", vbExclamation + vbOKOnly
       Me.txtCode(0).SetFocus
       txtCode_GotFocus 0
       Exit Function
   End If

   If IsEmptyText(txtCode(0)) = False Then
      If txtCode(0) <> cp(1) Then
         Cancel = True
         MsgBox "轉本所案號必須與原本所案號之系統別相同 !", vbCritical
         txtCode_GotFocus 0
      End If
   End If

   '若非執行轉本所案號
   If Me.txtCode(0).Text = "" Or Me.txtCode(1).Text = "" Then
       For Each objTxt In Me.txtCaseField
          If objTxt.Enabled = True And objTxt.Visible = True Then
             Cancel = False
             'Modify by Amy 2015/01/22
             txtCaseField_Validate objTxt.Index, Cancel
             If Cancel = True Then
                Exit Function
             End If
          End If
       Next
          '92.02.24 nick 邱小姐後來說要檢查一定要輸入
          If cmdCountry.Visible = True Then
               If Trim(strCountry) = "" Then
                   MsgBox "指定國家一定要輸入！", vbCritical
                   cmdCountry.SetFocus
                   Exit Function
               End If
          End If
       
       For Each objTxt In Me.txtCode
          If objTxt.Enabled = True Then
             Cancel = False
             txtCode_Validate objTxt.Index, Cancel
             If Cancel = True Then
                Exit Function
             End If
          End If
       Next
       
       For Each objTxt In Me.Text1
          If objTxt.Enabled = True Then
             Cancel = False
             Text1_Validate objTxt.Index, Cancel
             If Cancel = True Then
                Exit Function
             End If
          End If
       Next
       
      'Add by Morgan 2004/9/23
      '檢查申請人減免身分
      For ii = 1 To 5
         If txtAD(ii).Enabled = True Then
            Cancel = False
            txtAD_Validate ii, Cancel
            If Cancel = True Then
               txtAD(ii).SetFocus
               Exit Function
            End If
         End If
      Next
      
       'Add By Cheng 2002/06/21
       '若案件性質為"答辨"(107),"選取"(208),"公開費"(217),"延期"(404),"請求繼續審查"(424),"訴願"(501),"領證"(601),"年費"(605),"維持費"(606)
       'Modified by Lydia 2016/08/26 +126 期末拋棄, +438 再考量試行計畫(AFCP 2.0)
       'modify by sonia 2024/10/9 +408面詢
       If Me.txtCaseField(7).Text = "107" Or _
          Me.txtCaseField(7).Text = "208" Or _
          Me.txtCaseField(7).Text = "217" Or _
          Me.txtCaseField(7).Text = "404" Or _
          Me.txtCaseField(7).Text = "424" Or _
          Me.txtCaseField(7).Text = "501" Or _
          Me.txtCaseField(7).Text = "601" Or _
          Me.txtCaseField(7).Text = "605" Or _
          Me.txtCaseField(7).Text = "606" Or _
          Me.txtCaseField(7).Text = "126" Or _
          Me.txtCaseField(7).Text = "408" Or _
          Me.txtCaseField(7).Text = "438" Then
          '檢查本所期限
          If Me.txtCaseField(4).Text = "" Then
             MsgBox "本所期限不可空白!!!", vbExclamation + vbOKOnly
             If Me.txtCaseField(4).Enabled Then
                Me.txtCaseField(4).SetFocus
                TextInverse Me.txtCaseField(4)
             End If
             Exit Function
          End If
          '檢查法定期限
          If Me.txtCaseField(9).Text = "" Then
             MsgBox "法定期限不可空白!!!", vbExclamation + vbOKOnly
             If Me.txtCaseField(9).Enabled Then
                Me.txtCaseField(9).SetFocus
                TextInverse Me.txtCaseField(9)
             End If
             Exit Function
          End If
       End If
       
       '92.11.20還原,P與CFP同時收文時先以P之收文日為優先權日,優先權號可不輸
       'Modify By Cheng 2002/12/09
       '取消控制, 因為若P與CFP同時收文尚無P之優先權號
   '    '91.11.3 add by sonia
      'Modify by Morgan 2004/10/14
       'If Me.txtCaseField(7).Text = 主張優先權 And strPriority2 = "" Then
       '20140211START Modify By eric 改為--詢問是否先輸入優先權資料,若是,則顯示優先權輸入畫面讓user輸入,否則繼續往下執行
       If (Me.txtCaseField(7).Text = 主張優先權 Or txtCaseField(7).Text = "121") And strPriority1 = "" Then
          strTit = "檢核資料"
          If Me.txtCaseField(7).Text = 主張優先權 Then
          
            strMsg = "案件性質為" & lblCaseProperty & ", 是否先輸入優先權資料? 若選擇「否」,系統於 7 天後會通知智權同仁補資料 "
            If MsgBox(strMsg, vbYesNo + vbQuestion, strTit) = vbYes Then
               'Modify by Amy 2014/04/14 加 strPriority5
               ModifyPriority strPriority1, strPriority2, strPriority3, field(8), , field(1) & field(2) & field(3) & field(4), field(9), , strPriority4, strPriority5
               Exit Function
               
            'Added by Morgan 2022/9/14
            ElseIf txtCaseField(4) = "" Then
               MsgBox "案件性質為" & lblCaseProperty & ", 若不輸入優先權資料則本所期限不可為空白！", vbExclamation
               SSTab1.Tab = 0
               txtCaseField(4).SetFocus
               Exit Function
            'end 2022/9/14
            End If
          Else
             'Modified by Morgan 2022/9/14
             'MsgBox "案件性質為主張優先權, 請先輸入優先權資料 !", vbExclamation + vbOKOnly
             MsgBox "案件性質為" & lblCaseProperty & ", 請先輸入優先權資料 !", vbExclamation + vbOKOnly
             'end 2022/9/14
             Me.txtCaseField(7).SetFocus
             TextInverse Me.txtCaseField(7)
             Exit Function
          End If
          
       End If
   '    If (Me.txtCaseField(7).Text = 主張優先權 Or txtCaseField(7).Text = "121") And strPriority1 = "" Then
   '       MsgBox "案件性質為主張優先權, 請先輸入優先權資料 !", vbExclamation + vbOKOnly
   '       Me.txtCaseField(7).SetFocus
   '       TextInverse Me.txtCaseField(7)
   '       Exit Function
   '    End If
       '20140211END
   '    '91.11.3 end
       
       'Add By Cheng 2003/08/13
       '若案件性質為延期, 則不可點選本案期限
       If Me.txtCaseField(7).Text = "404" Then
           For ii = 1 To Me.grdDataList.Rows - 1
               If Me.grdDataList.TextMatrix(ii, 0) <> "" Then
                   MsgBox "此案僅收文<延期>，不可點選下一程序期限資料，" & vbCrLf & "否則無法管制下一程序的期限!!!", vbExclamation + vbOKOnly
                   Exit Function
               End If
           Next ii
       End If
                       
      'Add by Morgan 2004/3/23
      '專利分案當有顯示是否取消閉卷但未輸入'Y'時，提示並取消！
      If txtCaseField(6).Visible = True And Trim(txtCaseField(6).Text) <> "Y" Then
         MsgBox "本案已閉卷且未取消!!!", vbCritical
         txtCaseField(6).SetFocus
         txtCaseField_GotFocus 6
         Exit Function
      End If
      
      'Add by Morgan 2006/5/24
      If txtCaseField(15).Visible = True Then
         If cp(27) = "" Then
            '檢查PCT資料
            If txtCaseField(15) <> "" Or txtCaseField(16) <> "" Then
               If txtCaseField(15) = "" Then
                  MsgBox "有PCT優先權日時PCT申請日不可空白！", vbExclamation
                  txtCaseField(15).SetFocus
                  Exit Function
               ElseIf Val(txtCaseField(16)) > Val(txtCaseField(15)) Then
                  MsgBox "PCT優先權日不可晚於PCT申請日！", vbExclamation
                  txtCaseField(16).SetFocus
                  Exit Function
               Else
                  SetPCTDate
               End If
            End If
            
            
            'Add by Morgan 2009/12/24
            'Modified by Morgan 2019/11/27 接續案除外
            'If txtCaseField(15) <> "" And Text1(23) = "" Then
            If txtCaseField(15) <> "" And Text1(23) = "" And m_bolXCACase = False Then
            'end 2019/11/27
               MsgBox "PCT案之PCT申請號不可空白！", vbExclamation
               Text1(23).SetFocus
               Exit Function
            End If
      
            'PCT 案必須要有期限
            If txtCaseField(7) = 發明申請 And txtCaseField(15) <> "" Then
               If txtCaseField(9) = "" Or txtCaseField(4) = "" Then
                  MsgBox "PCT 案必須要有期限 !", vbCritical
                  Exit Function
               End If
            End If
         End If
      End If
      'end 2006/5/24
      
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
      
      'Add by Morgan 2010/3/17
      If txtFavDt.Visible Then
         If txtFavDt.Text = "" Then
            MsgBox "新穎性優惠期日期不可空白！"
            txtFavDt.SetFocus
            Exit Function
         'Modified by Morgan 2018/3/16
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
      
      'Add by Morgan 2010/3/19
      If txtCaseField(17).Enabled And txtCaseField(17).Visible Then
         If txtCaseField(17) <> "" And txtCaseField(4) <> "" Then
            'Modify by Morgan 2010/8/11 百年蟲
            'If txtCaseField(17) > txtCaseField(4) Then
            If Val(txtCaseField(17)) > Val(txtCaseField(4)) Then
               MsgBox "承辦期限不可晚於本所期限！"
               txtCaseField(17).SetFocus
               Exit Function
            End If
         End If
      End If
      'Add by Morgan 2010/6/3
      '非美日德則詢問是否要直譯本，若要且未收文其他翻譯則Mail通知智權人員
      If txtCaseField(7) = "106" And InStr("101,011,231", txtCaseField(2)) = 0 Then
         If MsgBox("是否需直譯本？", vbYesNo) = vbYes Then
            m_bol106Chk = True
            If PUB_ChkCPExist(cp, "927") = False Then
               m_bolMail927Inform = True
            End If
         End If
      End If
      
      'Added by Morgan 2012/3/12 等舊資料補齊後改強制
      'Modified by Morgan 2012/3/19 改強制要輸入
      If txtEngGroup.Visible = True Then
         If txtEngGroup = "" Then
            MsgBox "國外部收文案件，工程師組別不可空白！", vbExclamation
            Exit Function
         End If
      End If
      
      'Added by Morgan 2012/7/20
      If txtCP147 <> txtCP147.Tag Then
         If txtCP147 <> GetCP147Default() Then
            Me.SSTab1.Tab = 1
            txtCP147.SetFocus
            If MsgBox("是否為複雜或特殊案件已變更為 [ " & IIf(txtCP147 = "Y", "是", "否") & " ] 與預設值不同是否確定要繼續?", vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            End If
         End If
      End If
      'end 2012/7/20
      
      'Add By Sindy 2014/5/22 CFP申請案
      If cp(1) = "CFP" And InStr(NewCasePtyList, Me.txtCaseField(7).Text) > 0 Then
         strExc(0) = "select cm05,cm06,cm07,cm08 from casemap" & _
                     " where cm01='" & cp(1) & "' and cm02='" & cp(2) & "'and cm03='" & cp(3) & "' and cm04='" & cp(4) & "'" & _
                     " and cm10='0'"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
         strExc(1) = ""
         If intI = 1 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
               stInCNo(1) = rsTmp.Fields("cm05")
               stInCNo(2) = rsTmp.Fields("cm06")
               stInCNo(3) = rsTmp.Fields("cm07")
               stInCNo(4) = rsTmp.Fields("cm08")
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
                  'Modify By Sindy 2014/11/12
                  If strSrvDate(1) >= 專利發明人檔啟用日 Then
                     strExc(0) = "select pi06 from patentinventor,inventor where pi01='" & stInCNo(1) & "' and pi02='" & stInCNo(2) & "' and pi03='" & stInCNo(3) & "' and pi04='" & stInCNo(4) & "' and substr(pi06,1,8)=in01(+) and substr(pi06,9,2)=in02(+) and in11='020'"
                  Else
                  '2014/11/12 END
                     'Memo by Lydia 2021/08/17 刪除舊程式碼：專利發明人在專利基本檔60~69
                  End If
                  intI = 1
                  Set adoRecordset = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If adoRecordset.RecordCount > 0 Then
                        strExc(1) = strExc(1) & vbCrLf & stInCNo(1) & stInCNo(2) & "-" & stInCNo(3) & "-" & stInCNo(4)
                     End If
                  End If
               End If
               rsTmp.MoveNext
            Loop
            If strExc(1) <> "" Then
               MsgBox "請與智權同仁確認大陸案是否要保密審查！" & strExc(1), vbInformation
            End If
         End If
         rsTmp.Close
         Set rsTmp = Nothing
      End If
      '2014/5/22 END
      
      'Added by Morgan 2019/11/27
      If m_bolXCACase Then
         If txtCaseField(15) = "" Then
            MsgBox Replace(Label3(1).Caption, "：", "") & "不可空白！", vbCritical
            txtCaseField(15).SetFocus
            Exit Function
         End If
      End If
      'end 2019/11/27
   End If
   
   'Add By Sindy 2010/10/29
   'Modify by Morgn 2010/11/2 加控制 101,102 才要
   'If Val(DBDATE(txtCaseField(14))) >= 20101102 And m_CP31 = "Y" And Left(PUB_GetStaffST15(Trim(m_CP13), 1), 1) <> "F"  Then
   If Val(DBDATE(txtCaseField(14))) >= 20101102 And m_CP31 = "Y" And Left(cp(12), 1) <> "F" And (txtCaseField(7) = "101" Or txtCaseField(7) = "102") Then
      If Combo3 = "" Then
         MsgBox "案件屬性不可空白！"
         Combo3.SetFocus
         Exit Function
      End If
   End If
   If Combo3 <> "" Then
      Cancel = False
      Combo3_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
      'Add by Morgan 2010/11/1
      If PUB_ChkRefCasePA158(cp(1), cp(2), cp(3), cp(4), Left(Combo3, 1)) = False Then
         Exit Function
      End If
   End If
   '2010/10/29 End

   'Add By Sindy 2013/12/16
   '檢查是否有客戶不開發票
   If Val(m_CP31isYGetCP05) >= Val(InvoiceStartDate) Then 'Add By Sindy 2014/1/29 +if
      If txtPA161.Visible = True And txtPA161 = "J" Then
         For ii = 1 To 5
            If Trim(field(26 + ii - 1)) <> "" Then
               'Modify By Sindy 2024/9/25 增加傳入公司別做判斷
               If PUB_ChkCU144isN(Left(ChangeCustomerL(field(26 + ii - 1)), 8), Right(ChangeCustomerL(field(26 + ii - 1)), 1), "", txtPA161, False) = True Then
                  MsgBox Left(ChangeCustomerL(field(26 + ii - 1)), 8) & Right(ChangeCustomerL(field(26 + ii - 1)), 1) & "此客戶為不開發票，因此特殊出名公司不可選智權公司 !", vbCritical
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
   
   'Added by Morgan 2013/3/20
   If OptChoose(2).Enabled = True Then
      If txtCaseField(2) = "101" And (txtCaseField(7) = "101" Or txtCaseField(7) = "103") Then
         If OptChoose(0).Value = False And OptChoose(1).Value = False And OptChoose(2).Value = False Then
            If SSTab1.Tab <> 0 Then SSTab1.Tab = 0
            MsgBox "CFP美國新申請案請點選個體別！", vbInformation
            Exit Function
         ElseIf OptChoose(2).Value = True Then
            For intI = 0 To 4
               If field(26 + intI) <> "" Then
                  If PUB_CheckMicroEntity(field(26 + intI), 1, 1, field(1) & field(2) & field(3) & field(4)) = False Then
                     If SSTab1.Tab <> 0 Then SSTab1.Tab = 0
                     Exit Function
                  End If
               End If
            Next
            For intI = 0 To 9
               If field(60 + intI) <> "" Then
                  If PUB_CheckMicroEntity(field(60 + intI), 3, 1, field(1) & field(2) & field(3) & field(4)) = False Then
                     If SSTab1.Tab <> 0 Then SSTab1.Tab = 0
                     Exit Function
                  End If
               End If
            Next
         End If
      End If
   End If
   'end 2013/3/20
   
   'Add by Amy 2014/09/05 若承辦加乘註記或承辦人基數有改則註記修改理由一定要輸
   If (Val(txtCP97) <> Val(cp(97)) Or Val(txtCP98) <> Val(cp(98))) And (txtCP99 = cp(99) Or txtCP99 = "") Then
        MsgBox "承辦加乘註記 或 承辦人基數有修改" & vbCrLf & "「註記修改理由」一定要輸入!", vbExclamation + vbOKOnly
        'SSTab1.Tab = 1
        'txtCP99.SetFocus
        'txtCP99_GotFocus
        Exit Function
   End If
   'end 2014/09/05
   
   'Added by Morgan 2018/12/6
   '美國正式申請案主張暫時申請案優先權時若暫時申請案有建其他多國案關聯時提醒User及王副總應改與正式申請案關聯
   'Modified by Morgan 2023/10/27
   'If field(9) = "101" And field(8) = "1" And cp(10) = "121" Then
   If txtCaseField(2) = "101" And Text1(21) = "1" And txtCaseField(7) = "121" Then
   'end 2023/10/27
      strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='101' and cp159=0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Modified by Morgan 2020/1/21 加檢查正式申請案沒有建國內案關聯才要--禧佩 Ex:CFP-31552
         strExc(0) = "select cm01||'-'||cm02||decode(cm03||cm04,'000','','-'||cm03||'-'||cm04) c1,cm05||'-'||cm06||decode(cm07||cm08,'000','','-'||cm07||'-'||cm08) c2" & _
            " from patent,caseprogress,casemap m1 where pa11 in ('" & Replace(ChgSQL(strPriority3), "，", "','") & "')" & _
            " and pa01='CFP' and pa09='101' and pa08='1' and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
            " and cp10='118' and cp159=0 and cm01(+)=pa01 and cm02(+)=pa02 and cm03(+)=pa03 and cm04(+)=pa04 and cm10='0'" & _
            " and not exists(select * from casemap m2 where m2.cm05=m1.cm05 and m2.cm06=m1.cm06 and m2.cm07=m1.cm07 and m2.cm08=m1.cm08" & _
            " and m2.cm01='" & cp(1) & "' and m2.cm02='" & cp(2) & "' and m2.cm03='" & cp(3) & "' and m2.cm04='" & cp(4) & "' and m2.cm10='0')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modified by Lydia 2023/04/24 修改王副總退休之相關控制
            'strExc(1) = "本案主張 " & RsTemp("c1") & "(暫時申請案) 國內優先權，該案目前與國內案 " & RsTemp("c2") & " 關聯。為免將來漏提IDS，國內案應改與正式申請案關聯！" & vbCrLf & vbCrLf & "請送回給王副總更改關聯後再分案！"
            If strSrvDate(1) >= "20230501" Then
                strExc(3) = "李經理"
            Else
                strExc(3) = "王副總"
            End If
            strExc(1) = "本案主張 " & RsTemp("c1") & "(暫時申請案) 國內優先權，該案目前與國內案 " & RsTemp("c2") & " 關聯。為免將來漏提IDS，國內案應改與正式申請案關聯！" & vbCrLf & vbCrLf & "請送回給" & strExc(3) & " 更改關聯後再分案！"
            'end 2023/04/24
            MsgBox strExc(1), vbExclamation
            Exit Function
         End If
      End If
   End If
   'end 2018/12/6
   
   'Added by Morgan 2021/1/26
   '多國子案正常應該不計件，若誤設要計件可能因國內案已作業而自動上齊備日，故加提醒 Ex:CFP-032119
   If txtCaseField(2) = "011" And txtCaseField(3) = "Y" And txtCaseField(1) = "" Then
      If MsgBox("日本多國子案是否確定要計件？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         txtCaseField(1).SetFocus
         Exit Function
      End If
   End If
   'end 2021/1/26
   
    m_IDSCP44 = "": m_IDSCP45 = "" 'Added by Morgan 2024/3/8
   
    'Added by Morgan 2023/12/13
    'CFP美國發明案=936(回覆委任代理人)或957(詢問代理人)內部收文/分案時，若曾發文IDS時必須要設相關總收文號以確保後續預設代理人時不會抓錯
    If field(1) = "CFP" And txtCaseField(2) = "101" And Text1(21) = "1" And (txtCaseField(7) = "936" Or txtCaseField(7) = "957") Then
      If txtCaseField(5) = "" Then
         strExc(0) = "select cp44 from caseprogress where cp01='" & field(1) & "' and cp02='" & field(2) & "' and cp03='" & field(3) & "' and cp04='" & field(4) & "' and cp10='214' and cp27>0 and cp159=0 order by cp27 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "CFP美國發明案有IDS已發文時，" & lblCaseProperty & "的【相關總收文號】不可空白！", vbExclamation
            txtCaseField(5).SetFocus
            Exit Function
         End If
      
      'Added by Morgan 2024/3/8 若相關收文號為IDS時將CF代理人及彼號帶入以便發文時能正確預設
      Else
         strExc(0) = "select cp44,cp45 from caseprogress where cp09='" & txtCaseField(5) & "' and cp10='214'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_IDSCP44 = "" & RsTemp("cp44")
            m_IDSCP45 = "" & RsTemp("cp45")
         End If
      'end 2024/3/8
      End If
    End If
    'end 2023/12/13
   
    'Added by Morgan 2024/1/5
    If Frame1.Visible And Frame1.Enabled And OptSendType(3).Value = True And txtCP142.Enabled = True Then
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
         
         If Frame2.Visible = True Then
            If Option1(0).Value = False And Option1(1).Value = False And Option1(2).Value = False Then
               MsgBox "有輸入指定日期，當天或之前或之後請擇一。", vbExclamation
               Exit Function
            End If
         End If
      End If
    End If
    'end 2024/1/5
   
   'Added by Morgan 2024/3/1
   '若子案已有進度時不可更改指定國家，否則案號會跟基本檔對不上(照理說發文後就不會再來重新指定)
'   If cmdCountry.Visible = True And strCountry <> strCountryOld Then
'      If txtCaseField(7) <> "215" Then
'         strExc(0) = "select * from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04<>'00'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(inst, strex(0))
'         If intI = 1 Then
'            MsgBox "已有子案進度資料，不可更改指定國家！", vbCritical
'            Exit Function
'         End If
'      End If
'   End If
   'end 2024/3/1
   
   'Added by Morgan 2025/10/22
   '外翻編號要檢查薪資的公司別與案件是否一致
   If Left(txtCaseField(0), 1) = "F" Then
      If txtPA161 <> "" Then
         strExc(1) = txtPA161
      Else
         strExc(1) = "" & PUB_GetReceiptComp(cp(1), cp(2), cp(3), cp(4), True)
      End If
      strExc(0) = "select sd19,st26 from salarydata,staff where sd01='" & txtCaseField(0) & "' and st01(+)=sd01"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) <> strExc(1) Then
            strExc(2) = txtCaseField(0) & "所屬公司別(" & RsTemp(0) & ")與案件出名公司(" & strExc(1) & ")不同"
            strExc(0) = "select st01 from staff,salarydata where st26='" & RsTemp(1) & "' and st04='1' and sd01(+)=st01 and sd19='" & strExc(1) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox strExc(2) & "，將自動更改為" & RsTemp(0) & "。", vbInformation
               txtCaseField(0) = RsTemp(0)
               txtCaseField_Validate 0, Cancel
               If Cancel = True Then
                  txtCaseField(0).SetFocus
                  Exit Function
               End If
            Else
               MsgBox strExc(2) & "且該翻譯人員尚無" & strExc(1) & "公司的編號，請先建檔！", vbCritical
               Exit Function
            End If
         End If
      End If
   End If
   'end 2025/10/22
   
   TxtValidate = True
   
End Function

'Add By Cheng 2003/08/12
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

'Add By Cheng 2003/08/05
'若為分所員工, 則寄給郭雅娟(79075)
Private Function ReGetStaffCode(StrST01 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

ReGetStaffCode = StrST01
StrSQLa = "Select ST06 From Staff Where ST01='" & StrST01 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    If "" & rsA.Fields(0).Value <> "1" Then
        ReGetStaffCode = "79075"
    End If
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'Add By Cheng 2003/10/21
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

'Add by Morgan 2004/3/29
'讀取分割案資料
Private Function GetDivCase() As Boolean

   Dim stSQL As String, rsQuery As String
   
On Error GoTo flgErr

   stSQL = "SELECT DC05, DC06, DC07, DC08 FROM DIVISIONCASE" & _
      " WHERE DC01='" & field(1) & "' AND DC02='" & field(2) & "' AND DC03='" & field(3) & "' AND DC04='" & field(4) & "'"
      
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
   If txtCaseField(7) = "307" Then
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

'Add by Amy 2014/09/05
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

'Add by Morgan 2005/3/9
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

Private Sub txtCP99_Validate(Cancel As Boolean)
   Cancel = Not CheckLengthIsOK(txtCP99, txtCP99.MaxLength)
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
      If (txtDivCaseNo(1) <> "" Or txtDivCaseNo(2) <> "" Or txtDivCaseNo(3) <> "" Or txtDivCaseNo(4) <> "") Then CheckDivCase
   End If
End Sub
'Add by Morgan 2004/3/30
'檢查母案是否存在
Private Function CheckDivCase() As Boolean
Dim dc(4) As String   '2009/8/24 add by sonia

On Error GoTo flgErr

   Dim stSQL As String, rsQuery As New ADODB.Recordset, stPA08 As String, stPA09 As String
   
   If (txtDivCaseNo(1) = "" Or txtDivCaseNo(2) = "") Then
      MsgBox "分割母案本所案號輸入錯誤！", vbExclamation
      Exit Function
   End If
   
   txtDivCaseNo(1) = Trim(txtDivCaseNo(1))
   txtDivCaseNo(2) = Right("00000" & txtDivCaseNo(2), 6)
   txtDivCaseNo(3) = Right("0" & txtDivCaseNo(3), 1)
   txtDivCaseNo(4) = Right("00" & txtDivCaseNo(4), 2)
   
   'Add by Morgan 2004/4/29
   If (txtDivCaseNo(1) = field(1) And txtDivCaseNo(2) = field(2) And txtDivCaseNo(3) = field(3) And txtDivCaseNo(4) = field(4)) Then
      MsgBox "分割案不可為母案！", vbExclamation
      Exit Function
   End If
   
   stSQL = "select PA08, PA09 from patent where pa01='" & ChgSQL(txtDivCaseNo(1)) & "' and pa02='" & ChgSQL(txtDivCaseNo(2)) & "' and  pa03='" & ChgSQL(txtDivCaseNo(3)) & "' and pa04='" & ChgSQL(txtDivCaseNo(4)) & "'"
   
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      stPA08 = "" & rsQuery.Fields(0)
      stPA09 = "" & rsQuery.Fields(1)
      'Add by Morgan 2004/4/29
      '分割案與母案的申請國家和專利種類需相同
      If stPA09 <> field(9) Then
         MsgBox "分割案與母案的申請國家需相同！", vbExclamation
      ElseIf stPA08 <> Text1(21) Then
         MsgBox "分割案與母案的專利種類需相同！", vbExclamation
      Else
         CheckDivCase = True
         '2009/8/24 add by sonia 分割案未輸入優先權時帶母案優先權
         '因發文及申請案號時都要再輸一次,故再考慮是否要做
         'If txtCaseField(7) = "307" And strPriority1 = "" Then
         '   dc(1) = txtDivCaseNo(1): dc(2) = txtDivCaseNo(2): dc(3) = txtDivCaseNo(3): dc(4) = txtDivCaseNo(4)
         '   If ClsPDReadPriority(dc(), strPriority1, strPriority2, strPriority3, strPriority4) = False Then GoTo flgErr
         'End If
         '2009/8/24 end
      End If
   Else
      MsgBox "分割母案本所案號不存在！", vbExclamation
   End If
   
flgErr:
   Set rsQuery = Nothing
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
    
End Function
'Add by Morgan 2005/6/10
'檢查是否為多國
Private Function CheckMutiNation() As Boolean

   '檢查其他相同國內案之國外案是否有多國案
   strSql = "SELECT 1 FROM CASEMAP,CASEPROGRESS" & _
      " WHERE CM10='0' AND CM05='" & txtCode(4) & "' AND CM06='" & txtCode(5) & "'" & _
      " AND CM07='" & txtCode(6) & "' AND CM08='" & txtCode(7) & "' AND NOT (CP01='" & cp(1) & "' AND CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "')" & _
      " AND CP01(+)=CM01 AND CP02(+)=CM02 AND CP03(+)=CM03 AND CP04(+)=CM04 AND CP01='CFP' AND ROWNUM<2"
   
On Error GoTo ErrHnd

   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         CheckMutiNation = True
      End If
   End With
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   
End Function
'Add by Morgan 2005/6/10
'新增多案相關
Private Function InsCR() As Integer
   Dim iEffRecs As Integer
   strSql = "insert into caserelation(CR01,CR02,CR03,CR04,CR05,CR06,CR07,CR08)" & _
      " select '" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "',CM01,CM02,CM03,CM04" & _
      " FROM casemap WHERE CM10='0'" & _
      " AND CM05='" & txtCode(4) & "' AND CM06='" & txtCode(5) & "'" & _
      " AND CM07='" & txtCode(6) & "' AND CM08='" & txtCode(7) & "'" & _
      " AND CM01='CFP'" & _
      " AND NOT (CM01='" & cp(1) & "' AND CM02='" & cp(2) & "' AND CM03='" & cp(3) & "' AND CM04='" & cp(4) & "')" & _
      " AND NOT EXISTS(SELECT * FROM CASERELATION" & _
      " WHERE CR01='" & cp(1) & "' AND CR02='" & cp(2) & "' AND CR03='" & cp(3) & "' AND CR04='" & cp(4) & "'" & _
      " AND CR05=CM01 AND CR06=CM02 AND CR07=CM03 AND CR08=CM04)" & _
      " UNION ALL select CM01,CM02,CM03,CM04,'" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'" & _
      " FROM casemap WHERE CM10='0'" & _
      " AND CM05='" & txtCode(4) & "' AND CM06='" & txtCode(5) & "'" & _
      " AND CM07='" & txtCode(6) & "' AND CM08='" & txtCode(7) & "'" & _
      " AND CM01='CFP'" & _
      " AND NOT (CM01='" & cp(1) & "' AND CM02='" & cp(2) & "' AND CM03='" & cp(3) & "' AND CM04='" & cp(4) & "')" & _
      " AND NOT EXISTS(SELECT * FROM CASERELATION" & _
      " WHERE CR05='" & cp(1) & "' AND CR06='" & cp(2) & "' AND CR07='" & cp(3) & "' AND CR08='" & cp(4) & "'" & _
      " AND CR01=CM01 AND CR02=CM02 AND CR03=CM03 AND CR04=CM04)"
   
   cnnConnection.Execute strSql, iEffRecs
   InsCR = iEffRecs
   
End Function
'Add by Morgan 2006/5/24
'計算PCT期限
Private Sub SetPCTDate(Optional ByVal p_Msg As Boolean = False)
   '未發文才要算期限
   If cp(27) = "" Then
      If m_bolXCACase = False Or m_bolPCTbyPass = True Then 'Added by Morgan 2019/11/27 PCT接續案也適用
         '優先權日
         If txtCaseField(16) <> "" And txtCaseField(15) <> "" Then
            strExc(0) = txtCaseField(16)
         ElseIf txtCaseField(15) <> "" Then
            strExc(0) = txtCaseField(15)
         Else
            strExc(0) = ""
         End If
         If strExc(0) <> "" Then
         'Add by Lydia 2014/11/24 PCT進各個國家階段之期限設定
            'PUB_GetPCTLimit strExc(0), strExc(1), strExc(2), txtCaseField(2)
            PUB_GetPCTLimit strExc(0), strExc(1), strExc(2), cp(1), cp(2), cp(3), cp(4), txtCaseField(2)
            If (txtCaseField(4) <> "" Or txtCaseField(9) <> "") And (txtCaseField(9) <> strExc(1) Or txtCaseField(4) <> strExc(2)) Then
               If p_Msg = True Then
                  'Modified by Morgan 2019/11/27
                  'If MsgBox("本案為PCT案，是否更新期限？(法限：" & strExc(1) & "；" & "所限：" & strExc(2) & ")", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                  If MsgBox("本案為PCT案" & IIf(m_bolPCTbyPass, " by pass", "") & "，是否更新期限？(法限：" & strExc(1) & "；" & "所限：" & strExc(2) & ")", vbQuestion + vbYesNo + vbDefaultButton1) = vbNo Then
                  'end 2019/11/27
                     Exit Sub
                  End If
               End If
            End If
            txtCaseField(9) = strExc(1)
            txtCaseField(4) = strExc(2)
         End If
      End If 'Added by Morgan 2019/11/27
   End If
End Sub

'Add by Morgan 2006/7/19 PCT進國家階段可請發明或新型
Private Sub SetPCTVisible()
   Label3(1).Caption = "PCT申請日：" 'Added by Morgan 2019/11/27
   If (txtCaseField(7) = "101" Or txtCaseField(7) = "102") And txtCaseField(2) <> "056" Then
      Label3(1).Visible = True: txtCaseField(15).Visible = True
      Label3(2).Visible = True: txtCaseField(16).Visible = True
      Label1(22).Visible = True: Text1(23).Visible = True 'Added by Morgan 2020/7/29
   'Added by Morgan 2019/11/27
   ElseIf cp(3) = "0" And (txtCaseField(7) = "113" Or txtCaseField(7) = "122") Then
      m_bolXCACase = True
      Label3(1).Visible = True: txtCaseField(15).Visible = True
      If m_bolPCTbyPass Then
         Label3(2).Visible = True: txtCaseField(16).Visible = True
         Label1(22).Visible = True: Text1(23).Visible = True 'Added by Morgan 2020/7/29
      Else
         Label3(1).Caption = "母案申請日："
         Label3(2).Visible = False: txtCaseField(16).Visible = False
         Label1(22).Visible = False: Text1(23).Visible = False 'Added by Morgan 2020/7/29
      End If
   'end 2019/11/27
   Else
      Label3(1).Visible = False: txtCaseField(15).Visible = False
      Label3(2).Visible = False: txtCaseField(16).Visible = False
      Label1(22).Visible = False: Text1(23).Visible = False 'Added by Morgan 2020/7/29
   End If
End Sub

'Add by Morgan 2009/7/30
'若CFP案與P案的承辦人 [相同] 則以P案的會稿日為CFP案的齊備日
'若CFP案與P案的承辦人 [不同] 則以P案的會稿完成日為CFP案的齊備日
'若CFP案無國內案則以該案的會稿完成日更新其他多國案的齊備日
'必要傳入欄位:cp(1)~cp(4),cp(9)-->值不會變的欄位
Private Sub UpdEp06BySameCase(cp() As String)
   Dim stSQL As String, intR As Integer
   Dim bUpdate As Boolean, stCP48 As String
   Dim adoRst As ADODB.Recordset
   Dim stPA09 As String, stCP06 As String, stCP10 As String, stCP14 As String
   Dim stCP29 As String, stCP29F As String 'Added by Morgan 2021/3/17
   
   'Modified by Morgan 2016/5/27 +判斷是主案或日本要計件案
   'Modified by Morgan 2019/8/27 +更正日本要計件案條件(原語法會排除非主案日本案)
   'Modified by Morgan 2021/2/23 日本案加判斷有北所分案日(因為分所案件可能會先上承辦人但是否要計件未上N)
   'Modified by Morgan 2021/3/17 +cp29
   stSQL = "select cp06,cp10,cp14,pa09,cp29 from engineerprogress e1,caseprogress c1,patent" & _
      " where e1.ep02='" & cp(9) & "' and e1.ep06 is null and c1.cp09(+)=e1.ep02" & _
      " and c1.cp10 in (" & CaseMapOut & ") and c1.cp27 is null and c1.cp57 is null and c1.cp14 is not null" & _
      " and pa01(+)=c1.cp01 and pa02(+)=c1.cp02 and pa03(+)=c1.cp03 and pa04(+)=c1.cp04 and (cp21 is null or (pa09='011' and cp157 is not null and cp26 is null))"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      stCP06 = "" & adoRst("cp06")
      stCP10 = "" & adoRst("cp10")
      stCP14 = "" & adoRst("cp14")
      stPA09 = "" & adoRst("pa09")
      stCP29 = "" & adoRst("cp29") 'Added by Morgan 2021/3/17
      '要有 group by 語法否則沒資料也會回傳
      'Modified by Morgan 2021/3/17 +cp29
      stSQL = "select max(c2.cp14) eng2,min(nvl(e2.ep07,0)) sam,min(nvl(e2.ep08,0)+nvl(c2.cp27,0)) dif,min(cp29) cp29" & _
         " from casemap,caseprogress c2,engineerprogress e2" & _
         " where cm01(+)='" & cp(1) & "' and cm02(+)='" & cp(2) & "'" & _
         " and cm03(+)='" & cp(3) & "' and cm04(+)='" & cp(4) & "' and cm10='0'" & _
         " and c2.cp01(+)=cm05 and c2.cp02(+)=cm06 and c2.cp03(+)=cm07 and c2.cp04(+)=cm08" & _
         " and c2.cp10 in (" & CaseMapIn & ") and c2.cp57 is null and e2.ep02(+)=c2.cp09 group by cm01,cm02,cm03,cm04"
         
      intR = 1
      Set adoRst = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         stCP29F = "" & adoRst("cp29") 'Added by Morgan 2021/3/17
         With adoRst
         '皆已發文或會稿完成
         If .Fields("dif") > 0 Then
            bUpdate = True
         '皆已會稿且承辦人相同
         ElseIf .Fields("sam") > 0 And stCP14 = .Fields("eng2") Then
            bUpdate = True
         End If
         End With
      '無國內案且有其他國外案已發文或會稿完成
      Else
         'Modified by Morgan 2021/3/17 +cp29
         'Modified by Moragn 2021/3/22 要有 group by 語法否則沒資料也會回傳
         stSQL = "select min(cp29) cp29 from caserelation,caseprogress c2,engineerprogress e2" & _
            " where cr01(+)='" & cp(1) & "' and cr02(+)='" & cp(2) & "'" & _
            " and cr03(+)='" & cp(3) & "' and cr04(+)='" & cp(4) & "'" & _
            " and c2.cp01(+)=cr05 and c2.cp02(+)=cr06 and c2.cp03(+)=cr07 and c2.cp04(+)=cr08" & _
            " and c2.cp10 in (" & CaseMapOut & ") and e2.ep02(+)=c2.cp09" & _
            " and (nvl(c2.cp27,0)+nvl(e2.ep08,0))>0 group by cr01,cr02,cr03,cr04"
            
         intR = 1
         Set adoRst = ClsLawReadRstMsg(intR, stSQL)
         If intR = 1 Then
            stCP29F = "" & adoRst("cp29") 'Added by Morgan 2021/3/17
            bUpdate = True
         End If
      End If
      
      If bUpdate = True Then
         '更新齊備日
         stSQL = "Update Engineerprogress set ep06=" & strSrvDate(1) & " where ep06 is null and ep02='" & cp(9) & "'"
         cnnConnection.Execute stSQL, intR
         
         'Added by Morgan 2021/3/17
         '更新繪圖人員
         If stCP29 = "" And stCP29F <> "" Then
            stSQL = "select st04,st06 from staff where st01='" & stCP29F & "'"
            intR = 1
            Set adoRst = ClsLawReadRstMsg(intR, stSQL)
            If intR = 1 Then
               If adoRst("st04") = "2" Then
                  stSQL = "SELECT ST01 FROM STAFF WHERE ST06='" & adoRst("st06") & "' AND ST05='81' and st04='1'"
                  intR = 1
                  Set adoRst = ClsLawReadRstMsg(intR, stSQL)
                  If intR = 1 Then
                     stCP29F = adoRst("st01")
                  End If
               End If
               stSQL = "Update caseprogress set cp29='" & stCP29F & "'  where cp09='" & cp(9) & "' and cp29 is null"
               cnnConnection.Execute stSQL, intR
            End If
         End If
         'end 2021/3/17
         
         If PUB_IfSetCP48() Then 'Add by Morgan 2010/10/1
            stCP48 = Pub_GetHandleDay(cp(1), stPA09, stCP10, , stCP06, cp(9))
            If stCP48 <> "" Then
               '更新承辦期限
               stSQL = "Update caseprogress set cp48=" & stCP48 & " where cp09='" & cp(9) & "'"
               cnnConnection.Execute stSQL, intR
            End If
         End If
      End If
      
   'Added by Morgan 2021/2/5
   Else
      PUB_SetEP06ByCR cp(1), cp(2), cp(3), cp(4)
   'end 2021/2/5
   End If
   Set adoRst = Nothing
End Sub

Private Sub txtEngGroup_Change()
   If txtEngGroup <> "" Then
      Label3(13) = PUB_GetFCPGrpName(txtEngGroup, True)
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
''所限=法限-1週
'Private Sub txtFavDt_Validate(Cancel As Boolean)
'   Dim stDate As String
'   If txtFavDt <> "" Then
'      If ChkDate(txtFavDt) Then
'         stDate = TransDate(CompDate(1, 6, txtFavDt), 1)
'         If txtCaseField(9) = "" Or Val(txtCaseField(9)) > Val(stDate) Then
'            txtCaseField(9) = stDate
'            txtCaseField(4) = TransDate(PUB_GetWorkDay1(CompDate(2, -7, stDate), True), 1)
'            If Val(txtCaseField(4)) < Val(strSrvDate(2)) Then
'               txtCaseField(9) = strSrvDate(2)
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
      If txtCaseField(7) = "605" Then
         MsgBox "繳費年度不可空白！"
      Else
         MsgBox "繳費次數不可空白！"
      End If
      txtFeeYear(Index).SetFocus
      Cancel = True
      Exit Sub
   Else
      If Index = 1 Then
         strNext = PUB_Getnexttimes(field(1), field(2), field(3), field(4), strYear)
         If strNext <> "" Then
            If txtCaseField(7) = "605" Then '繳費年度
               If Val(txtFeeYear(Index)) <> Val(strYear) Then
                  MsgBox "繳費(起)年度有誤，應為" & strYear & "！"
                  txtFeeYear(Index).SetFocus
                  Cancel = True
                  Exit Sub
               End If
            Else '繳費次數
               If Val(txtFeeYear(Index)) <> Val(strNext) Then
                  'add by sonia 2017/6/2 CFP-020159
                  'Modified by Morgan 2017/7/3 俄羅斯5/2改國家代碼 "233"->"023"
                  'Modified by Morgan 2022/6/15 俄羅斯設計案 2015/1/1以前提申案件除了延展費外仍要繳年費，2015/1/1以後提申案件僅須繳延展費--禧佩
                  'If ((field(9) = "018" And field(8) = "2") Or (field(9) = "023" And field(8) = "3")) And cp(10) = "607" Then
                  If ((field(9) = "018" And field(8) = "2") Or (field(9) = "023" And field(8) = "3" And Val(field(10)) < 20150101)) And cp(10) = "607" Then
                  'end 2022/6/15
                     If MsgBox(lblNation & Label3(11) & "延展費無法檢查繳費紀錄，是否確定次數資料無誤？", vbYesNo + vbDefaultButton2) = vbNo Then
                        Cancel = True 'Added by Morgan 2022/6/15
                        Exit Sub
                     End If
                  Else
                  'end 2017/6/2
                     MsgBox "繳費(起)次數有誤，應為" & strNext & "！"
                     txtFeeYear(Index).SetFocus
                     Cancel = True
                     Exit Sub
                  End If  'add by sonia 2017/6/2
               End If
            End If
         Else
            If txtCaseField(7) = "605" Then  '繳費年度
               MsgBox "無下次繳費年度！"
            Else '繳費次數
               MsgBox "無下次繳費次數！"
            End If
            txtFeeYear(Index).SetFocus
            Cancel = True
            Exit Sub
         End If
      Else
         If Val(txtFeeYear(1)) > Val(txtFeeYear(2)) Then
            If txtCaseField(7) = "605" Then  '繳費年度
               MsgBox "繳費(迄)年度不可小於(起)年度！"
            Else '繳費次數
               MsgBox "繳費(迄)次數不可小於(起)次數！"
            End If
            txtFeeYear(Index).SetFocus
            Cancel = True
            Exit Sub
         End If
      End If
   End If
End Sub

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
           'Modify by Amy 2017/07/13 +服務業務 非台灣才可輸,只能輸J或空白
           If intCaseKind <> 專利 And field(9) <> "000" And KeyAscii <> 8 And KeyAscii <> Asc("J") Then
                 KeyAscii = 0
                 Beep
           '專利
           ElseIf intCaseKind = 專利 And KeyAscii <> 8 And KeyAscii <> Asc("T") And KeyAscii <> Asc("J") And field(9) <> "000" Then
              KeyAscii = 0
              Beep
           End If
        Else
        '2013/12/15 END
           If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
              KeyAscii = 0
              Beep
           End If
        End If
   End If 'Added by Lydia 2020/03/31
End Sub
'Add by Lydia 2014/11/24 PCT進各個國家階段之期限設定
'414恢復專利權
Public Sub Set414Date()
   Dim stSQL As String, adoRst As ADODB.Recordset, adoRst1 As ADODB.Recordset, iR As Integer
   Dim bUpdate As Boolean, strCP133 As String, MonDiff As Integer
   '是否為PCT案件
   If m_field46 = "Y" And txtCaseField(7) = 申請復活 And txtCaseField(5) <> "" Then
      'Modified by Morgan 2018/5/16 +未發文新案條件 CP31='Y' and cp27 is null ( Ex:CFP-25836 非新案不適用 )
      stSQL = "select cp06,cp07,cp05,cp10,cp31,na01,NVL(na75,30) na75,NVL(na76,30) NA76 from caseprogress,patent,nation " & _
              "where cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa09=na01(+) and cp09='" & txtCaseField(5) & "' and cp31='Y' and cp27 is null"
      iR = 1
      Set adoRst = ClsLawReadRstMsg(iR, stSQL)
      If iR = 1 Then
         strExc(1) = "": strExc(2) = ""
         MonDiff = adoRst("na76") - adoRst("na75")
         '新案
         If adoRst("cp31") = "Y" Then
            If IsNull(adoRst("cp07")) Then
               MsgBox "相關收文號尚無期限，請先做該程序分案作業！"
               GoTo ExitPort
            Else
                '(最初)優先權日=PCT優先權日>PCT申請日
                If txtCaseField(16) <> "" And txtCaseField(15) <> "" Then
                   strExc(0) = txtCaseField(16)
                ElseIf txtCaseField(15) <> "" Then
                   strExc(0) = txtCaseField(15)
                End If
                   '新案與414恢復的算法一致
                   PUB_GetPCTLimit strExc(0), strExc(1), strExc(2), cp(1), cp(2), cp(3), cp(4), txtCaseField(2)
            End If
         End If 'end '新案

            If DBDATE(strExc(2)) < strSrvDate(1) Then strExc(2) = strSrvDate(1) '所限不可小於系統日
            
            '轉民國年
            strExc(1) = TransDate(strExc(1), 1)
            strExc(2) = TransDate(PUB_GetWorkDay1(strExc(2), True), 1) '抓工作日
            
            If txtCaseField(9) = "" And txtCaseField(4) = "" Then
               bUpdate = True
               If MonDiff <= 0 Then MsgBox "恢復原狀不會延長月數!!"
            ElseIf strExc(2) <> txtCaseField(4) Or strExc(1) <> txtCaseField(9) Then
               If MonDiff <= 0 Then MsgBox "恢復原狀不會延長月數!!"
               strExc(0) = "期限將變更(所限 " & txtCaseField(4) & "->" & strExc(2) & ",法限 " & txtCaseField(9) & "->" & strExc(1) & "),是否要更新?"
               If MsgBox(strExc(0), vbYesNo, "") = vbYes Then
                  bUpdate = True
               End If
            End If
            
            If bUpdate Then
               txtCaseField(4) = TransDate(strExc(2), 1) '所限
               txtCaseField(9) = TransDate(strExc(1), 1) '法限
            End If
         End If 'If iR = 1 Then
      
   End If

ExitPort:

   Set adoRst = Nothing
   Set adoRst1 = Nothing
'end 2014/11/24
End Sub

'Add by Lydia 2015/02/02 輸入新穎性優惠期公開事實 (多筆)
Private Sub CmdFav_Click()
'Modified by Morgan 2023/10/27
'If cp(10) = "123" Then
If txtCaseField(7) = "123" Then
'end 2023/10/27
   Set frm880020.m_PrevF = Me
   frm880020.m_dbCheck = False 'Modified by Lydia 2015/02/25  發文DoubleCheck
   frm880020.strFPD01 = field(1):   frm880020.strFPD02 = field(2)
   frm880020.strFPD03 = field(3):   frm880020.strFPD04 = field(4)
   frm880020.strLimit1 = txtCaseField(4)
   frm880020.strLimit2 = txtCaseField(9)
   frm880020.strNation = field(9)
   frm880020.strPA08 = Text1(21) 'Added by Morgan 2018/3/16
   frm880020.strPA140 = IIf(txtFavDt.Text = "", IIf(field(140) <> "", Val(field(140)) - 19110000, ""), txtFavDt.Text)
   frm880020.Show
End If
End Sub

'Add by Amy 2022/10/21
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

'Add By Sindy 2023/3/31
Private Sub txtPA61_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 78 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Added by Morgan 2024/1/5
Private Sub OptSendType_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim oOpt As OptionButton
   If OptSendType(Index).Tag = "1" Then
      OptSendType(Index).Value = False
      OptSendType(Index).Tag = "0"
      If Index = 3 Then
         txtCP142.Text = ""
         txtCP142.Enabled = False
         If Frame2.Visible = True Then
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False
            Option1(0).Enabled = False
            Option1(1).Enabled = False
            Option1(2).Enabled = False
         End If
      End If
      
   Else
      For Each oOpt In OptSendType
         If oOpt.Index = Index Then
            oOpt.Tag = "1"
         Else
            oOpt.Tag = "0"
         End If
      Next
      If Index = 3 And OptSendType(Index).Value Then
         txtCP142.Enabled = True
         txtCP142.SetFocus
         If Frame2.Visible = True Then
            Option1(0).Enabled = True
            Option1(1).Enabled = True
            Option1(2).Enabled = True
         End If
      Else
         txtCP142.Text = ""
         txtCP142.Enabled = False
         If Frame2.Visible = True Then
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False
            Option1(0).Enabled = False
            Option1(1).Enabled = False
            Option1(2).Enabled = False
         End If
      End If
   End If
End Sub

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
      ElseIf Val(txtCaseField(9)) > 0 And Val(txtCP142) > Val(txtCaseField(9)) Then
         MsgBox "指定送件日期不可晚於法定期限！", vbExclamation
         Cancel = True
      End If
   End If
End Sub
