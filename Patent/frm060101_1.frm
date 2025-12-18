VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060101_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專分案"
   ClientHeight    =   5772
   ClientLeft      =   132
   ClientTop       =   972
   ClientWidth     =   8808
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5772
   ScaleWidth      =   8808
   Begin VB.CommandButton Command2 
      Caption         =   "多國案卷號(&F)"
      Height          =   350
      Index           =   6
      Left            =   1500
      TabIndex        =   110
      Top             =   10
      Width           =   1410
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4608
      Left            =   96
      TabIndex        =   54
      Top             =   1152
      Width           =   8628
      _ExtentX        =   15219
      _ExtentY        =   8128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   529
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm060101_1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(43)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(41)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(40)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(39)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(35)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(34)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(32)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(31)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(30)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(29)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(3)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(36)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(37)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(38)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label3(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(11)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(13)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblDivCase"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblTF05"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblTF18"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(14)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(42)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblEP06"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(15)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblDesignCase"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label1(16)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label3(0)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lblTF23"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lblTF19"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cboCP14"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label5(12)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label5(2)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label5(3)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label5(4)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label5(5)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label5(6)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label5(7)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "text1(11)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "text1(10)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "text1(7)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "text1(23)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "cmdSetDate"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtDivCaseNo(4)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "text1(22)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtDivCaseNo(3)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtDivCaseNo(2)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtDivCaseNo(1)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "text1(18)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "text1(17)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "text1(16)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "text1(15)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "text1(14)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "text1(13)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "text1(12)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "text1(8)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "text1(6)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "text1(4)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "text1(27)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "text1(26)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "text1(25)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "text1(2)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "MSHFlexGrid1"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "text1(9)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "text1(5)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "text1(3)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "text1(1)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "txtTF05"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "txtTF18"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "Combo2"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "text1(28)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "Combo3"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "text1(29)"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "txtDesignCaseNo(1)"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "txtDesignCaseNo(2)"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "txtDesignCaseNo(3)"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "txtDesignCaseNo(4)"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "text1(30)"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "text1(0)"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "txtTF23"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "txtTF19"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "FraLOS"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "Option1(1)"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "Option1(0)"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "Option1(2)"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "chkCP176"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).ControlCount=   89
      TabCaption(1)   =   "備註"
      TabPicture(1)   =   "frm060101_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(10)"
      Tab(1).Control(1)=   "Label1(9)"
      Tab(1).Control(2)=   "Label2(0)"
      Tab(1).Control(3)=   "Label2(1)"
      Tab(1).Control(4)=   "Label2(2)"
      Tab(1).Control(5)=   "Label2(3)"
      Tab(1).Control(6)=   "textCP64"
      Tab(1).Control(7)=   "textPA91"
      Tab(1).ControlCount=   8
      Begin VB.CheckBox chkCP176 
         Caption         =   "暫不送件"
         ForeColor       =   &H000000C0&
         Height          =   220
         Left            =   4350
         TabIndex        =   26
         Top             =   1830
         Width           =   1060
      End
      Begin VB.OptionButton Option1 
         Caption         =   "之後"
         Height          =   195
         Index           =   2
         Left            =   3630
         TabIndex        =   25
         Top             =   1830
         Width           =   705
      End
      Begin VB.OptionButton Option1 
         Caption         =   "當天"
         Height          =   195
         Index           =   0
         Left            =   2250
         TabIndex        =   23
         Top             =   1830
         Width           =   705
      End
      Begin VB.OptionButton Option1 
         Caption         =   "之前"
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   24
         Top             =   1830
         Width           =   705
      End
      Begin VB.Frame FraLOS 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '沒有框線
         Caption         =   "Frame2"
         Height          =   285
         Left            =   4380
         TabIndex        =   108
         Top             =   40
         Width           =   3225
         Begin VB.TextBox txtLOSagree 
            Height          =   270
            Left            =   1890
            MaxLength       =   1
            TabIndex        =   47
            Top             =   -8
            Width           =   405
         End
         Begin VB.Label LBL6 
            Caption         =   "是否需要法律所配合：　　　(Y: 配合) "
            Height          =   195
            Left            =   0
            TabIndex        =   109
            Top             =   30
            Width           =   3135
         End
      End
      Begin VB.TextBox txtTF19 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   7440
         MaxLength       =   6
         TabIndex        =   38
         Top             =   2660
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTF23 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   5340
         MaxLength       =   6
         TabIndex        =   37
         Top             =   2660
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   0
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   104
         Top             =   30
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   30
         Left            =   1350
         MaxLength       =   7
         TabIndex        =   22
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtDesignCaseNo 
         Height          =   270
         Index           =   4
         Left            =   7635
         MaxLength       =   2
         TabIndex        =   43
         Top             =   3210
         Width           =   390
      End
      Begin VB.TextBox txtDesignCaseNo 
         Height          =   270
         Index           =   3
         Left            =   7275
         MaxLength       =   1
         TabIndex        =   42
         Top             =   3210
         Width           =   390
      End
      Begin VB.TextBox txtDesignCaseNo 
         Height          =   270
         Index           =   2
         Left            =   6570
         MaxLength       =   6
         TabIndex        =   41
         Top             =   3210
         Width           =   705
      End
      Begin VB.TextBox txtDesignCaseNo 
         Height          =   270
         Index           =   1
         Left            =   6180
         MaxLength       =   3
         TabIndex        =   40
         Top             =   3210
         Width           =   390
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   29
         Left            =   5730
         MaxLength       =   1
         TabIndex        =   39
         Top             =   2940
         Width           =   330
      End
      Begin VB.ComboBox Combo3 
         Height          =   276
         Left            =   3240
         Style           =   2  '單純下拉式
         TabIndex        =   100
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   28
         Left            =   7695
         MaxLength       =   7
         TabIndex        =   9
         Top             =   780
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         Left            =   6210
         Style           =   2  '單純下拉式
         TabIndex        =   12
         Top             =   1020
         Width           =   780
      End
      Begin VB.TextBox txtTF18 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   7440
         MaxLength       =   3
         TabIndex        =   36
         Top             =   2400
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTF05 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   5340
         MaxLength       =   3
         TabIndex        =   35
         Top             =   2400
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   1
         Left            =   5340
         MaxLength       =   4
         TabIndex        =   1
         Top             =   345
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   3
         Left            =   5340
         MaxLength       =   1
         TabIndex        =   6
         Top             =   585
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   5
         Left            =   5340
         MaxLength       =   7
         TabIndex        =   8
         Top             =   825
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   9
         Left            =   7695
         MaxLength       =   7
         TabIndex        =   15
         Top             =   1305
         Width           =   855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1050
         Left            =   60
         TabIndex        =   44
         Top             =   3510
         Width           =   8505
         _ExtentX        =   15007
         _ExtentY        =   1842
         _Version        =   393216
         Cols            =   12
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
         _Band(0).Cols   =   12
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   2
         Left            =   1350
         MaxLength       =   3
         TabIndex        =   2
         Top             =   585
         Width           =   390
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   25
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   3
         Top             =   585
         Width           =   705
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   26
         Left            =   2445
         MaxLength       =   1
         TabIndex        =   4
         Top             =   585
         Width           =   390
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   27
         Left            =   2835
         MaxLength       =   2
         TabIndex        =   5
         Top             =   585
         Width           =   390
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   4
         Left            =   1350
         MaxLength       =   7
         TabIndex        =   7
         Top             =   825
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   6
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   10
         Top             =   1065
         Width           =   1095
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   8
         Left            =   1350
         TabIndex        =   13
         Top             =   1305
         Width           =   2805
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   12
         Left            =   3300
         MaxLength       =   8
         TabIndex        =   17
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   13
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   29
         Top             =   2025
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   14
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   30
         Top             =   2265
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   15
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   31
         Top             =   2505
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   16
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   32
         Top             =   2745
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   17
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   33
         Top             =   2985
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   18
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   34
         Top             =   3210
         Width           =   855
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   270
         Index           =   1
         Left            =   5835
         MaxLength       =   3
         TabIndex        =   18
         Top             =   1575
         Width           =   390
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   270
         Index           =   2
         Left            =   6225
         MaxLength       =   6
         TabIndex        =   19
         Top             =   1575
         Width           =   705
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   270
         Index           =   3
         Left            =   6930
         MaxLength       =   1
         TabIndex        =   20
         Top             =   1575
         Width           =   390
      End
      Begin VB.TextBox text1 
         Enabled         =   0   'False
         Height          =   270
         Index           =   22
         Left            =   5340
         MaxLength       =   6
         TabIndex        =   28
         Top             =   2115
         Width           =   855
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   270
         Index           =   4
         Left            =   7290
         MaxLength       =   2
         TabIndex        =   21
         Top             =   1575
         Width           =   390
      End
      Begin VB.CommandButton cmdSetDate 
         Caption         =   "設定期限"
         Height          =   255
         Left            =   2160
         TabIndex        =   95
         Top             =   825
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   23
         Left            =   5340
         MaxLength       =   7
         TabIndex        =   11
         Top             =   1065
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   7
         Left            =   5340
         MaxLength       =   7
         TabIndex        =   14
         Top             =   1305
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   10
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   11
         Left            =   6710
         MaxLength       =   1
         TabIndex        =   27
         Top             =   1850
         Width           =   410
      End
      Begin MSForms.Label Label5 
         Height          =   180
         Index           =   7
         Left            =   2220
         TabIndex        =   119
         Top             =   3270
         Width           =   2055
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3625;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label5 
         Height          =   180
         Index           =   6
         Left            =   2220
         TabIndex        =   118
         Top             =   3030
         Width           =   2055
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3625;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label5 
         Height          =   180
         Index           =   5
         Left            =   2220
         TabIndex        =   117
         Top             =   2790
         Width           =   2055
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3625;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label5 
         Height          =   180
         Index           =   4
         Left            =   2220
         TabIndex        =   116
         Top             =   2550
         Width           =   2055
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3625;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label5 
         Height          =   180
         Index           =   3
         Left            =   2220
         TabIndex        =   115
         Top             =   2310
         Width           =   2055
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3625;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label5 
         Height          =   180
         Index           =   2
         Left            =   2220
         TabIndex        =   114
         Top             =   2070
         Width           =   2055
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3625;317"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label5 
         Height          =   195
         Index           =   12
         Left            =   6240
         TabIndex        =   113
         Top             =   2160
         Width           =   795
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1402;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboCP14 
         Height          =   285
         Left            =   1350
         TabIndex        =   0
         Top             =   300
         Width           =   1860
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3281;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA91 
         Height          =   1170
         Left            =   -73980
         TabIndex        =   46
         Top             =   1560
         Width           =   7515
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "13256;2064"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   1170
         Left            =   -73980
         TabIndex        =   45
         Top             =   360
         Width           =   7515
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "13256;2064"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblTF19 
         Caption         =   "相似度                        %"
         Height          =   180
         Left            =   6600
         TabIndex        =   107
         Top             =   2705
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblTF23 
         Caption         =   "原文字數"
         Height          =   180
         Left            =   4365
         TabIndex        =   106
         Top             =   2705
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   180
         Index           =   0
         Left            =   2580
         TabIndex        =   105
         Top             =   30
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "指定送件日期"
         Height          =   180
         Index           =   16
         Left            =   150
         TabIndex        =   103
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label lblDesignCase 
         AutoSize        =   -1  'True
         Caption         =   "衍生設計母案本所案號"
         Height          =   180
         Left            =   4365
         TabIndex        =   102
         Top             =   3255
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否向客戶收款           (N: 不收)"
         Height          =   180
         Index           =   15
         Left            =   4365
         TabIndex        =   101
         Top             =   2985
         Width           =   2445
      End
      Begin VB.Label lblEP06 
         AutoSize        =   -1  'True
         Caption         =   "文件齊備日"
         Height          =   180
         Left            =   6705
         TabIndex        =   99
         Top             =   810
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "取消收文日"
         Height          =   180
         Index           =   42
         Left            =   6705
         TabIndex        =   58
         Top             =   1365
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦期限"
         Height          =   180
         Index           =   14
         Left            =   4365
         TabIndex        =   98
         Top             =   1110
         Width           =   720
      End
      Begin VB.Label lblTF18 
         Caption         =   "加成比率                    %"
         Height          =   180
         Left            =   6600
         TabIndex        =   97
         Top             =   2430
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label lblTF05 
         Caption         =   "相似折扣                       %"
         Height          =   180
         Left            =   4365
         TabIndex        =   96
         Top             =   2430
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.Label lblDivCase 
         AutoSize        =   -1  'True
         Caption         =   "分割母案本所案號"
         Height          =   180
         Left            =   4365
         TabIndex        =   94
         Top             =   1620
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "核稿人"
         Height          =   180
         Index           =   13
         Left            =   4365
         TabIndex        =   93
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(N:不算)"
         Height          =   180
         Index           =   11
         Left            =   1740
         TabIndex        =   88
         Top             =   1620
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   180
         Index           =   1
         Left            =   6300
         TabIndex        =   84
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   3
         Left            =   -72000
         TabIndex        =   79
         Top             =   3624
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   2
         Left            =   -72120
         TabIndex        =   78
         Top             =   2484
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   1
         Left            =   -72120
         TabIndex        =   77
         Top             =   1524
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   0
         Left            =   -72120
         TabIndex        =   76
         Top             =   504
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人"
         Height          =   180
         Index           =   38
         Left            =   150
         TabIndex        =   75
         Top             =   3225
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質"
         Height          =   180
         Index           =   37
         Left            =   4365
         TabIndex        =   74
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人5"
         Height          =   180
         Index           =   36
         Left            =   150
         TabIndex        =   73
         Top             =   2985
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   72
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "轉本所案號"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   71
         Top             =   585
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   70
         Top             =   825
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "相關總收文號"
         Height          =   180
         Index           =   3
         Left            =   150
         TabIndex        =   69
         Top             =   1065
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號"
         Height          =   180
         Index           =   4
         Left            =   150
         TabIndex        =   68
         Top             =   1305
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否算案件數"
         Height          =   180
         Index           =   29
         Left            =   150
         TabIndex        =   67
         Top             =   1575
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收文日"
         Height          =   180
         Index           =   30
         Left            =   2670
         TabIndex        =   66
         Top             =   1620
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人1"
         Height          =   180
         Index           =   31
         Left            =   150
         TabIndex        =   65
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人2"
         Height          =   180
         Index           =   32
         Left            =   150
         TabIndex        =   64
         Top             =   2265
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人3"
         Height          =   180
         Index           =   34
         Left            =   150
         TabIndex        =   63
         Top             =   2505
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人4"
         Height          =   180
         Index           =   35
         Left            =   150
         TabIndex        =   62
         Top             =   2745
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公告日"
         Height          =   180
         Index           =   39
         Left            =   4365
         TabIndex        =   61
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限"
         Height          =   180
         Index           =   40
         Left            =   4365
         TabIndex        =   60
         Top             =   870
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "卷宗性質                          (1.申請 2.異議 3.舉發)"
         Height          =   180
         Index           =   41
         Left            =   4365
         TabIndex        =   59
         Top             =   630
         Width           =   3585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否取消閉卷          (Y : 取消閉卷)"
         Height          =   180
         Index           =   43
         Left            =   5570
         TabIndex        =   57
         Top             =   1860
         Width           =   2720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "進度備註"
         Height          =   180
         Index           =   9
         Left            =   -74760
         TabIndex        =   56
         Top             =   444
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件備註"
         Height          =   180
         Index           =   10
         Left            =   -74760
         TabIndex        =   55
         Top             =   1524
         Width           =   720
      End
   End
   Begin VB.TextBox text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   21
      Left            =   3768
      MaxLength       =   1
      TabIndex        =   92
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一筆(&N)"
      Height          =   350
      Index           =   5
      Left            =   3300
      TabIndex        =   48
      Top             =   10
      Width           =   900
   End
   Begin VB.CommandButton Command2 
      Caption         =   "優先權資料(&P)"
      Height          =   350
      Index           =   4
      Left            =   150
      TabIndex        =   49
      Top             =   10
      Visible         =   0   'False
      Width           =   1300
   End
   Begin VB.TextBox text1 
      Height          =   270
      Index           =   24
      Left            =   4944
      TabIndex        =   87
      Top             =   360
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.CommandButton Command2 
      Caption         =   "相關卷號(&F)"
      Height          =   350
      Index           =   0
      Left            =   4188
      TabIndex        =   50
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "案件進度(&C)"
      Height          =   350
      Index           =   1
      Left            =   5412
      TabIndex        =   51
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   2
      Left            =   6636
      TabIndex        =   52
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   3
      Left            =   7464
      TabIndex        =   53
      Top             =   10
      Width           =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   204
      Index           =   2
      Left            =   6960
      TabIndex        =   123
      Top             =   600
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "工程師組別："
      Height          =   180
      Index           =   18
      Left            =   5820
      TabIndex        =   122
      Top             =   600
      Width           =   1080
   End
   Begin MSForms.Label Label5 
      Height          =   192
      Index           =   11
      Left            =   7944
      TabIndex        =   121
      Top             =   396
      Width           =   792
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1397;339"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "程序："
      Height          =   180
      Index           =   17
      Left            =   7344
      TabIndex        =   120
      Top             =   396
      Width           =   540
   End
   Begin MSForms.Label Label5 
      Height          =   192
      Index           =   10
      Left            =   6384
      TabIndex        =   112
      Top             =   396
      Width           =   792
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1411;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   288
      Left            =   996
      TabIndex        =   111
      Top             =   816
      Width           =   7740
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "13652;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   11
      Left            =   4176
      TabIndex        =   91
      Top             =   408
      Width           =   468
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   12
      Left            =   2856
      TabIndex        =   90
      Top             =   408
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  '透明
      Caption         =   "Label4"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   4500
      TabIndex        =   89
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   9
      Left            =   1344
      TabIndex        =   86
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   8
      Left            =   1350
      TabIndex        =   85
      Top             =   390
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號"
      Height          =   180
      Index           =   5
      Left            =   144
      TabIndex        =   83
      Top             =   408
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號"
      Height          =   180
      Index           =   6
      Left            =   144
      TabIndex        =   82
      Top             =   648
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱"
      Height          =   180
      Index           =   7
      Left            =   144
      TabIndex        =   81
      Top             =   888
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權："
      Height          =   180
      Index           =   8
      Left            =   5820
      TabIndex        =   80
      Top             =   396
      Width           =   540
   End
End
Attribute VB_Name = "frm060101_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/11 Form2.0已修改
'Modify by Amy 2014/04/14 尋問靜芳確定優先權資料不會在此輸,故優先權資料鈕拿掉
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
'2005/7/14整理
Option Explicit

Dim StrTot1(0 To 500) As String, StrTot2(0 To 500) As String
Dim IntNow As Integer, IntTot As Integer 'Memo by Lydia 2016/06/21 前一畫面點選的進度
Dim m_PrevForm As Form 'Added by Lydia 2018/05/21前一畫面(表單)
Dim strReceiveNo As String, strKind As String
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String, intWhere As Integer, intLastRow As Integer, intNowRow As Integer
Dim pa() As String, intWhere As Integer, intLastRow As Integer, intNowRow As Integer

'Dim strPriority(1 To 3) As String 'Mark by Amy 2014/04/14 原優先權資料使用
Dim m_CP60 As String
'Add By Cheng 2002/06/10
Dim m_strCP06 As String '記錄原始的本所期限
Dim m_strCP07 As String '記錄原始的法定期限
Dim m_strCP07_1 As String '記錄原始的法定期限  add by sonia 2015/10/13
Dim m_strCP48 As String '記錄原始的承辦期限
'NICK 900606 ***********
' 轉本所案號  本來一個  TEXT  物件
'  換成 4 個物件
'***********************
'Add By Cheng 2002/08/22
Dim m_strCust1 As String '申請人1
Dim m_strCust2 As String '申請人2
Dim m_strCust3 As String '申請人3
Dim m_strCust4 As String '申請人4
Dim m_strCust5 As String '申請人5
'Add By Cheng 2002/11/04
Dim m_strCP10 As String
'Add by Morgan 2004/3/22
'本所期限,法定期限,實審收文號,申請國家
'Memo by Lydia 2022/09/15 m_stVar(0)=(暫存)所限, m_stVar(3)=(暫存)法限
Dim m_stVar(0 To 3) As String, m_st416CP09 As String, m_stPA09 As String
'2005/6/15 ADD BY SONIA
Dim m_PA08 As String
'Add by Morgan 2006/5/1
Dim m_bol108 As Boolean '是否有收文申請寄存
'Add by Morgan 2007/6/11
Dim m_CP27 As String '發文日
Dim m_CP57 As String '取消收文日
Dim m_ExpCP14 As String, m_RelCP10n As String '已收文未發文且需為相同承辦人的承辦人代碼,案件性質名稱
Dim m_203CP09 As String '實審分案時,有主動修正(203,206)未發文
Dim m_PA143 As String 'Add by Morgan 2008/3/18
Dim m_bActive As Boolean
Dim m_EP06 As String '文件齊備日
Dim m_CP66 As String 'Added by Lydia 2018/01/08 建檔日
'Remove by Morgan 2010/4/29 改加欄位顯示
'Dim m_CP20 As String '是否向客戶收款 Add by Morgan 2009/10/6
Dim m_bolActivated As Boolean 'Add by Morgan 2010/3/23
Dim m_CP20Default As String '案件性質預設值 Add by Morgan 2010/4/29
Dim m_CP14 As String, m_EP04 As String  'Add by Morgan 2010/6/17
Dim m_CP30 As String 'Add by Morgan 2011/4/22
Dim m_CP31 As String 'Added by Lydia 2020/08/18
Dim m_CP122 As String 'Added by Morgan 2013/1/4
Dim m_203CP48 As String 'Added by Morgan 2013/1/7 主動修正的預設承辦期限
Dim m_307PA08 As String, m_DCCode As Integer 'Add by Lydia 2014/10/21 當分割案與母案的專利種類不同，自動更新為相同種類
Dim BolIsInventorUpd As Boolean 'Add By Sindy 2014/12/23 檢查是否要更新發明人資料
Dim m_bol435 As Boolean 'Added by Morgan 2015/9/9
Dim m_PA163 As String 'Added by Morgan 2015/9/30
Dim strSetLimitDT As String 'Add By Sindy 2015/12/16 有約定期限或指定日期
Dim m_EP09 As String '完稿日
Dim m_FCPTeam As String 'Add By Sindy 2016/1/28
Dim m_CP118 As String, n_CP118 As String 'Added by Lydia 2017/12/14 記錄是否電子送件
Private Const cAutoCP924 As String = "209,235" 'Addded by Lydia 2018/05/31 會稿自動掛承辦人的案件性質
'Added by Lydia 2018/08/07 命名作業的資料
Dim m_TCT01 As String '新案收文號(PK)
Dim m_TCT10 As String '命名人員
Dim m_TCT27 As String '欲翻譯此案件者/指定翻譯
Dim m_TCT28 As String '其他指定翻譯
Dim mTransKind As String 'Added by Lydia 2018/08/08 翻譯分案->只能上班翻譯 (Y後面+逗號,再加上有折扣或固定報價)
Dim mLimitDate As String 'Added by Lydia 2018/09/12 翻譯分案->交稿期限
'Added by Lydia 2018/08/07 自動掛承辦人的案件性質
Private Const cAutoCPMList As String = "107再審申請,203主動修正,204修正,220補充元件符號,239擇一申復,402更正,407請求面詢,408面詢,422加速審查,428差異說明,431高速審查,433誤譯訂正,501訴願,503行政訴訟,803舉發,804舉發答辯,901告知代理人,902回覆代理人"
Dim FCP檢視中說必輸原文字數 As String 'Added by Lydia 2019/06/28
'Added by Lydia 2020/05/20 法律所案源收文
Dim m_LOS15 As String '案源單號
Dim m_LOS01 As String '案源總收文號
Dim m_LOS02 As String 'Added by Lydia 2020/06/09 案源案件類型
Dim m_LOS07 As String '放棄日期
Dim strMurgitroyd As String 'Added by Lydia 2021/01/06 Murgitroyd案的代理人
Dim m_pAgreeOnDate As String 'Modify By Sindy 2021/4/27
Dim bUpdPA150 As String 'Added by Lydia 2022/09/29 變更案件之工程師組別 ( 新增工程師分組控管)
'Modified by Lydia 2025/06/05 更改名稱
'Dim m_strBASF As String 'Added by Lydia 2023/04/19 BASF集團的X編號
Dim m_str所內譯 As String
Dim m_str所內譯例外 As String 'Added by Lydia 2025/07/01
Dim m_strCP06Update As String '更新後的本所期限
Dim m_strCP07Update As String '更新後的法定期限
'Add By Sindy 2023/12/6
Dim m_CP43CP08 As String '相關總收文號的資料
Dim m_CP43CP64 As String
Dim m_CP13 As String
'2023/12/6 END
Dim m_strSubject As String 'Added by Morgan 2024/5/21 特定的承辦人通知信主旨(訴願的補充說明)
Dim m_NA16Na79 As String 'Added by Lydia 2024/10/04 程序管制人
Dim strMsgCloseCancel As String 'Added by Lydia 2025/06/27 取消閉卷時，若下一程序有未過期且已上N之年費605，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。

'Add By Sindy 2016/7/26
'承辦人
Private Sub cboCP14_Change()
   SetTF
End Sub
Private Sub CboCP14_GotFocus()
   cboCP14.SelStart = 0
   cboCP14.SelLength = Len(cboCP14.Text)
End Sub
Private Sub CboCP14_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCP14_LostFocus()
   '有輸入本所期限 或 有輸入指定送件日期
   If (Text1(4).Text <> "" And Text1(4).Tag <> Text1(4).Text) Or _
      (Text1(30).Text <> "" And Text1(30).Tag <> Text1(30).Text _
        And (Option1(0).Value = True Or Option1(1).Value = True Or Option1(2).Value = True) _
      ) Then
      'Modified by Lydia 2016/06/21 改成模組PUB_GetFCPsetCP48Limit
      'Call SetCP48Limit
      Text1(0) = Left(Trim(cboCP14.Text), 5)
      Label3(0) = Trim(Mid(Trim(cboCP14.Text), 6))
      'Modify By Sindy 2021/9/2 + , Text1(1):案件性質
      'Modify By Sindy 2021/9/2 + , IIf(Option1(0).Value = True, "1", IIf(Option1(1).Value = True, "2", "3")), pa(1)
      'Added by Lydia 2021/11/05 判斷C類來函有指定送件日，不變更預設承辦期限；ex.FCP-58246變更承辦人
      If Text1(30) <> "" And Left(Label3(8), 1) = "C" Then
      Else
      'end 2021/11/05
           'Modify By Sindy 2024/12/19 + , Label3(8)=收文號
           Call PUB_GetFCPsetCP48Limit(strSetLimitDT, Text1(0), Text1(4), Text1(23), Text1(30), Text1(1), IIf(Option1(0).Value = True, "1", IIf(Option1(1).Value = True, "2", "3")), pa(1), Label3(8))
           'Added by Lydia 2024/03/08 因應內專支援機械組OA (含P案) 收文: 927其他翻譯
           If mTransKind = "＃" Then
              '承辦期限: 分案日+14日曆天
              Text1(23) = TransDate(CompDate(2, 14, strSrvDate(1)), 1)
           End If
           'end 2024/03/08
      End If 'Added by Lydia 2021/11/05
   End If
End Sub

'Add By Sindy 2021/10/21
Private Sub Option1_Click(Index As Integer)
   '有輸入本所期限 或 有輸入指定送件日期
   If (Text1(4).Text <> "" And Text1(4).Tag <> Text1(4).Text) Or _
      (Text1(30).Text <> "" And Text1(30).Tag <> Text1(30).Text _
        And (Option1(0).Value = True Or Option1(1).Value = True Or Option1(2).Value = True) _
      ) Then
      'Modified by Lydia 2016/06/21 改成模組PUB_GetFCPsetCP48Limit
      'Call SetCP48Limit
      Text1(0) = Left(Trim(cboCP14.Text), 5)
      Label3(0) = Trim(Mid(Trim(cboCP14.Text), 6))
      'Modify By Sindy 2021/9/2 + , Text1(1):案件性質
      'Modify By Sindy 2021/9/2 + , IIf(Option1(0).Value = True, "1", IIf(Option1(1).Value = True, "2", "3")), pa(1)
      'Added by Lydia 2021/11/05 判斷C類來函有指定送件日，不變更預設承辦期限；ex.FCP-58246變更承辦人
      If Text1(30) <> "" And Left(Label3(8), 1) = "C" Then
      Else
      'end 2021/11/05
           'Modify By Sindy 2024/12/19 + , Label3(8)=收文號
           Call PUB_GetFCPsetCP48Limit(strSetLimitDT, Text1(0), Text1(4), Text1(23), Text1(30), Text1(1), IIf(Option1(0).Value = True, "1", IIf(Option1(1).Value = True, "2", "3")), pa(1), Label3(8))
      End If 'Added by Lydia 2021/11/05
   End If
End Sub

Private Sub CboCP14_Validate(Cancel As Boolean)
Dim m_Team As String
Dim strText As String
Dim bolRunOK As Boolean
   
   If Trim(cboCP14.Text) <> "" Then
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(Trim(cboCP14.Text), 5), False) = True Then
         'Modify By Sindy 2021/8/30 [新增]當分案分到上一道是離職承辦工程師，承辦人請自動帶該組別副理
         '（日文組案件時，化學案就帶簡副理(99037)；電機案就帶林副理(94012)）
         'Modified by Lydia 2023/01/12 日文組只需單純回傳離職人員的主管
         'cboCP14.Text = PUB_GetFCPEngSup(Left(Trim(cboCP14.Text), 5), True)
         cboCP14.Text = PUB_GetFCPEngSup(Left(Trim(cboCP14.Text), 5), True, True)
         strText = GetPrjSalesNM(Left(cboCP14.Text, 5))
         If strText <> "" Then
            cboCP14.Text = Left(cboCP14.Text, 5) & " " & strText
         Else
         '2021/8/30 END
            MsgBox "此人員已離職！！", , "人員錯誤！！"
            Cancel = True
            cboCP14.Text = ""
            If cboCP14.Enabled And cboCP14.Visible = True Then 'Added by Lydia 2018/08/08 FCP-058299分主動修正，自動帶入工程師已離職，因為畫面未顯示造成程式出錯
               cboCP14.SetFocus
               Call CboCP14_GotFocus
               Exit Sub
            End If
         End If
         '2021/8/30 END
      End If
      
      '判斷承辦人和智權人員
      Text1(0) = Left(Trim(cboCP14.Text), 5)
      Label3(0) = Trim(Mid(Trim(cboCP14.Text), 6))
      If PUB_FCPGetCP14EP04("CP14", pa, Text1(0), Label3(0)) = False Then
         Cancel = True
      Else
         cboCP14.Text = Trim(Text1(0)) & " " & Trim(Label3(0))
      End If
      If PUB_FCPCheckCP14(pa, Text1(1), Text1(0), Text1(22), Label5(12)) = False Then
         Cancel = True
      End If
      
      'Add by Morgan 2008/8/20
      'Modify By Sindy 2021/10/20 + Not (Text1(23).Text <> "" And pa(1) = "FG")
      '亭妙:FG案件的一些性質(EX.其他翻譯、專利調查、提供情報..) 在改承辦人的時候，期限不要更新。
      If Cancel = False And Left(Trim(cboCP14.Text), 5) <> Trim(cboCP14.Tag) And _
         Not (Text1(23).Text <> "" And pa(1) = "FG") Then
         
            'Modified by Lydia 2016/06/21 改成模組
            'SetCP48
            Call PUB_GetFCPsetCP48(Me.Visible, pa, m_CP27, Text1(1), Text1(0), m_CP122, Text1(4), Text1(5), Text1(23), Text1(28), Combo2, Text1(12))
           'Added by Lydia 2024/03/08 因應內專支援機械組OA (含P案) 收文: 927其他翻譯
           If mTransKind = "＃" Then
              '承辦期限: 分案日+14日曆天
              Text1(23) = TransDate(CompDate(2, 14, strSrvDate(1)), 1)
           End If
           'end 2024/03/08
            'Added by Lydia 2018/09/12 翻譯分案->比對交稿期限和預設承辦期限,以最早期限為準
            If Text1(23).Text <> "" And mLimitDate <> "" Then
                If TransDate(mLimitDate, 1) < Text1(23).Text Then
                    Text1(23).Text = TransDate(mLimitDate, 1)
                End If
            'Added by Lydia 2018/09/28 如果沒有承辦期限
            ElseIf Text1(23).Text = "" And mLimitDate <> "" Then
                Text1(23).Text = TransDate(mLimitDate, 1)
            'end 2018/09/27
            End If
            'end 2018/09/12
      End If
      
      '依員工姓名抓取員工編號
      strText = GetPrjSalesNM_2(cboCP14.Text)
      If strText <> "" Then
         cboCP14.Text = strText & " " & cboCP14.Text
      Else
         '依員工編號抓取員工姓名
         strText = GetPrjSalesNM(Left(cboCP14.Text, 5))
         If strText <> "" Then
            cboCP14.Text = Left(cboCP14.Text, 5) & " " & strText
         End If
      End If
   End If
   
   If Trim(cboCP14.Text) = "" Then
      '若清除承辦人時核稿人也要清
      'Removed by Morgan 2022/2/7 核稿人改只顯示不可修改--Sharon
      'text1(22) = ""
      'Label5(12) = ""
      'end 2022/2/7
      Text1(0) = "": Label3(0) = ""
      'Cancel = True
      If cboCP14.Visible = True Then cboCP14.SetFocus
      Call CboCP14_GotFocus
      Exit Sub
   End If
   
   Cancel = False 'Added by Lydia 2018/09/12
End Sub
'2016/7/26 END

Private Sub cmdSetDate_Click()
   Dim strCP06 As String
   Text1(4) = "": Text1(5) = ""
   Text1(4).Tag = "" 'Add By Sindy 2015/12/16
'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
'   strExc(0) = "Select NVL(CF04,0) From CaseFee where cf01='" & pa(1) & "' and cf02='" & pa(9) & "' and cf03='" & Text1(1).Text & "'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      strCP06 = PUB_GetWorkDay(RsTemp.Fields(0) + 1)
'   End If
   '2008/8/27 modify by sonia 算本所期限不可抓casefee之工作天,改為系統日+6個工作天
   'strCP06 = Pub_GetHandleDay(pa(1), pa(9), Text1(1), , , Label3(8))
   'strCP06 = CompWorkDay(2, strCP06)
   'Modified by Morgan 2016/3/22 改12個工作天(公告通知函也要同步修改)
   'strCP06 = CompWorkDay(6, strSrvDate(1), 0)
   'Modified by Lydia 2016/06/17 改成模組
   'strCP06 = CompWorkDay(12, strSrvDate(1), 0)
   strCP06 = PUB_GetFCPsetDate(pa(1), Text1(1))
   'end 2016/3/22
   '2008/8/27 end
'end 2007/10/11
   If strCP06 <> "" Then
      Text1(4).Text = ChangeWStringToTString(strCP06)
      Text1(4).Tag = Text1(4).Text 'Add By Sindy 2015/12/16
      Text1(5).Text = Text1(4).Text
      'Add by Morgan 2008/9/3
      '當有點選設定期限時，核對已准專利的承辦期限=本所期限
      If Text1(1) = "926" Then
         Text1(23) = Text1(4).Text
      End If
   End If
End Sub

Private Sub Combo2_Click()
   If Combo2.ListIndex >= 0 Then
      Text1(23) = TransDate(PUB_GetWorkDay1(CompDate(2, Combo2.ItemData(Combo2.ListIndex), strSrvDate(1)), False), 1)
      If Text1(23) <> "" And Text1(4) <> "" And Val(Text1(23)) > Val(Text1(4)) Then
         Text1(23) = Text1(4)
      End If
   End If
End Sub

'Add by Morgan 2009/10/1
Private Sub Combo3_Click()
   If Combo3.ListIndex >= 0 Then
      '最大實審通知日(考慮有再審的情形)
      If Combo3.Tag = "" Then
         SetStartDate2Tag
      End If
      If Combo3.Tag <> "" Then
         SetDueDate Combo3.ItemData(Combo3.ListIndex), Combo3.Tag
      End If
   End If
End Sub

Private Sub Command2_Click(Index As Integer)
Dim strCP06 As String
Dim strCPM34 As String 'Add By Sindy 2021/4/29
Dim dbTfRate As Double, bolIsHigher As Boolean  'Added by Lydia 2021/07/29 判斷翻譯費折扣率＞30%
Dim m_StrTo As String, m_StrSub As String, m_StrCont As String 'Add by Sindy 2023/12/6

   Select Case Index
      Case 0
         Where1103ComeFrom Me, pa(1), pa(2), pa(3), pa(4)
      Case 1
         Set frm060101_2.fmParent = Me
         frm060101_2.Show
         Me.Hide
         
      Case 2 '確定
      
         '轉案號
         If Text1(2) <> "" And Text1(25) <> "" Then
            strExc(1) = Text1(2)
            strExc(2) = Text1(25)
            strExc(3) = Text1(26)
            If strExc(3) = "" Then strExc(3) = "0"
            strExc(4) = Text1(27)
            If strExc(4) = "" Then strExc(4) = "00"
            strExc(5) = Text1(1).Text '案件性質
            strExc(6) = Label3(1) '案件性質名稱
            strExc(7) = Text1(12) '收文日
            strExc(8) = Label3(8) '總收文號
            strExc(9) = pa(26)
            'edit by nickc 2007/02/05 不用 dll 了
            'If Not objLawDll.ChkSameCase(strExc) Then Exit Sub
            If Not ClsLawChkSameCase(strExc) Then Exit Sub
            'Added by Lydia 2020/08/18 更新相關卷號前,先檢查是否有重複
            If m_CP31 = "Y" Then
                If PUB_ChkUpdCR(pa(1), pa(2), pa(3), pa(4), strExc(1), strExc(2), strExc(3), strExc(4)) = False Then
                    Exit Sub
                End If
            End If
            'end 2020/08/18
            MsgBox "本次只更新轉本所案號的資料，不更新其他欄位的修改!!!", vbCritical
            
         '分案
         Else
            'Add by Morgan 2007/4/20
            '檢查下一程序是否有相同的案件性質但未勾選
            If MSHFlexGrid1.Rows > 1 Then
               With MSHFlexGrid1
                  For intI = 1 To .Rows - 1
                     If .TextMatrix(intI, 0) = "v" Then
                        Exit For
                     End If
                  Next
                  If intI = .Rows Then
                     For intI = 1 To .Rows - 1
                        If .TextMatrix(intI, 8) = Text1(1) Then
                           If MsgBox("有相同案件性質的下一程序未勾選，是否確定要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                              Exit Sub
                           End If
                           Exit For
                        End If
                     Next
                     
                  End If
               End With
            End If
            'end 2007/4/20
            
            'Add by Morgan 2007/10/18 檢查相關收文號
            If CheckRefNo = False Then
               Text1(6).SetFocus
               Exit Sub
            End If
            
            'Add By Sindy 2017/11/28 修正案時,檢查是否有收過431高速審查
            '                        自動填入相關總收文號==>PPH修正案
            If Text1(6) = "" And (Text1(1) = "204" Or Text1(1) = "203") Then
               If PUB_ChkCPExist(pa, "431", , strExc(10)) = True Then
                  Text1(6) = strExc(10)
               End If
            End If
            '2017/11/28 END
            
            '94.3.7 ADD BY SONIA
            If Text1(1) = 自請撤回 And Text1(6) = "" Then
               MsgBox "案件性質為自請撤回時，相關總收文號不可空白 !", vbCritical
               Exit Sub
            End If
            '94.3.7 END
            
'CANCEL BY SONIA 2014/7/25 不知為何要控管FCP-50343故取消
'            'ADD BY SONIA 2014/6/23
'            If text1(1) = "949" And text1(6) = "" Then
'               MsgBox "案件性質為寄中說時，相關總收文號不可空白 !", vbCritical
'               Exit Sub
'            End If
'            '2014/6/23 END
'END 2014/7/25
            
            If CheckPromAndCK = False Then Exit Sub
            
            'Modify By Sindy 2021/9/30
            If Text1(4).Tag <> Text1(4) Then
               If MsgBox("是否確定修改本所期限 ?", vbQuestion + vbYesNo) = vbNo Then
                  Text1(4) = Text1(4).Tag
                  Exit Sub
               End If
            End If
            If Text1(5).Tag <> Text1(5) Then
               If MsgBox("是否確定修改法定期限 ?", vbQuestion + vbYesNo) = vbNo Then
                  Text1(5) = Text1(5).Tag
                  Exit Sub
               End If
            End If
            
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
         
'Remove by Morgan 2006/4/28 不必再控制--靜芳
'           If (text1(1) = 請求公告 Or text1(1) = 延緩公告) And text1(6) = "" Then
'              MsgBox "案件性質為請求公告或延緩公告時，相關總收文號不可空白 !", vbCritical
'              Exit Sub
'           End If
'2006/4/28 end
            
            'Add by Morgan 2004/3/17
            '分割案提示
            If Text1(1) = "307" Then
               Erase m_stVar(): m_st416CP09 = "": m_stPA09 = ""
               m_bol435 = False 'Added by Morgan 2015/9/9 是否管制續行母案再審
               'Add by Morgan 2004/3/30
               Dim strTmp1(0 To 4) As String, strTmp(1 To 3) As String, i As Integer
               
               
               'FCP及P的國內案件,當專利種類為'發明'且案件性質為'分割'時,若無收文未取消收文之'實體審查'則顯示'此分割案尚未收文實體審查，期限為XXXXXX，請提醒智權人員 !!
               If txtDivCaseNo(1).Text = "" And txtDivCaseNo(2).Text = "" And txtDivCaseNo(3).Text = "" And txtDivCaseNo(4).Text = "" Then
                  If pa(1) = "FCP" And Text1(21) = "1" Then
                     MsgBox "國內發明分割案必須輸入分割母案本所案號！", vbCritical
                     txtDivCaseNo_GotFocus 1
                     txtDivCaseNo(1).SetFocus
                     Exit Sub
                  ElseIf MsgBox("本進度的案件性質為分割，確定不輸入分割母案本所案號？", vbExclamation + vbYesNo) = vbNo Then
                     txtDivCaseNo_GotFocus 1
                     txtDivCaseNo(1).SetFocus
                     Exit Sub
                  End If
               '檢查母案本所案號是否存在
               'Modified by Morgan 2012/11/8 改呼叫公用函數檢查
               'ElseIf CheckDivCase(m_stPA09) = False Then
               'Add by Lydia 2014/10/21 .begin
'               ElseIf PUB_CheckDivCase(txtDivCaseNo, pa, m_stPA09) = False Then
'                    txtDivCaseNo_GotFocus 1
'                    txtDivCaseNo(1).SetFocus
'                    Exit Sub
'                ElseIf pa(1) = "FCP" And Text1(21) = "1" Then
               Else
               
                 'Add by Lydia 2014/10/21 當分割案(307)與母案的專利種類不同，自動更新為相同種類
                 m_307PA08 = "": m_DCCode = 0
                 If pa(1) = "FCP" Then
                   m_DCCode = -1
                   pa(8) = Text1(21).Text  '將最初讀取的暫存陣列值變更為現在輸入，才能判斷
                 End If
                 
                 If PUB_CheckDivCase(txtDivCaseNo, pa, m_stPA09, m_307PA08, m_DCCode) = False Then
                    If pa(1) = "FCP" And Text1(1).Text = "307" And m_DCCode = 1 Then
                      Text1(21).Text = m_307PA08  '帶入母案的專利種類
                      pa(8) = Text1(21).Text  '將最初讀取的暫存陣列值變更為現在輸入，才能判斷
                      m_307PA08 = "": m_DCCode = 0
                      If PUB_CheckDivCase(txtDivCaseNo, pa, m_stPA09, m_307PA08, m_DCCode) = False Then '後續check
                        txtDivCaseNo_GotFocus 1
                        txtDivCaseNo(1).SetFocus
                        Exit Sub
                      End If
                
                    Else
                      txtDivCaseNo_GotFocus 1
                      txtDivCaseNo(1).SetFocus
                      Exit Sub
                    End If
                 End If
               End If
               
               If pa(1) = "FCP" And Text1(21) = "1" Then
                  'Add by Lydia 2014/10/21 .end
                  If m_stPA09 <> "000" Then
                     MsgBox "發明分割案的母案申請國家必須為台灣！", vbCritical
                     Exit Sub
                  End If
                  For i = 1 To 4
                     strTmp1(i) = txtDivCaseNo(i)
                  Next
                  'Modified by Morgan 2013/9/17
                  '若已收文435(續行母案再審)則不必再管制實審期限--靜芳
                  If PUB_ChkCPExist(pa(), "435", , m_st416CP09) = False Then
                     'Modified by Morgan 2015/9/9
                     'If PUB_GetDivCaseState(pa(), strSrvDate(1), True) = "N" Then
                     If m_PA163 = "" Then
                        m_PA163 = PUB_GetDivCaseState(strTmp1(), strSrvDate(1), True, True)
                        If m_PA163 = "" Then
                           intI = MsgBox("資訊不足無法判斷!!請問本案是否為初審階段提分割??", vbYesNoCancel + vbQuestion + vbDefaultButton3)
                           If intI = vbYes Then
                              m_PA163 = "Y"
                           ElseIf intI = vbNo Then
                              m_PA163 = "N"
                           Else
                              Exit Sub
                           End If
                        End If
                     End If
                     If m_PA163 = "N" Then
                        m_bol435 = True
                     'end 2015/9/9
                        MsgBox "此分割案尚未收文435(續行母案再審)，請提醒智權人員!!!", vbExclamation
                     Else
                     'end 2013/9/17
                     
                        '讀取實體審查得法定期限
                        If GetMoneyDate(4, m_stPA09, strTmp1, strTmp(1), strTmp(2), strTmp(3)) = True Then
                           If strTmp(3) <> "" Then
                              strTmp(3) = CompDate(2, 1, strTmp(3))
                              '法定期限
                              m_stVar(3) = PUB_Get416LawLimit(Text1(12), strTmp(3))
                              'Modified by Morgan 2014/11/20 外專改回舊規則
                              ''Added by Morgan 2014/10/29
                              'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                              '   m_stVar(0) = PUB_GetOurDeadline(m_stVar(3))
                              'Else
                              ''end 2014/10/29
                              'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
                              If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
                                 'Modify By Sindy 2021/4/27 + m_pAgreeOnDate
                                 m_stVar(0) = PUB_GetFCPOurDeadline(m_stVar(3), 4, , m_pAgreeOnDate)
                              Else
                              'end 2019/7/11
                                 '本所期限= 法定期限-4天
                                 'Modify by Morgan 2004/12/17 不用考慮假日
                                 m_stVar(0) = CompDate(2, -4, m_stVar(3))
                              End If 'Added by Morgan 2019/7/11
                              'End If 'Added by Morgan 2014/10/29
                              'end 2014/11/20
                              
                              '檢查有收'實體審查'否，有則抓收文號-->m_st416CP09
                              If PUB_Get416CP09(m_st416CP09, ChangeWStringToTString(m_stVar(0)), pa()) = False Then
                                 Exit Sub
                              End If
                           Else
                              MsgBox "無法讀取實體審查的法定期限！", vbCritical
                              Exit Sub
                           End If
                        Else
                           MsgBox "無法讀取實體審查的法定期限！", vbCritical
                           Exit Sub
                        End If
                        
                  'Added by Morgan 2013/9/17
                     End If
                  'Added by Morgan 2015/9/9
                  Else
                     m_bol435 = True
                  'end 2015/9/9
                  End If
                  'end 2013/9/17
                  
                  'Removed by Morgan 2022/5/4 續行母案再審改在分案時更新(分割未發文:同分割期限,分割已發文:發文日+4個月)--陳亭妙
                  ''Added by Morgan 2015/9/9
                  'if m_bol435 Then
                  '   '收文日/發文日+30天
                  '   m_stVar(3) = PUB_Get416LawLimit(Text1(12), Text1(12))
                  '   m_stVar(0) = CompDate(2, -4, m_stVar(3))
                  'End If
                  ''end 2015/9/9
                  'end 2022/5/4
                  
               End If
              
            End If
         
            'Add by Morgan 2004/3/23
            '專利分案當有顯示是否取消閉卷但未輸入'Y'時，提示並取消！
            If Text1(11).Visible = True And Trim(Text1(11).Text) <> "Y" Then
               'Add by Morgan 2009/9/9 退費且相關總收文號為實審或再審的不彈訊息 -靜芳
               strExc(1) = ""
               If Text1(1) = "908" Then
                  'modify by sonia 2021/10/4  再加續行母案再審435(FCP-064306)
                  '***** 加案件性質時，發文frm060104_1也要改
                  strExc(0) = "select 1 from caseprogress where cp09='" & Text1(6) & "' and cp10 in ('416','107','435')"
                  'Added by Morgan 2013/8/22 加判斷再審的延期
                  'modify by sonia 2021/10/4  再加續行母案再審435(FCP-064306)
                  strExc(0) = strExc(0) & " union select 2 from caseprogress,nextprogress where cp09='" & Text1(6) & "' and cp10='404' and cp84>0 and np01(+)=cp43 and to_char(np22)=cp30 and np07 in ('107','435')"
                  strExc(0) = strExc(0) & " union select 3 from caseprogress a,caseprogress b where a.cp09='" & Text1(6) & "' and a.cp10='404' and a.cp84>0 and b.cp09(+)=a.cp43 and b.cp10 in ('107','435')"
                  'end 2013/8/22
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strExc(1) = "N"
                  End If
               'Added by Morgan 2023/4/27
               '已閉卷案【回代】及【告代】仍可進行分案--Sharon
               ElseIf Text1(1) = "901" Or Text1(1) = "902" Then
                  strExc(1) = "N"
               'end 2023/4/27
               End If
               If strExc(1) = "" Then
                  MsgBox "本案已閉卷且未取消!!!", vbCritical
                  Text1(11).SetFocus
                  Exit Sub
               End If
            End If
            
            'Add by Morgan 2006/4/25
            If Text1(1) = "202" Then
               With MSHFlexGrid1
               If .Recordset.RecordCount > 0 Then
                  For intI = 1 To .Rows - 1
                     '若有勾選補文件時跳離
                     If .TextMatrix(intI, 0) = "v" And .TextMatrix(intI, 8) = "202" Then
                        Exit For
                     End If
                  Next
                  '未勾選
                  If intI = .Rows Then
                     If MsgBox("請確認是否要勾選下一程序？", vbYesNo + vbDefaultButton1) = vbYes Then
                        Exit Sub
                     End If
                  End If
               End If
               End With
               If textCP64 = "" Then
                  MsgBox "【補文件】進度備註不可空白！", vbExclamation
                  Exit Sub
               Else
                  
               End If
            End If
            '2006/4/25 end
            
            'Add by Morgan 2007/4/3
            If cmdSetDate.Visible = True Then
               If DBDATE(Text1(4).Text) <> PUB_GetWorkDay1(Text1(4).Text, True) Then
                  If MsgBox("本所期限非工作天，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                     Text1(4).SetFocus
                     Exit Sub
                  End If
               End If
            End If
            'end 2007/4/3
            
'cancel by sonia 2016/5/26 先收實審但不發文, 後續再收主動修正, 則會因此點限制而不能分案
'            'Add by Morgan 2007/6/12 實審與(分割或主動修正)同時收文時提醒承辦人必須同一人
'            If m_CP27 = "" And text1(0) <> "" And (text1(1) = "203" Or text1(1) = "307" Or text1(1) = "416") Then
'               If AssignNote = False Then
'                  text1(0).SetFocus
'                  Text1_GotFocus 0
'                  Exit Sub
'               End If
'            End If
'            'end 2007/6/12
'end 2016/5/26
            
            'Add by Morgan 2007/8/2
            'Remove by Lydia 2016/06/21 改成模組PUB_CheckFCPtxtValidate
            'If text1(1) = "201" And text1(0) <> "" And text1(0) < "F" Then
            '   If MsgBox("承辦人為外專工程師之員工編號是否要繼續?", vbYesNo + vbDefaultButton2) = vbNo Then
            '      text1(0).SetFocus
            '      Text1_GotFocus 0
            '      Exit Sub
            '   End If
            'End If
            
            'Add by Morgan 2007/8/29 實審分案時若有主動修正未發文時提醒更新期限
            'Remove by Lydia 2016/06/21 改成模組PUB_CheckFCPtxtValidate
            'm_203CP09 = ""
            'If text1(1) = "416" And text1(5) <> "" And pa(10) <> "" And m_CP27 = "" Then
            '   'Modify by Morgan 2008/10/9 改判斷非來函通知的補充說明206
            '   'strExc(0) = "select cp09 from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='203' and cp27 is null and cp57 is null"
            '   strExc(0) = "select cp09,cpm03 from caseprogress,nextprogress a,casepropertymap where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in ('203','206') and cp27 is null and cp57 is null" & _
            '      " and np02(+)=cp43 and np07(+)=cp10 and np09 is null and cpm01(+)=cp01 and cpm02(+)=cp10"
            '   intI = 1
            '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            '   If intI = 1 Then
            '      '主動修正收文號
            '      m_203CP09 = RsTemp.Fields(0)
            '      MsgBox "有收文" & RsTemp.Fields(1) & "，期限將改與實審相同！", vbInformation
            '   End If
            'End If
            'end 2007/8/29
            
            'Add by Morgan 2008/3/18 年費若無申請程序且基本檔未設定申請人不出名時提醒存檔時將自動上N(剔除96.6.12以後有非林律師出名程序者)
            m_PA143 = pa(143)
            If Text1(1) = "605" And pa(143) = "" Then
               If PUB_ChkCPExist(pa(), "401", 1) = False Then 'Added by Morgan 2020/3/11 剔除有變更未發文者--
                  'Modify by Morgan 2008/5 不管是否中間案件改判斷有非年費且非林律師出名的AB類發文
                  'strExc(0) = "select cp09 from caseprogress" & _
                     " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                     " and cp10 in ('101','102','103')" & _
                     " union select cp09 from caseprogress" & _
                     " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                     " and cp27>=20070612 and cp110<>'65002'"
                  strExc(0) = "select cp09 from caseprogress" & _
                     " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                     " and cp09<'C' and cp10<>'605' and cp110<>'65002' and cp27>0" & _
                     " union select cp09 from caseprogress" & _
                     " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                     " and cp27>=20070612 and cp110<>'65002'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 0 Then
                     'Modified by Morgan 2025/10/20 改欄位說明--Sharon
                     'MsgBox "年費申請人是否出名應為【N】，存檔時將自動更新！"
                     MsgBox "【年費本所是否出名】應為【N】，存檔時將自動更新！"
                     'end 2025/10/20
                     m_PA143 = "N"
                  End If
               End If
            End If
            
            'Add by Morgan 2008/3/18
            '當案件性質為主動修正(203)、更正(402)、專利權延長(415)或舉發答辯(804)時若未辦或不辦重新委任時提醒
            If Text1(1) = "203" Or Text1(1) = "402" Or Text1(1) = "415" Or Text1(1) = "804" Then
               If PUB_Check928NotOk(pa) = True Then
                  MsgBox "本案下一程序有重新委任之補文件未辦理！", vbInformation, "注意"
               End If
            End If
            
            'Add by Sindy 2011/3/11
            '當案件性質為檢視中說(209)、製作中說(210)、創作說明(223)時, 若有會稿(924)尚未分案, 請相互提示另一案件性質尚未分案
            'Modified by Morgan 2013/11/6 +235核對中說格式
            'MODIFY BY SONIA 2014/6/23 +949寄中說
            If Text1(1) = "209" Or Text1(1) = "235" Or Text1(1) = "210" Or Text1(1) = "223" Or Text1(1) = "924" Or Text1(1) = "949" Then
               'Added by Lydia 2018/05/31 排除檢視中說(209)和核對中說格式(235),會自動上會稿
               If Trim(cboCP14.Text) <> "" And InStr(cAutoCP924, Text1(1).Text) > 0 Then
                   strExc(1) = "'209','210','223','235','949'"
               Else
                   strExc(1) = "'209','210','223','235','924','949'"
               End If
               'Modified by Lydia 2018/05/31 cp10 in('209','210','223','235','924','949') => cp10 in (" & strExc(1) & ")
               strExc(0) = "select cpm03 from caseprogress,casepropertymap where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
                  " and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in (" & strExc(1) & ")" & _
                  " and cp14 is null and cp10<>'" & Text1(1) & "' and cp01=cpm01(+) and cp10=cpm02(+)"
               'end 2018/05/31
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                     If Not IsNull(RsTemp.Fields("cpm03")) Then
                        MsgBox "本案尚有" & RsTemp.Fields("cpm03") & "未分案！", vbInformation, "注意"
                     End If
                     RsTemp.MoveNext
                  Loop
               End If
            End If
            '2011/3/11 End
         End If
         'Add by Lydia 2015/01/28 判斷收文的案件性質在轉本所案號,是否有同樣的下一程序
         If Text1(2) <> "" And Text1(25) <> "" Then
            strExc(0) = "select NP01,NP07 from nextprogress where np02='" & Text1(2) & "' and np03='" & Text1(25) & "' " & _
               "and np04='" & IIf(Text1(26) = "", "0", Text1(26)) & "' and np05='" & IIf(Text1(27) = "", "00", Text1(27)) & "' " & _
               "and np07='" & Text1(1) & "' AND (NP06<>'Y' OR NP06 IS NULL) "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If MsgBox("下一程序檔有相同性質，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
                  Text1(2).SetFocus
                  Exit Sub
               End If
            End If
         End If
         'end 2015/01/28
         
         'Added by Lydia 2017/05/17 提醒有相似案
         If txtTF19.Visible = True And Val(txtTF19) > 0 Then
            MsgBox Label3(1).Caption & "有相似案" & txtTF19.Tag & " !"
         End If
         'end 2017/05/17
         
         'Added by Lydia 2024/04/10 分案和工作進度維護點選不可查閱工程師需要彈訊息
         If Left(Trim(cboCP14), 1) > "6" And Left(Trim(cboCP14), 1) < "F" And cboCP14.Text <> cboCP14.Tag Then
            If PUB_ChkCufaByCaseNo(Trim(Left(cboCP14, 6)), Me.Name, pa(1) & pa(2) & pa(3) & pa(4), "2") = False Then
               Exit Sub
            End If
         End If
         'end 2024/04/10
         
         'Added by Lydia 2017/12/11 FCP案件命名電子化：中說輸入相關設定
         'Modified by Lydia 2018/08/07 改成先讀命名作業
'         If strSrvDate(1) >= FCP案件命名啟用日 And pa(1) = "FCP" And InStr(FcpTctPtys, text1(1)) > 0 Then
'            strExc(0) = "select TCT01,TCT10,TCT27,TCT28 from TRANSCASETITLE,caseprogress where TCT01=cp09(+) " & _
'                              "and cp01= '" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and nvl(tct05,0)> 0 "
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'                strExc(1) = "" & RsTemp.Fields("TCT27")
'                If strExc(1) <> "" Then
'                    strExc(2) = ""
'                    strExc(2) = Pub_GetTct27ID("" & RsTemp.Fields("TCT10"), strExc(1), "" & RsTemp.Fields("TCT28"), strExc(3))
'
'                    If strExc(2) <> "" And Trim(Left(cboCP14.Text, 6)) <> strExc(2) And Trim(cboCP14.Text) <> "" Then
'                        If MsgBox("命名記錄的預設承辦人為" & strExc(2) & " " & strExc(3) & "，是否繼續存檔？", vbYesNo) = vbNo Then
'                           cboCP14.SetFocus
'                           Exit Sub
'                        End If
'                    End If
'                End If
'            End If
'         End If
'         'end 2017/12/11
         
         'Added by Lydia 2023/04/19 外專翻譯分案承辦人不得為翻譯社及外譯人員
         If pa(1) = "FCP" And Text1(1) = "927" And Trim(Left(cboCP14.Text, 6)) <> cboCP14.Tag Then
            'Modified by Lydia 2025/07/01 增加例外案件設定InStr(m_str所內譯例外,  pa(1) & pa(2) & pa(3) & pa(4)) = 0 And
            If InStr(m_str所內譯例外, pa(1) & pa(2) & pa(3) & pa(4)) = 0 And (InStr(m_str所內譯, ChangeCustomerL(pa(26))) > 0 Or InStr(m_str所內譯, ChangeCustomerL(pa(75))) > 0) Then
               'Modified by Lydia 2024/04/18 已與F5523宗家澔簽約外翻---Sharon口述
               If Trim(Left(cboCP14.Text, 1)) = "F" And InStr("F5523,", Trim(Left(cboCP14.Text, 6))) = 0 Then
                   strExc(2) = PUB_GetMapID(Trim(Left(cboCP14.Text, 6)), 1)
                   If strExc(2) = "" Then
                       'Modified by Lydia 2025/06/05 「BASF集團公司為申請人的所有專利案件」改為「本案所有」
                       MsgBox "本案所有相關翻譯事宜（201新案翻譯/927其他翻譯）皆須由本所工程師翻譯/處理，不得委外。", vbExclamation + vbOKOnly
                       cboCP14.SetFocus
                       Exit Sub
                   End If
               End If
            End If
         End If
         'end 2023/04/19
         'Modified by Lydia 2021/08/09 判斷有修改承辦人才做檢查 Trim(Left(cboCP14.Text, 6)) <> cboCP14.Tag
         If m_TCT01 <> "" And pa(1) = "FCP" And Text1(1) = "201" And Trim(Left(cboCP14.Text, 6)) <> cboCP14.Tag Then
             'Added Lydia 2023/04/19 外專翻譯分案承辦人不得為翻譯社及外譯人員
             'Modified by Lydia 2025/07/01 增加例外案件設定InStr(m_str所內譯例外,  pa(1) & pa(2) & pa(3) & pa(4)) = 0 And
             If InStr(m_str所內譯例外, pa(1) & pa(2) & pa(3) & pa(4)) = 0 And (InStr(m_str所內譯, ChangeCustomerL(pa(26))) > 0 Or InStr(m_str所內譯, ChangeCustomerL(pa(75))) > 0) Then
                'Modified by Lydia 2024/04/18 已與F5523宗家澔簽約外翻---Sharon口述
                If Trim(Left(cboCP14.Text, 1)) = "F" And InStr("F5523,", Trim(Left(cboCP14.Text, 6))) = 0 Then
                    strExc(2) = PUB_GetMapID(Trim(Left(cboCP14.Text, 6)), 1)
                    If strExc(2) = "" Then
                        'Modified by Lydia 2025/06/05 「BASF集團公司為申請人的所有專利案件」改為「本案所有」
                        MsgBox "本案所有相關翻譯事宜（201新案翻譯/927其他翻譯）皆須由本所工程師翻譯/處理，不得委外。", vbExclamation + vbOKOnly
                        cboCP14.SetFocus
                        Exit Sub
                    End If
                End If
                dbTfRate = 0
             Else
             'end 2023/04/19
               'Added by Lydia 2021/07/29 判斷翻譯費折扣率
                dbTfRate = PUB_GetTransFeeRate(pa(1), pa(2), pa(3), pa(4), , bolIsHigher, True)
             End If    'Added by 2023/04/19
             '控制翻譯費折扣率＞30%客戶案件之承辦人只能為所內人員上班譯編號。
             If dbTfRate > 30 Then
                 If Left(cboCP14, 1) = "F" Then
                     'Modified by Lydia 2021/12/02 因為案件經過呈報王協理後可以"非所內人員上班譯"，所以改成彈訊息詢問---Sharon 口頭協商
                     'MsgBox "該案件之承辦人只能為所內人員上班譯編號！", vbExclamation
                     If MsgBox("該案件之承辦人只能為所內人員上班譯編號，是否繼續分案？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                        cboCP14.SetFocus
                         Exit Sub
                     End If 'Added by Lydia 2021/12/02
                 End If
             ElseIf bolIsHigher = True Then  '折扣率＞30%但是例外控制的客戶
                  '不受限
             End If
             'end 2021/07/29

             'If m_TCT27 <> "" Then 'Remove by Lydia 2018/09/27
                   'Modified by Lydia 2018/09/27 認領翻譯人員
                   'strExc(2) = Pub_GetTct27ID(m_TCT10, m_TCT27, m_TCT28, strExc(3))
                   strExc(2) = Pub_GetTct27ID(m_TCT10, m_TCT27, m_TCT28, Me.Label3(8).Caption, strExc(3))
                   If strExc(2) <> "" And Trim(Left(cboCP14.Text, 6)) <> strExc(2) And Trim(cboCP14.Text) <> "" Then
                        'Modified by Lydia 2018/09/27
                        'If MsgBox("命名記錄的預設承辦人為" & strExc(2) & " " & strExc(3) & "，是否繼續存檔？", vbYesNo) = vbNo Then
                        If MsgBox("認領翻譯人員為" & strExc(2) & " " & strExc(3) & "，是否繼續存檔？", vbYesNo) = vbNo Then
                           cboCP14.SetFocus
                           Exit Sub
                        End If
                   'Added by Lydia 2018/09/27
                   Else
                        If strExc(3) <> "" And InStr(strExc(3), "下班") > 0 And Trim(Left(cboCP14.Text, 1)) <> "F" Then
                            If MsgBox("認領翻譯人員為" & strExc(2) & " " & strExc(3) & "，輸入非下班譯者編號，是否繼續存檔？", vbYesNo) = vbNo Then
                               cboCP14.SetFocus
                               Exit Sub
                            End If
                        End If
                        'end 2018/09/27
                   End If
             'End If 'Remove by Lydia 2018/09/27
         End If
         'end 2018/08/07

         'Added by Lydia 2018/08/08 若輸入下班翻譯在存檔前先提醒"只能上班翻譯是否繼續存檔"
         'Modified by Lydia 2025/03/13 改用模組取得
         'If Left(mTransKind, 1) = "Y" And text1(1) = "201" And Trim(Left(cboCP14.Text, 1)) = "F" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, Trim(Left(cboCP14.Text, 6))) = 0 Then
         If Left(mTransKind, 1) = "Y" And Text1(1) = "201" And Trim(Left(cboCP14.Text, 1)) = "F" And InStr(Pub_SetF51Order("F", ""), Trim(Left(cboCP14.Text, 6))) = 0 Then
             If MsgBox("只能上班翻譯：" & Mid(mTransKind, 3) & vbCrLf & "，是否繼續存檔？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                  cboCP14.SetFocus
                  Exit Sub
             End If
         End If
         'end 2018/08/07
         
         'Added by Lydia 2021/04/14 外專翻譯承辦及核稿期限控管：
         'Modified by Lydia 2025/03/13 改用模組取得
         'If pa(1) = "FCP" And text1(1) = "201" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, Trim(Left(cboCP14.Text, 6))) = 0 Then
         If pa(1) = "FCP" And Text1(1) = "201" And InStr(Pub_SetF51Order("F", ""), Trim(Left(cboCP14.Text, 6))) = 0 Then
             '工程師認領翻譯時，查詢該認領人員，新案翻譯未上完稿日案件,請彈提醒: 尚未完稿案件FCPxxxx , 承辦期限
             strExc(4) = Pub_GetEngEP09List(Trim(Left(cboCP14.Text, 6)))
             If strExc(4) <> "" Then
                 MsgBox "尚未完稿案件：" & strExc(4), vbCritical
             End If
         End If
         'end 2021/04/14
                          
         'Add By Sindy 2021/4/29 主管機關期限
         CheckOC3
         strCPM34 = ""
         strSql = "select cpm34 from casepropertymap where cpm01='" & pa(1) & "' and cpm02='" & Text1(1) & "'"
         AdoRecordSet3.CursorLocation = adUseClient
         AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If AdoRecordSet3.RecordCount > 0 Then
            strCPM34 = "" & AdoRecordSet3.Fields(0)
         End If
'分案時:
'(1) 若有法限，需與承辦確認是否為需掛法限
'(2) 若修改本所期限，自動備註: 修改本所期限為yyy/mm/dd(本所期限)
         'Add By Sindy 2021/4/29 原有法定期限，若將法定期限拿掉，增加提醒【是否確定刪除法定期限，本所期限將變更為承辦期限+5個工作天】
         If Val(m_CP27) = 0 And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
            If strCPM34 = "Y" And Val(Text1(5).Text) = 0 Then
               If MsgBox("此案件性質屬有主管機關期限，確定沒有法定期限嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                  Text1(5) = m_strCP07
                  Text1(5).SetFocus
                  Exit Sub
               End If
            ElseIf strCPM34 = "N" And Val(Text1(5).Text) > 0 Then
               If MsgBox("此案件性質屬非主管機關期限，確定有法定期限嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                  Text1(5).SetFocus
                  Exit Sub
               End If
            ElseIf strCPM34 = "N" Then
               If Val(Text1(5).Tag) > 0 And Val(Text1(5).Text) = 0 Then
                  If MsgBox("是否確定刪除法定期限，本所期限將變更為承辦期限+5個工作天" & vbCrLf & "，是否繼續存檔？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                     Text1(5).SetFocus
                     Exit Sub
                  Else
                     Text1(4).Text = TransDate(PUB_GetFCPOurDeadline(DBDATE(Text1(23)), , , , "N"), 1)
                  End If
               End If
            End If
         End If
         '2021/4/29 END
         
         'Add By Sindy 2019/10/17
         If pa(1) = "FCP" And Text1(1) = "201" Then
            '會稿收文有掛本所和法定期限
            strExc(0) = "SELECT cp09,cp06 FROM CASEPROGRESS WHERE cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                        " and cp10='924' and cp06>0 and cp07>0 and cp27 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strCP06 = RsTemp.Fields("cp06")
               '本所期限往前推7天
               strCP06 = DBDATE(DateAdd("d", -7, ChangeWStringToWDateString(strCP06)))
               '分案新案翻譯時,若有會稿且會稿有掛本所和法定期限
               '新案翻譯承辦期限大於會稿本所期限往前7天(例如會稿本所6/20,則新案翻譯承辦期限不可大於6/13),
               '若大於則彈提醒 "注意會稿期限為:XXXX,承辦期限是否確定?"
               If DBDATE(Text1(23)) > strCP06 Then
                  If MsgBox("注意會稿期限為:" & ChangeWStringToTDateString(strCP06) & "，承辦期限是否確定？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                     Text1(23).SetFocus
                     Exit Sub
                  End If
               End If
            End If
         End If
         '2019/10/17 END
         
         'Added by Lydia 2020/01/21 所有系統別分案，若該收文號存在於下一程序(NP24)，若修改案件性質則彈訊息
         'Modified by Lydia 2020/11/18 debug-原.Tag已改變
         'If Text1(1).Tag <> Text1(1).Text Then
         If m_strCP10 <> Text1(1).Text Then
             If Pub_CheckNP24Exists(Label3(8).Caption) = True Then
             End If
         End If
         'end 2020/01/21
         
        'Added by Lydia 2020/06/19 法律所案源收文：C類案源的案件性質若 "是否需要法律所配合"設定與來不同時提醒。
        If strSrvDate(1) >= 法律所案源收文啟用日 And pa(1) = "FCP" Then
             strExc(1) = "" 'Added by Lydia 2021/07/12 清空預設
             'Modified by Lydia 2020/07/23 重新整理: 因為案源收文已設定不可變更案件性質和申請國家,所以只要判斷非案源收文
             If m_LOS02 = "" And Text1(1).Tag <> Text1(1).Text Then
                strExc(1) = PUB_GetLOSkind(pa(1), Text1(1).Text, "000")
                strExc(1) = Replace(strExc(1), "P", "")
                '準備程序在輸入接洽單已決定是否為案源的補收款, 所以不用另外判斷
                If strExc(1) <> "" Then
                     MsgBox "收文不可修改為法務案源的案件性質！", vbCritical
                     Exit Sub
                End If
             End If
             'end 2020/07/23
        
             If m_LOS01 = "" And m_LOS07 = "" And FraLOS.Visible = True Then
                If (Left(strExc(1), 1) = "C" And m_LOS15 = "" And txtLOSagree = "Y") Or (Left(strExc(1), 1) = "C" And m_LOS15 <> "" And txtLOSagree <> "Y") _
                     Or (strExc(1) = "" And Left(m_LOS02, 1) = "C" And m_LOS15 <> "" And txtLOSagree <> "Y") Then
                   If MsgBox(" ""是否需要法律所配合"" 與接洽單不同，是否繼續作業？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                       txtLOSagree.SetFocus
                       txtLOSagree_GotFocus
                       Exit Sub
                   End If
                End If
             End If
        End If
        'end 2020/06/19
        
         'Add By Sindy 2021/4/20 檢查指定送件日相關欄位
         If Val(Text1(30).Text) > 0 And Option1(0).Visible = True Then
            'Modify By Sindy 2021/10/20 + And Option1(2).Value = False
            If Option1(0).Value = False And Option1(1).Value = False And Option1(2).Value = False Then
               MsgBox "有輸入指定送件日，當天或之前或之後請擇一。", vbExclamation
               Exit Sub
            End If
         Else
            Option1(0).Value = False
            Option1(1).Value = False
            Option1(2).Value = False 'Add By Sindy 2021/10/20
         End If
         '2021/4/20 END
         
         'Added by Lydia 2022/09/29 新增工程師分組控管
         bUpdPA150 = False
         If Trim(Left(cboCP14.Text, 6)) <> cboCP14.Tag And InStr("927,", Text1(1)) = 0 Then '除了「其他翻譯927」
             'Added by Lydia 2022/12/21 基本檔無工程師組別需要預設組別; ex.FG-001437
             If ((pa(1) = "FCP" And pa(150) = "") Or (pa(1) = "FG" And pa(79) = "")) And PUB_GetST03(Trim(Left(cboCP14.Text, 6))) = "F21" Then
                'Modified by Lydia 2023/07/27 排除尚未認領階段; ex.FCP-070028在取消暫不認領(TCN16)前先分案回代902
                'bUpdPA150 = True
                If m_TCT01 = "" Then
                   bUpdPA150 = True
                Else
                   strExc(0) = "select tct04,tcn23,tcn16 from transcasetitle, trackingcasename where tct01='" & m_TCT01 & "' and tct01=tcn05(+) "
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                      If "" & RsTemp.Fields("tcn16") = "Y" Or ("" & RsTemp.Fields("tcn23") <> "" And "" & RsTemp.Fields("tct04") = "") Then
                         '本案處於暫不認領/認領階段，請勿變更工程師組別!
                      Else
                         bUpdPA150 = True
                      End If
                   End If
                End If
                'end 2023/07/27
             Else
             'end 2022/12/21
                '當承辦人為工程師時，所輸入的工程師組別與原來的工程師組別不一致時，談視窗詢問：是否變更工程師組別，是: 工程師組別改為此次輸入的工程師組別／否: 維持原來的工程師組別。
                If PUB_GetStaffST16(Trim(Left(cboCP14.Text, 6))) <> PUB_GetStaffST16(Trim(Left(cboCP14.Tag, 6))) And _
                    PUB_GetST03(Trim(Left(cboCP14.Text, 6))) = "F21" And PUB_GetST03(Trim(Left(cboCP14.Tag, 6))) = "F21" Then
                    If MsgBox("是否變更工程師組別？", vbExclamation + vbYesNo + vbDefaultButton2, "工程師分組控管") = vbYes Then
                        bUpdPA150 = True
                    End If
                End If
             End If 'Added by Lydia 2022/12/21
         End If
         'end 2022/09/29
         
         ' 90.07.06 modify by louis
         ' 設定滑鼠游標為等待狀態
         Screen.MousePointer = vbHourglass
         If FormSave = False Then
            ' 設定滑鼠游標為預設
            Screen.MousePointer = vbDefault
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         End If
         ' 設定滑鼠游標為預設
         Screen.MousePointer = vbDefault
         If Trim(Text1(2)) & Trim(Text1(25)) & Trim(Text1(26)) & Trim(Text1(27)) <> "" Then
            strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(Label3(9))
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If RsTemp.Fields(0) < 1 Then
               MsgBox "原本所案號 " & pa(1) & pa(2) & pa(3) & pa(4) & "已無案件進度資料，請通知收文人員刪號！", vbInformation
            Else
               MsgBox "原本所案號為 " & pa(1) & pa(2) & pa(3) & pa(4) & "，請自行更新原本所案號之下一程序資料 !", vbInformation
            End If
         End If
         
         'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之年費605，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。
         If strMsgCloseCancel <> "" Then
            MsgBox "已還原「" & strMsgCloseCancel & "」期限", vbInformation, "取消閉卷"
         End If
   
         'Added by Lydia 2018/06/15 翻譯分案無紙化:設定承辦人後，自動發mail
         'Modified by Lydia 2018/06/28 判斷有變更
         If TypeName(m_PrevForm) = "frm060122" And Trim(cboCP14.Text) <> "" And Trim(Left(cboCP14.Text, 6)) <> cboCP14.Tag Then
                m_PrevForm.nKeyNo = Trim(Left(cboCP14.Text, 6))
                m_PrevForm.bolNextDone = True
         End If
         'end 2018/06/15
         
         'Add By Sindy 2023/12/4
         If (Text1(1) = 準備程序 Or Text1(1) = 言詞辯論) And Text1(4) <> "" And Text1(5) <> "" _
            And pa(9) = "000" Then
            '取得更新後的本所期限
            m_strCP06Update = GetCP06(Me.Label3(8).Caption)
            '取得更新後的法定期限
            m_strCP07Update = GetCP07(Me.Label3(8).Caption)
            'Modify By Sindy 2025/1/13 增加第一次分案時也要發Mail通知
            If m_CP122 = "" Or (Text1(0).Text <> m_CP14) Or (Text1(4).Text <> Text1(4).Tag) Or (Text1(5).Text <> Text1(5).Tag) Then
               strSql = "select CP08,CP64 from CASEPROGRESS where CP09='" & Text1(6) & "'"
               CheckOC
               adoRecordset.CursorLocation = adUseClient
               adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset.RecordCount > 0 Then
                  m_CP43CP08 = CheckStr(adoRecordset.Fields(0))
                  m_CP43CP64 = CheckStr(adoRecordset.Fields(1))
               End If
               'Modify By Sindy 2023/12/8 法律所調整內專行政訴訟開庭通知之系統通知信也請一併轉陳亮之; 商標一併調整
               'Modified by Lydia 2024/10/30 串法律所案源資料，抓出法律所案件有承辦人且收文日最大的進度，抓承辦人及所有出庭律師。
               'm_StrTo = Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & m_CP13 & IIf(Text1(0) <> "", ";" & Text1(0), "")
               m_StrTo = PUB_GetLosCL02list(pa(1), pa(2), pa(3), pa(4))
               'Modify By Sindy 2025/5/8 系統自動發信副本增加: 程序、程序主管
               m_StrTo = IIf(m_StrTo <> "", m_StrTo & ";", "") & Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & m_CP13 & IIf(Text1(0) <> "", ";" & Text1(0), "") & _
                         ";" & m_NA16Na79 & ";" & GetST52SelfList(m_NA16Na79)
               'end 2024/10/30
               
               m_StrSub = "開庭通知--分案案件：" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)
               m_StrCont = "本所案號：" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & vbCrLf & _
                           "案件名稱：" & Me.Combo1.Text & vbCrLf & _
                           "案件性質：" & Me.Label3(1).Caption & vbCrLf & _
                           "申請人　：" & Me.Label5(2).Caption & vbCrLf & _
                           "承辦人　：" & Me.Label3(0).Caption & vbCrLf & _
                           "智權人員　：" & Me.Label5(10).Caption & vbCrLf & _
                           "法定期限：" & DBYEAR(m_strCP07Update) - 1911 & " 年 " & DBMONTH(m_strCP07Update) & " 月 " & DBDAY(m_strCP07Update) & " 日 " & vbCrLf & _
                           "時間地點：" & m_CP43CP64 & vbCrLf & _
                           "法院案號：" & m_CP43CP08
               PUB_SendMail strUserNum, m_StrTo, Label3(8), m_StrSub, m_StrCont
            End If
         End If
         '2023/12/4 END
         
         If IntNow <> IntTot Then
            GetData IntNow
         Else
            ' 90.07.06 modify by louis
            ' 設定滑鼠游標為等待狀態
            If TypeName(m_PrevForm) = "frm060101" Then 'Added by Lydia 2018/05/21 回到分案前畫面
                Screen.MousePointer = vbHourglass
                ' 90.07.06 modify by louis (重新搜尋資料)
                frm060101.Show
                frm060101.RefreshData
                frm060101.Show
                ' 設定滑鼠游標為預設
                Screen.MousePointer = vbDefault
            End If
            Unload Me
         End If
      Case 3 'Memo by Lydia 2018/05/21 回前畫面
         If TypeName(m_PrevForm) = "frm060101" Then 'Added by Lydia 2018/05/21 回到分案前畫面
            frm060101.Show
            frm060101.ComUCase2
            frm060101.Show
         End If
         Unload Me
      
      Case 5 'Memo by Lydia 2018/05/21 下一筆
         If IntNow <> IntTot Then
            GetData IntNow
         Else
            If TypeName(m_PrevForm) = "frm060101" Then 'Added by Lydia 2018/05/21 回到分案前畫面
                frm060101.Show
                frm060101.ComUCase2
            End If
            Unload Me
         End If
         
      Case 6 '多國案卷號 Added by Morgan 2021/2/25
         frm1104.intWhereComeFrom = 1
         Set frm1104.m_form = Me
         frm1104.Show
         frm1104.txtSystem = pa(1)
         frm1104.txtCode(0) = pa(2)
         frm1104.txtCode(1) = pa(3)
         frm1104.txtCode(2) = pa(4)
         frm1104.GetRelation
         Me.Hide
         
   End Select
End Sub

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

Private Function FormSave() As Boolean
   Dim strTxt(1 To 20) As String, intStep As Integer, strTmp(1 To 3) As String
   Dim j As Integer, i As Integer
   Dim strST15 As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim sPA() As String
   Dim sSP() As String
   Dim m_CP16 As Long
   Dim m_CP17 As Long
   'Modified by Morgan 2017/4/11 點數會有小數
   'Dim m_CP18 As Long
   Dim m_CP18 As String
   'end 2017/4/11
   Dim stUpdate As String
   Dim st307Msg As String
   Dim arrInv
   Dim strTo As String, strSubject As String, strContent As String 'Add By Sindy 2015/12/15
   Dim strTemp As String, strEP08 As String 'Add By Sindy 2015/12/14
   Dim strDivCaseNo(1 To 4) As String 'Add By Sindy 2018/5/9 母案
   Dim msgTxt As String 'Added by Lydia 2018/05/31 存檔後彈訊息
   Dim str307CP06 As String, str307CP07 As String 'Added by Morgan 2021/8/20
   
   FormSave = True
   
On Error GoTo CheckingErr

   cnnConnection.BeginTrans
   
   intStep = 1
   '若有輸入轉本所案號
   If Me.Text1(2).Text <> "" And Me.Text1(25).Text <> "" Then
   
      'Modify by Morgan 2010/12/28 要先新增基本檔,否則紀錄原FC代理人的 Trigger 會錯
      '判斷是否新增專利或服務業務基本案
      Select Case pa(1)
         Case "P", "CFP", "FCP":
            StrSQLa = "SELECT * FROM PATENT WHERE " & ChgPatent(Me.Text1(2).Text & Me.Text1(25).Text & Me.Text1(26).Text & Me.Text1(27).Text)
         Case Else:
            StrSQLa = "SELECT * FROM SERVICEPRACTICE WHERE " & ChgService(Me.Text1(2).Text & Me.Text1(25).Text & Me.Text1(26).Text & Me.Text1(27).Text)
      End Select
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount <= 0 Then
         Select Case pa(1)
            Case "P", "CFP", "FCP":
               ReDim sPA(1 To TF_PA) As String
               
               If PUB_ReadPatentData(sPA(), pa(1), pa(2), pa(3), pa(4)) Then
                  sPA(1) = Me.Text1(2).Text
                  sPA(2) = Me.Text1(25).Text
                  sPA(3) = Left(Me.Text1(26).Text & "0", 1)
                  sPA(4) = Left(Me.Text1(27).Text & "00", 2)
                  If PUB_AddNewPatent(sPA()) Then
                    Else
                        GoTo CheckingErr
                  End If
               End If
            Case Else:
               ReDim sSP(1 To tf_SP) As String
               
               If PUB_ReadServicePracticeData(sSP(), pa(1), pa(2), pa(3), pa(4)) Then
                  sSP(1) = Me.Text1(2).Text
                  sSP(2) = Me.Text1(25).Text
                  sSP(3) = Left(Me.Text1(26).Text & "0", 1)
                  sSP(4) = Left(Me.Text1(27).Text & "00", 2)
                  If PUB_AddNewServicePractice(sSP()) Then
                    Else
                        GoTo CheckingErr
                  End If
               End If
         End Select

      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      
      'Modify by Morgan 2004/6/8 新案旗標要清除 CP31=NULL
      '2005/7/14 MODIFY BY SONIA 若該案號無基本檔則 CP31='Y' 否則 NULL
      'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP01='" & Me.text1(2).Text & "',CP02='" & Left(Me.text1(25).Text & "000000", 6) & "' ,CP03='" & Left(Me.text1(26).Text & "0", 1) & "' ,CP04='" & Left(Me.text1(27).Text & "00", 2) & "' ,CP43='', CP31=NULL WHERE CP09='" & Me.Label3(8).Caption & "'"
      'Modified by Morgan 2016/3/21 基本檔一定會先存在,改判斷進度檔
       'strTxt(intStep) = "UPDATE CASEPROGRESS  SET CP01='" & Me.text1(2).Text & "',CP02='" & Left(Me.text1(25).Text & "000000", 6) & "' ,CP03='" & Left(Me.text1(26).Text & "0", 1) & "' ,CP04='" & Left(Me.text1(27).Text & "00", 2) & "' ,CP43=''" & _
         ", CP31=(select DECODE(MAX(1),NULL,'Y',NULL) from PATENT where PA01='" & Me.text1(2).Text & "' AND PA02='" & Left(Me.text1(25).Text & "000000", 6) & "' AND PA03='" & Left(Me.text1(26).Text & "0", 1) & "' AND PA04='" & Left(Me.text1(27).Text & "00", 2) & "')" & _
         " WHERE CP09 = '" & Me.Label3(8).Caption & "'"
      strTxt(intStep) = "UPDATE CASEPROGRESS a SET CP01='" & Me.Text1(2).Text & "',CP02='" & Left(Me.Text1(25).Text & "000000", 6) & "' ,CP03='" & Left(Me.Text1(26).Text & "0", 1) & "' ,CP04='" & Left(Me.Text1(27).Text & "00", 2) & "' ,CP43=''" & _
         ", CP31=(select DECODE(count(*),0,'Y','') from caseprogress b where b.cp01='" & Me.Text1(2).Text & "' AND b.cp02='" & Left(Me.Text1(25).Text & "000000", 6) & "' AND b.cp03='" & Left(Me.Text1(26).Text & "0", 1) & "' AND b.cp04='" & Left(Me.Text1(27).Text & "00", 2) & "' and b.cp09<>a.cp09)" & _
         " WHERE CP09 = '" & Me.Label3(8).Caption & "'"
      'end 2016/3/21
      '2005/7/14 END
      cnnConnection.Execute strTxt(intStep), intI
      intStep = intStep + 1
      
      'Added by Lydia 2020/08/18 更新CaseRelation1和DivisionCase
      If m_CP31 = "Y" Then
          Call PUB_UpdateCaseRelation1(pa(1), pa(2), pa(3), pa(4), Me.Text1(2).Text, Left(Me.Text1(25).Text & "000000", 6), Left(Me.Text1(26).Text & "0", 1), Left(Me.Text1(27).Text & "00", 2))
      End If
      'end 2020/08/18
      
      'Add by Morgan 2010/8/12
      '更正財務相關資料
      PUB_UpdateAccData strReceiveNo, pa(1) & pa(2) & pa(3) & pa(4)
      
   '若未輸入轉本所案號
   Else
      'Add by Morgan 2007/6/11
      If (Text1(1) = "201" Or Text1(1) = "927") Then
         If txtTF05.Enabled = True Then
            If RTrim(txtTF05) = "" Or Val(txtTF05) = 100 Then
               strExc(10) = "Null"
            Else
               strExc(10) = Val(txtTF05)
            End If
            'Add by Morgan 2007/8/9 加成比率
            If RTrim(txtTF18) = "" Or Val(txtTF18) = 100 Then
               strExc(9) = "Null"
            Else
               strExc(9) = Val(txtTF18)
            End If
            'Added by Lydia 2017/05/17 原文字數、相似度
            If RTrim(txtTF23) = "" Then
               strExc(7) = "Null"
            Else
               strExc(7) = Val(txtTF23)
            End If
            If RTrim(txtTF19) = "" Or Val(txtTF19) = 100 Then
               strExc(8) = "Null"
            Else
               strExc(8) = Val(txtTF19)
            End If
            'end 2017/05/17
            
            'Modified by Lydia 2017/05/17
            'strSql = "update transfee set TF05=" & strExc(10) & ",TF18=" & strExc(9) & " where TF01='" & strReceiveNo & "' and tf07 is null"
            strSql = "update transfee set TF05=" & strExc(10) & ",TF18=" & strExc(9) & ",TF23=" & strExc(7) & ",TF19=" & strExc(8) & " where TF01='" & strReceiveNo & "' and tf07 is null"
            'end 2017/05/17
            cnnConnection.Execute strSql, intI
            If intI = 0 Then
               'Modified by Lydia 2017/05/17
               'strSql = "insert into transfee(TF01,TF05,TF18) values('" & strReceiveNo & "'," & strExc(10) & "," & strExc(9) & ")"
               strSql = "insert into transfee(TF01,TF05,TF18,TF23,TF19) values('" & strReceiveNo & "'," & strExc(10) & "," & strExc(9) & "," & strExc(7) & "," & strExc(8) & ")"
               'end 2017/05/17
               cnnConnection.Execute strSql, intI
            End If
         End If
      Else
         'Memo by Lydia 2018/01/09 改案件性質，刪除翻譯費用檔
         'Remove by Lydia 2018/09/26 取消刪除
         'strSql = "delete transfee where tf01='" & strReceiveNo & "' and tf07 is null"
         'cnnConnection.Execute strSql, intI
         'end 2018/09/26
      End If
      'end 2007/6/11
      
      Select Case pa(1)
         Case "FCP"
            'Modify by Morgan 2008/3/18 +PA143
            strTxt(intStep) = "UPDATE PATENT SET PA08=" & CNULL(Text1(21).Text) & _
               ",PA75=" & CNULL(ChangeCustomerL(Text1(18))) & ",PA91=" & CNULL(ChgSQL(textPA91)) & _
               ",PA23=" & CNULL(Text1(3)) & ",PA48=" & CNULL(ChgSQL(Text1(8))) & ",PA143='" & m_PA143 & "'"
            
            'Added by Lydia 2020/01/20 更換代理人，發Email通知程序管制人和承辦管制人。
            'Email內文：本案代理人由原Y22457000(美國THE DOW CHEMICAL COMPANY)更改為Y34232000(日本YASUTOMI & ASSOCIATES)，請以最新資料之管制人為主，謝謝。
            'Email收件者: 更新後該區的程序及承辦
            '副本: 更新後該區的程序及承辦之主管 , 更新前該區的程序及承辦
            If Text1(18).Text <> Text1(18).Tag Then
                StrSQLa = ""
                If Text1(18).Text <> "" Then  '更新後的程序/承辦管制人和主管
                    StrSQLa = "select '1' as ord1, fa01||fa02 as fno,na03,nvl(fa05,nvl(fa04,fa06)) as fname, na16, nvl(s1.st52,nvl(s1.st53,s1.st54)) na16m, na51, nvl(s2.st52,nvl(s2.st53,s2.st54)) na51m " & _
                                     " from fagent ,nation, staff s1,staff s2 where fa01||fa02='" & ChangeCustomerL(Text1(18).Text) & "' and fa10=na01(+) and na16=s1.st01(+) and na51=s2.st01(+) "
                End If
                If Text1(18).Tag <> "" Then  '更新前的程序/承辦管制人和主管
                    If StrSQLa <> "" Then StrSQLa = StrSQLa & " Union all "
                    StrSQLa = StrSQLa & "select '2' as ord1, fa01||fa02 as fno,na03,nvl(fa05,nvl(fa04,fa06)) as fname, na16, nvl(s1.st52,nvl(s1.st53,s1.st54)) na16m, na51, nvl(s2.st52,nvl(s2.st53,s2.st54)) na51m " & _
                                     " from fagent ,nation, staff s1,staff s2 where fa01||fa02='" & ChangeCustomerL(Text1(18).Tag) & "' and fa10=na01(+) and na16=s1.st01(+) and na51=s2.st01(+) "
                End If
                StrSQLa = StrSQLa & " order by ord1 "
                intI = 1
                Set rsA = ClsLawReadRstMsg(intI, StrSQLa)
                If intI = 1 Then
                    strExc(0) = ""
                    strExc(1) = "": strExc(2) = ""
                    strExc(3) = "": strExc(4) = ""
                    rsA.MoveFirst
                    Do While Not rsA.EOF
                        If "" & rsA.Fields("ord1") = "1" Then '更新後的程序/承辦管制人, CC: 主管
                            strExc(1) = strExc(1) & rsA.Fields("na16") & ";" & rsA.Fields("na51") & ";"
                            strExc(2) = strExc(2) & rsA.Fields("na16m") & ";" & rsA.Fields("na51m") & ";"
                            If "" & rsA.Fields("fno") <> "" Then strExc(3) = "" & rsA.Fields("fno") & "(" & Trim(rsA.Fields("na03")) & rsA.Fields("fname") & ")"
                        Else    'CC:更新前的程序/承辦管制人和主管
                            strExc(2) = strExc(2) & rsA.Fields("na16") & ";" & rsA.Fields("na51") & ";"
                            If "" & rsA.Fields("fno") <> "" Then strExc(4) = "" & rsA.Fields("fno") & "(" & Trim(rsA.Fields("na03")) & rsA.Fields("fname") & ")"
                        End If
                        rsA.MoveNext
                    Loop
                    strExc(1) = Replace(strExc(1), ";;", ";")
                    strExc(2) = Replace(strExc(2), ";;", ";")
                    If Len(strExc(1)) > 5 Or Len(strExc(2)) > 5 Then
                        If Len(strExc(1)) < 5 Then
                            strExc(1) = strExc(2): strExc(2) = ""
                        End If
                        strExc(0) = "本案代理人" & IIf(strExc(4) <> "", "由原" & strExc(4), "由空白") & "更改為" & IIf(strExc(3) <> "", strExc(3), "空白") & "，請以最新資料之管制人為主，謝謝。"
                        strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                   " VALUES ( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd')" & _
                                   ",to_char(sysdate,'hh24miss'),'" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "新案立卷" & "','" & strExc(0) & "'," & CNULL(strExc(2)) & ")"
                        cnnConnection.Execute strSql
                    End If
                End If
                Set rsA = Nothing
            End If
            'end 2020/01/20
            
            If Text1(11) = "Y" Then
               strTxt(intStep) = strTxt(intStep) & ",PA57=NULL,PA58=NULL,PA59=NULL"
            End If
            
            If Text1(1) = 異議_專 Then
               strTxt(intStep) = strTxt(intStep) & ",PA14=" & CNULL(TransDate(Text1(7), 2))
            End If
            
            If Left(Text1(1), 1) = 5 Then
               strTxt(intStep) = strTxt(intStep) & ",PA18='Y'"
            End If
            If Left(Text1(1), 1) = 8 Then
               strTxt(intStep) = strTxt(intStep) & ",PA19='Y'"
            End If
            
            strTxt(intStep) = strTxt(intStep) & ",pa163='" & m_PA163 & "'" 'Added by Morgan 2015/9/30
            
            If Text1(13) <> "" Then
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetCustomerNameAndAddress(Text1(13).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
               If ClsPDGetCustomerNameAndAddress(Text1(13).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
                  '修改申請人時
                  If InStr(ChangeCustomerL(pa(26)), ChangeCustomerL(Text1(13))) = 0 Then
                     If m_CP60 <> "" Then
                        strExc(1) = pa(1)
                        strExc(2) = pa(2)
                        strExc(3) = pa(3)
                        strExc(4) = pa(4)
                        strExc(5) = m_CP60
                        strExc(6) = Text1(13)
                        strExc(7) = strExc(0)
                        strExc(8) = pa(26)
                        If Not ClsLawUpdAcc0k0(strExc(), True) Then
                           Text1(13).SetFocus
                           GoTo CheckingErr
                        End If
                     End If
                     strTxt(intStep) = strTxt(intStep) & _
                        ",PA26=" & CNULL(ChangeCustomerL(Text1(13))) & _
                        ",PA31=" & CNULL(ChgSQL(strTmp(1))) & ",PA36=" & CNULL(ChgSQL(strTmp(2))) & _
                        ",PA41=" & CNULL(ChgSQL(strTmp(3))) & _
                        ",PA79=NULL,PA80=NULL,PA81=NULL,PA82=NULL,PA83=NULL,PA84=NULL"
                  End If
               End If
            Else
               strTxt(intStep) = strTxt(intStep) & ",PA26=NULL,PA31=NULL,PA36=NULL," & _
                  "PA41=NULL,PA79=NULL,PA80=NULL,PA81=NULL,PA82=NULL,PA83=NULL,PA84=NULL"
            End If
            
            If Text1(14) <> "" Then
               If ClsPDGetCustomerNameAndAddress(Text1(14).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
                  '修改申請人時
                  If InStr(pa(27), Text1(14)) = 0 Then
                     strTxt(intStep) = strTxt(intStep) & _
                        ",PA27=" & CNULL(ChangeCustomerL(Text1(14))) & _
                        ",PA32=" & CNULL(ChgSQL(strTmp(1))) & ",PA37=" & CNULL(ChgSQL(strTmp(2))) & _
                        ",PA42=" & CNULL(ChgSQL(strTmp(3))) & _
                        ",PA109=NULL,PA110=NULL,PA111=NULL,PA112=NULL,PA113=NULL,PA114=NULL"
                  End If
               End If
            Else
               strTxt(intStep) = strTxt(intStep) & ",PA27=NULL,PA32=NULL,PA37=NULL," & _
                  "PA42=NULL,PA109=NULL,PA110=NULL,PA111=NULL,PA112=NULL,PA113=NULL,PA114=NULL"
            End If
            
            If Text1(15) <> "" Then
               If ClsPDGetCustomerNameAndAddress(Text1(15).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
                  '修改申請人時
                  If InStr(pa(28), Text1(15)) = 0 Then
                     strTxt(intStep) = strTxt(intStep) & _
                        ",PA28=" & CNULL(ChangeCustomerL(Text1(15))) & _
                        ",PA33=" & CNULL(ChgSQL(strTmp(1))) & ",PA38=" & CNULL(ChgSQL(strTmp(2))) & _
                        ",PA43=" & CNULL(ChgSQL(strTmp(3))) & _
                        ",PA115=NULL,PA116=NULL,PA117=NULL,PA118=NULL,PA119=NULL,PA120=NULL"
                  End If
               End If
            Else
               strTxt(intStep) = strTxt(intStep) & ",PA28=NULL,PA33=NULL,PA38=NULL," & _
                  "PA43=NULL,PA115=NULL,PA116=NULL,PA117=NULL,PA118=NULL,PA119=NULL,PA120=NULL"
            End If
            
            If Text1(16) <> "" Then
               If ClsPDGetCustomerNameAndAddress(Text1(16).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
                  '修改申請人時
                  If InStr(pa(29), Text1(16)) = 0 Then
                     strTxt(intStep) = strTxt(intStep) & _
                        ",PA29=" & CNULL(ChangeCustomerL(Text1(16))) & _
                        ",PA34=" & CNULL(ChgSQL(strTmp(1))) & ",PA39=" & CNULL(ChgSQL(strTmp(2))) & _
                        ",PA44=" & CNULL(ChgSQL(strTmp(3))) & _
                        ",PA121=NULL,PA122=NULL,PA123=NULL,PA124=NULL,PA125=NULL,PA126=NULL"
                  End If
               End If
            Else
               strTxt(intStep) = strTxt(intStep) & ",PA29=NULL,PA34=NULL,PA39=NULL," & _
                  "PA44=NULL,PA121=NULL,PA122=NULL,PA123=NULL,PA124=NULL,PA125=NULL,PA126=NULL"
            End If
            
            If Text1(17) <> "" Then
               If ClsPDGetCustomerNameAndAddress(Text1(17).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
                  '修改申請人時
                  If InStr(pa(30), Text1(17)) = 0 Then
                     strTxt(intStep) = strTxt(intStep) & _
                        ",PA30=" & CNULL(ChangeCustomerL(Text1(17))) & _
                        ",PA35=" & CNULL(ChgSQL(strTmp(1))) & ",PA40=" & CNULL(ChgSQL(strTmp(2))) & _
                        ",PA45=" & CNULL(ChgSQL(strTmp(3))) & _
                        ",PA127=NULL,PA128=NULL,PA129=NULL,PA130=NULL,PA131=NULL,PA132=NULL"
                  End If
               End If
            Else
               strTxt(intStep) = strTxt(intStep) & ",PA30=NULL,PA35=NULL,PA40=NULL," & _
                  "PA45=NULL,PA127=NULL,PA128=NULL,PA129=NULL,PA130=NULL,PA131=NULL,PA132=NULL"
            End If
            
            'Modify By Sindy 2014/11/12
            If strSrvDate(1) >= 專利發明人檔啟用日 Then
'               '申請人有變更且未重新點選發明人資料時清除原發明人資料
'               '申請案才要
'               If InStr(NewCasePtyList, text1(1)) > 0 Then
'                  '串列申請人資料
'                  strExc(1) = text1(13)
'                  For intI = 1 To 4
'                     If text1(intI + 13) <> "" Then
'                        strExc(1) = strExc(1) & "," & text1(3)
'                     End If
'                  Next
'                  strExc(2) = "" 'Added by Morgan 2014/12/19
'                  '串列發明人資料
'                  strSql = "SELECT pi06 from PatentInventor where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
'                           " order by pi05 asc"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 1 Then
'                     RsTemp.MoveFirst
'                     strExc(2) = RsTemp.Fields(0)
'                     RsTemp.MoveNext
'                     Do While Not RsTemp.EOF
'                        strExc(2) = strExc(2) & "," & RsTemp.Fields(0)
'                        RsTemp.MoveNext
'                     Loop
'                  End If
'                  '檢查是否要更新
'                  If PUB_ChkInventor(strExc(2), strExc(1), True) = False Then
                  If BolIsInventorUpd = True Then 'Add By Sindy 2014/12/23 要更新發明人資料
                     '全部刪除
                     strSql = "delete from patentInventor where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4))
                     Pub_SeekTbLog strSql 'Add By Sindy 2017/8/23
                     cnnConnection.Execute strSql
                     arrInv = Split(strExc(2), ",")
                     For intI = 0 To UBound(arrInv)
                        '重新新增
                        strSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
                                 CNULL(pa(1)) & "," & CNULL(pa(2)) & "," & CNULL(pa(3)) & "," & CNULL(pa(4)) & "," & intI + 1 & ",'" & arrInv(intI) & "')"
                        Pub_SeekTbLog strSql 'Add By Sindy 2017/8/23
                        cnnConnection.Execute strSql
                     Next intI
                  End If
'               End If
            Else
            '2014/11/12 END
               For intI = 0 To 9
                  strTxt(intStep) = strTxt(intStep) & ",pa" & (60 + intI) & "='" & pa(60 + intI) & "'"
               Next
            End If
            
            strTxt(intStep) = strTxt(intStep) & " WHERE " & ChgPatent(Label3(9))
            Pub_SeekTbLog strTxt(intStep) 'Added by Lydia 2023/07/05
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
                  
         Case "FG"
            
            strTxt(intStep) = "UPDATE SERVICEPRACTICE SET SP29=" & CNULL(Text1(8)) & _
               ",SP26=" & CNULL(ChangeCustomerL(Text1(18))) & ",SP18=" & CNULL(textPA91)
            
            If Text1(11) = "Y" Then
               strTxt(intStep) = strTxt(intStep) & ",SP15=NULL,SP16=NULL,SP17=NULL"
            End If
            
            If Text1(13) <> "" Then
               If ClsPDGetCustomerNameAndAddress(Text1(13).Text, strExc(0), strTmp(1), strTmp(2), strTmp(3)) Then
                  If m_CP60 <> "" And InStr(ChangeCustomerL(pa(26)), ChangeCustomerL(Text1(13))) = 0 Then
                     strExc(1) = pa(1)
                     strExc(2) = pa(2)
                     strExc(3) = pa(3)
                     strExc(4) = pa(4)
                     strExc(5) = m_CP60
                     strExc(6) = Text1(13)
                     strExc(7) = strExc(0)
                     strExc(8) = pa(26)
                     If Not ClsLawUpdAcc0k0(strExc(), True) Then
                        Text1(13).SetFocus
                        GoTo CheckingErr
                     End If
                  End If
               End If
            End If
            
            strTxt(intStep) = strTxt(intStep) & ",SP08=" & CNULL(ChangeCustomerL(Text1(13))) & _
               ",SP58=" & CNULL(ChangeCustomerL(Text1(14))) & ",SP59=" & CNULL(ChangeCustomerL(Text1(15)))
   
            strTxt(intStep) = strTxt(intStep) & " WHERE " & ChgService(Label3(9))
            Pub_SeekTbLog strTxt(intStep) 'Added by Lydia 2023/07/05
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
      End Select
      
      'Added by Lydia 2025/06/30 取消閉卷時，若下一程序有未過期且已上N之年費605，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。
      If Text1(11) = "Y" Then
         strMsgCloseCancel = PUB_GetCaseCloseCancel(pa(1), pa(2), pa(3), pa(4), pa(9))
      End If
      
      'Add by Morgan 2008/8/28
      '承辦期限不可大於所限
      If Text1(23) <> "" And Text1(4) <> "" And Val(Text1(23)) > Val(Text1(4)) Then
         Text1(23) = Text1(4)
      End If
      
      'Modify by Morgan 2008/8/19 +CP48
      'Modified  by Lydia 2017/12/14 + CP118
      'Memo
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP13=" & CNULL(Text1(24)) & _
         ",CP14=" & CNULL(Left(Trim(cboCP14.Text), 5)) & ",CP10=" & CNULL(Text1(1)) & ",CP06=" & CNULL(TransDate(Text1(4), 2)) & _
         ",CP07=" & CNULL(TransDate(Text1(5), 2)) & ",CP43=" & CNULL(Text1(6)) & ",CP57=" & CNULL(TransDate(Text1(9), 2)) & _
         ",CP26=" & CNULL(Text1(10)) & ",CP05=" & CNULL(TransDate(Text1(12), 2)) & _
         ",CP48=" & CNULL(TransDate(Text1(23), 2), True) & ",CP118=" & CNULL(n_CP118)
         
      If strKind <> Text1(1) Then
         strExc(0) = "SELECT CF13,CF14 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & Text1(1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            i = 0: j = 0
            If Not IsNull(RsTemp.Fields(0)) Then i = RsTemp.Fields(0)
            If Not IsNull(RsTemp.Fields(1)) Then j = RsTemp.Fields(1)
            strTxt(intStep) = strTxt(intStep) & ",CP33=" & i & ",CP34=" & j
         End If
      End If
      
      'Add By Sindy 2021/4/29 若有修改本所期限，自動加備註
      'Modify By Sindy 2021/9/30 一開始就是無期限,不用記錄 => m_strCP06 <> ""
      If m_strCP06 <> "" And Val(m_strCP06) <> Val(DBDATE(Text1(4))) And Val(m_CP27) = 0 And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
         If InStr(textCP64, "原本所期限為" & ChangeWStringToTDateString(DBDATE(m_strCP06)) & "已修改;") = 0 Then
            textCP64 = "原本所期限為" & ChangeWStringToTDateString(DBDATE(m_strCP06)) & "已修改;" & textCP64
         End If
      End If
      '2021/4/29 END
      'Add By Sindy 2015/12/16 存指定日期
      If m_CP27 = "" And m_CP57 = "" Then
         If Text1(30).Text <> "" Then
            'If Left(Label3(8), 1) <> "C" Then 'Added by Lydia 2021/11/05 排除C類來函：客戶指定送件日 'Mark by Lydia 2023/05/25 當時排除的原因已無法確定(Sharon: 不應該排除)
                strTxt(intStep) = strTxt(intStep) & ",CP141='3'" '3.指定日期送件 'Memo by Lydia 2021/11/03 對智慧局
            'End If   'Added by Lydia 2021/11/05
            strTxt(intStep) = strTxt(intStep) & ",CP142=" & DBDATE(Text1(30))
            'Modify By Sindy 2021/4/20
            If Option1(0).Visible = True Then
               'Modify By Sindy 2021/10/20 + , IIf(Option1(1).Value = True, "2", "3")
               strTxt(intStep) = strTxt(intStep) & ",CP164='" & IIf(Option1(0).Value = True, "1", IIf(Option1(1).Value = True, "2", "3")) & "'"
            End If
         Else
            '取消指定日
            'Modify By Sindy 2021/4/20 + ,CP164=null
            strTxt(intStep) = strTxt(intStep) & ",CP141=null,CP142=null,CP164=null"
         End If
         'Add By Sindy 2024/4/25
         If chkCP176.Value = 1 Then
            strTxt(intStep) = strTxt(intStep) & ",CP176='Y'"
         Else
            strTxt(intStep) = strTxt(intStep) & ",CP176=null"
         End If
         If chkCP176.Value = 1 And chkCP176.Tag <> "Y" Then
            textCP64 = ChangeWStringToTDateString(strSrvDate(1)) & "需待客戶最終指示" & IIf(textCP64 <> "", ";", "") & textCP64
         ElseIf chkCP176.Value = 0 And chkCP176.Tag = "Y" Then
            textCP64 = "於" & ChangeWStringToTDateString(strSrvDate(1)) & "取消待客戶最終指示" & IIf(textCP64 <> "", ";", "") & textCP64
         End If
         '2024/4/25 END
      End If
      '2015/12/16 END
      strTxt(intStep) = strTxt(intStep) & ",CP64=" & CNULL(ChgSQL(textCP64)) '進度備註
      
      'Modify by Morgan 2010/4/29 改抓畫面欄位值
      'Add by Morgan 2009/10/6
      '退審查費要請款
      'If text1(1).Text = "908" And Combo3.Enabled = True And m_CP20 = "N" And m_CP27 = "" Then
      '   strTxt(intStep) = strTxt(intStep) & ",cp20=null"
      'End If
         strTxt(intStep) = strTxt(intStep) & ",cp20='" & Text1(29) & "'"
      'end 2010/4/29
   
      'Add by Morgan 2011/4/22 延期要紀錄NP22
      If Text1(1) = "404" Then
         strTxt(intStep) = strTxt(intStep) & ",CP30='" & m_CP30 & "'"
      
      'Added by Morgan 2012/12/21
      ElseIf pa(1) = "FCP" And (Text1(1) = "125" Or Text1(1) = "308") Then
         If txtDesignCaseNo(2) <> "" Then
            '紀錄衍生設計母案本所案號
            strTxt(intStep) = strTxt(intStep) & ",CP30='" & txtDesignCaseNo(1) & txtDesignCaseNo(2) & txtDesignCaseNo(3) & txtDesignCaseNo(4) & "'"
            '新增相關案號
            strExc(0) = "select * from caserelation1 where cr01='" & pa(1) & "' and cr02='" & pa(2) & "' and cr03='" & pa(3) & "' and cr04='" & pa(4) & "'" & _
               " and cr05='" & txtDesignCaseNo(1) & "' and cr06='" & txtDesignCaseNo(2) & "' and cr07='" & txtDesignCaseNo(3) & "' and cr08='" & txtDesignCaseNo(4) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               strSql = "insert into caserelation1(CR01,CR02,CR03,CR04,CR05,CR06,CR07,CR08) " & _
                  " select '" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','" & txtDesignCaseNo(1) & "','" & txtDesignCaseNo(2) & "','" & txtDesignCaseNo(3) & "','" & txtDesignCaseNo(4) & "'" & _
                  " From DUAL" & _
                  " UNION select '" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "',CR05,CR06,CR07,CR08" & _
                  " from caserelation1 where cr01='" & txtDesignCaseNo(1) & "' and cr02='" & txtDesignCaseNo(2) & "' AND cr03='" & txtDesignCaseNo(3) & "' and cr04='" & txtDesignCaseNo(4) & "'" & _
                  " UNION select '" & txtDesignCaseNo(1) & "','" & txtDesignCaseNo(2) & "','" & txtDesignCaseNo(3) & "','" & txtDesignCaseNo(4) & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'" & _
                  " From DUAL" & _
                  " UNION select CR05,CR06,CR07,CR08,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'" & _
                  " from caserelation1 where cr01='" & txtDesignCaseNo(1) & "' and cr02='" & txtDesignCaseNo(2) & "' AND cr03='" & txtDesignCaseNo(3) & "' and cr04='" & txtDesignCaseNo(4) & "'"
               cnnConnection.Execute strSql, intI
            End If
            
            'Add By Sindy 2018/5/9 FCP案母案資料複製到子案
            If Text1(1) <> "308" Then 'Added by Morgan 2020/2/26 改請衍生設計不用(要維持原申請案資料)--淑華
               strDivCaseNo(1) = txtDesignCaseNo(1)
               strDivCaseNo(2) = txtDesignCaseNo(2)
               strDivCaseNo(3) = txtDesignCaseNo(3)
               strDivCaseNo(4) = txtDesignCaseNo(4)
               Call PUB_FCPCopyDataToCase(pa(), strDivCaseNo())
            End If
            '2018/5/9 END
            
            'Added by Lydia 2019/07/04 衍生設計新案(母案已提申者)不須重跑新案命名流程，直接沿用母案名稱
            'Remove by Lydia 2019/07/30 改成衍生設計新案發文時檢查命名記錄尚未分組,才刪除
            'If Text1(1) = "125" Then
            '    strSql = "delete from transcasetitle where tct01='" & Label3(8) & "' and tct05||tct08||tct11 is null "
            '    cnnConnection.Execute strSql, intI
            ''End If
         End If
      End If
      
      strTxt(intStep) = strTxt(intStep) & " WHERE CP09='" & Label3(8) & "'"

      'Added by Lydia 2016/07/07 若承辦人變更時紀錄異動人員日期時間
      If Trim(cboCP14.Tag) <> "" And Trim(cboCP14.Text) <> Trim(cboCP14.Tag) Then
         Call PUB_ChgEmpUpdEEP05(Label3(8), Trim(cboCP14.Tag), Left(Trim(cboCP14.Text), 5), "1")
         
         '要寫 Log 並改觸發 Trigger 更新修改人員日期時間
         Pub_SeekTbLog strTxt(intStep)
         strTxt(intStep) = "begin user_data.user_enabled:=1; " & strTxt(intStep) & "; end;"
      'Added by Lydia 2018/11/22 若變更案件性質,寫log方便追蹤
      ElseIf strKind <> Text1(1).Text Then
          Pub_SeekTbLog strTxt(intStep)
          strTxt(intStep) = "begin user_data.user_enabled:=1; " & strTxt(intStep) & "; end;"
      'end 2018/11/22
      End If
      'end 2016/07/07
      
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      'Added by Lydia 2024/02/27 外專機械設計組人員異動調整程式：承辦人為虛帳號人員時加發EMAIL
      'Modified by Lydia 2024/03/08 排除外翻編號F +Left(Trim(cboCP14.Text), 1) <> "F"
      If Left(Trim(cboCP14.Text), 1) <> "F" And Trim(cboCP14.Text) <> Trim(cboCP14.Tag) And (Mid(Trim(cboCP14.Tag) & ",", 4, 1) = "9" Or Mid(Trim(cboCP14.Text) & ",", 4, 1) = "9") Then
         'Added by Lydia 2024/06/06 CC改用變數
         If Mid(Trim(cboCP14.Text) & ",", 4, 1) = "9" And Trim(Left(cboCP14.Text, 6)) <> "99097" Then 'Added by Lydia 2024/05/23 分案給內專工程師時，請新增副本人員: 國外部對接主管---Winfrey
            strExc(1) = PUB_GetFCPEngSup(Trim(Left(cboCP14.Text, 6)), , , True)
         Else
            strExc(1) = ""
         End If
         'Modified by Lydia 2024/06/06 CC改用變數strExc(1)
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
                  " VALUES ( '" & strUserNum & "','" & Trim(Left(cboCP14.Text, 6)) & IIf(Trim(cboCP14.Tag) <> "", ";" & Trim(Left(cboCP14.Tag, 6)), "") & "',to_char(sysdate,'yyyymmdd')" & _
                  ",to_char(sysdate,'hh24miss'),'" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "變更承辦人'," & _
                  CNULL(IIf(Trim(cboCP14.Tag) <> "", "原承辦人：" & GetStaffName(Trim(Left(cboCP14.Tag, 6)), True) & vbCrLf & "新承辦人：" & GetStaffName(Trim(Left(cboCP14.Text, 6)), True) & vbCrLf, "")) & ",'" & strExc(1) & "' ,'" & Label3(8) & "')"
         cnnConnection.Execute strSql
      End If
      'end 2024/02/27
         
      'Added by Lydia 2022/09/29 新增工程師分組控管
      '當承辦人為工程師時，所輸入的工程師組別與原來的工程師組別不一致時，談視窗詢問：是否變更工程師組別，是: 工程師組別改為此次輸入的工程師組別
      If bUpdPA150 = True Then
           If pa(1) = "FCP" Then
              strSql = "Update Patent Set PA150=" & CNULL(PUB_GetStaffST16(Trim(Left(cboCP14.Text, 6)))) & " Where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "' "
           Else 'FG
              strSql = "Update ServicePractice Set SP79=" & CNULL(PUB_GetStaffST16(Trim(Left(cboCP14.Text, 6)))) & " Where sp01='" & pa(1) & "' and sp02='" & pa(2) & "' and sp03='" & pa(3) & "' and sp04='" & pa(4) & "' "
           End If
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
      End If
      'end 2022/09/29
      
      'Added by Morgan 2015/10/7
      '審查意見或核駁修改承辦人時一併修改相關收文號之告代承辦人 Ex.FCP-45516
      'Remove by Lydia 2016/06/21 改成模組PUB_SaveFCPcp14
      'If text1(0) <> "" And (text1(1) = "1202" Or text1(1) = "1002" Or text1(1) = "1227") Then
      '   strSql = "update caseprogress set  cp14='" & text1(0) & "' where cp43='" & Label3(8) & "' and cp10='901' and cp27 is null"
      '   cnnConnection.Execute strSql, intI
      'End If
      ''end 2015/10/7
      
      '若有輸入分割母案本所案號則更新 DIVISIONCASE
      If Text1(1) = "307" Then
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
            Pub_SeekTbLog strTxt(intStep), , True   'Added by Lydia 2021/04/08 留下修改記錄, 但是查詢dml_log時無法用本所案號
            intStep = intStep + 1
            'Add By Sindy 2018/5/9 FCP案母案資料複製到子案
            strDivCaseNo(1) = txtDivCaseNo(1)
            strDivCaseNo(2) = txtDivCaseNo(2)
            strDivCaseNo(3) = txtDivCaseNo(3)
            strDivCaseNo(4) = txtDivCaseNo(4)
            Call PUB_FCPCopyDataToCase(pa(), strDivCaseNo())
            '2018/5/9 END
         End If
         
         'Added by Morgan 2012/11/12
         If m_CP27 = "" Then
            st307Msg = PUB_Update307RefTw(strReceiveNo, str307CP06, str307CP07)
            
            'Added by Morgan 2021/8/20 若分割案未發文則實審期限為分割案相同--敏莉
            'Modified by Morgan 2022/5/4 +續行母案再審(考慮後收文情形,435的分案也會更新)
            'If str307CP06 <> "" And m_stVar(0) <> "" Then
            If str307CP06 <> "" And (m_stVar(0) <> "" Or m_bol435) Then
            'end 2022/5/4
               m_stVar(0) = str307CP06
               m_stVar(3) = str307CP07
               'Add By Sindy 2021/9/13 少了約定期限
               'm_stVar(0) = PUB_GetFCPOurDeadline(m_stVar(3), 4, , m_pAgreeOnDate)
               Call PUB_GetFCPOurDeadline(m_stVar(3), 4, , m_pAgreeOnDate)
               '2021/9/13 END
            End If
            'end 2021/8/20
         End If
         'end 2012/11/12
         
         If m_stVar(0) <> "" Then 'Added by Morgan 2013/9/17
            'FCP及P案件，若專利種類為'發明'且案件性質為'分割'時，實審期限=母案之申請日＋3年<=分割案之收文日＋1月(110/8/20 若母案實審已過且分割案未發文則實審期限設定為相同--敏莉)
            '若有收文未取消收文之'實體審查'，則更新該筆'實體審查'之期限，若無則新增下一程序'實體審查'期限，並顯示'此分割案尚未收文實體審查，期限為XXXXXX，請提醒智權人員 !!'
            'Modified by Morgan 2015/9/10 若有收文"續行母案再審"435或母案已進入再審階段則改更新/新增"續行母案再審"435期限
            If pa(1) = "FCP" And Text1(21) = "1" Then
               '有收'實體審查'
               If m_st416CP09 <> "" Then
                  strTxt(intStep) = " UPDATE CASEPROGRESS SET CP06=" & m_stVar(0) & ",CP07=" & m_stVar(3) & " WHERE CP09='" & m_st416CP09 & "'"
                  cnnConnection.Execute strTxt(intStep)
                  intStep = intStep + 1
               '沒有收'實體審查'
'Removed by Morgan 2021/9/23 取消，改分割發文時再管制就好，以免催期限時不小心催到客戶--葉敏莉
'               Else
'                  If m_bol435 Then
'                     strExc(1) = "435"
'                  Else
'                     strExc(1) = "416"
'                  End If
'
'                  'Modify By Sindy 2021/4/27 + ,NP23=" & CNULL(DBDATE(m_pAgreeOnDate)):約定期限
'                  strTxt(intStep) = _
'                     " DECLARE" & _
'                        " V_NP22 NUMERIC(10,0);" & _
'                        " V_NP02 VARCHAR2(9);" & _
'                     " BEGIN" & _
'                        " SELECT MAX(NP02) INTO V_NP02 FROM NEXTPROGRESS WHERE NP01='" & Label3(8) & "' AND NP07='" & strExc(1) & "';" & _
'                        " IF V_NP02 IS NULL THEN" & _
'                           " SELECT NVL(MAX(NP22),0)+1 INTO V_NP22 FROM NEXTPROGRESS;" & _
'                           " INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22,NP23)" & _
'                           " VALUES ('" & Label3(8) & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'" & _
'                           " ,'" & strExc(1) & "'," & m_stVar(0) & "," & m_stVar(3) & ",'" & text1(24) & "',V_NP22," & CNULL(DBDATE(m_pAgreeOnDate)) & ");" & _
'                        " ELSE" & _
'                           " UPDATE NEXTPROGRESS SET NP08=" & m_stVar(0) & ",NP09=" & m_stVar(3) & ",NP23=" & CNULL(DBDATE(m_pAgreeOnDate)) & " WHERE NP01='" & Label3(8) & "' AND NP07='" & strExc(1) & "';" & _
'                        " END IF;" & _
'                     " END;"
'                  cnnConnection.Execute strTxt(intStep)
'                  intStep = intStep + 1
'end 2021/9/23
               End If
               
               'Add by Morgan 2006/5/1
               '分割案有收文申請寄存108之存活證明221期限管制
               If m_bol108 = True Then
                  'Added by Lydia 2022/09/15 抓約定期限
                  If Trim(m_stVar(0)) = "" Then m_stVar(0) = DBDATE(Text1(4))
                  If Trim(m_stVar(3)) = "" Then m_stVar(3) = DBDATE(Text1(5))
                  If m_pAgreeOnDate = "" Then
                      Call PUB_GetFCPOurDeadline(m_stVar(3), 4, , m_pAgreeOnDate)
                  End If
                  'end 2022/09/15
                  'Modify By Sindy 2021/4/27 + , DBDATE(m_pAgreeOnDate)
                  strSql = PUB_Get221SQL(pa, m_stVar(0), m_stVar(3), Text1(24), strReceiveNo, DBDATE(m_pAgreeOnDate))
                  cnnConnection.Execute strSql
               End If
            End If 'Added by Morgan 2013/9/17
         End If
      
      'Added by Morgan 2022/5/4
      '續行母案再審期限:分割未發文>>同分割期限(分案),分割已發文>>發文日+4個月(發文)--陳亭妙
      ElseIf Text1(1) = "435" And m_CP27 = "" Then
         strSql = "update caseprogress a set (cp06,cp07,cp48)=(select cp06,cp07,cp48 from caseprogress b where cp01=a.cp01 and cp02=a.cp02 and cp03=a.cp03 and cp04=a.cp04 and cp10='307')" & _
               " where cp09='" & strReceiveNo & "' and exists(select * from caseprogress b where cp01=a.cp01 and cp02=a.cp02 and cp03=a.cp03 and cp04=a.cp04 and cp10='307' and cp27 is null)"
         cnnConnection.Execute strSql, intI
      'end 2022/5/4
      
      'Add by Morgan 2006/5/1
      '申請寄存108之存活證明221期限管制
      ElseIf Text1(1) = "108" Then
         strSql = "select cp06,cp07,1 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='416' union all select np08,np09,2 from nextprogress where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np07='416' order by 3"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)  'edit by nickc 2007/02/05 不用 dll 了  objLawDll.ReadRstMsg(intI, strSQL)
         If intI = 1 Then
            With RsTemp
               If Val("" & .Fields(0)) > 0 And Val("" & .Fields(1)) > 0 Then
                  'Added by Lydia 2022/09/15 抓約定期限; ex.FCP-67719的申請寄存108沒有先抓到約定期限
                  If m_pAgreeOnDate = "" Then
                      Call PUB_GetFCPOurDeadline("" & .Fields(1), 4, , m_pAgreeOnDate)
                  End If
                  'end 2022/09/15
                  'Modify By Sindy 2021/4/27 + , DBDATE(m_pAgreeOnDate)
                  strSql = PUB_Get221SQL(pa, .Fields(0), .Fields(1), Text1(24), strReceiveNo, DBDATE(m_pAgreeOnDate))
                  cnnConnection.Execute strSql
               End If
            End With
         End If
      End If
      
      'Add by Morgan 2010/6/30
      '異議答辯、舉發答辯更新對造號數名稱為被異議(舉發)之C類來函資料
      If Text1(1) = "802" Or Text1(1) = "804" Then
         If Text1(6) <> "" And Text1(6) <> Text1(6).Tag Then
            strSql = "update caseprogress a set (cp36,cp37,cp38,cp39,cp40,cp41,cp42)=(select b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42 from caseprogress b where b.cp09=a.cp43) where CP09='" & Label3(8) & "'  and exists(select * from caseprogress b where b.cp09=a.cp43 and b.cp10 in ('1801','1802'))"
            cnnConnection.Execute strSql, intI
         End If
      End If
      'end 2010/6/30
      
      With MSHFlexGrid1
         For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "v" Then
               strTmp(1) = .TextMatrix(i, 7)
               strTmp(2) = .TextMatrix(i, 8)
               strTmp(3) = .TextMatrix(i, 9)
               
               'Add By Sindy 2021/8/16
               If Text1(1) = "404" Then '延期:點選下一程序性質時，若該性質續辦狀態是N，請將恢復期限管制，同時將解除期限日期及原因清除。
                  strSql = "UPDATE NEXTPROGRESS SET NP06='',NP11=null,NP12=null WHERE NP01='" & strTmp(1) & "' AND " & _
                     "NP07=" & strTmp(2) & " AND NP22=" & strTmp(3) & " AND NP06='N'"
                  cnnConnection.Execute strSql, intI
               Else
               '2021/8/16 END
                  'Modified by Lydia 2021/08/31 +NP24
                  'Modify By Sindy 2021/9/3 亭妙:控管924.會稿點選下一程序是不能沖銷期限→只是為帶相關總收文號而點選
                  If Text1(1) <> "924" Then
                  '2021/9/3 END
                     strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y',NP24='" & Label3(8) & "'  WHERE NP01='" & strTmp(1) & "' AND " & _
                        "NP07=" & strTmp(2) & " AND NP22=" & strTmp(3)
                     Pub_SeekTbLog strTxt(intStep) 'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業，若畫面勾選下一程序期限且存檔有上續辦Y的都寫Log以便事後能追蹤
                     cnnConnection.Execute strTxt(intStep)
                     intStep = intStep + 1
                  End If
               End If
            End If
         Next
      End With
      
      If Trim(Text1(2)) & Trim(Text1(25)) & Trim(Text1(26)) & Trim(Text1(27)) <> "" Then '轉本所案號
         Dim strKey1 As String
         Dim StrKey2 As String
         Dim strKey3 As String
         Dim strKey4 As String
         strKey1 = Text1(2)
         StrKey2 = Text1(25)
         strKey3 = Text1(26)
         If IsEmptyText(strKey3) Then strKey3 = "0"
         strKey4 = Text1(27)
         If IsEmptyText(strKey4) Then strKey4 = "00"
         
         strTxt(intStep) = "UPDATE CASEPROGRESS SET CP01='" & strKey1 & "'," & _
            "CP02='" & StrKey2 & "',CP03='" & strKey3 & "',CP04='" & strKey4 & "'" & _
            " WHERE CP09='" & Label3(8) & "'"
            
         cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
         
         Select Case pa(1)
            Case "FCP"
               strExc(0) = "SELECT COUNT(*) FROM PATENT WHERE " & ChgPatent(strKey1 & StrKey2 & strKey3 & strKey4)
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If RsTemp.Fields(0) = 0 Then
                  strTmp(1) = ""
                  For i = 5 To 99
                     strTmp(1) = strTmp(1) & "PA" & Format(i, "00") & ","
                  Next
                  For i = 100 To 132
                     strTmp(1) = strTmp(1) & "PA" & Format(i) & ","
                  Next
                  strTmp(1) = Left(strTmp(1), Len(strTmp(1)) - 1)
                  strTxt(intStep) = "INSERT INTO PATENT (PA01,PA02,PA03,PA04," & strTmp(1) & ") " & _
                     "SELECT '" & strExc(1) & "','" & strExc(2) & "','" & strExc(3) & "','" & strExc(4) & "'," & _
                     strTmp(1) & " FROM PATENT WHERE " & ChgPatent(Label3(9))
                     
                  cnnConnection.Execute strTxt(intStep)
                  intStep = intStep + 1
               End If
            Case "FG"
               strExc(0) = "SELECT COUNT(*) FROM SERVICEPRACTICE WHERE " & ChgService(strKey1 & StrKey2 & strKey3 & strKey4)
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If RsTemp.Fields(0) = 0 Then
                  strTmp(1) = ""
                  For i = 5 To tf_SP
                     Select Case i
                        Case 52, 53, 54, 55, 56, 57
                        Case Else
                           strTmp(1) = strTmp(1) & "SP" & Format(i, "00") & ","
                     End Select
                  Next
                  strTmp(1) = Left(strTmp(1), Len(strTmp(1)) - 1)
                  strTxt(intStep) = "INSERT INTO SERVICEPRACTICE (SP01,SP02,SP03,SP04," & strTmp(1) & ") " & _
                     "SELECT '" & strExc(1) & "','" & strExc(2) & "','" & strExc(3) & "','" & strExc(4) & "'," & _
                     strTmp(1) & " FROM SERVICEPRACTICE WHERE " & ChgService(Label3(9))
                  
                  cnnConnection.Execute strTxt(intStep)
                  intStep = intStep + 1
               End If
         End Select
      End If
   
   'Remove by Morgan 2008/9/23 統一改由收文設定或畫面輸入
   '   Select Case Me.Text1(1).Text
   '      '告知代理人, 回覆代理人, 專利調查, 調卷, 鑑定報告
   '      Case "901", "902", "903", "904", "906"
   '         '若原承辦期限為NULL且有輸入承辦人
   '         If m_strCP48 = "" And Me.Text1(0).Text <> "" Then
   '            'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
   '            m_strCP48 = Pub_GetHandleDay(pa(1), pa(9), Text1(1).Text, , m_strCP06)
   '                     " WHERE CP09='" & Label3(8) & "'"
   '            cnnConnection.Execute strSQL
   '         End If
   '   End Select
      
'Remove by Lydia 2016/06/21 改成模組PUB_SaveFCPcp14
'      'Add by Morgan 2008/9/4 更新齊備日
'      If text1(28).Locked = False Then
'         stUpdate = ",EP06=" & CNULL(DBDATE(text1(28).Text), True)
'      End If
'      'Add By Sindy 2015/12/16 有約定期限或指定日期
'      If m_CP27 = "" And m_CP57 = "" Then
'         If strSetLimitDT <> "" And text1(1) = "201" Then
'            '新案翻譯的核稿期限設定為前7個工作天,不能早於完稿日
'            strEP08 = DBDATE(CompWorkDay(8, DBDATE(strSetLimitDT), 1))
'            m_EP09 = DBDATE(m_EP09) '原完稿日
'            If Val(strEP08) < Val(m_EP09) Then
'               strEP08 = m_EP09
'            End If
'            stUpdate = ",EP08=" & CNULL(strEP08, True)
'            'e核稿人及工程師主管
'            If Me.text1(22).Text <> "" Then
'               strTo = Me.text1(22).Text & ";"
'            End If
'            If Left(strTo, 1) = "F" Then
'               strExc(0) = "SELECT st01,st15,st52 FROM staff WHERE st26 in(" & _
'                           " SELECT st26 FROM staff WHERE st01='" & Left(strTo, 5) & "'" & _
'                           " and st26 is not null)" & _
'                           " and substr(st01,1,1)<>'F'"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  strTo = RsTemp.Fields("st01") & ";"
'               End If
'            End If
'            If m_FCPTeam <> "" Then
'               strTemp = IIf(m_FCPTeam = "1", Pub_GetSpecMan("T"), IIf(m_FCPTeam = "2", Pub_GetSpecMan("R"), IIf(m_FCPTeam = "3", Pub_GetSpecMan("S"), Pub_GetSpecMan("T1"))))
'               If InStr(strTo, strTemp) = 0 Then strTo = strTo & strTemp
'            End If
'            strSubject = "請儘速辦理中說(翻譯/核稿)"
'            strContent = "FCP" & pa(2) & "因客戶催辦，指定於" & ChangeTStringToTDateString(strSetLimitDT) & IIf(text1(4) <> "" And DBDATE(text1(4)) <> strSetLimitDT, "(本所期限：" & ChangeTStringToTDateString(text1(4)) & ")", "") & vbCrLf & _
'                         "前呈送智慧局，請儘速辦理，謝謝。"
'            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                     " VALUES ( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
'                     ",to_char(sysdate,'hh24miss'),'" & strSubject & "','" & strContent & "')"
'            cnnConnection.Execute strSql
'         End If
'      End If
'      '2015/12/16 END
'      '更新核稿人
'      strSql = "Update EngineerProgress Set EP04='" & Me.text1(22).Text & "'" & stUpdate & " Where EP02='" & Me.Label3(8).Caption & "'"
'      cnnConnection.Execute strSql
'end 2016/06/21

      '2005/6/8 ADD BY SONIA 改案件性質或專利種類時重新抓費用及規費
      'Modified by Lydia 2017/12/14 +是否電子送件,有變更
      'If m_strCP10 <> text1(1) Or (Me.text1(21).Enabled And m_PA08 <> text1(21)) Then
      If m_strCP10 <> Text1(1) Or (Me.Text1(21).Enabled And m_PA08 <> Text1(21)) Or m_CP118 <> n_CP118 Then
         ' 規費
         '2010/8/17 MODFI BY SONIA
         'm_CP17 = Val(GetPatentOfficialFee(pa(1), Text1(1), Text1(5), Text1(21), "000", pa(16)))
         'Modified by Lydia 2017/12/14
         'm_CP17 = Val(GetPatentOfficialFee(pa(1), text1(1), text1(5), text1(21), "000", pa(16), pa(14), pa(2), pa(3), pa(4)))
         m_CP17 = Val(GetPatentOfficialFee(pa(1), Text1(1), Text1(5), Text1(21), "000", pa(16), pa(14), pa(2), pa(3), pa(4), n_CP118))
         ' 費用
         If Val(GetFCPFee(pa(1), Text1(1))) + Val(m_CP17) > 0 Then
            m_CP16 = Val(GetFCPFee(pa(1), Text1(1))) + Val(m_CP17)
            '點數
            m_CP18 = Format((Val(m_CP16) - Val(m_CP17)) / 1000, "0.0")
         End If
         strSql = " Update CaseProgress Set CP16=" & m_CP16 & ",CP17=" & m_CP17 & ",CP18=" & Val(m_CP18) & ",CP79=" & m_CP16 & "  WHERE CP09='" & Me.Label3(8).Caption & "'"
         cnnConnection.Execute strSql
         'Added by Lydia 2018/08/27 新案翻譯(201)預設個案的固定報價(PA62)
         If Text1(1) = "201" Then
             strExc(1) = Pub_GetPa62Flag(pa(1) & pa(2) & pa(3) & pa(4))
             strSql = "Update Patent set pa62='" & strExc(1) & "' where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "' "
             cnnConnection.Execute strSql
         ElseIf m_strCP10 = "201" Then
             strSql = "Update Patent set pa62=null where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "' "
             cnnConnection.Execute strSql
         End If
         'end 2018/08/27
      End If
      '2005/6/8 END
      
'Remove by Lydia 2016/06/21 改成模組PUB_SaveFCPcp14
'      'Add by Morgan 2007/8/29
'      If m_203CP09 <> "" And text1(5) <> "" Then
'         strSql = "update caseprogress set cp06=" & DBDATE(text1(4)) & ",cp07=" & DBDATE(text1(5)) & " where cp09='" & m_203CP09 & "'"
'         cnnConnection.Execute strSql, intI
'      End If
'      'end 2007/8/29
'
'      'Added by Morgan 2012/4/10
'      '若實審分案時有主動修正未發文且承辦期限較早者更新為相同
'      'Modified by Morgan 2012/5/1 +判斷申請案已送件--靜芳
'      If m_CP27 = "" And (text1(1) = "416" Or text1(1) = "203") And text1(23) <> "" Then
'         'Modified by Morgan 2013/1/4
'         '當主動修正後收文時,承辦期限更新至實審,若實審後收文則以實審承辦期限更新至主動修正-->以後收文者為準,不用比較大小
'         'strExc(0) = "select cp09,cp06 from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
'            " and cp10='" & IIf(text1(1) = "203", "416", "203") & "' and cp27||cp57 is null and nvl(cp48,0)<" & DBDATE(text1(23)) & _
'            " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp10='101' and b.cp27>0)"
'         strExc(0) = "select cp09,cp06 from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
'            " and cp10='" & IIf(text1(1) = "203", "416", "203") & "' and cp27||cp57 is null" & _
'            " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp10='101' and b.cp27>0)"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            strExc(1) = DBDATE(text1(23))
'            If Not IsNull(RsTemp.Fields("cp06")) And Val("" & RsTemp.Fields("cp06")) < strExc(1) Then
'               strExc(1) = RsTemp.Fields("cp06")
'            End If
'            strSql = "update caseprogress set cp48=" & strExc(1) & " where cp09='" & RsTemp("cp09") & "'"
'            cnnConnection.Execute strSql, intI
'         End If
'      End If
'
'      'Add by Morgan 2010/6/17 若已開請款單則換承辦人或核稿人時發Mail通知靜芳
'      'Modified by Lydia 2016/03/18 收件人改為FCP管制人
'      'Modified by Lydia 2016/06/15 若該筆請款單已收款(即nvl(a1k30,0)>0)則改發83002，副本才給FCP管制人；
'      If m_CP60 > "X" Then
'         PUB_PointReAssignInform Label3(9), m_CP60, m_CP14, text1(0), m_EP04, text1(22)
'      End If
'
'      'Add By Sindy 2015/12/15
'      If m_CP27 = "" And m_CP57 = "" Then
'         '已達法定期限前一個工作天
'         'Modify By Sindy 2016/1/8 改 達法定期限當天
'         'If text1(5) <> "" And strSrvDate(1) >= CompWorkDay(2, DBDATE(text1(5)), 1) Then
'         If text1(5) <> "" And strSrvDate(1) = DBDATE(text1(5)) Then
'         '2016/1/8 END
'            strExc(0) = "SELECT s1.st15 aST15,s1.st52 aST52,na01,na16,s2.st01 bST01,s2.st52 bST52" & _
'                        " FROM staff s1,nation,staff s2,fagent" & _
'                        " WHERE s1.st01='" & text1(0) & "'" & _
'                        " and fa01='" & Left(ChangeCustomerL(text1(18)), 8) & "' and fa02='" & Mid(ChangeCustomerL(text1(18)), 9, 1) & "'" & _
'                        " and na01(+)=fa10" & _
'                        " and s2.st01(+)=na16"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            strTo = ""
'            If intI = 1 Then
'               '工程師
'               If "" & RsTemp.Fields("aST15") = "F21" Then
'                  strTo = text1(0) & ";"
'                  If m_FCPTeam <> "" Then
'                     strTemp = IIf(m_FCPTeam = "1", Pub_GetSpecMan("T"), IIf(m_FCPTeam = "2", Pub_GetSpecMan("R"), IIf(m_FCPTeam = "3", Pub_GetSpecMan("S"), Pub_GetSpecMan("T1"))))
'                     If InStr(strTo, strTemp) = 0 Then strTo = strTo & strTemp & ";"
'                  End If
'                  '加發管制人
'                  If "" & RsTemp.Fields("na16") <> "" Then
'                     If InStr(strTo, RsTemp.Fields("na16")) = 0 Then strTo = strTo & RsTemp.Fields("na16") & ";"
'                     If "" & RsTemp.Fields("bST52") <> "" Then
'                        If InStr(strTo, RsTemp.Fields("bST52")) = 0 Then strTo = strTo & RsTemp.Fields("bST52") & ";"
'                     End If
'                  End If
'                  If InStr(strTo, Pub_GetSpecMan("N")) = 0 Then strTo = strTo & Pub_GetSpecMan("N")
'               '程序人員
'               ElseIf "" & RsTemp.Fields("aST15") = "F22" Then
'                  strTo = text1(0) & ";"
'                  '一級主管
'                  If "" & RsTemp.Fields("aST52") <> "" Then
'                     If InStr(strTo, RsTemp.Fields("aST52")) = 0 Then strTo = strTo & RsTemp.Fields("aST52") & ";"
'                  End If
'                  '分案程序主管
'                  If InStr(strTo, Pub_GetSpecMan("N")) = 0 Then strTo = strTo & Pub_GetSpecMan("N")
'               End If
'               If strTo <> "" Then
'                  strSubject = "已達法定當天"
'                  strContent = "本所案號：" + pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) + vbCrLf + _
'                               "案件名稱：" + Combo1.Text + vbCrLf + _
'                               "案件性質：" + Label3(1) + vbCrLf + _
'                               "本所期限：" + ChangeTStringToTDateString(text1(4)) + vbCrLf + _
'                               "法定期限：" + ChangeTStringToTDateString(text1(5)) + vbCrLf + _
'                               "承 辦 人：" + Label3(0) + vbCrLf
'                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                     " VALUES ( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
'                     ",to_char(sysdate,'hh24miss'),'" & strSubject & "','" & strContent & "')"
'                  cnnConnection.Execute strSql
'               End If
'            End If
'         '已達本所前二個工作天
'         ElseIf text1(4) <> "" And strSrvDate(1) >= CompWorkDay(3, DBDATE(text1(4)), 1) Then
'            strExc(0) = "SELECT st15,st52" & _
'                        " FROM staff" & _
'                        " WHERE st01='" & text1(0) & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            strTo = ""
'            If intI = 1 Then
'               '工程師
'               If "" & RsTemp.Fields("st15") = "F21" Then
'                  '工程師本人
'                  strTo = text1(0) & ";"
'                  If m_FCPTeam = "" Then
'                     '分案程序主管
'                     If InStr(strTo, Pub_GetSpecMan("N")) = 0 Then strTo = strTo & Pub_GetSpecMan("N")
'                  Else
'                     '工程師主管
'                     strTemp = IIf(m_FCPTeam = "1", Pub_GetSpecMan("T"), IIf(m_FCPTeam = "2", Pub_GetSpecMan("R"), IIf(m_FCPTeam = "3", Pub_GetSpecMan("S"), Pub_GetSpecMan("T1"))))
'                     If InStr(strTo, strTemp) = 0 Then strTo = strTo & strTemp
'                  End If
'               End If
'               If strTo <> "" Then
'                  strSubject = "已達本所前二個工作天"
'                  strContent = "本所案號：" + pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) + vbCrLf + _
'                               "案件名稱：" + Combo1.Text + vbCrLf + _
'                               "案件性質：" + Label3(1) + vbCrLf + _
'                               "本所期限：" + ChangeTStringToTDateString(text1(4)) + vbCrLf + _
'                               "法定期限：" + ChangeTStringToTDateString(text1(5)) + vbCrLf + _
'                               "承 辦 人：" + Label3(0) + vbCrLf
'                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                     " VALUES ( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
'                     ",to_char(sysdate,'hh24miss'),'" & strSubject & "','" & strContent & "')"
'                  cnnConnection.Execute strSql
'               End If
'            End If
'         '達本所期限當天
'         ElseIf text1(4) <> "" And strSrvDate(1) = DBDATE(text1(4)) Then
'            strExc(0) = "SELECT st15,st52" & _
'                        " FROM staff" & _
'                        " WHERE st01='" & text1(0) & "'"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            strTo = ""
'            If intI = 1 Then
'               '程序人員
'               If "" & RsTemp.Fields("st15") = "F22" Then
'                  strTo = text1(0) & ";"
'                  '一級主管
'                  If "" & RsTemp.Fields("st52") <> "" Then
'                     If InStr(strTo, RsTemp.Fields("st52")) = 0 Then strTo = strTo & RsTemp.Fields("st52") & ";"
'                  Else
'                     '分案程序主管
'                     If InStr(strTo, Pub_GetSpecMan("N")) = 0 Then strTo = strTo & Pub_GetSpecMan("N")
'                  End If
'               End If
'               If strTo <> "" Then
'                  strSubject = "已達本所當天"
'                  strContent = "本所案號：" + pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) + vbCrLf + _
'                               "案件名稱：" + Combo1.Text + vbCrLf + _
'                               "案件性質：" + Label3(1) + vbCrLf + _
'                               "本所期限：" + ChangeTStringToTDateString(text1(4)) + vbCrLf + _
'                               "法定期限：" + ChangeTStringToTDateString(text1(5)) + vbCrLf + _
'                               "承 辦 人：" + Label3(0) + vbCrLf
'                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                     " VALUES ( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
'                     ",to_char(sysdate,'hh24miss'),'" & strSubject & "','" & strContent & "')"
'                  cnnConnection.Execute strSql
'               End If
'            End If
'         End If
'      End If
'      '2015/12/15 END
      'Added by Lydia 2016/06/21 模組PUB_SaveFCPcp14
      'Modified by Morgan 2024/5/21 +m_strExSubj
      Call PUB_SaveFCPcp14(pa, Text1(23), strReceiveNo, Text1(1).Text, Left(Trim(cboCP14.Text), 5), m_CP14, Text1(4).Text, Text1(5).Text, m_CP27, m_CP57, m_CP60, _
                 Text1(18).Text, m_FCPTeam, Text1(22).Text, m_EP04, Text1(28), strSetLimitDT, m_203CP09, Label3(1).Caption, Combo1.Text, m_strSubject)
     
      'Added by Morgan 2012/8/8
      If Trim(cboCP14.Text) <> "" Then
         strSql = "update caseprogress set cp122='Y' where cp09='" & strReceiveNo & "' and cp122 is null"
         cnnConnection.Execute strSql, intI
         'Added by Lydia 2018/01/04 分案時，新案翻譯承辦人為舜禹、捷恩凱、迅達時，控管若原文字數欄位為空白，發一Email給程序管制人員，cc:Sharon
         'Modified by Lydia 2018/01/08 +m_cp66>="20180101" =>  從107/1/1開始控管 by Sharon
         'Modified by Lydia 2019/06/28 +固定請款對象之帳單（LEDES帳單）其請款項目209檢視中說英文敘述後方加上+英文字數
         'If m_CP66 >= "20180101" And Me.Text1(1).Text = "201" And txtTF23.Visible = True And Val(txtTF23.Text) = 0 _
                    And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, Trim(Left(cboCP14.Text, 6))) > 0 Then
         'Modified by Lydia 2025/03/13 改用模組取得
         'If txtTF23.Visible = True And Val(txtTF23.Text) = 0 And _
                    ((m_CP66 >= "20180101" And Me.text1(1).Text = "201" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, Trim(Left(cboCP14.Text, 6))) > 0) Or _
                    (m_CP60 = "" And Me.text1(1) = "209" And Me.text1(18) <> "" And InStr(FCP檢視中說必輸原文字數, ChangeCustomerL(Me.text1(18))) > 0)) Then
         If txtTF23.Visible = True And Val(txtTF23.Text) = 0 And _
                    ((m_CP66 >= "20180101" And Me.Text1(1).Text = "201" And InStr(Pub_SetF51Order("F", ""), Trim(Left(cboCP14.Text, 6))) > 0) Or _
                    (m_CP60 = "" And Me.Text1(1) = "209" And Me.Text1(18) <> "" And InStr(FCP檢視中說必輸原文字數, ChangeCustomerL(Me.Text1(18))) > 0)) Then
             strExc(1) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
             If strExc(1) <> "" Then
                   strExc(2) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "原文字數為空白,請輸入原文字數"
                   strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                              " VALUES ( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd')" & _
                              ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','同主旨','86013')"
                   cnnConnection.Execute strSql
             End If
         End If
         'end 2018/01/04
         'Added by Lydia 2018/05/31 檢視中說209及核對中說格式235分案時檢查有無會稿, 會稿自動掛承辦人和承辦期限
         If InStr(cAutoCP924, Text1(1).Text) > 0 Then
              msgTxt = PUB_Update924CP(pa(1), pa(2), pa(3), pa(4), Left(cboCP14.Text, 6), Text1(4).Text)
         End If
         'end 2018/05/31
      End If
      
      'Added by Lydia 2020/05/20 法律所案源收文：如果案件性質或申請國家有變化,則需要對應分案; 5/28 +配合開庭
      If strSrvDate(1) >= 法律所案源收文啟用日 And pa(1) = "FCP" And m_LOS07 = "" Then '排除已放棄的案源
          'Modified by Lydia 2020/07/23 重新整理: 因為案源收文已設定不可變更案件性質和申請國家,所以只要判斷有案源
          'If text1(1).Tag <> text1(1).Text Or (m_LOS15 = "" And txtLOSagree = "Y") Then
          '    Call PUB_UpdateCP10toPT(pa(1), pa(2), pa(3), pa(4), Label3(8), text1(1).Tag, "000", text1(1).Text, "000", text1(4).Text, text1(24).Text, pa(26), IIf(m_LOS15 = "" And txtLOSagree = "Y", True, False))
          'End If
          '
          'If m_CP14 = "" And Left(Trim(cboCP14.Text), 5) <> "" Then
          '  strSql = PUB_GetLOSkind(pa(1), text1(1).Text, "000")
          '  'Modified by Lydia 2020/06/09 判斷是否為補收文
          '  'If strSql <> "" Then
          '  strExc(1) = ""
          '  If m_LOS15 <> "" And strSql = "" Then strExc(1) = PUB_GetLOSplus(pa(1), pa(2), pa(3), pa(4), text1(1).Text, "000", m_LOS02)
          '  If strSql <> "" Or strExc(1) <> "" Then
          '  'end 2020/06/09
          '      Call PUB_UpdateLOS01(pa(1), pa(2), pa(3), pa(4), Label3(8), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30), txtLOSagree)
          '  End If
          'End If
          'Modified by Lydia 2022/03/03 增加判斷LOS01為空白時，要更新; ex.FCP-047557行政訴訟在收文時判斷有來函的期限,所以預設CP14
          'If m_LOS15 <> "" And m_CP14 = "" And Left(Trim(cboCP14.Text), 5) <> "" Then
          If m_LOS15 <> "" And (m_LOS01 = "" Or (m_CP14 = "" And Left(Trim(cboCP14.Text), 5) <> "")) Then
               Call PUB_UpdateLOS01(pa(1), pa(2), pa(3), pa(4), Label3(8), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30), txtLOSagree)
          End If
          'end 2020/07/23
      End If
      'end 2020/05/20
            
      'add by sonia 2015/10/13
      Select Case Me.Text1(1).Text
         Case "601" '領證
            '法定期限
            If Val(Me.Text1(5).Text) > 0 Then
               '小於收文日
               If Val(Me.Text1(5).Text) < Val(Me.Text1(12).Text) Then
                  '抓案件收費表的下次期限
                  'Modified by Lydia 2022/02/07 debug: strKey1=>pa(1) ; ex.FCP-64525
                  'If GetCF12(strKey1, pa(9), Me.text1(1).Text) <> 0 Then
                  '   m_strCP07 = DBDATE(CompDate(2, (GetCF12(strKey1, pa(9), Me.text1(1).Text)), Format(m_strCP07)))
                  'Else
                  '   m_strCP07 = DBDATE(CompDate(1, (GetCF28(strKey1, pa(9), Me.text1(1).Text)), Format(m_strCP07)))
                  If GetCF12(pa(1), pa(9), Me.Text1(1).Text) <> 0 Then
                     m_strCP07 = DBDATE(CompDate(2, (GetCF12(pa(1), pa(9), Me.Text1(1).Text)), Format(m_strCP07)))
                  Else
                     m_strCP07 = DBDATE(CompDate(1, (GetCF28(pa(1), pa(9), Me.Text1(1).Text)), Format(m_strCP07)))
                  'end 2022/02/07
                     'add by sonia 2018/2/21 原法定若為月底則延期後也要是月底 FCP-042474
                     'Modified by Lydia 2022/02/07 +DBDATE
                     PUB_LastDayConvert DBDATE(Me.Text1(5).Tag), m_strCP07
                  End If
                  '本所期限=系統日 2015/10/15 再加承辦期限=系統日
                  strSql = "UPDATE CASEPROGRESS SET CP48=" & strSrvDate(1) & ",CP06=" & strSrvDate(1) & ",CP07=" & CNULL(m_strCP07) & _
                           " WHERE CP09='" & strReceiveNo & "'"
                  cnnConnection.Execute strSql
               End If
            End If
      
         Case "605" '繳年費
            'modify by sonia 2015/10/15 重算原法定期限,以免下一程序已改為逾期6個月的期限 FCP-026640,故改以收文日與原法定期限比較
            'If Val(Me.text1(5).Text) < Val(Me.text1(12).Text) Then
            m_strCP07_1 = PUB_GetNextFeeDate(pa)
            If DBDATE(Text1(12)) > Val(m_strCP07_1) Then
            'end 2015/10/15
               If GetCF12(pa(1), pa(9), Me.Text1(1).Text) <> 0 Then
                  m_strCP07 = DBDATE(CompDate(2, (GetCF12(pa(1), pa(9), Me.Text1(1).Text)), Format(m_strCP07_1)))
               Else
                  m_strCP07 = DBDATE(CompDate(1, (GetCF28(pa(1), pa(9), Me.Text1(1).Text)), Format(m_strCP07_1)))
                  'add by sonia 2018/2/21 原法定若為月底則延期後也要是月底 FCP-042474
                  PUB_LastDayConvert m_strCP07_1, m_strCP07
               End If
               '102新法:年費收文已逾過期的6個月(跨102年)自動設法限=原法限(m_strCP07_1)+18月,所限=系統日
               If pa(9) = "000" And Val(m_strCP07) >= 20130101 And m_strCP07 < DBDATE(Text1(12)) Then
                  m_strCP07 = CompDate(1, 18, m_strCP07_1)
                  'add by sonia 2018/2/21 原法定若為月底則延期後也要是月底 FCP-042474
                  PUB_LastDayConvert m_strCP07_1, m_strCP07
               End If
               '本所期限=系統日 2015/10/15 再加承辦期限=系統日
               'Added by Lydia 2019/06/14 輸入指定日期，依照指定日期計算承辦期限。
               If Text1(30).Text <> "" Then
                    strSql = "UPDATE CASEPROGRESS SET CP06=" & strSrvDate(1) & ",CP07=" & CNULL(m_strCP07) & _
                             " WHERE CP09='" & strReceiveNo & "'"
               Else
               'end 2019/06/12
                    strSql = "UPDATE CASEPROGRESS SET CP48=" & strSrvDate(1) & ",CP06=" & strSrvDate(1) & ",CP07=" & CNULL(m_strCP07) & _
                             " WHERE CP09='" & strReceiveNo & "'"
               End If 'end 2019/06/14
               cnnConnection.Execute strSql
            'add by sonia 2017/4/6 FCP-038816下一程序改為逾期6個月的期限,但後來在原法定前又收文
            Else
               '更新原法定及本所,並參考收文程式重算承辦期限
               'Modified by Morgan 2023/7/5 法限不同才要依照新規則重算所限
               'm_strCP06 = CompDate(2, -2, m_strCP07_1)
               If m_strCP07 <> m_strCP07_1 Then
                  m_strCP06 = PUB_GetFCPOurDeadline(m_strCP07_1, 2)
               End If
               'end 2023/7/5
               If pa(75) = "Y45697" Then   'Ciba Y45697的年費承辦期限掛15個工作天
                  m_strCP48 = CompWorkDay(15, strSrvDate(1))
                  If Val(m_strCP06) > 0 And Val(m_strCP48) > Val(m_strCP06) Then
                     m_strCP48 = m_strCP06
                  End If
               Else
                  m_strCP48 = Pub_GetHandleDay("FCP", "000", Me.Text1(1).Text, , m_strCP06)
               End If
               'Added by Lydia 2019/06/14 輸入指定日期，依照指定日期計算承辦期限。
                        '1.起因: FCP-50441年費(AA7012152)在3/22有輸入指定日期108/5/20但是在3/31拿掉承辦期限,造成國外部專利期限彈跳在6/4才顯示在早收文階段
                        '2.修改: 於維護記錄log中發現3/22分案有更新2次承辦期限,第1次是依照畫面輸入108/5/20,第2次是領證或年費會更新原法限及所限並參考收文程式重算承辦期限(108/4/3)
                        '            與Sharon 確認:有輸入指定日期，依照指定日期計算承辦期限。
               If Text1(30).Text <> "" Then
                    strSql = "UPDATE CASEPROGRESS SET CP06=" & CNULL(m_strCP06) & ",CP07=" & CNULL(m_strCP07_1) & _
                             " WHERE CP09='" & strReceiveNo & "'"
               Else
               'end 2019/06/14
                    strSql = "UPDATE CASEPROGRESS SET CP48=" & CNULL(m_strCP48) & ",CP06=" & CNULL(m_strCP06) & ",CP07=" & CNULL(m_strCP07_1) & _
                             " WHERE CP09='" & strReceiveNo & "'"
               End If 'end 2019/06/14
               cnnConnection.Execute strSql
            'end 2017/4/6
            End If
            
      End Select
      'end 2015/10/13
   End If
   'Added by Lydia 2024/03/08 因應內專支援機械組OA (含P案) 收文: 927其他翻譯
   If mTransKind = "＃" And Trim(cboCP14.Text) <> "" Then
      'Added by Lydia 2024/07/03 參考106101901-OA委外翻譯點數控管：承辦人為外翻，進度備註請改為 OA委外翻譯
      If Trim(Left(cboCP14, 1)) = "F" Then
         strSql = "Update CaseProgress Set CP64='OA委外翻譯" & "'||decode(cp64,null,null,';')||cp64 where cp09=" & CNULL(Label3(8))
      Else
      'end 2024/07/03
         strSql = "Update CaseProgress Set CP64='" & Replace(GetStaffName(Trim(Left(cboCP14, 6))), "翻譯", "") & "翻譯" & Pub_GetNoToCPM("2", Label3(8)) & ";" & "'||cp64 where cp09=" & CNULL(Label3(8))
      End If
      cnnConnection.Execute strSql
   End If
   'end 2024/03/08
   
   cnnConnection.CommitTrans
   
   'Add by Morgan 2012/11/12
   If st307Msg <> "" Then MsgBox st307Msg
   'Added by Lydia 2018/05/31 彈會稿訊息
   If msgTxt <> "" Then MsgBox msgTxt, vbCritical
   
   Exit Function
CheckingErr:
   cnnConnection.RollbackTrans
   FormSave = False

End Function

'Added by Morgan 2012/6/18
'重讀資料
Public Sub QueryMainFile()
   IntNow = IntNow - 1
   GetData IntNow, True
End Sub

Private Sub GetData(intSitu As Integer, Optional bolNoMainForm As Boolean)
   Dim rsTmp1 As New ADODB.Recordset, i As Integer, txt As TextBox, Lbl As Object
   ' 90.10.18 modify by louis (記錄總收文號)
   Dim strCP09 As String
   'Add By Cheng 2001/12/25
   Dim ii As Integer '回圈流水號
   'Add By Cheng 2002/12/16
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
   'Add By Cheng 2002/06/10
   Me.Text1(21).Enabled = False
   '2005/6/15 ADD BY SONIA
   m_PA08 = ""

   For Each txt In Text1
      txt.Text = ""
      txt.Tag = "" 'Added by Lydia 2020/01/20
   Next
   'Add By Sindy 2021/5/11
   textCP64.Text = "": textCP64.Tag = ""
   textPA91.Text = "": textPA91.Tag = ""
   '2021/5/11 END
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   'Add By Sindy 2021/5/11
   For Each Lbl In Label5
      Lbl.Caption = ""
   Next
   '2021/5/11 END
   strSetLimitDT = "" 'Added by Lydia 2021/08/27 發生前一畫面資料保留; ex.FCP-065514有指定日期會發email,因為先分案造成後續案都比照辦理
   
   '總收文號
   Label3(8) = StrTot2(intSitu)
   ' 90.10.18 modify by louis (記錄總收文號)
   strCP09 = StrTot2(intSitu)
   strReceiveNo = strCP09
   '本所案號
   Label3(9) = StrTot1(intSitu)
   
   ChgCaseNo Label3(9), pa
   
   i = Len(Label3(9)) - 9
   pa(1) = Left(Label3(9), i)
   pa(2) = Mid(Label3(9), i + 1, 6)
   pa(3) = Mid(Label3(9), i + 7, 1)
   pa(4) = Right(Label3(9), 2)
   Combo1.Clear
   m_FCPTeam = "" 'Add By Sindy 2016/1/28
   m_NA16Na79 = "" 'Added by Lydia 2024/10/04
   If pa(1) = "FCP" Then '讀取Patent專利檔
      If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         If pa(23) = "1" Then
            Combo1.AddItem "中 : " & pa(5)
            Combo1.AddItem "英 : " & pa(6)
            'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
            Combo1.AddItem "外 : " & pa(7)
         Else
            strExc(0) = "SELECT CP37,CP38,CP39 FROM CASEPROGRESS WHERE " & ChgCaseprogress(StrTot1(intSitu))
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            With RsTemp
               If intI = 1 Then
                  If Not IsNull(.Fields(0)) Then
                     Combo1.AddItem "中 : " & .Fields(0)
                  Else
                     Combo1.AddItem "中 : "
                  End If
                  If Not IsNull(.Fields(1)) Then
                     Combo1.AddItem "英 : " & .Fields(1)
                  Else
                     Combo1.AddItem "英 : "
                  End If
                   'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
                  If Not IsNull(.Fields(2)) Then
                     Combo1.AddItem "外 : " & .Fields(2)
                  Else
                     Combo1.AddItem "外 : "
                  End If
               End If
            End With
         End If
         Combo1.ListIndex = 0
         Text1(3) = pa(23)
         Text1(8) = pa(48)
         m_FCPTeam = pa(150) 'Add By Sindy 2016/1/28
         If pa(57) = "Y" Then
            Label4 = "已閉卷"
            Label1(43).Visible = True
            Text1(11).Visible = True
         Else
            Label4 = ""
            Label1(43).Visible = False
            Text1(11).Visible = False
         End If
         
         Text1(16).Enabled = True
         Text1(17).Enabled = True
         For i = 26 To 30
            If pa(i) <> "" Then Text1(i - 13) = pa(i): ChgType (i - 13)
         Next
         If pa(75) <> "" Then Text1(18) = pa(75): ChgType (18)
         Text1(18).Tag = Text1(18).Text 'Added by Lydia 2020/01/20
         textPA91 = pa(91)
         'Add By Cheng 2002/06/10
         'Me.Text1(21).Enabled = True 'Removed by Morgan 2014/11/3 移到下面
         If pa(8) <> "" Then
            Me.Text1(21).Text = pa(8)
            '2005/6/15 ADD BY SONIA
            m_PA08 = pa(8)
            Me.Label3(11).Caption = "" & PUB_GetPatentKindName(Me.Text1(21).Text, pa(9)) 'Added by Lydia 2024/10/04
         End If
         'Add By Cheng 2002/08/22

         m_strCust1 = "" & Me.Text1(13).Text
         m_strCust2 = "" & Me.Text1(14).Text
         m_strCust3 = "" & Me.Text1(15).Text
         m_strCust4 = "" & Me.Text1(16).Text
         m_strCust5 = "" & Me.Text1(17).Text
         
         m_PA163 = pa(163)
      End If
      
   ElseIf pa(1) = "FG" Then
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
         Combo1.AddItem "中 : " & pa(5)
         Combo1.AddItem "英 : " & pa(6)
         'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
         Combo1.AddItem "外 : " & pa(7)
         Combo1.ListIndex = 0
'            Text1(3) = pa(23)
         Text1(8) = pa(29)
         m_FCPTeam = pa(79) 'Add By Sindy 2016/1/28
         
         If pa(15) = "Y" Then
            Label4 = "已閉卷"
            Label1(43).Visible = True
            Text1(11).Visible = True
         Else
            Label4 = ""
            Label1(43).Visible = False
            Text1(11).Visible = False
         End If
         
         If pa(8) <> "" Then Text1(13) = pa(8): ChgType (13)
         If pa(58) <> "" Then Text1(14) = pa(58): ChgType (14)
         If pa(59) <> "" Then Text1(15) = pa(59): ChgType (15)
         Text1(16).Enabled = False
         Text1(17).Enabled = False
         
         If pa(26) <> "" Then Text1(18) = pa(26): ChgType (18)
         'pa(26) = pa(8) 'Mark by Lydia 2024/06/19
         textPA91 = pa(18)
         'Add By Cheng 2002/08/22
         m_strCust1 = "" & Me.Text1(13).Text
         m_strCust2 = "" & Me.Text1(14).Text
         m_strCust3 = "" & Me.Text1(15).Text
         m_strCust4 = "" & Me.Text1(16).Text
         m_strCust5 = "" & Me.Text1(17).Text
      End If
   End If
   
   'Added by Lydia 2024/10/04
   Label3(2).Caption = PUB_GetFCPGrpName(m_FCPTeam)  '工程師組別
   m_NA16Na79 = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) '程序人員
   Label5(11) = GetStaffName(m_NA16Na79, True)
   'end 2024/10/04
   
   '2007/8/13 ADD BY SONIA銷卷提醒
   CheckCaseDestroy pa(1), pa(2), pa(3), pa(4)
   '2007/8/13 END
   
   'Add By Cheng 2002/06/10
   m_strCP06 = ""
   m_strCP07 = "": m_strCP07_1 = ""
   m_strCP48 = ""
   'Add By Cheng 2002/11/04
   m_strCP10 = ""
    
    'Modify by Morgan 2004/3/16
    '增加分割案件資料
    'strExc(0) = "SELECT CP13,CP14,CP10,CP06,CP07,CP43,CP57,CP26,CP05," & _
      "CP64,CP48,CP60 FROM CASEPROGRESS WHERE CP09='" & StrTot2(intSitu) & "'"
   'Modify by Morgan 2011/4/22 +CP30
   'Modified by Morgan 2013/1/4 +CP122
   'Modified by Morgan 2013/11/3 +CP31
   'Modify by Sindy 2015/12/29 + CP141,CP142
   'Modified by Lydia 2017/12/14 +CP118
   'Modified by Lydia 2018/01/08 +CP66
   'Modified by Lydia 2020/08/18 +CP31
   'Modify By Sindy 2024/5/28 +,CP176
   strExc(0) = "SELECT CP13,CP14,CP10,CP06,CP07,CP43,CP57,CP26,CP05," & _
      "CP64,CP48,CP60, DC05, DC06, DC07, DC08,CP27,CP20,CP30,CP122,CP31,CP141,CP142,CP118,CP66,CP31,cp164,CP176 " & _
      "FROM CASEPROGRESS, DIVISIONCASE WHERE DC01(+)=CP01 AND DC02(+)=CP02 AND DC03(+)=CP03 AND DC04(+)=CP04 AND CP09='" & StrTot2(intSitu) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
    If intI = 1 Then
      'Added by Morgan 2014/11/3
      'Modified by Morgan 2014/12/11 +3字頭
      If pa(1) = "FCP" And (.Fields("cp31") = "Y" Or Left("" & .Fields("cp10"), 1) = "3") Then
         Me.Text1(21).Enabled = True
      End If
      'end 2014/11/3
      m_CP122 = "" & .Fields("CP122")
       If Not IsNull(.Fields(0)) Then
          Text1(24) = .Fields(0)
          ChgType (24)
       End If
       'Added by Lydia 2024/03/08 927其他翻譯，承辦人為程序人員
       If mTransKind = "＃" Then
          cboCP14.Text = ""
       Else
       'end 2024/03/08
         'Modify By Sindy 2016/7/28
         Call Frm060101_1_SetCboCP14(Text1(1), IIf(pa(1) = "FCP", pa(150), IIf(pa(1) = "FG", pa(79), "")), cboCP14)
         If Not IsNull(.Fields("CP14")) Then
           cboCP14.Text = .Fields("CP14"): CboCP14_Validate False
         End If
         '2016/7/28 EMd
       End If 'Added by Lydia 2024/03/08
       m_CP14 = "" & .Fields("CP14") 'Add by Morgan 2010/6/17
       m_CP13 = "" & .Fields("CP13") 'Add by Sindy 2023/12/6
       If Not IsNull(.Fields(2)) Then
          'Modified by Lydia 2018/11/22 記錄原案件性質
          'strKind = Text1(1)
          strKind = "" & .Fields(2)
          Text1(1) = .Fields(2)
          ChgType (1)
       End If
       
       If Not IsNull(.Fields(3)) Then Text1(4) = TransDate(.Fields(3), 1): m_strCP06 = .Fields(3)
       If Not IsNull(.Fields(4)) Then Text1(5) = TransDate(.Fields(4), 1): m_strCP07 = .Fields(4): m_strCP07_1 = .Fields(4)
       If Not IsNull(.Fields(5)) Then Text1(6) = .Fields(5)
       Text1(6).Tag = Text1(6) 'Add by Morgan 2010/6/30
       If Not IsNull(.Fields(6)) Then Text1(9) = TransDate(.Fields(6), 1)
       If Not IsNull(.Fields(7)) Then Text1(10) = .Fields(7)
       If Not IsNull(.Fields(8)) Then Text1(12) = TransDate(.Fields(8), 1)
       If Not IsNull(.Fields(9)) Then textCP64 = .Fields(9)
       
       'Added by Lydia 2024/05/30 勘誤公報控管：分案「變更401、更改403」的時候，判斷此案已有公告號，在進分案作業畫面中，詢問
       If pa(1) = "FCP" And Text1(6).Text = "" And (Text1(1) = 變更 Or Text1(1) = 更改) Then
          Text1(6) = Pub_GetProcCRC("2", pa(1), pa(2), pa(3), pa(4), Text1(1))
       End If
       'end 2024/05/30
       
       '910626 Sieg
       If Not IsNull(.Fields("CP60")) Then
          m_CP60 = .Fields("CP60")
       Else
          m_CP60 = ""
       End If
       'Add By Cheng 2002/06/10
       If Not IsNull(.Fields("CP48")) Then m_strCP48 = .Fields("CP48").Value
        'Add By Cheng 2002/11/04
        m_strCP10 = "" & .Fields("CP10").Value
        
        m_CP30 = "" & .Fields("CP30") 'Add by Morgan 2011/4/22
        m_CP31 = "" & .Fields("CP31") 'Added by Lydia 2020/08/18
        
        'Added by Morgan 2012/12/21
        SetDesignCase
         'End 2012/12/21
   
        'Add by Morgan 2004/3/17
        '分割母案本所案號
        txtDivCaseNo(1) = "" & .Fields("DC05").Value: txtDivCaseNo(1).Tag = txtDivCaseNo(1)
        txtDivCaseNo(2) = "" & .Fields("DC06").Value: txtDivCaseNo(2).Tag = txtDivCaseNo(2)
        txtDivCaseNo(3) = "" & .Fields("DC07").Value: txtDivCaseNo(3).Tag = txtDivCaseNo(3)
        txtDivCaseNo(4) = "" & .Fields("DC08").Value: txtDivCaseNo(4).Tag = txtDivCaseNo(4)
        If Text1(1) = "307" Then
            DivVisibleSwitch True
        Else
            DivVisibleSwitch False
        End If
        m_CP27 = "" & .Fields("CP27") 'Add by Morgan 2007/6/11
        m_CP57 = "" & .Fields("CP57") 'Add by Sindy 2015/12/15
        m_CP66 = "" & .Fields("CP66") 'Added by Lydia 2018/01/08
        
        'Modify by Morgan 2010/4/29 改加欄位顯示
        'm_CP20 = "" & .Fields("CP20") 'Add by Morgan 2009/10/6
        Text1(29) = "" & .Fields("CP20")
        m_CP20Default = PUB_GetCP20(pa(1), Text1(1))
        'end 2010/4/29
        
        'Add by Morgan 2008/8/19
        Text1(23) = TransDate("" & .Fields("CP48"), 1)
        If m_CP27 <> "" Then
            Text1(23).Locked = True
            Combo2.Enabled = False
        End If
        'end 2008/8/19
        
        'Add By Sindy 2015/12/29
        'Modified by Lydia 2021/11/03 C類來函客戶指定送件日：不會有CP141 (對智慧局)
        'If "" & .Fields("CP141") = "3" And Val("" & .Fields("CP142")) > 0 Then
        If Val("" & .Fields("CP142")) > 0 Then
           Text1(30).Text = Val("" & .Fields("CP142")) - 19110000
        End If
        '2015/12/29 END
        'Add By Sindy 2021/4/20
        If "" & .Fields("CP164") = "1" Then
           Option1(0).Value = True
        ElseIf "" & .Fields("CP164") = "2" Then
           Option1(1).Value = True
        'Add By Sindy 2021/10/20
        ElseIf "" & .Fields("CP164") = "3" Then
           Option1(2).Value = True
        End If
        '2021/4/20 END
        
        'Add By Sindy 2024/4/25 暫不送件
        chkCP176.Value = 0
        If "" & .Fields("CP176") = "Y" Then
           chkCP176.Value = 1
        End If
        chkCP176.Tag = "" & .Fields("CP176")
        '2024/4/25 END
        
        'Add By Sindy 2024/4/25 鎖住欄位
        If m_CP27 <> "" Or m_CP57 <> "" Then
            chkCP176.Enabled = False
            Text1(30).Locked = True
            Option1(0).Enabled = False
            Option1(1).Enabled = False
            Option1(2).Enabled = False
        End If
        '2024/4/25 END
        
        If Text1(1) = "605" Then chkCP176.Enabled = False 'Added by Morgan 2025/7/23 年費分案不可設定暫不送件（一律交由程序至年費維護中上暫不繳納Ｙ）---Winfrey
        
        'Added by Lydia 2017/12/14 參考內專分案,預設是否電子送件
        m_CP118 = "" & .Fields("CP118")
        n_CP118 = m_CP118
        'end 2017/12/14
        
        'Add By Sindy 2022/7/13
        '將案件性質: 124 回復優先權主張，在分案的時候掛上期限管制規則：
        '法定期限: 第1優先權日+16個月
        '本所期限及承辦期限:依新期限管制規則帶入
        If Text1(1) = "124" And Text1(5) = "" Then
            strExc(2) = ""
            strExc(0) = "select pd05 from pridate where pd01='" & pa(1) & "' and pd02='" & pa(2) & "' and pd03='" & pa(3) & "' and pd04='" & pa(4) & "' and pd05>0 order by pd05 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(2) = RsTemp.Fields(0)
            End If
            If strExc(2) <> "" Then
               strExc(5) = CompDate(1, 16, strExc(2)) '法限=第1優先權日+16個月
               If Val(DBDATE(strExc(5))) < Val(strSrvDate(1)) Then
                  MsgBox "已超過" & Label3(1) & "期限(" & TransDate(strExc(5), 1) & ")，不掛期限！", vbInformation
               Else
                  'Modify By Sindy 2022/11/28
                  'strExc(6) = PUB_GetWorkDay1(strExc(5), False)
                  'Text1(5) = TransDate(strExc(6), 1) '法限
                  Text1(5) = TransDate(strExc(5), 1) '法限
                  strExc(4) = PUB_GetFCPOurDeadline(strExc(5), 2) '所限
                  strExc(6) = PUB_GetWorkDay1(strExc(4), False) '所限:若為假日，故自動順延一天
                  Text1(4) = TransDate(strExc(6), 1) '所限
                  '承辦期限
                  Call PUB_GetFCPsetCP48(True, pa, m_CP27, Text1(1), Text1(0), m_CP122, Text1(4), Text1(5), Text1(23), Text1(28), Combo2, Text1(12))
                  MsgBox vbCrLf & "法定期限將設定為 " & Text1(5) & " ！" '& IIf(strExc(6) = strExc(5), "", vbCrLf & vbCrLf & "( 原期限 " & TransDate(strExc(5), 1) & " 為假日，故自動順延！ )")
                  '2022/11/28 END
               End If
            End If
        End If
        '2022/7/13 END
    End If
   End With
   
   'Added by Lydia 2017/12/14 參考內專分案,預設是否電子送件
   'Modify By Sindy 2024/11/11 1=有主管機關者
   ' 依操作的案件性質檢查是否屬於有呈送主管機關(不管是否為經濟部智慧財產局)，則"電子送件"欄位，請自動上"Y"，以防人員當紙本送件
   If PUB_ChkhadCF10forEMP_46(pa(1), pa(9), Trim(Text1(1))) = 1 _
      And pa(9) = "000" And m_CP27 = "" And m_CP118 = "" Then
      'Modify By Sindy 2024/12/2 敏莉說803舉發預設的電子送件"Y"請拿掉
      If Trim(Text1(1)) <> "803" Then
      '2024/12/2 END
         n_CP118 = "Y"
      End If
   End If
'   'Modified by Lydia 2018/05/17 排除對象非智慧局(告知代理人901,會稿924,回覆代理人902,其他翻譯927)
'   'Modified by Lydia 2019/11/20 FCP案的領證和年費已經統一用電子送件(by Phoebe)
'   'If pa(1) = "FCP" And pa(9) = "000" And InStr("601,605,232,421,807,941,501,503,803,804,901,924,902,927", Trim(Text1(1))) = 0 And m_CP27 = "" _
'           And InStr(NewCasePtyList, Text1(1)) = 0 And m_CP118 = "" Then
'   'Modified by Sindy 2021/11/5 增加排除(969提供本所意見)
'   'Modified by Sindy 2022/1/13 增加排除(968回覆說明書校閱)
'   If pa(1) = "FCP" And pa(9) = "000" And InStr("232,421,807,941,501,503,803,804,901,924,902,927,969,968", Trim(text1(1))) = 0 And m_CP27 = "" _
'           And InStr(NewCasePtyList, text1(1)) = 0 And m_CP118 = "" Then
'               strExc(0) = "select 1 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10 IN (" & NewCasePtyList & ") and cp118 is not null"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  'Modify By Sindy 2021/8/30 亭妙,淑華:現在普遍都電子送件，所以不用再提醒此案要電子送件了。
'                  'MsgBox "本案為電子送件案，本程序將預設為電子送件！", vbExclamation
'                  n_CP118 = "Y"
'               End If
'    End If
'   'end 2017/12/14
   '2024/11/11 END
   
   'Move by Lydia 2019/09/06 從最下方移上來; 因為FCP-61857的實審416承辦期限設定(PUB_GetFCPsetCP48),要判斷是否有修改案件性質。
   '案件性質
   Text1(1).Tag = Text1(1).Text
   '本所期限
   Text1(4).Tag = Text1(4).Text
   '法定期限
   Text1(5).Tag = Text1(5).Text
   '核稿人
   Text1(22).Tag = Text1(22).Text
   '承辦期限
   Text1(23).Tag = Text1(23).Text
   'end 2019/09/06
   
   GetGrid StrTot2(intSitu), 0
   
   IntNow = IntNow + 1
   
   If bolNoMainForm = False Then 'Added by Morgan 2012/6/18
       'Modify By Cheng 2002/11/04
       '若案件性質為"416","203","204"時, 不必檢查基本檔是否有申請案號
       'Modify By Cheng 2002/12/16
       '再加若案件性質為"901","902"時, 不必檢查基本檔是否有申請案號
       'Modify by Morgan 2004/4/2
       '加案件性質為"307"時, 不必檢查基本檔是否有申請案號
       'If m_strCP10 <> "416" And m_strCP10 <> "203" And m_strCP10 <> "204" And m_strCP10 <> "901" And m_strCP10 <> "902" Then
       'Modify by Morgan 2005/11/25 加924會稿
       '2010/5/24 MODIFY BY SONIA 加940工程師提申
       'MODIFY BY SONIA 2014/6/23 +949寄中說
       'Modified by Morgan 2022/5/4 +435續行母案再審
       If m_strCP10 <> "435" And m_strCP10 <> "416" And m_strCP10 <> "203" And m_strCP10 <> "204" And m_strCP10 <> "901" And m_strCP10 <> "902" And m_strCP10 <> "307" And m_strCP10 <> "924" And m_strCP10 <> "940" And m_strCP10 <> "949" Then
           ' 90.10.18 modify by louis (顯示專利基本檔)
           ShowMaintainForm strCP09, , , Me
       End If
      
'Removed by Morgan 2012/6/18 前畫面有控制
'      'Add By Cheng 2001/12/25
'      For ii = 0 To Forms.Count - 1
'         '專利案件基本資料維護(frm050701)
'         If Forms(ii).Name = "frm050701" Then
'            frm060101_1.ZOrder 1
'
'            'Add By Cheng 2002/01/03
'            frm050701.SelectToolbarButtom
'
'            DoEvents
'            Exit For
'         End If
'      Next ii
   End If
   
    strMurgitroyd = Pub_GetSpecMan("外專MURGITROYD設定") 'Added by Lydia 2021/01/06
    
   'Add by Morgan 2008/9/15 檢視中說209,製作中說210可輸齊備日以便計算承辦期限
   Text1(28).Locked = True
   m_EP06 = ""
   'Modified by Morgan 2013/11/6 +235核對中說格式
   'Modified by Lydia 2021/01/29 +系統別判斷
   'If Text1(1) = "209" Or Text1(1) = "235" Or Text1(1) = "210" Then
   If pa(1) = "FCP" And (Text1(1) = "209" Or Text1(1) = "235" Or Text1(1) = "210") Then
      strExc(0) = "Select EP06 From EngineerProgress, Staff Where EP04=ST01(+) And EP02='" & Me.Label3(8).Caption & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_EP06 = TransDate("" & RsTemp.Fields("ep06"), 1)
         Text1(28).Text = m_EP06
         Text1(28).Tag = Text1(28).Text
         Text1(28).Locked = False
      End If
   End If
   
   'Added by Lydia 2021/01/06 排除Murgitroyd案的檢視中說
   'Modified by Lydia 2021/01/29 +分別判斷 ex.FG-001253在分案時pa(75)非代理人
   'If pa(1) = "FCP" And strMurgitroyd <> "" And pa(75) <> "" And InStr(strMurgitroyd, ChangeCustomerL(pa(75))) > 0 And text1(1).Text = "209" Then
   '     text1(28).Locked = True
   If pa(1) = "FCP" And strMurgitroyd <> "" And pa(75) <> "" Then
      If InStr(strMurgitroyd, ChangeCustomerL(pa(75))) > 0 And Text1(1).Text = "209" Then Text1(28).Locked = True
   'end 2021/01/29
   End If
   'end 2021/01/06
   
   'Add By Sindy 2015/12/16 新案翻譯
   m_EP09 = "" '原完稿日
   'Added by Lydia 2018/12/06
   cboCP14.Locked = False
   'Modified by Morgan 2022/2/7 核稿人改只顯示不可修改--Sharon
   'text1(22).Locked = False
   Text1(22).Locked = True
   'end 2022/2/7
   'end 2018/12/06
   If Text1(1) = "201" Then
      strExc(0) = "Select EP09 From EngineerProgress Where EP02='" & Me.Label3(8).Caption & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_EP09 = TransDate("" & RsTemp.Fields("ep09"), 1)
      End If
      'Added by Lydia 2018/12/06 新案翻譯在輸入翻譯完稿日後,分案及發文之承辦人,不能被修改(Sharon)
      If m_EP09 <> "" Then
           cboCP14.Locked = True
           Text1(22).Locked = True
      End If
      'end 2018/12/06
   End If
   '2015/12/16 END
   
   'Modify by Morgan 2004/3/23
   '只有發明,新型的翻譯，檢視中說，製作中說會有核稿人
   'Modified by Morgan 2013/6/19 +927其他翻譯
   'Modified by Morgan 2013/11/6 +235核對中說格式
   'Modified by Lydia 2016/06/17 改成常數
   'If Text1(21) <> "3" And (Text1(1) = "201" Or Text1(1) = "209" Or Text1(1) = "235" Or Text1(1) = "210" Or Text1(1) = "927") Then
   If Text1(21) <> "3" And InStr(FCPHaveEP04, Text1(1)) > 0 And Text1(1) <> "" Then
      'Add By Cheng 2002/12/16
      '取得核稿人
      StrSQLa = "Select EP04,ST02 From EngineerProgress, Staff Where EP04=ST01(+) And EP02='" & Me.Label3(8).Caption & "'"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         Me.Text1(22).Text = "" & rsA.Fields(0).Value
         m_EP04 = "" & rsA.Fields(0).Value 'Add by Morgan 2010/6/17
         Me.Label5(12).Caption = "" & rsA.Fields(1).Value
      Else
          Me.Text1(22).Text = ""
          Me.Label5(12).Caption = ""
      End If
   Else
      Me.Text1(22).Text = ""
      Me.Label5(12).Caption = ""
   End If
   
'Removed by Morgan 2013/1/3 102新法
'   'Add by Morgan 2007/8/29 發明或新型的主動修正203若未發文且已有申請日時設定期限
'   '1.實審未發文時期限同該程序
'   '2.發明法限=申請日(最早優先權日)+15個月,新型法限=申請日+2個月;所限=法限-4天
'   'Modify by Morgan 2007/10/15 已有期限時不必更新--靜芳
'   'If (pa(8) = "1" Or pa(8) = "2") And m_CP27 = "" And pa(10) <> "" And Text1(1) = "203" Then
'   'Modify by Morgan 2008/10/9 +206 補充說明
'   'If (pa(8) = "1" Or pa(8) = "2") And m_CP27 = "" And pa(10) <> "" And text1(1) = "203" And text1(5) = "" Then
'   If (pa(8) = "1" Or pa(8) = "2") And m_CP27 = "" And pa(10) <> "" And (text1(1) = "203" Or text1(1) = "206") And text1(5) = "" Then
'
'      Else
'      'end 2012/11/21
'         strExc(4) = "": strExc(5) = ""
'         '發明
'         If pa(8) = "1" Then
'            strExc(0) = "select cp06,cp07 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='416' and cp07>0 and cp27 is null and cp57 is null"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            '有實審未發文
'            If intI = 1 Then
'               If Not IsNull(RsTemp.Fields(0)) Then
'                  MsgBox "有收文實體審查會將" & Label3(1) & "期限改與實審相同！", vbInformation
'                  text1(5) = TransDate("" & RsTemp.Fields(1), 1) '法限
'                  text1(4) = TransDate("" & RsTemp.Fields(0), 1) '所限
'               End If
'            Else
'               strExc(0) = "select pd05 from pridate where pd01='" & pa(1) & "' and pd02='" & pa(2) & "' and pd03='" & pa(3) & "' and pd04='" & pa(4) & "' and pd05>0 order by pd05"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  strExc(2) = RsTemp.Fields(0)
'               Else
'                  strExc(2) = DBDATE(pa(10))
'               End If
'               strExc(5) = CompDate(1, 15, strExc(2)) '法限=申請日(最早優先權日)+15月
'            End If
'         '新型
'         Else
'            strExc(2) = DBDATE(pa(10))
'            strExc(5) = CompDate(1, 2, strExc(2))  '法限=申請日+2月
'         End If
'
'         If strExc(5) <> "" And text1(5) = "" Then
'            If Val(DBDATE(strExc(5))) < Val(strSrvDate(1)) Then
'               MsgBox "已超過" & Label3(1) & "期限(" & TransDate(strExc(5), 1) & ")，不掛期限！", vbInformation
'            Else
'               strExc(6) = PUB_GetWorkDay1(strExc(5), False)
'               '所限=法限-4日
'               'Modify by Morgan 2009/10/16 本所用原來的法限推算--靜芳
'               'strExc(4) = CompDate(2, -4, strExc(6))
'               strExc(4) = CompDate(2, -4, strExc(5))
'
'               text1(5) = TransDate(strExc(6), 1)  '法限
'               text1(4) = TransDate(strExc(4), 1)  '所限
'               'Add by Morgan 2009/7/21
'               MsgBox vbCrLf & "法定期限將設定為 " & text1(5) & " ！" & IIf(strExc(6) = strExc(5), "", vbCrLf & vbCrLf & "( 原期限 " & TransDate(strExc(5), 1) & " 為假日，故自動順延！ )")
'            End If
'         End If
'      End If 'Added by Morgan 2012/11/15
'
'   End If
'   'end 2007/8/29
 'end 2013/1/4
   
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
    
   'Add by Morgan 2007/4/3
   If Text1(1) = "926" Then
      cmdSetDate.Visible = True
   End If
   'end 2007/4/3
   
'Removed by Morgan 2013/1/7 102新法,主動修正無期限,改以下面規則控制
'
'   '2009/10/15 add by sonia FCP主動修正203補充說明206收文日>=新案申請日時,承辦期限改為本所期限FCP-040067
'   If (Text1(1) = "203" Or Text1(1) = "206") And DBDATE(Text1(12)) >= DBDATE(pa(10)) And Text1(4) <> "" And DBDATE(Text1(23)) <> DBDATE(Text1(4)) Then
'      'Added by Morgan 2012/4/6 +控制d(或檢視中說或製作中說)未發文
'      strExc(0) = "select cp06,cp07 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in('201','209','210') and cp27>0"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 0 Then
'      'end 2012/4/6
'         MsgBox "申請案已發文，主動修正或補充說明的承辦期限改為本所期限，原承辦期限為" & ChangeTStringToTDateString(Text1(23)) & "！", vbInformation
'         Text1(23) = Text1(4)
'      End If
'   End If
'   '2009/10/15 end
'
'end 2013/1/7
      
'Remove by Lydia 2016/06/21 改成模組PUB_CheckFCPshowMsg
'   'Added by Morgan 2013/1/7
'   '102新法,主動修正無期限,但若新案已發文且201,209,210未發文則更新承辦期限預設為上述程序之所限
'   m_203CP48 = ""
'   If m_CP27 = "" And pa(10) <> "" And (Text1(1) = "203" Or Text1(1) = "206") Then
'      'Modified by Morgan 2013/11/6 +235核對中說格式
'      strExc(0) = "select cp06,cp07,cpm03 from caseprogress,casepropertymap where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
'         " and cp10 in('201','209','235','210') and cp27||cp57 is null and cp06>0 and cpm01(+)=cp01 and cpm02(+)=cp10"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         MsgBox "有" & RsTemp("cpm03") & "未發文，" & Label3(1) & "的承辦期限將設定為" & RsTemp("cpm03") & "的本所期限！", vbInformation
'         Text1(23) = TransDate("" & RsTemp.Fields("cp06"), 1)
'         m_203CP48 = Text1(23)
'      '發明若有實審未發文則所限與法限設相同
'      ElseIf pa(8) = "1" And Text1(5) = "" Then
'         strExc(0) = "select cp06,cp07,cpm03 from caseprogress,casepropertymap where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
'            " and cp10='416' and cp27||cp57 is null and cp06>0 and cpm01(+)=cp01 and cpm02(+)=cp10"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            MsgBox "有" & RsTemp("cpm03") & "未發文，" & Label3(1) & "的本所期限與法定期限將設定為相同！", vbInformation
'            Text1(4) = TransDate("" & RsTemp.Fields("cp06"), 1)
'            Text1(5) = TransDate("" & RsTemp.Fields("cp07"), 1)
'         End If
'      End If
'   End If
'   'end 2013/1/7
'end 2016/06/21

   'AssignNote 'Add by Morgan 2007/6/12  'cancel by sonia 2016/5/26 先收實審但不發文, 後續再收主動修正, 則會因此點限制而不能分案
   CheckRefNo
   FormActivate    '2008/11/27 ADD BY SONIA 上前畫一次選多筆時才會做 'Memo by Lydia 2016/06/21 單筆也會執行
   
   'Add by Morgan 2009/12/25 延期不可改案件性質
   MSHFlexGrid1.Enabled = True
   If Text1(1) = "404" Then
      Text1(1).Enabled = False
      If m_CP27 <> "" Then
         MSHFlexGrid1.Enabled = False '已發文不可再點選
      End If
   Else
      Text1(1).Enabled = True
   End If

   'Added by Morgan 2012/6/15
   '若為併號請以連絡單通知電腦中心處理
   If m_CP27 <> "" Or strReceiveNo > "C" Then
      Text1(2).Enabled = False
      Text1(25).Enabled = False
      Text1(26).Enabled = False
      Text1(27).Enabled = False
   Else
      Text1(2).Enabled = True
      Text1(25).Enabled = True
      Text1(26).Enabled = True
      Text1(27).Enabled = True
   End If
   'end 2012/6/15
   If cboCP14.Enabled And cboCP14.Visible = True Then
      cboCP14.SetFocus
   End If
   
'Removed by Morgan 2013/4/11 取消--靜芳,FCP案狀況不同無需提醒 Ex.FCP-46826
'   'Added by Morgan 2012/11/14
'   If text1(1) = "413" And m_CP27 = "" Then
'      strExc(1) = PUB_GetFirstPriDate(pa)
'      If strExc(1) = "" Then strExc(1) = DBDATE(pa(10))
'      strExc(2) = CompDate(1, 15, strExc(1))
'      If strSrvDate(1) > strExc(2) Then
'         MsgBox "已超過申請日(優先權日)起算15個月！", vbExclamation
'      End If
'   End If
'   'end 2012/11/14
'end 2013/4/11
   
    'Added by Lydia 2018/08/07 讀取命名作業資料
    'Modified by Lydia 2023/07/28 拿掉條件and nvl(tct05,0)> 0
    strExc(0) = "select TCT01,TCT10,TCT27,TCT28 from TRANSCASETITLE,caseprogress where TCT01=cp09(+) " & _
                      "and cp01= '" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        m_TCT01 = "" & RsTemp.Fields("TCT01")
        m_TCT10 = "" & RsTemp.Fields("TCT10")
        m_TCT27 = "" & RsTemp.Fields("TCT27")
        m_TCT28 = "" & RsTemp.Fields("TCT28")
    Else
        m_TCT01 = ""
        m_TCT10 = ""
        m_TCT27 = ""
        m_TCT28 = ""
    End If
    'end 2018/08/07
    
   '記錄原始值
   '承辦人
   cboCP14.Tag = Left(Trim(cboCP14.Text), 5)
   
   'Added by Morgan 2022/5/9
   '續行母案再審:進度檔有分割/主動修正未發文，承辦人帶該道進度的工程師，否則帶該區的程序人員。
   If Text1(1) = "435" And cboCP14.Text = "" Then
      strExc(0) = "select cp14 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in ('307','203') and cp27||cp57 is null order by cp09 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         cboCP14.Text = "" & RsTemp("CP14")
      Else
         cboCP14.Text = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
      End If
      If cboCP14 <> "" Then CboCP14_Validate True
   End If
   'end 2022/5/9
   
   'Added by Lydia 2018/05/31 檢視中說209及核對中說格式235若無承辦人，自動帶最新一道程序的工程師(F21)若無則帶命名人員
   'Modified by Lydia 2018/08/07 +其他案件性質
   'If InStr(cAutoCP924, Text1(1)) > 0 And cboCP14.Text = "" Then
   If (InStr(cAutoCP924, Text1(1)) > 0 Or InStr(cAutoCPMList, Text1(1)) > 0) And cboCP14.Text = "" Then
        'Modified by Lydia 2018/08/07 已有命名人員資料,只抓工程師
        'strExc(0) = "select 1 ord1,cp05, cp14 from caseprogress,staff where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp159=0 and cp14=st01(+) and st03='F21' " & _
                          "union all select 2 ord2,tct113, tct10 from caseprogress, transcasetitle where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp31='Y' and cp09=tct01(+) " & _
                          "order by 1, 2 desc"
        'Modified by Lydia 2018/08/08 新案翻譯201因為有外翻，所以抓核稿人
        'strExc(0) = "select cp05, cp14 from caseprogress,staff where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp159=0 and cp14=st01(+) and st03='F21' " & _
                          "order by 1 desc"
        'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
        strExc(0) = "select cp09,cp05, cp14 from caseprogress,staff where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp159=0 and cp14=st01(+) and cp10 <> '201' and st03='F21' and cp14<>'F4102' "
        strExc(0) = strExc(0) & "union all select cp09,cp05, ep04 from caseprogress, engineerprogress,staff where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp159=0 and cp10='201' and cp09=ep02 and ep04=st01(+) and st03='F21' and cp14<>'F4102' "
        strExc(0) = strExc(0) & "order by cp05 desc"
        'end 2018/08/08
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
           RsTemp.MoveFirst
           Do While Not RsTemp.EOF
                If "" & RsTemp.Fields("cp14") <> "" Then
                     cboCP14.Text = "" & RsTemp.Fields("cp14")
                     cboCP14.Text = PUB_SetEng(cboCP14.Text)  'Added by Lydia 2024/02/29 外專機械設計組人員異動調整程式
                     CboCP14_Validate True
                     Exit Do
                End If
                RsTemp.MoveNext
           Loop
        End If
        
        'Added by Lydia 2018/08/07 檢視中說209及核對中說格式235若無工程師,預設為命名人員
        'Modified by Lydia 2018/09/19 +其他特定性質 (ex.FCP-59191)
        'If InStr(cAutoCP924, Text1(1)) > 0 And cboCP14.Text = "" And m_TCT10 <> "" Then
        If (InStr(cAutoCP924, Text1(1)) > 0 Or InStr(cAutoCPMList, Text1(1)) > 0) And cboCP14.Text = "" And m_TCT10 <> "" Then
            cboCP14.Text = m_TCT10
            cboCP14.Text = PUB_SetEng(cboCP14.Text)  'Added by Lydia 2024/02/29 外專機械設計組人員異動調整程式
            CboCP14_Validate True
        End If
        
        'Added by Lydia 2018/08/07 其他特定性質若無工程師,預設為工程師主管
        If InStr(cAutoCPMList, Text1(1)) > 0 And cboCP14.Text = "" And m_FCPTeam <> "" Then
            cboCP14.Text = IIf(m_FCPTeam = "1", Pub_GetSpecMan("T"), IIf(m_FCPTeam = "2", Pub_GetSpecMan("R"), IIf(m_FCPTeam = "3", Pub_GetSpecMan("S"), Pub_GetSpecMan("T1"))))
            cboCP14.Text = PUB_SetEng(cboCP14.Text)  'Added by Lydia 2024/02/29 外專機械設計組人員異動調整程式
            CboCP14_Validate True
        End If

   End If
   'end 2018/05/31
   'Added by Lydia 2018/08/07 製作中說210預設為命名人員
   If Text1(1) = "210" And cboCP14.Text = "" And m_TCT10 <> "" Then
        cboCP14.Text = m_TCT10
        CboCP14_Validate True
   End If
   'end 2018/08/07
   
   'Added by Lydia 2020/05/20 法律所案源收文
   Call ReadLOS
   Call SetLOSagree
     
End Sub

'Add by Morgan 2007/10/18 讓與、合併、繼承、申請人更名(變更)[未閉卷未銷案]
Private Function CheckRefNo() As Boolean
   If (Text1(1) = "401" Or Text1(1) = "701" Or Text1(1) = "702" Or Text1(1) = "703") And pa(57) = "" And pa(108) = "" Then
      strSql = "select np01 from nextprogress,caseprogress" & _
         " where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'" & _
         " and np06 is null and np07='202' and cp09(+)=np01 and cp10='928'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Text1(6) = "" Then
            Text1(6) = RsTemp(0)
            MsgBox "重新委任之補文件未收文，此道程序會自動帶重新委任之相關收文號！"
         ElseIf Text1(6) <> RsTemp(0) Then
            MsgBox "重新委任之補文件未收文，此道程序之相關收文號必須為重新委任之收文號！"
            Exit Function
         End If
      End If
   End If
   CheckRefNo = True
End Function
'Add by Morgan 2007/5/22
'只要是翻譯都要顯示，不必控制是否有輸承辦人，因為有可能會先紀錄相似折扣
Private Sub SetTF()
   'Modified by Lydia 2019/06/28 +209檢視中說
   If (Text1(1) = "201" Or Text1(1) = "927" Or Text1(1) = "209") Then
      lblTF05.Visible = True
      txtTF05.Visible = True
      lblTF18.Visible = True
      txtTF18.Visible = True
      'Added by Lydia 2017/05/17
      lblTF23.Visible = True
      txtTF23.Visible = True
      lblTF19.Visible = True
      txtTF19.Visible = True
      'end 2017/05/17
      strExc(0) = "select * from TransFee where TF01='" & Label3(8) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtTF05.Text = "" & RsTemp.Fields("TF05")
         txtTF18.Text = "" & RsTemp.Fields("TF18")
         'Added by Lydia 2017/05/17 原文字數TF23、相似度TF19、相似案號TF20
         txtTF23.Text = "" & RsTemp.Fields("TF23")
         txtTF19.Text = "" & RsTemp.Fields("TF19")
         txtTF19.Tag = "" & RsTemp.Fields("TF20")
         'end 2017/05/17
         If IsNull(RsTemp.Fields("TF07")) Then
            txtTF05.Enabled = True
            txtTF18.Enabled = True
         Else
            txtTF05.Enabled = False
            txtTF18.Enabled = False
         End If
      End If
   Else
      lblTF05.Visible = False
      txtTF05.Visible = False
      lblTF18.Visible = False
      txtTF18.Visible = False
      'Added by Lydia 2017/05/17
      lblTF23.Visible = False
      txtTF23.Visible = False
      lblTF19.Visible = False
      txtTF19.Visible = False
      'end 2017/05/17
   End If
End Sub

'Add by Morgan 2008/8/20
Private Sub FormActivate()
   Command2(2).Enabled = True 'Added by Morgan 2013/1/7
   
'Modified by Lydia 2016/06/18 改成模組
'   'If m_bActive = False Then     '2008/11/27 CANCEL BY SONIA
'   '   m_bActive = True
'      If m_CP27 = "" Then
'         'Modified by Morgan 2013/1/7 改判斷未分案確認(有些程序會在收文時自動分案)
'         'If Text1(23) = "" Then SetCP48
'         If m_CP122 = "" Then SetCP48
'         '告代分案檢查若新案已發文提醒是否重新計算承辦期限
'         If pa(1) = "FCP" And text1(1) = "901" And text1(0) = "" Then
'            strExc(0) = "select cp06 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp31='Y' and cp27>0"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               If MsgBox("新案已發文，是否重算【" & Label3(1) & "】承辦期限？", vbYesNo) = vbYes Then
'                  text1(23) = TransDate(Pub_GetHandleDay("FCP", "000", text1(1), , DBDATE(text1(4))), 1)
'               End If
'            End If
'         End If
'
'         'Add by Morgan 2008/9/2
'         '主動修正分案檢查若新案未發文提醒是否更改承辦期限與提申期限一致
'         '2009/1/5 modify by sonia 取消承辦人條件控制
'         'If pa(1) = "FCP" And InStr("203,204", text1(1)) > 0 And text1(0) = "" Then
'         If pa(1) = "FCP" And InStr("203,204", text1(1)) > 0 Then
'            strExc(0) = "select cp06 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp31='Y' and cp27 is null and cp06>0"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               If text1(23) = "" Or Val(DBDATE(text1(23))) > Val(RsTemp.Fields(0)) Then
'                  If MsgBox("新案未發文，是否更新【" & Label3(1) & "】承辦期限與提申期限一致？", vbYesNo) = vbYes Then
'                     text1(23) = TransDate(RsTemp.Fields(0), 1)
'                  End If
'               End If
'            End If
'         End If
'      End If
'   'End If
'
'      'Add by Morgan 2010/3/23
'      If text1(1) = "201" And text1(5) = "" Then
'         strExc(0) = "select cp27 from caseprogress where cp01='" & pa(1) & "'" & _
'            " and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
'            " and cp10='101' and cp27>0"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            MsgBox "翻譯無期限！"
'            '法限=申請案發文日+4個月
'            strExc(1) = CompDate(1, 4, RsTemp(0))
'            text1(5) = TransDate(strExc(1), 1)
'
'            'Modified by Morgan 2014/11/20 外專改回舊規則
'            ''Added by Morgan 2014/10/29
'            'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
'            '   Text1(4) = TransDate(PUB_GetOurDeadline(strExc(1)), 1)
'            'Else
'            ''end 2014/10/29
'               '所限=法限-4天
'               strExc(2) = CompDate(2, -4, strExc(1))
'               text1(4) = TransDate(strExc(2), 1)
'               text1(4).Tag = text1(4).Text 'Add By Sindy 2015/12/16
'            'End If 'Added by Morgan 2014/10/29
'            'end 2014/11/20
'         End If
'      End If
'      'end 2010/3/23
      Text1(0) = Left(Trim(cboCP14.Text), 5)
      Label3(0) = Trim(Mid(Trim(cboCP14.Text), 6))
      'Modify By Sindy 2021/9/8 + , textCP64
      Call PUB_CheckFCPshowMsg(Me.Visible, pa, m_CP27, Text1(1), Text1(0), m_CP122, Text1(4), Text1(5), Text1(23), Label3(1).Caption, Text1(28), Combo2, Text1(12), m_203CP48, textCP64)
      'Modify By Sindy 2021/9/6 會稿要預帶承辦人
      If Text1(0).Text <> Left(Trim(cboCP14.Text), 5) Then
         cboCP14.Text = Text1(0)
         strExc(0) = GetPrjSalesNM(Left(cboCP14.Text, 5))
         If strExc(0) <> "" Then
            cboCP14.Text = Left(cboCP14.Text, 5) & " " & strExc(0)
         End If
      End If
      '2021/9/6 END
'end 2016/06/18
      If Text1(1) = "103" Then
         If Not PUB_ChkCPExist(pa, "210") Then
            MsgBox "無收文製作中說！"
         End If
      End If
      
      'Added by Morgan 2012/12/21
      If pa(9) = "000" And pa(8) = "3" And (Text1(1) = "701" Or Text1(1) = "708" Or Text1(1) = "704" Or Text1(1) = "705") And (Len(pa(11)) = 9 Or Mid(pa(11), 10, 1) = "D") Then
         strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)" & _
            " from patent,caseprogress where substr(pa11,1,9)='" & Left(pa(11), 9) & "' and nvl(pa17,'Y')='Y' and pa11<>'" & pa(11) & "'" & _
            " and nvl(substr(pa11,10,1),'D')='D' and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
            " and cp10(+)='" & Text1(1) & "' and cp27(+) is null and cp57(+) is null and cp09 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            MsgBox "衍生設計案收文【 " & Label3(1) & " 】時母案及其他衍生設計案也需" & vbCrLf & "一併收文，但目前有未收文案號如下：" & vbCrLf & vbCrLf & RsTemp.GetString(adClipString, , , vbCrLf), vbExclamation
         End If
      End If
      'end 2012/12/21
   
      'Added by Morgan 2012/12/27
      If pa(9) = "000" Then
         If Text1(1) = "601" Then
            If pa(20) <> "" Then
               strExc(1) = CompDate(2, 1, pa(20))
               strExc(1) = CompDate(1, 3, strExc(1))
               strExc(1) = CompDate(2, -1, strExc(1))
               If strExc(1) < "20130101" And strExc(1) < strSrvDate(1) Then
                  MsgBox "原領證法限已逾期且早於 102/1/1，不可分案！"
                  Command2(2).Enabled = False
               End If
            End If
         End If
         
         If Text1(1) = "605" And pa(14) <> "" Then
            strExc(2) = Right(pa(72), 2)
            If Left(strExc(2), 1) = "," Then strExc(2) = Mid(strExc(2), 2)
            strExc(1) = CompDate(0, Val(strExc(2)), pa(14))
            strExc(1) = CompDate(2, -1, strExc(1))
            
            If strExc(1) < "20120701" And strExc(1) < strSrvDate(1) Then
               MsgBox "原年費法限已逾期且早於 101/7/1，不可分案！"
               Command2(2).Enabled = False
            End If
         End If
      End If
      'end 2012/12/27
End Sub

'Added by Morgan 2015/10/7
Private Sub Form_Activate()
   Static bolActivated As Boolean
   If Not bolActivated Then
      '檢查是否有來函的告代,回代或工程師承辦之C類來函未發文
      strExc(0) = "select cp09,cpm03 from caseprogress,staff,casepropertymap where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
         " and st01(+)=cp14 and cp09<>'" & strReceiveNo & "' and ((cp10 in ('901','902') and cp43>'C') or (cp09>'C' and st03='F21')) and cp27||cp57 is null and cpm01(+)=cp01 and cpm02(+)=cp10"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         
         With RsTemp
         strExc(1) = "收文號 " & .Fields("cp09") & " " & .Fields("cpm03") & "尚未發文!"
         .MoveNext
         Do While Not .EOF
            strExc(1) = strExc(1) & vbCrLf & "收文號 " & .Fields("cp09") & " " & .Fields("cpm03") & "尚未發文!"
            .MoveNext
         Loop
         End With
         MsgBox strExc(1), vbExclamation
      End If
      
      bolActivated = True
   End If
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
   ReDim pa(1 To TF_PA) As String
End Sub

'Added by Lydia 2018/05/21 改從SetParent傳收文號
'Modified by Lydia 2018/08/08 +只能上班翻譯 pTransKind
'Modified by Lydia 2018/09/12 +交稿期限pLimitDate
Public Sub SetParent(ByRef pForm As Form, ByVal pCnt As Integer, ByVal pCaseNo As String, ByVal pCpNo As String, Optional ByRef pTransKind As String, Optional ByRef pLimitDate As String)
Dim tmpArr As Variant, tmpArr2 As Variant
Dim intA As Variant
    Set m_PrevForm = pForm
    IntTot = pCnt
    tmpArr = Split(pCaseNo, ",")
    tmpArr2 = Split(pCpNo, ",")
    mTransKind = pTransKind 'Added by Lydia 2018/08/08
    mLimitDate = pLimitDate 'Added by Lydia 2018/09/12
    For intA = 0 To UBound(tmpArr)
        If Trim(tmpArr(intA)) <> "" Then
            StrTot1(intA) = tmpArr(intA)
            StrTot2(intA) = tmpArr2(intA)
        End If
    Next
End Sub

Private Sub Form_Load()
 Dim i As Integer
   MoveFormToCenter Me
   intWhere = 國外_FC
   'Remove by Lydia 2018/05/21 改從SetParent傳收文號
'   With frm060101.MSHFlexGrid1
'      intTot = 0
'      If .Rows < 2 Then Exit Sub
'      For i = 1 To .Rows - 1
'         If .TextMatrix(i, 0) = "v" Then
'            StrTot1(intTot) = .TextMatrix(i, 7) '本所案號
'            StrTot2(intTot) = .TextMatrix(i, 1) '收文號
'            intTot = intTot + 1
'         End If
'      Next
'   End With
   'end 2018/05/21
   
   'Added by Lydia 2020/05/20 法律所案源收文
   FraLOS.Visible = False
   FraLOS.BackColor = &H8000000F
   txtLOSagree.Text = ""
   FraLOS.Top = 3225
   'end 2020/05/20
   
   IntNow = 0
   SetCombo2 'Add by Morgan 2008/8/18
   SetCombo3 'Add by Morgan 2009/10/1
   GetData IntNow
   
   FCP檢視中說必輸原文字數 = Pub_GetSpecMan("FCP檢視中說必輸原文字數")
   'Modified by Lydia 2025/06/05 更改名稱
   'm_strBASF = Pub_GetSpecMan("外專翻譯分案-BASF") & ","  'Added by Lydia 2023/04/19
   m_str所內譯 = Pub_GetSpecMan("外專翻譯分案-所內譯") & ","
   m_str所內譯例外 = Pub_GetSpecMan("外專翻譯分案-所內譯例外") & "," 'Added by Lydia 2025/07/01
   
   Me.SSTab1.Tab = 0 'Added by Lydia 2021/09/09
End Sub

Private Sub SetCombo2()
   Combo2.Clear
   Combo2.AddItem "90 天", 0
   Combo2.ItemData(0) = 90
   Combo2.AddItem "75 天", 0
   Combo2.ItemData(0) = 75
   Combo2.AddItem "60 天", 0
   Combo2.ItemData(0) = 60
   Combo2.AddItem "45 天", 0
   Combo2.ItemData(0) = 45
   Combo2.AddItem "30 天", 0
   Combo2.ItemData(0) = 30
   Combo2.AddItem "3 週", 0
   Combo2.ItemData(0) = 21
   Combo2.AddItem "2 週", 0
   Combo2.ItemData(0) = 14
   Combo2.AddItem "1 週", 0
   Combo2.ItemData(0) = 7
End Sub

Private Sub SetCombo3()
   Combo3.Clear
   Combo3.AddItem "36 個月", 0
   Combo3.ItemData(0) = 36
   Combo3.AddItem "30 個月", 0
   Combo3.ItemData(0) = 30
   Combo3.AddItem "27 個月", 0
   Combo3.ItemData(0) = 27
   Combo3.AddItem "24 個月", 0
   Combo3.ItemData(0) = 24
   Combo3.AddItem "18 個月", 0
   Combo3.ItemData(0) = 18
   Combo3.AddItem "12 個月", 0
   Combo3.ItemData(0) = 12
   Combo3.ListIndex = -1
   Combo3.Enabled = False
End Sub

Private Function ChgType(i As Integer) As Boolean
Dim strTempName As String
Dim m_Team As String '2011/11/30 add by sonia

   ChgType = False
   Select Case i
      'Case 0, 24
      Case 24
'Modified  by Lydia 2016/06/21 改成模組
'         If Text1(i) <> "" Then
'           'edit by nickc 2007/02/02 不用 dll 了
'            'If objPublicData.GetStaff(Text1(i), strTempName) Then
'            If ClsPDGetStaff(Text1(i), strTempName) Then
'               If i = 0 Then
'                   '2011/11/30 ADD BY SONIA 林信昌因分組故自動帶與案件組別的編號
'                  If InStr(strTempName, "林信昌") > 0 Then
'                     If pa(1) = "FCP" Then m_Team = pa(150)
'                     If pa(1) = "FG" Then m_Team = pa(79)
'                     Select Case m_Team
'                        Case "1"
'                           If Left(Text1(0), 1) = "6" Then Text1(0) = "68091"
'                           If Left(Text1(0), 1) = "F" Then Text1(0) = "F5644"
'                        Case "2"
'                           If Left(Text1(0), 1) = "6" Then Text1(0) = "68092"
'                           If Left(Text1(0), 1) = "F" Then Text1(0) = "F5645"
'                        Case Else
'                           If Left(Text1(0), 1) = "6" Then Text1(0) = "68007"
'                           If Left(Text1(0), 1) = "F" Then Text1(0) = "F5162"
'                     End Select
'                     If ClsPDGetStaff(Text1(i), strTempName) Then
'                     End If
'                  End If
'                  '2011/11/30 END
'                  Label3(0) = strTempName
'               ElseIf i = 24 Then
'                  Label5(10) = strTempName
'               End If
'               ChgType = True
'                'Add By Cheng 2003/08/14
'                '若案件性質為翻譯(201), 檢視中說(209), 製作中說(210)時, 若承辦人為所內之員工時, 則核稿人預設為承辦人, 若非所內之員工時, 則核稿人不預預, 可空白
'               If i = 0 Then
'                    'Modify by Morgan 2004/3/17
'                    '當更改承辦時，核稿人要同步更新
'                    '專利種類<>非設計
''                    If Me.text1(22).Text = "" Then
''                        If Me.text1(1).Text = "201" Or Me.text1(1).Text = "209" Or Me.text1(1).Text = "210" Then
''                            If ChkStaffDepIsF51(Me.text1(0).Text) = False Then
''                                Me.text1(22).Text = Me.text1(0).Text
''                                Me.Label5(12).Caption = GetStaffName(Me.text1(0).Text)
''                            End If
''                        End If
''                    End If
'
'                  ChgType = CheckCP14
'               End If
'            Else
'               If i = 0 Then
'                  Label3(0) = ""
'               ElseIf i = 24 Then
'                  Label5(10) = ""
'               End If
'            End If
'         End If
         '判斷承辦人和智權人員
         If PUB_FCPGetCP14EP04(IIf(i = 0, "CP14", ""), pa, Text1(i), IIf(i = 0, Trim(Mid(Trim(cboCP14.Text), 6)), Label5(10))) Then
             ChgType = True
         End If
'         If i = 0 Then
'            ChgType = PUB_FCPCheckCP14(pa, text1(1), Left(Trim(cboCP14.Text), 5), text1(22), Label5(12))
'         End If
         
      Case 1 '案件性質
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCaseProperty(pA(1), Text1(i), strTempName, False) Then
         If ClsPDGetCaseProperty(pa(1), Text1(i), strTempName, False) Then
            Label3(1) = strTempName
            'Add by Morgan 2004/3/23
            '案件性質為分割時，顯示分割母案本所案號
            If Text1(1) = "307" Then
                DivVisibleSwitch True
            Else
                DivVisibleSwitch False
            End If
            'Modify By Sindy 2016/8/1
            Text1(0) = Left(Trim(cboCP14.Text), 5)
            Label3(0) = Trim(Mid(Trim(cboCP14.Text), 6))
            If Text1(0) = "" Then
               ChgType = True
            Else
            '2016/8/1 END
               'Modified by Lydia 2016/06/21
               'ChgType = CheckCP14
               ChgType = PUB_FCPCheckCP14(pa, Text1(1), Text1(0), Text1(22), Label5(12))
            End If
            'Add end
         End If
        
      Case 13, 14, 15, 16, 17
         strExc(1) = Text1(i).Text
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetCustomer(strExc(1), strTempName) Then
         If ClsPDGetCustomer(strExc(1), strTempName) Then
            Text1(i).Text = strExc(1)
            Label5(i - 11) = strTempName
            
            '910626 Sieg
            If i = 13 Then
               If m_CP60 <> "" And InStr(ChangeCustomerL(pa(26)), ChangeCustomerL(strExc(1))) = 0 Then
                  strExc(1) = pa(1)
                  strExc(2) = pa(2)
                  strExc(3) = pa(3)
                  strExc(4) = pa(4)
                  strExc(5) = m_CP60
                  strExc(6) = Text1(13)
                  strExc(7) = strTempName
                  '911118 nick 新增申請人
                  strExc(8) = pa(26)
                  'edit by nickc 2007/02/05 不用 dll 了
                  'If Not objLawDll.UpdAcc0k0(strExc()) Then
                  If Not ClsLawUpdAcc0k0(strExc()) Then
                     Label5(i - 11) = ""
                  Else
                     ChgType = True
                  End If
               Else
                  ChgType = True
               End If
            Else
               ChgType = True
            End If
         Else
            Label5(i - 11) = ""
         End If
         
      Case 18
         strExc(1) = Text1(i).Text
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetAgent(strExc(1), strTempName) Then
         If ClsPDGetAgent(strExc(1), strTempName) Then
            Text1(i) = strExc(1)
            'Modify By Cheng 2003/08/14
'            Label5(7) = strTempName
            Label5(7) = Replace(strTempName, "&", "&&")
            ChgType = True
         Else
            Label5(7) = ""
         End If
   End Select
End Function
'Add by Morgan 2004/3/17
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
Private Function GetGrid(ByVal strRecive As String, ByVal intSitu As Integer) As Boolean
   GetGrid = True
   If intSitu = 0 Then
      strExc(1) = Label3(9)
   Else
      strExc(1) = Trim(Text1(2)) & Trim(Text1(25)) & Trim(Text1(26)) & Trim(Text1(27))
   End If
   'Modify by Morgan 2011/6/10 +排除程序管制的案件性質
   If pa(1) = "FCP" Then
      'Modify by Morgan 2004/8/5
      '催審不續辦者不出現
      strExc(0) = "SELECT '' C0,DECODE(PA09,'" & 台灣國家代號 & "',CPM03,CPM04) C1," & _
         SQLDate("NP08") & " C2," & SQLDate("NP09") & " C3,NP13,NP14," & SQLDate("NP11") & _
         " C6,NP01,NP07,NP22,NP15 FROM NEXTPROGRESS,CASEPROPERTYMAP,PATENT WHERE " & _
         ChgNextProgress(strExc(1)) & " AND (NP06<>'Y' OR NP06 IS NULL) AND " & _
         "CPM01=NP02 AND CPM02=NP07 AND " & ChgPatent(strExc(1)) & strNpSqlOfNoSalesDuty
   Else
      strExc(0) = "SELECT '' C0,DECODE(SP09,'" & 台灣國家代號 & "',CPM03,CPM04) C1," & _
         SQLDate("NP08") & " C2," & SQLDate("NP09") & " C3,NP13,NP14," & SQLDate("NP11") & _
         "C6,NP01,NP07,NP22,NP15 FROM NEXTPROGRESS,CASEPROPERTYMAP,SERVICEPRACTICE WHERE " & _
         ChgNextProgress(strExc(1)) & " AND (NP06<>'Y' OR NP06 IS NULL) AND " & _
         "CPM01=NP02 AND CPM02=NP07 AND " & ChgService(strExc(1)) & strNpSqlOfNoSalesDuty
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   If intSitu = 1 Then
      If pa(1) = "FCP" Then
         strExc(0) = "SELECT count(*) FROM PATENT WHERE " & ChgPatent(Trim(Text1(2)) & Trim(Text1(25)) & Trim(Text1(26)) & Trim(Text1(27)))
      Else
         strExc(0) = "SELECT count(*) FROM SERVICEPRACTICE WHERE " & ChgService(Trim(Text1(2)) & Trim(Text1(25)) & Trim(Text1(26)) & Trim(Text1(27)))
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If RsTemp.Fields(0) = 0 Then
         'edit by nickc 2007/02/02 不用 dll 了
         'If objPublicData.GetMaxNumber(pA(1), strExc(1)) Then
         If ClsPDGetMaxNumber(pa(1), strExc(1)) Then
            If Text1(2) > pa(1) & String(6 - Len(strExc(1)), "0") & strExc(1) Then
               MsgBox "新本所案號不可大於自動編號，請重新輸入 !", vbCritical
               GetGrid = False
            Else
               If MsgBox("此本所案號不存在 ( " & Trim(Text1(2)) & Trim(Text1(25)) & Trim(Text1(26)) & Trim(Text1(27)) & " ) ，請確認 ?", vbQuestion + vbYesNo) = vbNo Then
                  GetGrid = False
               End If
            End If
         End If
      End If
   End If
   GridHead
End Function

Private Sub MSHFlexGrid1_Click()
 Dim i As Integer
   With MSHFlexGrid1
      .col = 0
      If .CellBackColor = &HFFC0C0 Then
         For i = 0 To .Cols - 1
            .col = i
            .CellBackColor = .BackColor
         Next
         .col = 0: .Text = ""
         'Add By Cheng '2001/12/07
         .col = 2
         Text1(4) = ""
         Text1(4).Tag = "" 'Add By Sindy 2015/12/16
         .col = 3
         Text1(5) = ""
         .col = 7
         Text1(6) = ""
         .col = 10
         textCP64 = ""
         m_CP30 = "" 'Add by Morgan 2011/4/22
      Else
         'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
         If Pub_CheckNpTheSameShow(pa(1), Text1(1), Trim("" & .TextMatrix(.row, 8))) = False Then
             Exit Sub
         End If
         'end 2021/08/31
         For i = 0 To .Cols - 1
            .col = i
            .CellBackColor = &HFFC0C0
         Next
         'Modify By Sindy 2021/8/16 mark
'         'Modified by Morgan 2015/10/12
'         '延期只更新期限及相關收文號不上續辦, Ex.FCP-46191
'         If Text1(1) <> "404" Then
            .col = 0: .Text = "v"
'         End If
'         'end 2015/10/12
         .col = 2
         Text1(4) = FCDate(.Text)
         Text1(4).Tag = Text1(4).Text 'Add By Sindy 2015/12/16
         .col = 3
         Text1(5) = FCDate(.Text)
         .col = 7
         Text1(6) = .Text
         ' 90.07.06 modify by louis (備註帶到進度備註欄位)
         .col = 10
         textCP64 = .Text
         m_CP30 = .TextMatrix(.row, 9) 'Add by Morgan 2011/4/22
      End If
      'If .Rows > 0 Then
      '   .Row = 1
      '   .Col = 7
      '   Text1(6) = .Text
      'End If
   End With
End Sub

''Add By Sindy 2021/9/3
''亭妙:我這邊提到的會稿是客戶單獨來會稿指示，要求205.申復或107.再審的924.會稿，
''這種類型的會稿我認為應該要和告代一樣，分案點進去後，自動掛本所期限及承辦期限還有承辦工程師。
''她和淑華討論後,會稿點選下一程序期限若為申復或再審就計算
''所限=系統日+14個工作天(不可大於申復或再審的所限)
''承辦期限=所限-2個工作天
'Private Sub PUB_GetFCPsetCP48_924() '會稿.924
'Dim i As Integer
'
'   If text1(1) = "924" Then
'      With MSHFlexGrid1
'         For i = 1 To .Rows - 1
'            If .TextMatrix(i, 0) = "v" And _
'               (.TextMatrix(i, 1) = "205" Or .TextMatrix(i, 1) = "107") Then
'               text1(4) = CompWorkDay(14, CompDate(2, 1, strSrvDate(1)), 0)
'               '不可大於點選的所限
'               If Val(.TextMatrix(i, 2)) > 0 Then
'                  If Val(DBDATE(text1(4))) > Val(DBDATE(.TextMatrix(i, 2))) Then
'                     text1(4) = Val(DBDATE(.TextMatrix(i, 2))) - 19110000
'                  End If
'               End If
'               text1(23) = CompWorkDay(2, CompDate(2, 1, DBDATE(text1(4))), 0)
'               If PUB_GetFCPCP14_F21(pa, "") = True Then '抓承辦人為工程師
'               End If
'               Exit For
'            End If
'         Next
'      End With
'   End If
'End Sub

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
      End If
   End If
   
'Add by Morgan 2009/10/1
Case 6
   Combo3.Tag = ""
   If Len(Text1(Index).Text) = 9 Then
      ChkCP43
   End If
   
'Add by Morgan 2007/5/22
Case 1
   SetTF
   SetLOSagree 'Added by Lydia 2020/05/20 法律所案源收文
End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
'   'add by nickc 2007/07/13 將輸入法改成使用API
'   If Index = 19 Or Index = 20 Then
'        OpenIme
'   End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 3
         KeyAscii = UpperCase(KeyAscii)
         If (KeyAscii > 51 Or KeyAscii < 49) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 10, 29
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 11
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         ElseIf KeyAscii = 89 Then
            If MsgBox("是否確定取消閉卷 ?", vbQuestion + vbYesNo) = vbNo Then
               KeyAscii = 0
               Beep
            End If
         End If
        'Modify By Cheng 2002/12/16
'      Case 0, 2, 6, 13, 14, 15, 16, 17, 18
      Case 2, 6, 13, 14, 15, 16, 17, 18, 22, 26
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Select Case Index
   Case 1 '案件性質
      CheckPromAndCK
   'Add By Sindy 2015/12/16
   Case 4, 30 '本所期限,指定日期
      '本所期限輸入約定期限 或 輸入指定日期
      If (Text1(4).Text <> "" And Text1(4).Tag <> Text1(4).Text) Or _
         (Text1(30).Text <> "" And Text1(30).Tag <> Text1(30).Text _
           And (Option1(0).Value = True Or Option1(1).Value = True Or Option1(2).Value = True) _
         ) Then
         'Modified by Lydia 2016/06/21 改成模組PUB_GetFCPsetCP48Limit
         'Call SetCP48Limit
         Text1(0) = Left(Trim(cboCP14.Text), 5)
         Label3(0) = Trim(Mid(Trim(cboCP14.Text), 6))
         'Modify By Sindy 2021/9/2 + , Text1(1):案件性質
         'Modify By Sindy 2021/9/2 + , IIf(Option1(0).Value = True, "1", IIf(Option1(1).Value = True, "2", "3")), pa(1)
         'Added by Lydia 2021/11/05 判斷C類來函有指定送件日，不變更預設承辦期限；ex.FCP-58246變更承辦人
         If Text1(30) <> "" And Left(Label3(8), 1) = "C" Then
         Else
         'end 2021/11/05
              'Modify By Sindy 2024/12/19 + , Label3(8)=收文號
              Call PUB_GetFCPsetCP48Limit(strSetLimitDT, Text1(0), Text1(4), Text1(23), Text1(30), Text1(1), IIf(Option1(0).Value = True, "1", IIf(Option1(1).Value = True, "2", "3")), pa(1), Label3(8))
         End If 'Added by Lydia 2021/11/05
      End If
   '2015/12/16 END
   Case 18
      If Text1(18) = "" And Text1(13) = "" And Text1(14) = "" And Text1(15) = "" And Text1(16) = "" And Text1(17) = "" Then
         MsgBox "申請人及代理人至少輸入一個 !", vbCritical
         Text1(13).SetFocus
      End If
   Case 27 '轉本所案號
      'Add By Cheng 2002/09/09
      If Me.Text1(2).Text <> "" And Me.Text1(25).Text <> "" Then
         MsgBox "不可同時做" & """轉本所案號""" & "及" & """修改其他欄位""" & "，否則資料可能有不一致的情形產生!!!", vbExclamation + vbOKOnly
      End If
End Select
End Sub

'Remove by Lydia 2016/06/21 改成模組PUB_GetFCPsetCP48Limit
''Add By Sindy 2015/12/16 本所期限輸入約定期限 或 加打指定日期 計算承辦期限
'Private Sub SetCP48Limit()
'   If Text1(0) <> "" Then '有輸入承辦人
''      strExc(0) = "SELECT st15,st52" & _
''                  " FROM staff" & _
''                  " WHERE st01='" & text1(0) & "'"
'      strExc(0) = "SELECT st01,st15,st52 FROM staff WHERE st01='" & Text1(0) & "' and substr(st01,1,1)<>'F'" & _
'                  " Union" & _
'                  " SELECT st01,st15,st52 FROM staff WHERE st26 in(" & _
'                  " SELECT st26 FROM staff WHERE st01='" & Text1(0) & "' and substr(st01,1,1)='F'" & _
'                  " and st26 is not null)" & _
'                  " and substr(st01,1,1)<>'F'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 0 Then
'         strExc(0) = "SELECT st01,st15,st52 FROM staff WHERE st01='" & Pub_GetSpecMan("M") & "'"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      End If
'      If intI = 1 Then
'         If "" & RsTemp.Fields("st15") = "F21" Or "" & RsTemp.Fields("st15") = "F22" Then
'            If "" & RsTemp.Fields("st15") = "F21" Then
'               '當承辦人為工程師,其承辦期限更新為指定日期前4個工作天
'               If Text1(30).Text <> "" Then
'                  strSetLimitDT = DBDATE(Text1(30).Text)
'                  Text1(23) = CompWorkDay(5, DBDATE(Text1(30)), 1) - 19110000
'               '當承辦人為工程師,其承辦期限更新為本所前4個工作天
'               Else
'                  strSetLimitDT = DBDATE(Text1(4).Text)
'                  Text1(23) = CompWorkDay(5, DBDATE(Text1(4)), 1) - 19110000
'               End If
'            ElseIf "" & RsTemp.Fields("st15") = "F22" Then
'               '當承辦人為程序同仁,其承辦期限更新為指定日期
'               If Text1(30).Text <> "" Then
'                  strSetLimitDT = DBDATE(Text1(30).Text)
'                  Text1(23) = Text1(30)
'               '當承辦人為程序同仁,其承辦期限更新為本所
'               Else
'                  strSetLimitDT = DBDATE(Text1(4).Text)
'                  Text1(23) = Text1(4)
'               End If
'            End If
'            '承辦期限小於等於系統日期時,承辦期限等於系統日
'            If Text1(23) <> "" And Val(Text1(23)) <= Val(strSrvDate(2)) Then
'               Text1(23) = strSrvDate(2)
'            End If
'         End If
'      End If
'   End If
'End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTempName As String, i As Integer, blnIsEmpty As Boolean

   Select Case Index
'      Case 0 '承辦人
'         If text1(Index) <> "" Then
'            If ChgType(Index) = False Then Cancel = True
'         '若清除承辦人時核稿人也要清
'         ElseIf Trim(cboCP14.Text) = "" Then
'            text1(22) = ""
'            Label5(12) = ""
'         End If
'
'         'Add by Morgan 2008/8/20
'         If Cancel = False And text1(Index).Text <> text1(Index).Tag Then
'           'Modified by Lydia 2016/06/21 改成模組
'           'SetCP48
'           Call PUB_GetFCPsetCP48(Me.Visible, pa, m_CP27, text1(1), Left(Trim(cboCP14.Text), 5), m_CP122, text1(4), text1(5), text1(23), text1(28), Combo2, text1(12))
'         End If
         
      Case 1 '案件性質
         If Text1(Index) <> "" Then
            If ChgType(Index) = False Then
               Cancel = True
            'Add by Morgan 2007/4/3
            ElseIf Text1(1) = "926" Then
               cmdSetDate.Visible = True
            Else
               cmdSetDate.Visible = False
            'end 2007/4/3
            End If
         Else
            MsgBox "案件性質不可空白 !", vbCritical
            Cancel = True
         End If
         'Add by Morgan 2008/8/20
         If Cancel = False And Text1(Index).Text <> Text1(Index).Tag Then
            'Add by Morgan 2008/9/4
            'Modify by Morgan 2008/9/10 +210 製作中說
            'Modified by Morgan 2013/11/6 +235核對中說格式
            'Modified by Lydia 2021/01/06 排除Murgitroyd案的檢視中說
            'Addd by Lydia 2021/01/29 +系統別判斷 ex.FG-001253在分案時pa(75)非代理人
            If pa(1) <> "FCP" Then
                Text1(28).Locked = True
            Else
            'end 2021/01/29
                If strMurgitroyd <> "" And pa(75) <> "" And InStr(strMurgitroyd, ChangeCustomerL(pa(75))) > 0 And Text1(Index).Text = "209" Then
                   Text1(28).Locked = True
                ElseIf Text1(Index).Text = "209" Or Text1(Index).Text = "235" Or Text1(Index).Text = "210" Then
                'end 2021/01/06
                   Text1(28).Locked = False
                Else
                   Text1(28).Locked = True
                End If
            End If 'end 2021/01/29
            Text1(28).Text = m_EP06
            'end 2008/9/4
            'Modified by Lydia 2016/06/21 改成模組
            'SetCP48
            Text1(0) = Left(Trim(cboCP14.Text), 5)
            Label3(0) = Trim(Mid(Trim(cboCP14.Text), 6))
            Call PUB_GetFCPsetCP48(Me.Visible, pa, m_CP27, Text1(1), Text1(0), m_CP122, Text1(4), Text1(5), Text1(23), Text1(28), Combo2, Text1(12))
            'Modify By Sindy 2021/9/6 會稿要預帶承辦人
            If Text1(0).Text <> Left(Trim(cboCP14.Text), 5) And Trim(cboCP14.Text) = "" Then
               cboCP14.Text = Text1(0)
               strExc(0) = GetPrjSalesNM(Left(cboCP14.Text, 5))
               If strExc(0) <> "" Then
                  cboCP14.Text = Left(cboCP14.Text, 5) & " " & strExc(0)
               End If
            End If
            '2021/9/6 END
            
            'Add by Morgan 2010/4/29
            If Text1(29) = m_CP20Default Then
               Text1(29) = PUB_GetCP20(pa(1), Text1(1))
               m_CP20Default = Text1(29)
            ElseIf MsgBox("原是否請款欄位已修改過，是否要重新設定？", vbYesNo + vbDefaultButton1) = vbYes Then
               Text1(29) = PUB_GetCP20(pa(1), Text1(1))
               m_CP20Default = Text1(29)
            End If
            SetDesignCase 'Added by Morgan 2012/12/27
         End If

      Case 3 '卷宗性質
         If Text1(3) = "" Then
            'Modify By Cheng 2003/02/25
            '若系統類別為"FG"則卷宗性質可空白
            If pa(1) <> "FG" Then
                MsgBox "卷宗性質不可空白 !", vbCritical
                Cancel = True
            End If
         Else
            If Text1(1) = 異議_專 Then
               If Text1(3) <> "2" Then
                  MsgBox "案件性質為異議時，卷宗性質必須為 2 !", vbCritical
                  Cancel = True
               End If
            ElseIf Text1(1) = 舉發 Then
               If Text1(3) <> "3" Then
                  MsgBox "案件性質為舉發時，卷宗性質必須為 3 !", vbCritical
                  Cancel = True
               End If
            'add by sonia 2021/2/26
            ElseIf Text1(1) = "807" Then
               If Text1(3) <> "3" Then
                  MsgBox "案件性質為第三人申請技術報告時，卷宗性質必須為 3 !", vbCritical
                  Cancel = True
               End If
            'end 2021/2/26
            End If
         End If
         
      Case 4, 5 '本所期限/法定期限
            'Modify By Sindy 2021/9/30 寫在此會詢問2次,改位置
'         If text1(Index).Tag <> text1(Index) Then
'            If Index = 4 Then
'               If MsgBox("是否確定修改本所期限 ?", vbQuestion + vbYesNo) = vbNo Then
'                  text1(Index) = text1(Index).Tag
'                  Exit Sub
'               End If
'            Else
'               If MsgBox("是否確定修改法定期限 ?", vbQuestion + vbYesNo) = vbNo Then
'                  text1(Index) = text1(Index).Tag
'                  Exit Sub
'               End If
'            End If
'         End If
         If Text1(Index) <> "" Then
            If ChkDate(Text1(Index)) Then
               If Index = 5 Then
                  If Val(Text1(4)) > Val(Text1(5)) Then
                      MsgBox "本所期限必須小於法定期限 !", vbCritical
                      Cancel = True
                  End If
               End If
            Else
               Cancel = True
            End If
            
            'Add By Sindy 2021/10/20 亭妙:面詢(408)分案時，請協助將面詢本所及法定期限設為與法定期限同一天。並備註面詢日期(___/___/___)
            If Text1(5) <> "" And Text1(4) <> Text1(5) And Text1(1) = "408" Then
               Text1(4) = Text1(5)
               If InStr(textCP64, ChangeTStringToTDateString(Text1(5).Text)) = 0 Then
                  textCP64 = "面詢日期(" & ChangeTStringToTDateString(Text1(5).Text) & ");" & textCP64.Text
               End If
            End If
            '2021/10/20 END
         Else
            If Text1(1) = 年費 Or Text1(1) = 延期 Then
               MsgBox "案件性質為年費或延期時，必須有期限 !", vbCritical
               Cancel = True
            End If
         End If
         'Text1(Index).Tag = Text1(Index) 'Add By Sindy 2015/12/17 Mark
         
      Case 6
         Combo3.ListIndex = -1
         Combo3.Enabled = False
         
         If Text1(Index) <> "" Then
            intI = 1
            strExc(0) = "SELECT CP01||CP02||CP03||CP04,CP10 FROM CASEPROGRESS WHERE CP09='" & Text1(6) & "'"
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) = pa(1) & pa(2) & pa(3) & pa(4) Then
                  If Text1(1) = 請求公告 Or Text1(1) = 延緩公告 Then
                     If Left(Text1(6), 1) <> "C" Then
                        MsgBox "案件性質為請求公告或延緩公告時，必須為 C 類之收文號 !", vbCritical
                        Cancel = True
                        GoTo gotoExit 'Add By Sindy 2021/9/6
                     End If
                  'Add by Morgan 2009/10/1 若為實審或再審的退費顯示期限設定選單
                  ElseIf Text1(1) = 退費 Then
                     'Modified by Morgan 2022/10/12 +435續行母案再審
                     If RsTemp.Fields(1) = "416" Or RsTemp.Fields(1) = "107" Or RsTemp.Fields(1) = "435" Then
                        Combo3.Enabled = True
                     'Added by Morgan 2015/4/30
                     ElseIf RsTemp.Fields(1) = "404" Then
                        strExc(0) = " select 2 from caseprogress,nextprogress where cp09='" & Text1(6) & "' and cp10='404' and cp84>0 and np01(+)=cp43 and to_char(np22)=cp30 and np07='107'"
                        strExc(0) = strExc(0) & " union select 3 from caseprogress a,caseprogress b where a.cp09='" & Text1(6) & "' and a.cp10='404' and a.cp84>0 and b.cp09(+)=a.cp43 and b.cp10='107'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           Combo3.Enabled = True
                        End If
                     'end 2015/4/30
                     End If
                  End If
               Else
                  MsgBox "必須為此本所案號之其他收文號 !", vbCritical
                  Cancel = True
                  GoTo gotoExit 'Add By Sindy 2021/9/6
               End If
            Else
               MsgBox "必須為此本所案號之其他收文號 !", vbCritical
               Cancel = True
               GoTo gotoExit 'Add By Sindy 2021/9/6
            End If
            
            'Add By Sindy 2021/9/6 會稿計算所限,承辦期限要檢查相關總收文號的所限
            'Modify By Sindy 2021/9/30 + And (text1(4) = "" Or cboCP14.Text = "" Or text1(23) = "")
            If Text1(1) = "924" And (Text1(4) = "" Or cboCP14.Text = "" Or Text1(23) = "") Then
               strExc(0) = " select cp06 from caseprogress where cp09='" & Text1(6) & "' and cp57||cp27 is null"
               intI = 1: strExc(10) = ""
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strExc(10) = "" & RsTemp.Fields("cp06")
                  If Val(strExc(10)) > 0 Then strExc(10) = Val(strExc(10)) - 19110000
               End If
               Text1(0) = Left(Trim(cboCP14.Text), 5)
               Label3(0) = Trim(Mid(Trim(cboCP14.Text), 6))
               Call PUB_GetFCPsetCP48(Me.Visible, pa, m_CP27, Text1(1), Text1(0), m_CP122, Text1(4), Text1(5), Text1(23), Text1(28), Combo2, Text1(12), strExc(10))
               '會稿要預帶承辦人
               If Text1(0).Text <> Left(Trim(cboCP14.Text), 5) And Trim(cboCP14.Text) = "" Then
                  cboCP14.Text = Text1(0)
                  strExc(0) = GetPrjSalesNM(Left(cboCP14.Text, 5))
                  If strExc(0) <> "" Then
                     cboCP14.Text = Left(cboCP14.Text, 5) & " " & strExc(0)
                  End If
               End If
            End If
            '2021/9/6 END
         End If
         
      Case 7
         If Text1(1) = 異議_專 Then
            If Text1(Index) = "" Then
               MsgBox "案件性質為異議時公告日不可空白 !", vbCritical
               Cancel = True
            Else
               If ChkDate(Text1(Index)) Then
                  Text1(5) = TransDate(CompDate(2, -1, CompDate(1, 3, TransDate(Text1(Index).Text, 2))), 1)
                  
                  'Modified by Morgan 2014/11/20 外專改回舊規則
                  ''Added by Morgan 2014/10/29
                  'If pa(9) = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                  '   Text1(4) = TransDate(PUB_GetOurDeadline(Text1(5)), 1)
                  'Else
                  ''end 2014/10/29
                  'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
                  If strSrvDate(1) >= 外專台灣案所限新規則啟用日 Then
                     'Modify By Sindy 2021/4/27 + m_pAgreeOnDate
                     Text1(4) = TransDate(PUB_GetFCPOurDeadline(Text1(5), 2, , m_pAgreeOnDate), 1)
                  Else
                  'end 2019/7/11
                     Text1(4) = TransDate(CompDate(2, -2, TransDate(Text1(5).Text, 2)), 1)
                  End If 'Added by Morgan 2019/7/11
                  
                  Text1(4).Tag = Text1(4).Text 'Add By Sindy 2015/12/16
                  'End If 'Added by Morgan 2014/10/29
                  'end 2014/11/20
               End If
            End If
         Else
            If Text1(Index) <> "" Then
               MsgBox "案件性質為異議時才可輸入 !", vbCritical
               Text1(Index) = ""
            End If
         End If
      'Added by Lydia 2017/05/05 客戶案件案號長度控制
      Case 8
          'Modified by Lydia 2017/06/14 改常數
          'Cancel = Not CheckLengthIsOK(text1(Index), 100)
          Cancel = Not CheckLengthIsOK(Text1(Index), 專利客戶案號max)
      'end 2017/05/05
      Case 10
      Case 12
         If Not ChkDate(Text1(Index)) Then Cancel = True
      Case 13, 14, 15, 16, 17 '申請人
         Label5(Index - 11) = ""
         If Text1(Index) <> "" Then If ChgType(Index) = False Then Cancel = True
         'Add By Cheng 2002/08/22
         If Cancel = False Then
            Select Case Index
            Case 13
               If Me.Text1(Index).Text <> m_strCust1 Then
                  If Not PUB_EditCustOk(Me.Label3(8).Caption, pa(1), pa(2), pa(3), pa(4)) Then Cancel = True
               End If
            Case 14
               If Me.Text1(Index).Text <> m_strCust2 Then
                  If Not PUB_EditCustOk(Me.Label3(8).Caption, pa(1), pa(2), pa(3), pa(4)) Then Cancel = True
               End If
            Case 15
               If Me.Text1(Index).Text <> m_strCust3 Then
                  If Not PUB_EditCustOk(Me.Label3(8).Caption, pa(1), pa(2), pa(3), pa(4)) Then Cancel = True
               End If
            Case 16
               If Me.Text1(Index).Text <> m_strCust4 Then
                  If Not PUB_EditCustOk(Me.Label3(8).Caption, pa(1), pa(2), pa(3), pa(4)) Then Cancel = True
               End If
            Case 17
               If Me.Text1(Index).Text <> m_strCust5 Then
                  If Not PUB_EditCustOk(Me.Label3(8).Caption, pa(1), pa(2), pa(3), pa(4)) Then Cancel = True
               End If
            End Select
         End If
      Case 18
         If Text1(Index) <> "" Then If ChgType(Index) = False Then Cancel = True
      'Add By Cheng 2002/06/11
      Case 21 '專利種類
         If Me.Text1(Index).Enabled Then
            Me.Label3(11).Caption = "" & PUB_GetPatentKindName(Me.Text1(21).Text, 台灣國家代號)
            If Me.Label3(11).Caption = "" Then
               MsgBox "專利種類輸入錯誤!!!", vbExclamation + vbOKOnly
               Cancel = True
               Me.Text1(Index).SetFocus
            Else
               If (Me.Text1(1).Text >= "101" And Me.Text1(1).Text <= "103") Or _
                  (Me.Text1(1).Text >= "301" And Me.Text1(1).Text <= "303") Then
                  If Mid(Me.Text1(1).Text, 3, 1) <> Me.Text1(Index).Text Then
                     MsgBox "專利種類必須與案件性質的第三碼相同!!!", vbExclamation + vbOKOnly
                     Cancel = True
                     Me.Text1(Index).SetFocus
                  End If
               End If
               
               'Added by Morgan 2012/12/27
               If Me.Text1(1).Text = "125" Or Me.Text1(1).Text = "308" Then
                  If Me.Text1(Index).Text <> "3" Then
                     MsgBox Label3(1) & "專利種類必須為 3 設計!!!", vbExclamation + vbOKOnly
                     Cancel = True
                     Me.Text1(Index).SetFocus
                  End If
               End If
               'end 2012/12/27
               
            End If
         End If
        
      Case 22 '核稿人
            Me.Label5(12).Caption = GetStaffName(Me.Text1(Index).Text)
            If Me.Text1(Index).Text <> "" And Me.Label5(12).Caption = "" Then
                MsgBox "核稿人輸入錯誤!!!", vbExclamation + vbOKOnly
                Cancel = True
                Me.Text1(Index).SetFocus
                Text1_GotFocus Index
            End If
      
      'Add by Morgan 2008/8/19
      Case 23 '承辦期限
         If Text1(Index) <> "" And Text1(Index).Locked = False Then
            If Not ChkDate(Text1(Index)) Then
               Cancel = True
               Text1(Index).SetFocus
               Text1_GotFocus Index
            Else
               'Modified by Morgan 2013/1/7 若承辦期限沒有改時不必抓系統日(當主動修正承辦期限為新案翻譯的本所期限時有可能是非工作日 Ex.FCP-046742)
               'text1(Index) = TransDate(PUB_GetWorkDay1(text1(Index), False), 1)
               If Text1(23) <> m_203CP48 Then
                  Text1(Index) = TransDate(PUB_GetWorkDay1(Text1(Index), False), 1)
               End If
            End If
         End If
         
      Case 27
         If Trim(Text1(2)) & Trim(Text1(25)) & Trim(Text1(26)) & Trim(Text1(27)) <> "" Then
            If (pa(1) = "FCP" And Left(Trim(Text1(2)) & Trim(Text1(25)) & Trim(Text1(26)) & Trim(Text1(27)), 3) <> "FCP") Or _
               (pa(1) = "FG" And Left(Trim(Text1(2)) & Trim(Text1(25)) & Trim(Text1(26)) & Trim(Text1(27)), 2) <> "FG") Then
               MsgBox "轉本所案號必須與原本所案號之系統別相同 !", vbCritical
               Cancel = True
            Else
               If GetGrid(Trim(Text1(2)) & Trim(Text1(25)) & Trim(Text1(26)) & Trim(Text1(27)), 1) = False Then Cancel = True
               Text1(6) = ""
            End If
         End If
         
      'Add by Morgan 2008/9/4
      Case 28 '文件齊備日
         If Text1(Index) <> "" And Text1(Index).Locked = False Then
            If Not ChkDate(Text1(Index)) Then
               Cancel = True
               Text1(Index).SetFocus
               Text1_GotFocus Index
            Else
               'Modified by Lydia 2016/06/21 改成模組
               'SetCP48
               Text1(0) = Left(Trim(cboCP14.Text), 5)
               Label3(0) = Trim(Mid(Trim(cboCP14.Text), 6))
               Call PUB_GetFCPsetCP48(Me.Visible, pa, m_CP27, Text1(1), Text1(0), m_CP122, Text1(4), Text1(5), Text1(23), Text1(28), Combo2, Text1(12))
               'Modify By Sindy 2021/9/6 會稿要預帶承辦人
               If Text1(0).Text <> Left(Trim(cboCP14.Text), 5) And Trim(cboCP14.Text) = "" Then
                  cboCP14.Text = Text1(0)
                  strExc(0) = GetPrjSalesNM(Left(cboCP14.Text, 5))
                  If strExc(0) <> "" Then
                     cboCP14.Text = Left(cboCP14.Text, 5) & " " & strExc(0)
                  End If
               End If
               '2021/9/6 END
            End If
         End If
      
      'Add by Sindy 2015/12/17
      Case 30 '指定日期
         If Text1(Index) <> "" And Text1(Index).Locked = False Then
            If Not ChkDate(Text1(Index)) Then
               Cancel = True
               Text1(Index).SetFocus
               Text1_GotFocus Index
            End If
            If ChkWorkDay(DBDATE(Text1(Index))) = False Then
               MsgBox "請輸入工作天！", vbInformation, "輸入指定日期錯誤"
               Cancel = True
               Text1(Index).SetFocus
               Text1_GotFocus Index
            End If
            If Val(Text1(Index)) < Val(strSrvDate(2)) Then
               MsgBox "指定日期不可小於系統日！", vbInformation, "輸入指定日期錯誤"
               Cancel = True
               Text1(Index).SetFocus
               Text1_GotFocus Index
            End If
            'Add By Sindy 2021/8/30
            '指定日期不可大於法定期限
            If Text1(Index) <> "" And Text1(5) <> "" And Val(Text1(Index)) > Val(Text1(5)) Then
               MsgBox "指定日期不可大於法定期限！", vbInformation, "輸入指定日期錯誤"
               Cancel = True
               Text1(Index).SetFocus
               Text1_GotFocus Index
            End If
            '2021/8/30 END
         End If
      '2015/12/17 END
   End Select
   
gotoExit:
   If Cancel Then TextInverse Text1(Index)
   'add by nickc 2007/07/13 將輸入法改成使用API
   If Cancel = False Then CloseIme
End Sub

Private Sub GridHead()
 Dim i As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 195: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 900: .Text = "下一程序"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 800: .Text = "本所期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 800: .Text = "法定期限"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1185: .Text = "機關文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 675: .Text = "相關人"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 1200: .Text = "解除期限日期"
      .col = 7: .ColWidth(7) = 0
      .col = 8: .ColWidth(8) = 0
      .col = 9: .ColWidth(9) = 0
      .col = 10: .ColWidth(10) = 2000: .Text = "備註"
      .CellAlignment = flexAlignCenterCenter
      '判斷是否有資料
      .Visible = True
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Modified by Lydia 2024/06/06 +預設帶出收文性質=True
   PUB_SendMailCache , , , True 'Add by Morgan 2010/6/18
   
   'Added by Lydia 2018/05/21 回到前畫面
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" And TypeName(m_PrevForm) <> "frm060101" Then
        m_PrevForm.Show
        If TypeName(m_PrevForm) = "frm060122" Then
             m_PrevForm.cmdState = 0
             Call m_PrevForm.PubShowNextData
        End If
   End If
   'end 2018/05/21
   
   Set frm060101_1 = Nothing
End Sub

'Add By Cheng 2002/05/09
'若案件性質為"601"或"605"且有輸入承辦人時
Private Function CheckPromAndCK() As Boolean
Dim Rs As New ADODB.Recordset
CheckPromAndCK = False
If (Me.Text1(1).Text = "601" Or Me.Text1(1).Text = "605") And Len(Trim(cboCP14.Text)) > 0 Then
   If Rs.State <> adStateClosed Then Rs.Close
   Set Rs = Nothing
   Rs.CursorLocation = adUseClient
   Rs.Open " Select ST15 From Staff Where ST01='" & Left(Trim(cboCP14.Text), 5) & "'", cnnConnection, adOpenStatic, adLockReadOnly
   If Rs.RecordCount > 0 Then
      If Rs.Fields(0).Value <> "F22" Then
         MsgBox "承辦人必須為F22部門人員!!!", vbExclamation + vbOKOnly, "輸入錯誤"
         Me.cboCP14.SetFocus
         Exit Function
      End If
   Else
      MsgBox "承辦人輸入錯誤!!!", vbExclamation + vbOKOnly
      Me.cboCP14.SetFocus
      Exit Function
   End If
   If Rs.State <> adStateClosed Then Rs.Close
   Set Rs = Nothing
End If
CheckPromAndCK = True
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
   Dim objTxt As Object
   Dim ii As Integer
   Dim Cancel As Boolean
   Dim arrInv
   Dim strMsg As String 'Add by Amy 2013/08/20
   
'   'Add By Sindy 2016/7/27
'   If Trim(cboCP14.Text) = "" Then
'      MsgBox "承辦人不可空白！"
'      cboCP14.SetFocus
'      Call cboCP14_GotFocus
'      Exit Function
'   End If
'   '2016/7/27 END
   
   TxtValidate = False
   
   For Each objTxt In Text1
      If objTxt.Enabled = True Then
         Cancel = False
         Text1_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   'Add By Sindy 2016/7/27
   'Modified by Lydia 2018/09/12 不改日期
   'Cancel = False
   Cancel = True
   CboCP14_Validate Cancel
   If Cancel = True Then
      'Added by Lydia 2020/01/20
      If Trim(cboCP14.Text) = "" Then
         'Modify By Sindy 2021/10/20 亭妙:分案作業中改案件性質的時候，請不要鎖定掛承辦人，可以詢問: 是否確定沒有承辦人。
         '（有時候承辦指示要改案件性質，但其實此案還不能掛承辦人，例如新案翻譯）
         If m_strCP10 <> Text1(1).Text Then
            If MsgBox("是否確定沒有承辦人？", vbQuestion + vbYesNo) = vbNo Then
               cboCP14.SetFocus
               Exit Function
            End If
         Else
         '2021/10/20 END
            MsgBox "請輸入承辦人!", vbCritical
            cboCP14.SetFocus
            Exit Function
         End If
      End If
      'end 2020/01/20
   End If
   '2016/7/27 END

   If Text1(18) = "" And Text1(13) = "" And Text1(14) = "" And Text1(15) = "" And Text1(16) = "" And Text1(17) = "" Then
      MsgBox "申請人及代理人至少輸入一個 !", vbCritical
      Text1(13).SetFocus
      Exit Function
   End If
   
   'Added by Lydia 2024/06/14 對申請人1~5的重複輸入檢查
   If Pub_ChkAppList(strExc(0), Text1(13) & "," & Text1(14) & "," & Text1(15) & "," & Text1(16) & "," & Text1(17)) = False Then
      SSTab1.Tab = 0
      Text1(Val(strExc(0)) + 12).SetFocus
      Text1_GotFocus Val(strExc(0)) + 12
      Exit Function
   End If
   'end 2024/06/14
   
   'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
   For ii = 13 To 18
      strExc(1) = ChangeCustomerL(Text1(ii))
      If ii < 18 Then
         'Added by Lydia 2024/06/19 區分FG案
         If pa(1) = "FCP" Or pa(1) = "CFP" Or pa(1) = "P" Then
            strExc(2) = ChangeCustomerL(pa(ii + 13))
         Else
            If ii = 13 Then strExc(2) = ChangeCustomerL(pa(8))
            If ii = 14 Then strExc(2) = ChangeCustomerL(pa(58))
            If ii = 15 Then strExc(2) = ChangeCustomerL(pa(59))
            If ii = 16 Then strExc(2) = ChangeCustomerL(pa(65))
            If ii = 17 Then strExc(2) = ChangeCustomerL(pa(66))
         End If
         If strExc(1) <> "" And strExc(1) <> strExc(2) Then
            If GetCustomerAndState(strExc(1), strExc(3), , , , pa(1), strExc(8), False, Me.Name, pa(2), pa(3), pa(4)) = False Then
               SSTab1.Tab = 0
               Text1(ii).SetFocus
               Text1_GotFocus ii
               Exit Function
            End If
         End If
      Else
         'Added by Lydia 2024/06/19 區分FG案
         If pa(1) = "FCP" Or pa(1) = "CFP" Or pa(1) = "P" Then
            strExc(2) = ChangeCustomerL(pa(75))
         Else
            strExc(2) = ChangeCustomerL(pa(26))
         End If
         If strExc(1) <> "" And strExc(1) <> strExc(2) Then
            If GetAgentAndState(strExc(1), strExc(3), , , , pa(1), strExc(8), False) = False Then
               SSTab1.Tab = 0
               Text1(ii).SetFocus
               Text1_GotFocus ii
               Exit Function
            End If
         End If
      End If
   Next ii
   'end 2024/06/13
   
   'Modify By Sindy 2021/8/16 Mark
'   'Add By Cheng 2003/08/13
'   '若案件性質為延期, 則不可點選本案期限
'   If Me.Text1(1).Text = "404" Then
'       For ii = 1 To Me.MSHFlexGrid1.Rows - 1
'           If Me.MSHFlexGrid1.TextMatrix(ii, 0) <> "" Then
'               MsgBox "此案僅收文<延期>，不可點選下一程序期限資料，" & vbCrLf & "否則無法管制下一程序的期限!!!", vbExclamation + vbOKOnly
'               Exit Function
'           End If
'       Next ii
'   End If
   
   'Add by Morgan 2008/1/7
   'Modified by Morgan 2013/6/19 +927其他翻譯
   'modify by sonia 2013/12/29 + 236翻譯－主管機關來函
   'Remove by Lydia 2016/06/21 改成模組PUB_CheckFCPtxtValidate
   'If Left(text1(0).Text, 1) = "F" And text1(1).Text <> "201" And text1(1).Text <> "927" And text1(1).Text <> "236" Then
   '   MsgBox "只有當案件性質為<翻譯>時承辦人才可輸入外譯編號！"
   '   text1(0).SetFocus
   '   Text1_GotFocus 0
   '   Exit Function
   'End If
   Cancel = False
   Text1(0) = Left(Trim(cboCP14.Text), 5)
   Label3(0) = Trim(Mid(Trim(cboCP14.Text), 6))
   If PUB_CheckFCPtxtValidate(pa, Text1(0), Text1(1), Text1(22), Text1(5), Cancel, m_203CP09, m_CP27) = False Then
      If Cancel Then
         cboCP14.SetFocus
      End If
      Exit Function
   End If
   
    'Add by Morgan 2004/3/5
    '設計的翻譯，檢視中說，製作中說不掛核稿人
    'Modified by Morgan 2013/11/6 +235核對中說格式
    'Remove by Lydia 2016/06/21 改成模組PUB_CheckFCPtxtValidate
    'If text1(21) = "3" And (text1(1) = "201" Or text1(1) = "209" Or text1(1) = "235" Or text1(1) = "210") Then
    '    If text1(22) <> "" Then
    '        If MsgBox("設計的翻譯、檢視中說、核對中說格式、製作中說不掛核稿人，是否清空核稿人？", vbQuestion + vbYesNo) = vbYes Then
    '            text1(22) = ""
    '        Else
    '            Exit Function
    '        End If
    '    End If
    'End If
    
   'Added by Morgan 2015/5/19 案件性質改為非新案翻譯時核稿人檢查 Ex.FCP-51295 --靜芳
   'Remove by Lydia 2016/06/21 改成模組PUB_CheckFCPtxtValidate
   'If text1(22) <> "" And (text1(1) = "209" Or text1(1) = "235" Or text1(1) = "210") Then
   '   If MsgBox("檢視中說、核對中說格式、製作中說不掛核稿人，是否清空核稿人？", vbQuestion + vbYesNo) = vbYes Then
   '      text1(22) = ""
   '   Else
   '      Exit Function
   '   End If
   'End If
   'end 2015/5/19
   
   'Add by Morgan 2006/5/1 申請寄存及存活證明控制
   If Text1(1) = "108" Then
      m_bol108 = True
   Else
      m_bol108 = False
   End If
   '實審,發明分割要檢查是否有收文申請寄存
   If (Text1(1).Text = "416" Or (pa(8) = "1" And Text1(1).Text = "307")) Then
      If txtDivCaseNo(2) = "" Then
         m_bol108 = PUB_ChkCPExist(pa, "108", 0)
      Else
         strExc(1) = txtDivCaseNo(1)
         strExc(2) = txtDivCaseNo(2)
         strExc(3) = txtDivCaseNo(3)
         strExc(4) = txtDivCaseNo(4)
         m_bol108 = PUB_ChkCPExist(strExc, "108", 0)
      End If
      '有收申請寄存108
      If m_bol108 = True Then
         If PUB_ChkCPExist(pa, "221") = False Then
            If txtDivCaseNo(2) = "" Then
               MsgBox "本案有收文【申請寄存】但尚未收文【存活證明】！", vbOKOnly
            Else
               MsgBox "本案為分割案，母案有收文【申請寄存】但本案尚未收文【存活證明】！", vbOKOnly
            End If
         End If
      End If
   End If
   '2006/5/1 end
   
   'Add by Morgan 2010/4/29 改加欄位顯示
   'Modified by Morgan 2015/4/30 修正 Text1(27) -> Text1(29)
   If Text1(1).Text = "908" And Combo3.Enabled = True And Text1(29) = "N" And m_CP27 = "" Then
      MsgBox "退審查費將自動設定為要請款"
      Text1(29) = ""
   End If
   
   'Modify By Sindy 2014/11/12
   BolIsInventorUpd = False 'Add By Sindy 2014/12/23 檢查是否要更新發明人資料
   If strSrvDate(1) >= 專利發明人檔啟用日 Then
      '申請人有變更且未重新點選發明人資料時清除原發明人資料
      '申請案才要
      If InStr(NewCasePtyList, Text1(1)) > 0 Then
         '串列申請人資料
         strExc(1) = Text1(13)
         For intI = 1 To 4
            If Text1(intI + 13) <> "" Then
               strExc(1) = strExc(1) & "," & Text1(3)
            End If
         Next
         strExc(2) = "" 'Added by Morgan 2014/12/19
         '串列發明人資料
         strSql = "SELECT pi06 from PatentInventor where pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
                  " order by pi05 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            RsTemp.MoveFirst
            strExc(2) = RsTemp.Fields(0)
            RsTemp.MoveNext
            Do While Not RsTemp.EOF
               strExc(2) = strExc(2) & "," & RsTemp.Fields(0)
               RsTemp.MoveNext
            Loop
         End If
         '檢查是否要更新
         If PUB_ChkInventor(strExc(2), strExc(1), True) = False Then
            BolIsInventorUpd = True 'Add By Sindy 2014/12/23 要更新發明人資料
         End If
      End If
   Else
   '2014/11/12 END
      'Add by Morgan 2010/5/3 申請人有變更且未重新點選發明人資料時清除原發明人資料
      'Modify by Morgan 2010/6/24 申請案才要
      If InStr(NewCasePtyList, Text1(1)) > 0 Then
         strExc(1) = Text1(13)
         For intI = 1 To 4
            If Text1(intI + 13) <> "" Then
               strExc(1) = strExc(1) & "," & Text1(3)
            End If
         Next
         
         strExc(2) = pa(60)
         For intI = 1 To 9
            If pa(intI + 60) <> "" Then
               strExc(2) = strExc(2) & "," & pa(intI + 60)
            End If
         Next
         If PUB_ChkInventor(strExc(2), strExc(1), True) = False Then
            arrInv = Split(strExc(2), ",")
            For intI = 0 To 9
               If intI <= UBound(arrInv) Then
                  pa(60 + intI) = arrInv(intI)
               Else
                  pa(60 + intI) = ""
               End If
            Next
         End If
      End If
      'end 2010/5/3
   End If
   
   'Add by Morgan 2010/8/12
   '改請案必須輸入相關收文號(發文時要取消原申請程序的催審)
   If Text1(1) >= "301" And Text1(1) <= "306" And Text1(6) = "" Then
      MsgBox "必須輸入相關收文號！", vbExclamation
      Text1(6).SetFocus
      Exit Function
   End If
   'end 2010/8/12
   
   'Added by Lydia 2024/09/02 422加速審查一定要掛相關總收文號，以利後續抓資料能正確處理
   'Modified by Morgan 2024/11/18 +447再審查加速審查
   If pa(1) = "FCP" And (Text1(1) = "422" Or Text1(1) = "447") And Text1(6) = "" Then
      MsgBox "必須輸入相關收文號！", vbExclamation
      Text1_GotFocus 6
      Text1(6).SetFocus
      Exit Function
   End If
   'end 2024/09/02
   
   'Added by Morgan 2012/12/21
   If pa(1) = "FCP" And (Text1(1) = "125" Or Text1(1) = "308") Then
      If txtDesignCaseNo(1) = "" Or txtDesignCaseNo(2) = "" Then
         MsgBox "請輸入衍生設計母案本所案號！"
         txtDesignCaseNo(1).SetFocus
         Exit Function
      Else
         If txtDesignCaseNo(3) = "" Then txtDesignCaseNo(3) = "0"
         If txtDesignCaseNo(4) = "" Then txtDesignCaseNo(4) = "00"
         'Modified by Morgan 2020/2/26 +pa10,pa26,pa27,pa28,pa29,pa30
         strExc(0) = "select sqldatet(pa14),pa11,pa08,pa10,pa14,pa26,pa27,pa28,pa29,pa30 from patent where pa01='" & txtDesignCaseNo(1) & "' and pa02='" & txtDesignCaseNo(2) & "' and pa03='" & txtDesignCaseNo(3) & "' and pa04='" & txtDesignCaseNo(4) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            'Modified by Morgan 2022/9/2 此案的申請日早於母案的公告日，則排除此限制--何淑華
            'If Not IsNull(RsTemp(0)) Then
            If Not IsNull(RsTemp(0)) And Not (Val(pa(10)) > 0 And DBDATE(pa(10)) < RsTemp("pa14")) Then
            'end 2022/9/2
               MsgBox "本衍生設計案之母案已於 " & RsTemp(0) & " 公告，不可分案！", vbCritical
               txtDesignCaseNo(1).SetFocus
               Exit Function
            ElseIf RsTemp("pa08") <> "3" Or Len(RsTemp("pa11")) > 9 Then
               MsgBox "本衍生設計案之母案必須為設計申請案！", vbCritical
               txtDesignCaseNo(1).SetFocus
               Exit Function
            'Added by Morgan 2020/2/26
            ElseIf RsTemp("pa10") > DBDATE(pa(10)) And Val(pa(10)) > 0 Then
               MsgBox "【原申請案】的申請日(" & pa(10) & ")不得早於【衍生設計母案】之申請日(" & TransDate(RsTemp("pa10"), 1) & ")，請確認！", vbCritical
               txtDesignCaseNo(1).SetFocus
               Exit Function
            ElseIf "" & RsTemp("pa26") <> ChangeCustomerL(pa(26)) Or "" & RsTemp("pa27") <> ChangeCustomerL(pa(27)) Or "" & RsTemp("pa28") <> ChangeCustomerL(pa(28)) Or "" & RsTemp("pa29") <> ChangeCustomerL(pa(29)) Or "" & RsTemp("pa30") <> ChangeCustomerL(pa(30)) Then
               MsgBox "【原申請案】的申請人與【衍生設計母案】之申請人不同，請確認！", vbCritical
               txtDesignCaseNo(1).SetFocus
               Exit Function
            'end 2020/2/26
            End If
         Else
            MsgBox "衍生設計母案不存在，請重新輸入！"
            txtDesignCaseNo(1).SetFocus
            Exit Function
         End If
         If txtDesignCaseNo(1).Tag & txtDesignCaseNo(2).Tag & txtDesignCaseNo(3).Tag & txtDesignCaseNo(4).Tag <> "" Then
            If txtDesignCaseNo(1).Tag & txtDesignCaseNo(2).Tag & txtDesignCaseNo(3).Tag & txtDesignCaseNo(4).Tag <> txtDesignCaseNo(1) & txtDesignCaseNo(2) & txtDesignCaseNo(3) & txtDesignCaseNo(4) Then
               MsgBox "衍生設計母案已變更，請自行維護與原母案之相關案資料!!!"
            End If
         End If
      End If
   End If
   'end 2012/12/21
   
   'Add by Amy 2013/08/20 退費的相關收文號，需選最後一道有發文規費的案件進度
   If Text1(1) = "908" Then
       If Text1(6) <> GetLastCP09(strMsg) Then
            MsgBox strMsg
            Text1(6).SetFocus
            Exit Function
       End If
   End If
   'end 2013/08/20
   
   'Added by Lydia 2017/05/17 檢查原文字數和相似度
   If txtTF19.Visible = True And Trim(Replace(txtTF23 & txtTF19, " ", "")) <> "" Then
      If Trim(txtTF23) = "" Then
         MsgBox "請輸入原文字數！", vbCritical
         txtTF23.SetFocus
         Exit Function
      End If
      'Modified by Lydia 2017/12/29 淑華要求107/1/1先開放單獨輸入原文字數
      'If Trim(txtTF19) = "" Then
      '   MsgBox "請輸入相似度！", vbCritical
      '   txtTF19.SetFocus
      '   Exit Function
      'ElseIf Val(txtTF19) > 100 Then
      If Val(txtTF19) > 100 Then
      'end 2017/12/29
         MsgBox "相似度不可大於100！"
         txtTF19.SetFocus
         Exit Function
      End If
   End If
   'end 2017/05/17
   
   'add by sonia 2019/5/17 改案件性質時(P108898申復改為改請)要提醒
   If Me.Text1(21).Enabled = False And _
      ((Me.Text1(1).Text >= "101" And Me.Text1(1).Text <= "103") Or _
      (Me.Text1(1).Text >= "301" And Me.Text1(1).Text <= "303")) Then
      If Mid(Me.Text1(1).Text, 3, 1) <> Me.Text1(21).Text Then
         MsgBox "專利種類必須與案件性質的第三碼相同, 請存檔後再分案一次修改專利種類!!!", vbExclamation + vbOKOnly
      End If
   End If
   'end 2019/5/17
   
   'add by sonia 2021/2/19 FCP-058462
   If Val(Me.Text1(5).Text) > 0 And Val(Me.Text1(4).Text) = 0 Then
      MsgBox "有法定期限，一定要有本所期限！", vbCritical
      Text1(4).SetFocus
      Exit Function
   End If
   'end 2021/2/19
   
   'Added by Morgan 2024/5/21 --Sharon
   '訴願案設定(排除日本部)
   '"智慧局答辯函"的相關收文號為"501訴願"， 設定本所期限為5個工作天，承辦期限往前-2天，在接洽單上備註"請務必在5個工作天內送件"
   '"206 補充說明"的相關收文號為"501訴願"， 設定本所期限為3個工作天，承辦期限往前-2天，於分案時，自動發一封email給承辦人，主旨: "此為訴願案補充說明，請務必在3個工作天內送件"
   '期限計算都不含當日
   m_strSubject = ""
   If m_CP27 = "" And pa(150) <> "3" And Text1(1) = "206" Then
      strExc(0) = "select * from caseprogress where cp09='" & Text1(6) & "' and cp10='501'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_strSubject = "此為訴願案補充說明，請務必在3個工作天內送件!!"
         strExc(1) = CompDate(2, 1, Text1(12))  '先抓次日以免是非工作日上班
         strExc(2) = TransDate(CompWorkDay(3, strExc(1)), 1) '本所期限=收文日+3個工作天
         strExc(3) = TransDate(CompWorkDay(1, strExc(1)), 1) '承辦期限=本所期限-2個工作天=收文日+1個工作天
         If Text1(4) <> strExc(2) Or Text1(23) <> strExc(3) Then
            Text1(4) = strExc(2)
            Text1(23) = strExc(3)
            MsgBox "【206 補充說明】的相關收文號為【501訴願】，本所期限已設定為收文日+3個工作天(不含收文日)，承辦期限已設定為本所期限-2個工作天！", vbExclamation
         End If
      End If
   End If
   'end 20214/5/20
   
   'Add by Sindy 2021/5/11 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/5/11 END
   
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
If rsA.RecordCount > 0 Then
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

'Remove by Morgan 2007/10/11 不再使用
''Add By Cheng 2002/06/10
''取得案件收費表的工作天數
'Private Function GetCF04(strCF01 As String, strCF02 As String, strCF03 As String) As String
'Dim rsA As New ADODB.Recordset
'Dim StrSQLa As String
'
'GetCF04 = "0"
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'StrSQLa = "Select CF04 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "' AND CF04 IS NOT NULL"
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'   GetCF04 = rsA.Fields(0).Value
'End If
'If rsA.State <> adStateClosed Then rsA.Close
'Set rsA = Nothing
'
'End Function

'Add By Cheng 2003/08/14
'判斷員工部門別是否為外專外翻
Private Function ChkStaffDepIsF51(StrST01 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

ChkStaffDepIsF51 = False
StrSQLa = "Select ST15 From Staff Where ST01='" & StrST01 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    If rsA.Fields(0).Value = "F51" Then ChkStaffDepIsF51 = True
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Sub textCP64_GotFocus()
   TextInverse textCP64
End Sub
Private Sub textPA91_GotFocus()
   TextInverse textPA91
End Sub

Private Sub txtDesignCaseNo_GotFocus(Index As Integer)
   TextInverse txtDesignCaseNo(Index)
   CloseIme
End Sub

Private Sub txtDesignCaseNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDivCaseNo_GotFocus(Index As Integer)
    TextInverse txtDivCaseNo(Index)
End Sub

Private Sub txtDivCaseNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtDivCaseNo_Validate(Index As Integer, Cancel As Boolean)
Dim strCP14 As String

   If Index = 4 Then
      If (txtDivCaseNo(1) <> "" Or txtDivCaseNo(2) <> "" Or txtDivCaseNo(3) <> "" Or txtDivCaseNo(4) <> "") Then
         'Modified by Morgan 2012/11/8 改呼叫公用函數檢查
         'Call CheckDivCase(m_stPA09)

         If PUB_CheckDivCase(txtDivCaseNo, pa) = True Then
            'Add By Sindy 2021/9/3
            strCP14 = Left(Trim(cboCP14.Text), 5)
            If PUB_GetFCPCP14_F21(txtDivCaseNo, strCP14) = True Then '抓承辦人為工程師
               cboCP14 = strCP14
               CboCP14_Validate True
            End If
            '2021/9/3 END
         End If
      End If
   End If
End Sub
'Remove by Lydia 2016/06/21 改成模組 PUB_FCPCheckCP14
''Add by Morgan 2004/3/23
''承辦人欄位控制
'Private Function CheckCP14() As Boolean
'Dim stDept As String
'
'
'   'Add by Morgan 2004/10/13 209,210 承辦只為工程師 F21
'   'Modified by Morgan 2013/11/6 +235核對中說格式
'   If text1(1).Text = "209" Or text1(1).Text = "235" Or text1(1).Text = "210" Then
'      If text1(0).Text <> "" And text1(0).Text <> text1(0).Tag Then
'         stDept = GetStaffDepartment(text1(0))
'         If stDept <> "F21" And stDept <> "F81" Then    '2008/4/8 MODIFY BY SONIA 加 F81
'            MsgBox "該案件性質的承辦人只可為工程師!!!"
'            Exit Function
'         End If
'      End If
'   End If
'   '2004/10/13 end
'   'Modified by Morgan 2013/11/6 +235核對中說格式
'   If text1(21) <> "3" And (Me.text1(1).Text = "201" Or Me.text1(1).Text = "209" Or Me.text1(1).Text = "235" Or Me.text1(1).Text = "210") Then
'       If text1(0).Text <> text1(0).Tag Or text1(22).Text <> text1(22).Tag Then
'           stDept = GetStaffDepartment(text1(0))
'           Select Case stDept
'               Case "F22"
'                   MsgBox "該案件性質的承辦人不可為程序!!!"
'                   Exit Function
'
'               'Modify by Morgan 2005/8/1 加F52 --靜芳
'               'Modify by Morgan 2007/6/11 F52另外檢查
'               Case "F51"
'                   text1(22).Text = ""
'                   Label5(12).Caption = ""
'
''Remove by Morgan 2007/8/1 取消
''               'Add by Morgan 2007/6/11
''               '若為內翻時核稿人帶所內員工編號
''               Case "F52"
''                  Text1(22).Text = PUB_GetMapID(Text1(0))
''                  If Text1(22).Text <> "" Then
''                     Label5(12).Caption = GetStaffName(Me.Text1(0).Text)
''                  Else
''                     Label5(12).Caption = ""
''                  End If
''end 2007/8/1
'
'               Case Else
'
'                  'Add by Morgan 2008/9/18 加判斷承辦人為國外部工程師的才要預設核稿人
'                  'Modify by Morgan 2010/6/14 所內員工不用考慮對照(新進同仁無外譯編號)
'                  'strExc(0) = "select 1 from staff_idmap,staff where '" & Text1(0) & "' in (sim01,sim02)" & _
'                     " and st01(+)=sim01 and ST15='F21' and st04='1'"
'
'                  If text1(1).Text = "201" Then 'Added by Morgan 2015/3/11 只有 201 要預設核稿人 --靜芳
'
'                     strExc(0) = "select st01 from staff where st01='" & text1(0) & "' and ST15='F21' and st04='1'" & _
'                        " union select st01 from staff_idmap,staff where sim02='" & text1(0) & "'" & _
'                        " and st01(+)=sim01 and ST15='F21' and st04='1'"
'
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        text1(22).Text = text1(0).Text
'                        Label5(12).Caption = GetStaffName(text1(0).Text)
'                     Else
'                        text1(22).Text = ""
'                        Label5(12).Caption = ""
'                     End If
'
'                  End If 'Added by Morgan 2015/3/11 只有 201 要預設核稿人
'               End Select
'       End If
'   Else
'       Me.text1(22).Text = ""
'       Me.Label5(12).Caption = ""
'       '2009/1/7 add by sonia 年費依有無年費代理人檢查承辦人且須為該國管制人
'       If text1(1).Text = "605" Then
'          Dim m_Fagent As String
'          'Modify by Morgan 2011/6/3 改比照期限管制規則(同信函收件人)
'          'm_Fagent = PUB_GetA1K28("" & ChgSQL(pa(1)), "" & ChgSQL(pa(2)), "" & ChgSQL(pa(3)), "" & ChgSQL(pa(4)), "605")
'          m_Fagent = PUB_GetReceiver("" & ChgSQL(pa(1)), "" & ChgSQL(pa(2)), "" & ChgSQL(pa(3)), "" & ChgSQL(pa(4)), "605", "1")
'          strExc(0) = "select na16 from fagent,nation where fa01='" & ChgSQL(Mid(m_Fagent, 1, 8)) & "' and fa02='" & ChgSQL(Mid(m_Fagent, 9, 1)) & "' and fa10=na01(+) "
'          intI = 1
'          Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'          If intI = 1 Then
'            If text1(0).Text <> "" And text1(0).Text <> "" & RsTemp.Fields(0) Then
'               If MsgBox("年費的承辦人錯誤, 應為 " & RsTemp.Fields(0) & GetStaffName(RsTemp.Fields(0)) & "，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
'                  Exit Function
'               End If
'            End If
'         End If
'       End If
'       '2009/1/7 end
'   End If
'
'   'Add by Morgan 2010/12/22
'   '重新核稿承辦人不可為原翻譯的核稿人
'   If text1(1).Text = "229" Then
'      If Left(text1(0), 1) = "F" Then
'         strExc(1) = "select * from staff_idmap where sim02='" & text1(0) & "' and sim01=ep04"
'      Else
'         strExc(1) = "select * from staff_idmap where sim01='" & text1(0) & "' and sim02=ep04"
'      End If
'      strExc(0) = "select ep04 from caseprogress,engineerprogress" & _
'         " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "'" & _
'         " and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='201' and ep02(+)=cp09" & _
'         " and (ep04='" & text1(0) & "' or exists(" & strExc(1) & "))"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         MsgBox "重新核稿承辦人不可為原翻譯的核稿人!!", vbExclamation
'         Exit Function
'      End If
'   End If
'
'   CheckCP14 = True
'End Function

'Removed by Morgan 2012/11/8 改寫為 PUB_CheckDivCase
''Add by Morgan 2004/3/30
''檢查母案是否存在
'Private Function CheckDivCase(ByRef stPA09) As Boolean
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
'   'Add by Morgan 2004/4/29
'   If (txtDivCaseNo(1) = PA(1) And txtDivCaseNo(2) = PA(2) And txtDivCaseNo(3) = PA(3) And txtDivCaseNo(4) = PA(4)) Then
'      MsgBox "分割案不可為母案！", vbExclamation
'      Exit Function
'   End If
'
'   stSQL = "select PA08, PA09 from patent where pa01='" & ChgSQL(txtDivCaseNo(1)) & "' and pa02='" & ChgSQL(txtDivCaseNo(2)) & "' and  pa03='" & ChgSQL(txtDivCaseNo(3)) & "' and pa04='" & ChgSQL(txtDivCaseNo(4)) & "'"
'
'   rsQuery.CursorLocation = adUseClient
'   rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsQuery.RecordCount > 0 Then
'      stPA08 = "" & rsQuery.Fields(0)
'      stPA09 = "" & rsQuery.Fields(1)
'      'Add by Morgan 2004/4/29
'      '分割案與母案的申請國家和專利種類需相同
'      If stPA09 <> PA(9) Then
'         MsgBox "分割案與母案的申請國家需相同！", vbExclamation
'      ElseIf stPA08 <> text1(21) Then
'         MsgBox "分割案與母案的專利種類需相同！", vbExclamation
'      Else
'         CheckDivCase = True
'      End If
'   Else
'      MsgBox "分割母案本所案號不存在！", vbExclamation
'   End If
'
'flgErr:
'   Set rsQuery = Nothing
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
'
'End Function

'cancel by sonia 2016/5/26 先收實審但不發文, 後續再收主動修正, 則會因此點限制而不能分案
''Add by Morgan 2007/6/12
''第一次分案時若有實審與(分割或主動修正)同時收文時提醒必須同一人承辦
'Private Function AssignNote() As Boolean
'   AssignNote = True
'   'Modify by Morgan 2007/10/29 新案未提申的不用檢查--靜芳
'   'If m_CP27 = "" And (text1(1) = "203" Or text1(1) = "307" Or text1(1) = "416") Then
'   If pa(10) <> "" And m_CP27 = "" And (text1(1) = "203" Or text1(1) = "307" Or text1(1) = "416") Then
'   'end 2007/10/29
'      strExc(0) = "select cp14,cpm03 from caseprogress,casepropertymap" & _
'         " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
'         " and cp27 is null and cp57 is null and cp10 in (" & IIf(text1(1) = "416", "'203','307'", "'416'") & ")" & _
'         " and cpm01(+)=cp01 and cpm02(+)=cp10"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         m_ExpCP14 = "" & RsTemp("cp14")
'         m_RelCP10n = "" & RsTemp("cpm03")
'         If RsTemp("cp14") <> "" Then
'            If text1(0) = "" Then
'               MsgBox "本案同時有收文【" & m_RelCP10n & "】，請注意需為同一承辦人！", vbExclamation
'            ElseIf text1(0) <> m_ExpCP14 Then
'               MsgBox "本案同時有收文【" & m_RelCP10n & "】，必需為同一承辦人！", vbExclamation
'               AssignNote = False
'            End If
'         End If
'      End If
'   End If
'End Function
'end 2016/5/26

Private Sub txtTF05_GotFocus()
   TextInverse txtTF05
End Sub

Private Sub txtTF05_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtTF05_Validate(Cancel As Boolean)
   If Val(txtTF05) > 100 Then
      MsgBox "不可大於100！", vbCritical
      Cancel = True
   End If
End Sub

Private Sub txtTF18_GotFocus()
   TextInverse txtTF18
End Sub

Private Sub txtTF18_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
      KeyAscii = 0
      Beep
   End If
End Sub
'Remove by Lydia 2016/06/21 改成模組 PUB_GetFCPsetCP48
''Add by Morgan 2008/8/20
''設定承辦期限
'Private Sub SetCP48()
'   Dim bolExcept As Boolean
'   '讀檔時不必觸發
'   'Modified by Morgan 2013/1/7 收文自動分案的也要執行
'   'If Me.Visible = True And m_CP27 = "" Then
'   If (Me.Visible = True Or (text1(0) <> "" And m_CP122 = "")) And m_CP27 = "" Then
'      '翻譯
'      If text1(1).Text = "201" Then
'         If text1(0).Text <> text1(0).Tag Then
'            '設定舜禹(F5588)時預設1個月
            'Modified by Lydia 2016/06/29 +捷恩凱(F5653)
'            If text1(0).Text = "F5588" Or text1(0).Text = "F5653" Then
'               text1(23) = TransDate(PUB_GetWorkDay1(CompDate(1, 1, strSrvDate(1)), False), 1)
'            'Add by Morgan 2008/10/20 內翻預設3個月
'            ElseIf PUB_GetMapID(text1(0)) <> "" Then
'               'Modify by Morgan 2010/8/25 改75天
'               'Text1(23) = TransDate(PUB_GetWorkDay1(CompDate(1, 3, strSrvDate(1)), False), 1)
'               text1(23) = TransDate(PUB_GetWorkDay1(CompDate(2, 75, strSrvDate(1)), False), 1)
'            Else
'               text1(23).Text = ""
'            End If
'         End If
'      'Add by Morgan 2008/9/4
'      '檢視中說
'      'Modify by Morgan 2008/9/10 +210 製作中說
'      'Modified by Morgan 2013/11/6 +235核對中說格式
'      ElseIf text1(1).Text = "209" Or text1(1).Text = "235" Or text1(1).Text = "210" Then
'         If text1(28).Text <> "" And text1(28).Text <> text1(28).Tag Then
'            text1(23) = TransDate(Pub_GetHandleDay(pa(1), pa(9), text1(1), DBDATE(text1(28))), 1)
'         End If
'         text1(28).Tag = text1(28).Text
'      '其他
'      Else
'
'         'Add by Morgan 2008/10/23
'         '若實審之承辦為程序且發明申請已發文時承辦期限為10個工作天
'         If text1(1) = "416" And text1(0) <> "" Then
'            'Modified by Morgan 2013/1/8
'            'If PUB_GetST03(text1(0)) = "F22" Then
'            '   If PUB_ChkCPExist(pa, "101", 2) = True Then
'            '      text1(23) = TransDate(CompWorkDay(10, DBDATE(text1(12))), 1)
'            '      bolExcept = True
'            '   End If
'            'End If
'            '改檢查若"沒有"發明或分割未發文設承辦期限15個工作天
'            strExc(0) = "select * from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and  cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in ('101','307') and cp57||cp27 is null"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 0 Then
'               text1(23) = TransDate(CompWorkDay(15, DBDATE(text1(12))), 1)
'               bolExcept = True
'            End If
'            'end 2013/1/8
'         End If
'         'end 2008/10/23
'
'         If text1(1).Text <> text1(1).Tag Then
'            If bolExcept = False Then
'               text1(23) = TransDate(Pub_GetHandleDay(pa(1), pa(9), text1(1), DBDATE(text1(12))), 1)
'            End If
'         End If
'      End If
'      If text1(23) <> "" And text1(4) <> "" And Val(text1(23)) > Val(text1(4)) Then
'         text1(23) = text1(4)
'      End If
'      Combo2.ListIndex = -1
'      text1(0).Tag = text1(0).Text
'      text1(1).Tag = text1(1).Text
'   End If
'   'end 2008/8/18
'End Sub

Private Sub SetStartDate2Tag()
   'Modified by Morgan 2022/10/12 +435續行母案再審
   strExc(0) = "select pa12,c1.cp05,nvl(nvl(c4.cp27,c2.cp27),c3.cp27) cp27" & _
      " from patent,caseprogress c1,caseprogress c2,caseprogress c3,caseprogress c4" & _
      " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "'" & _
      " and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'" & _
      " and c1.cp01(+)=pa01 and c1.cp02(+)=pa02 and c1.cp03(+)=pa03 and c1.cp04(+)=pa04 and c1.cp10='1204'" & _
      " and c2.cp01(+)=pa01 and c2.cp02(+)=pa02 and c2.cp03(+)=pa03 and c2.cp04(+)=pa04 and c2.cp10(+)='107'" & _
      " and c3.cp01(+)=pa01 and c3.cp02(+)=pa02 and c3.cp03(+)=pa03 and c3.cp04(+)=pa04 and c3.cp10(+)='416'" & _
      " and c4.cp01(+)=pa01 and c4.cp02(+)=pa02 and c4.cp03(+)=pa03 and c4.cp04(+)=pa04 and c4.cp10(+)='435'" & _
      " and c1.cp05>nvl(nvl(c4.cp27,c2.cp27),c3.cp27) order by c1.cp05 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      If IsNull(.Fields("pa12")) Then
         MsgBox "公開日尚未輸入無法判斷起算日!!"
      ElseIf .Fields("cp27") > .Fields("pa12") Then
         Combo3.Tag = .Fields("cp27") '實(再)審發文日
      Else
         Combo3.Tag = .Fields("cp05") '通知實審日
      End If
      End With
   End If
End Sub

Private Sub ChkCP43()
   Text1_Validate 6, False
   If Combo3.Enabled = True Then
      If Text1(4) = "" Then
         SetStartDate2Tag
         If Val(Combo3.Tag) < 20080110 Then
            SetDueDate 18, Combo3.Tag
         End If
      End If
   End If
End Sub

Private Sub SetDueDate(iMonths As Integer, strStartDate As String)
   If m_CP27 = "" Then
      'Modified by Lydia 2025/11/12 改抓最近工作天+PUB_GetWorkDay1
      Text1(4) = TransDate(PUB_GetWorkDay1(CompDate(1, iMonths, strStartDate), True), 1)
      Text1(4).Tag = Text1(4).Text 'Add By Sindy 2015/12/16
      Text1(5) = Text1(4)
   End If
End Sub

Private Sub SetDesignCase()
   If pa(1) = "FCP" And (Text1(1) = "125" Or Text1(1) = "308") Then
      lblDesignCase.Visible = True
      txtDesignCaseNo(1).Visible = True
      txtDesignCaseNo(2).Visible = True
      txtDesignCaseNo(3).Visible = True
      txtDesignCaseNo(4).Visible = True
      If m_CP30 <> "" Then
         txtDesignCaseNo(1) = Left(m_CP30, Len(m_CP30) - 9)
         txtDesignCaseNo(2) = Mid(m_CP30, Len(m_CP30) - 8, 6)
         txtDesignCaseNo(3) = Mid(m_CP30, Len(m_CP30) - 2, 1)
         txtDesignCaseNo(4) = Right(m_CP30, 2)
         txtDesignCaseNo(1).Tag = txtDesignCaseNo(1)
         txtDesignCaseNo(2).Tag = txtDesignCaseNo(2)
         txtDesignCaseNo(3).Tag = txtDesignCaseNo(3)
         txtDesignCaseNo(4).Tag = txtDesignCaseNo(4)
      End If
   Else
      lblDesignCase.Visible = False
      txtDesignCaseNo(1).Visible = False
      txtDesignCaseNo(2).Visible = False
      txtDesignCaseNo(3).Visible = False
      txtDesignCaseNo(4).Visible = False
   End If
End Sub

'Add by Amy 2013/08/20
'退費的相關收文號，是否為最後一道有發文規費的案件進度
Private Function GetLastCP09(ByRef strMsg As String) As String
    Dim strCaseType1 As String, strCaseType2 As String
    'Modified by Morgan 2022/11/23 435改比照416(可能會用補收款繳規費) Ex:FCP-067213--Winfrey
    strCaseType1 = "416,435"         '案件性質抓實審(不需判斷發文規費,以補文件繳)
    '2013/12/11 modify by sonia 再加合併702(FCP-035511)
    'modify by sonia 2018/12/28 再加領證601(FCP-056544)
    'modify by sonia 2021/10/4  再加續行母案再審435(FCP-064306)
    strCaseType2 = "107,404,417,605,702,601"   '案件性質抓再審、延期、提早公開、年費(需判斷發文規費)
    
    GetLastCP09 = ""
    strExc(0) = "Select CP09,CP27,CPM03 From CaseProgress,CasePropertyMap Where cpm01(+)=cp01 and cpm02(+)=cp10 " & _
                     "And (Instr('" & strCaseType1 & "',cp10)>0 Or (Instr('" & strCaseType2 & "',cp10)>0 and cp84 >0)) " & _
                     "And " & ChgCaseprogress(Label3(9)) & " Order by cp27 Desc"

    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        RsTemp.MoveFirst
        GetLastCP09 = RsTemp.Fields("CP09")
        strMsg = "相關收文號點選錯誤！應點選 相關收文號:" & RsTemp.Fields("CP09") & "  案件性質:" & RsTemp.Fields("CPM03")
    'Added by Moragn 2015/8/12
    '沒資料也要回傳,否則會彈空的訊息
    Else
       'modify by sonia 2021/10/4  再加續行母案再審435(FCP-064306)
       strMsg = "相關收文號點選錯誤！請點選 實審 或有發文規費的 再審、續行母案再審、延期、提早公開及年費。"
   'end 2015/8/12
    End If
End Function
'end 2013/08/20

'Added by Lydia 2017/05/17
Private Sub txtTF23_GotFocus()
   TextInverse txtTF23
End Sub

Private Sub txtTF23_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtTF19_GotFocus()
   TextInverse txtTF19
End Sub

Private Sub txtTF19_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtTF23_Validate(Cancel As Boolean)
   If txtTF23 <> "" Then
      txtTF23 = Val(txtTF23)
   End If
End Sub

Private Sub txtTF19_Validate(Cancel As Boolean)
   If txtTF19 <> "" Then
      If Val(txtTF19) > 100 Then
         MsgBox "相似度不可大於100！"
         Cancel = True
         TextInverse txtTF19
      End If
      txtTF19 = Val(txtTF19)
   End If
End Sub
'end 2017/05/17

'Added by Lydia 2020/05/20 法律所案源收文：案件性質=>案源案件類型
Private Sub SetLOSagree()
Dim m_LOSkind As String

    If strSrvDate(1) >= 法律所案源收文啟用日 And pa(1) = "FCP" Then
        'Modified by Lydia 2020/06/29 直接用案源檔的類型
        'm_LOSkind = PUB_GetLOSkind(pa(1), Text1(1).Text, "000")
        m_LOSkind = m_LOS02
        txtLOSagree.Text = ""
        FraLOS.Visible = False

        If Left(m_LOSkind, 1) = "C" And m_LOS01 = "" Then 'C類-未分案通知
             FraLOS.Visible = True
             txtLOSagree.Text = "Y"
        End If
    End If

End Sub

Private Sub txtLOSagree_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 89 And KeyAscii <> 8 Then
       KeyAscii = 0
       Beep
    End If
End Sub

Private Sub txtLOSagree_GotFocus()
   TextInverse txtLOSagree 'Added by Lydia 2020/05/29
End Sub

'Added by Lydia 2020/05/20 法律所案源收文：讀取法務案源檔
Private Sub ReadLOS()
Dim stSQL As String, intQ As Integer
Dim RsQ As ADODB.Recordset
   
   m_LOS01 = ""
   m_LOS07 = ""
   m_LOS15 = ""
   Text1(1).Locked = False
   If strSrvDate(1) >= 法律所案源收文啟用日 And pa(1) = "FCP" Then
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
           'If m_LOS01 <> "" Then  'Mark by Lydia 2020/07/14 都不可以變更
               Text1(1).Locked = True
           'End If
        End If
        Set RsQ = Nothing
   End If
End Sub
'Added by Morgan 2022/5/5
'更新續行母案再審期限:分割未發文>>同分割期限(分案),分割已發文>>發文日+4個月(發文)--陳亭妙
Private Sub Upd435Date()
   strExc(0) = "select cp06,cp07,cp48,cp27 from caseprogress where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='307'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      
   End If
End Sub


