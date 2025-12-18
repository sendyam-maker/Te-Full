VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060101_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "外專分案-FMP案"
   ClientHeight    =   5784
   ClientLeft      =   132
   ClientTop       =   972
   ClientWidth     =   8808
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5784
   ScaleWidth      =   8808
   Begin TabDlg.SSTab SSTab1 
      Height          =   4608
      Left            =   96
      TabIndex        =   36
      Top             =   1152
      Width           =   8628
      _ExtentX        =   15219
      _ExtentY        =   8128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   529
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm060101_3.frx":0000
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
      Tab(0).Control(28)=   "Label5(12)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label5(7)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label5(6)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label5(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label5(4)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label5(3)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label5(2)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cboCP14"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label3(0)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "text1(11)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "text1(10)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "text1(7)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "text1(23)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtDivCaseNo(4)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "text1(22)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtDivCaseNo(3)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtDivCaseNo(2)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtDivCaseNo(1)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "text1(18)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "text1(17)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "text1(16)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "text1(15)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "text1(14)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "text1(13)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "text1(12)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "text1(8)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "text1(6)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "text1(4)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "text1(27)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "text1(26)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "text1(25)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "text1(2)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "MSHFlexGrid1"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "text1(9)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "text1(5)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "text1(3)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "text1(1)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "txtTF05"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "txtTF18"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "text1(28)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "text1(29)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "text1(0)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "Combo2"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).ControlCount=   71
      TabCaption(1)   =   "備註"
      TabPicture(1)   =   "frm060101_3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(10)"
      Tab(1).Control(1)=   "Label1(9)"
      Tab(1).Control(2)=   "Label2(0)"
      Tab(1).Control(3)=   "Label2(1)"
      Tab(1).Control(4)=   "Label2(2)"
      Tab(1).Control(5)=   "Label2(3)"
      Tab(1).Control(6)=   "textPA91"
      Tab(1).Control(7)=   "textCP64"
      Tab(1).ControlCount=   8
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   6210
         Style           =   2  '單純下拉式
         TabIndex        =   96
         Top             =   1050
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   0
         Left            =   1740
         MaxLength       =   6
         TabIndex        =   94
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   29
         Left            =   5730
         MaxLength       =   1
         TabIndex        =   31
         Top             =   2940
         Width           =   330
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
      Begin VB.TextBox txtTF18 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   5340
         MaxLength       =   3
         TabIndex        =   30
         Top             =   2670
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTF05 
         Alignment       =   1  '靠右對齊
         Height          =   270
         Left            =   5340
         MaxLength       =   3
         TabIndex        =   29
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
         TabIndex        =   13
         Top             =   1305
         Width           =   855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1050
         Left            =   240
         TabIndex        =   37
         Top             =   3510
         Width           =   8115
         _ExtentX        =   14309
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
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1305
         Width           =   1935
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   12
         Left            =   1350
         MaxLength       =   8
         TabIndex        =   21
         Top             =   1815
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   13
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   23
         Top             =   2025
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   14
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   24
         Top             =   2265
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   15
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   25
         Top             =   2505
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   16
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   26
         Top             =   2745
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   17
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   27
         Top             =   2985
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   18
         Left            =   1350
         MaxLength       =   9
         TabIndex        =   28
         Top             =   3228
         Width           =   855
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   270
         Index           =   1
         Left            =   5835
         MaxLength       =   3
         TabIndex        =   15
         Top             =   1575
         Width           =   390
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   270
         Index           =   2
         Left            =   6225
         MaxLength       =   6
         TabIndex        =   16
         Top             =   1575
         Width           =   705
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   270
         Index           =   3
         Left            =   6930
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1575
         Width           =   390
      End
      Begin VB.TextBox text1 
         Enabled         =   0   'False
         Height          =   270
         Index           =   22
         Left            =   5340
         MaxLength       =   6
         TabIndex        =   22
         Top             =   2115
         Width           =   855
      End
      Begin VB.TextBox txtDivCaseNo 
         Height          =   270
         Index           =   4
         Left            =   7290
         MaxLength       =   2
         TabIndex        =   18
         Top             =   1575
         Width           =   390
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
         TabIndex        =   12
         Top             =   1305
         Width           =   855
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   10
         Left            =   1350
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1545
         Width           =   375
      End
      Begin VB.TextBox text1 
         Height          =   270
         Index           =   11
         Left            =   5625
         MaxLength       =   1
         TabIndex        =   20
         Top             =   1845
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   180
         Index           =   0
         Left            =   2640
         TabIndex        =   95
         Top             =   60
         Visible         =   0   'False
         Width           =   480
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
      Begin MSForms.Label Label5 
         Height          =   180
         Index           =   2
         Left            =   2250
         TabIndex        =   93
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
         Height          =   180
         Index           =   3
         Left            =   2250
         TabIndex        =   92
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
         Index           =   4
         Left            =   2250
         TabIndex        =   91
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
         Index           =   5
         Left            =   2250
         TabIndex        =   90
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
         Index           =   6
         Left            =   2250
         TabIndex        =   89
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
         Index           =   7
         Left            =   2250
         TabIndex        =   88
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
         Height          =   195
         Index           =   12
         Left            =   6210
         TabIndex        =   87
         Top             =   2130
         Width           =   795
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "1402;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   1170
         Left            =   -73980
         TabIndex        =   39
         Top             =   390
         Width           =   7515
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "13256;2064"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA91 
         Height          =   1170
         Left            =   -73980
         TabIndex        =   40
         Top             =   1590
         Width           =   7515
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "13256;2064"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否向客戶收款           (N: 不收)"
         Height          =   180
         Index           =   15
         Left            =   4380
         TabIndex        =   84
         Top             =   2985
         Width           =   2445
      End
      Begin VB.Label lblEP06 
         AutoSize        =   -1  'True
         Caption         =   "文件齊備日"
         Height          =   180
         Left            =   6705
         TabIndex        =   83
         Top             =   810
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "取消收文日"
         Height          =   180
         Index           =   42
         Left            =   6705
         TabIndex        =   43
         Top             =   1365
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦期限"
         Height          =   180
         Index           =   14
         Left            =   4380
         TabIndex        =   82
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label lblTF18 
         Caption         =   "加成比率                       %"
         Height          =   180
         Left            =   4380
         TabIndex        =   81
         Top             =   2670
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VB.Label lblTF05 
         Caption         =   "相似折扣                       %"
         Height          =   180
         Left            =   4380
         TabIndex        =   80
         Top             =   2400
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label lblDivCase 
         AutoSize        =   -1  'True
         Caption         =   "分割母案本所案號"
         Height          =   180
         Left            =   4380
         TabIndex        =   79
         Top             =   1575
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "核稿人"
         Height          =   180
         Index           =   13
         Left            =   4380
         TabIndex        =   78
         Top             =   2115
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(N:不算)"
         Height          =   180
         Index           =   11
         Left            =   2310
         TabIndex        =   73
         Top             =   1575
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   180
         Index           =   1
         Left            =   6300
         TabIndex        =   69
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   3
         Left            =   -72000
         TabIndex        =   64
         Top             =   3624
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   2
         Left            =   -72120
         TabIndex        =   63
         Top             =   2484
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   1
         Left            =   -72120
         TabIndex        =   62
         Top             =   1524
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   0
         Left            =   -72120
         TabIndex        =   61
         Top             =   504
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人"
         Height          =   180
         Index           =   38
         Left            =   150
         TabIndex        =   60
         Top             =   3225
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質"
         Height          =   180
         Index           =   37
         Left            =   4380
         TabIndex        =   59
         Top             =   345
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人5"
         Height          =   180
         Index           =   36
         Left            =   150
         TabIndex        =   58
         Top             =   2985
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   57
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "轉本所案號"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   56
         Top             =   585
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   55
         Top             =   825
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "相關總收文號"
         Height          =   180
         Index           =   3
         Left            =   150
         TabIndex        =   54
         Top             =   1065
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號"
         Height          =   180
         Index           =   4
         Left            =   150
         TabIndex        =   53
         Top             =   1305
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否算案件數"
         Height          =   180
         Index           =   29
         Left            =   150
         TabIndex        =   52
         Top             =   1575
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收文日"
         Height          =   180
         Index           =   30
         Left            =   150
         TabIndex        =   51
         Top             =   1815
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人1"
         Height          =   180
         Index           =   31
         Left            =   150
         TabIndex        =   50
         Top             =   2025
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人2"
         Height          =   180
         Index           =   32
         Left            =   150
         TabIndex        =   49
         Top             =   2265
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人3"
         Height          =   180
         Index           =   34
         Left            =   150
         TabIndex        =   48
         Top             =   2505
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人4"
         Height          =   180
         Index           =   35
         Left            =   150
         TabIndex        =   47
         Top             =   2745
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "公告日"
         Height          =   180
         Index           =   39
         Left            =   4380
         TabIndex        =   46
         Top             =   1305
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限"
         Height          =   180
         Index           =   40
         Left            =   4380
         TabIndex        =   45
         Top             =   825
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "卷宗性質                          (1.申請 2.異議 3.舉發)"
         Height          =   180
         Index           =   41
         Left            =   4380
         TabIndex        =   44
         Top             =   585
         Width           =   3585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否取消閉卷                  (Y : 取消閉卷)"
         Height          =   180
         Index           =   43
         Left            =   4365
         TabIndex        =   42
         Top             =   1860
         Width           =   2985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "進度備註"
         Height          =   180
         Index           =   9
         Left            =   -74760
         TabIndex        =   41
         Top             =   444
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件備註"
         Height          =   180
         Index           =   10
         Left            =   -74760
         TabIndex        =   38
         Top             =   1620
         Width           =   720
      End
   End
   Begin VB.TextBox text1 
      Enabled         =   0   'False
      Height          =   270
      Index           =   21
      Left            =   3768
      MaxLength       =   1
      TabIndex        =   77
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一筆(&N)"
      Height          =   350
      Index           =   5
      Left            =   3300
      TabIndex        =   32
      Top             =   10
      Width           =   990
   End
   Begin VB.TextBox text1 
      Height          =   270
      Index           =   24
      Left            =   4968
      TabIndex        =   72
      Top             =   336
      Visible         =   0   'False
      Width           =   564
   End
   Begin VB.CommandButton Command2 
      Caption         =   "案件進度(&C)"
      Height          =   350
      Index           =   1
      Left            =   5412
      TabIndex        =   33
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton Command2 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   350
      Index           =   2
      Left            =   6636
      TabIndex        =   34
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   3
      Left            =   7464
      TabIndex        =   35
      Top             =   10
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "程序："
      Height          =   180
      Index           =   17
      Left            =   7332
      TabIndex        =   100
      Top             =   408
      Width           =   540
   End
   Begin MSForms.Label Label5 
      Height          =   192
      Index           =   11
      Left            =   7932
      TabIndex        =   99
      Top             =   408
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
      Caption         =   "工程師組別："
      Height          =   180
      Index           =   18
      Left            =   5820
      TabIndex        =   98
      Top             =   636
      Width           =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   204
      Index           =   2
      Left            =   6948
      TabIndex        =   97
      Top             =   636
      Width           =   1380
   End
   Begin MSForms.Label Label5 
      Height          =   192
      Index           =   10
      Left            =   6408
      TabIndex        =   86
      Top             =   408
      Width           =   792
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "1397;339"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   288
      Left            =   936
      TabIndex        =   85
      Top             =   840
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
      TabIndex        =   76
      Top             =   408
      Width           =   468
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   12
      Left            =   2856
      TabIndex        =   75
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
      TabIndex        =   74
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   9
      Left            =   1344
      TabIndex        =   71
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   180
      Index           =   8
      Left            =   1350
      TabIndex        =   70
      Top             =   390
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號"
      Height          =   180
      Index           =   5
      Left            =   144
      TabIndex        =   68
      Top             =   408
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號"
      Height          =   180
      Index           =   6
      Left            =   144
      TabIndex        =   67
      Top             =   648
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱"
      Height          =   180
      Index           =   7
      Left            =   144
      TabIndex        =   66
      Top             =   888
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權："
      Height          =   180
      Index           =   8
      Left            =   5820
      TabIndex        =   65
      Top             =   408
      Width           =   540
   End
End
Attribute VB_Name = "frm060101_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/12 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Created by Morgan 2012/5/16
Option Explicit

Dim StrTot1(0 To 500) As String, StrTot2(0 To 500) As String
Dim m_PrevForm As Form 'Added by Lydia 2018/05/21前一畫面(表單)
Dim IntNow As Integer, IntTot As Integer
Dim strReceiveNo As String, strKind As String
Dim pa() As String, intWhere As Integer
Dim m_CP60 As String, m_CP14 As String  'Add by Morgan 2013/11/22
Dim m_CP27 As String '發文日
Dim m_CP122 As String 'Add by Sindy 2022/6/20
'Added by Lydia 2018/08/07 命名作業的資料
Dim m_TCT01 As String '新案收文號(PK)
Dim m_TCT10 As String '命名人員
Dim m_TCT27 As String '欲翻譯此案件者/指定翻譯
Dim m_TCT28 As String '其他指定翻譯
Dim mTransKind As String 'Added by Lydia 2018/08/08 只能上班翻譯 (Y後面+逗號,再加上有折扣或固定報價)
Dim mLimitDate As String 'Added by Lydia 2022/07/05 翻譯分案->交稿期限
Dim bUpdPA150 As String 'Added by Lydia 2022/09/29 變更案件之工程師組別 ( 新增工程師分組控管)
'Modified by Lydia 2025/06/05 更改名稱
'Dim m_strBASF As String 'Added by Lydia 2023/04/19 BASF集團的X編號
Dim m_str所內譯 As String
Dim m_str所內譯例外 As String 'Added by Lydia 2025/07/01
Dim m_CP31 As String 'Added by Lydia 2024/06/12 是否為新案
Dim m_NA16Na79 As String 'Added by Lydia 2024/10/04 程序管制人

'Add By Sindy 2016/7/26
'承辦人
Private Sub CboCP14_GotFocus()
   cboCP14.SelStart = 0
   cboCP14.SelLength = Len(cboCP14.Text)
End Sub
Private Sub CboCP14_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub CboCP14_Validate(Cancel As Boolean)
Dim m_Team As String
Dim strText As String
Dim bolRunOK As Boolean
   
   If cboCP14.Text <> "" Then
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
      
      '檢查人員是否存在或離職
      If ChkStaffST04(Left(Trim(cboCP14.Text), 5)) = True Then
         If Me.Visible Then 'Added by Morgan 2016/11/16 已分案後離職再分案會當 Ex.P113150
            Cancel = True
            cboCP14.Text = ""
            cboCP14.SetFocus
            Call CboCP14_GotFocus
            Exit Sub
         End If
      End If
      
      '2011/11/30 ADD BY SONIA 林信昌因分組故自動帶與案件組別的編號
      If InStr(Trim(Mid(Trim(cboCP14.Text), 6)), "林信昌") > 0 Then
         If pa(1) = "P" Or pa(1) = "CFP" Then m_Team = pa(150)
         If pa(1) = "PS" Or pa(1) = "CPS" Then m_Team = pa(79)
         Select Case m_Team
            Case "1"
               If Left(Trim(cboCP14.Text), 1) = "6" Then cboCP14.Text = "68091" & " " & GetPrjSalesNM("68091")
               If Left(Trim(cboCP14.Text), 1) = "F" Then cboCP14.Text = "F5644" & " " & GetPrjSalesNM("F5644")
            Case "2"
               If Left(Trim(cboCP14.Text), 1) = "6" Then cboCP14.Text = "68092" & " " & GetPrjSalesNM("68092")
               If Left(Trim(cboCP14.Text), 1) = "F" Then cboCP14.Text = "F5645" & " " & GetPrjSalesNM("F5645")
            Case Else
               If Left(Trim(cboCP14.Text), 1) = "6" Then cboCP14.Text = "68007" & " " & GetPrjSalesNM("68007")
               If Left(Trim(cboCP14.Text), 1) = "F" Then cboCP14.Text = "F5162" & " " & GetPrjSalesNM("F5162")
         End Select
      End If
      
      'Add by Sindy 2022/6/20
      If Cancel = False And Left(Trim(cboCP14.Text), 5) <> Trim(cboCP14.Tag) Then
         text1(0).Text = Trim(Left(cboCP14.Text, 6))  'Added by Lydia 2022/06/23 P.S.修改段落有兩人，是因為之前Sindy有針對"外專分案-FMP案"修改，但是陳亭妙使用的是內專分案作業，所以只有增加Text1(0)和引用模組，後續沒有撰寫。
         Call PUB_GetFCPsetCP48(Me.Visible, pa, m_CP27, text1(1), text1(0), m_CP122, text1(4), text1(5), text1(23), text1(28), Combo2, text1(12))
         'Added by Lydia 2024/03/08 因應內專支援機械組OA (含P案) 收文: 927其他翻譯
         If mTransKind = "＃" Then
            '承辦期限: 分案日+14日曆天
            text1(23) = TransDate(CompDate(2, 14, strSrvDate(1)), 1)
         End If
         'end 2024/03/08
         'Added by Lydia 2022/07/05 翻譯分案->比對交稿期限和預設承辦期限,以最早期限為準
         If text1(23).Text <> "" And mLimitDate <> "" Then
             If TransDate(mLimitDate, 1) < text1(23).Text Then
                 text1(23).Text = TransDate(mLimitDate, 1)
             End If
         '如果沒有承辦期限
         ElseIf text1(23).Text = "" And mLimitDate <> "" Then
              text1(23).Text = TransDate(mLimitDate, 1)
         End If
         'end 2022/07/05
      End If
      '2022/6/20 END
   End If
   
   If Trim(cboCP14.Text) = "" Then
      Cancel = True
      cboCP14.SetFocus
      Call CboCP14_GotFocus
'      SetcboCP10
      Exit Sub
   End If
   
   Cancel = False 'Added by Lydia 2022/07/05
End Sub
'2016/7/26 END

'Add By Sindy 2016/7/26 依工程師組別,以下拉方式預設該組之工程師名單(依員工編號大小)供分案者點選,若無,就不顯示名單
Private Sub SetCboCP10()
Dim m_Team As String
   
   If Not (text1(1) = "201" Or text1(1) = "927") Then '201.新案翻譯;927.其他翻譯除外
      If pa(1) = "P" Or pa(1) = "CFP" Then m_Team = pa(150)
      If pa(1) = "PS" Or pa(1) = "CPS" Then m_Team = pa(79)
      If m_Team = "" Then
         cboCP14.Clear
      Else
         cboCP14.Clear
         'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
         'modify by sonia 2021/1/22 再排除F4104,F4105
         strSql = "select st01,st02 from staff" & _
                  " where st03='F21' and st04='1' and substr(st01,1)>='6' and substr(st01,1)<'F'" & _
                  " and substr(st01,4,1)<>'9' and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' " & _
                  " and st16='" & m_Team & "'" & _
                  " order by st01 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            With RsTemp
               RsTemp.MoveFirst
               Do While RsTemp.EOF = False
                  cboCP14.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
                  RsTemp.MoveNext
               Loop
            End With
         End If
      End If
   End If
End Sub

Private Sub Combo2_Click()
   If Combo2.ListIndex >= 0 Then
      text1(23) = TransDate(PUB_GetWorkDay1(CompDate(2, Combo2.ItemData(Combo2.ListIndex), strSrvDate(1)), False), 1)
      If text1(23) <> "" And text1(4) <> "" And Val(text1(23)) > Val(text1(4)) Then
         text1(23) = text1(4)
      End If
   End If
End Sub

Private Sub Command2_Click(Index As Integer)
Dim strCP06 As String  'Added by Lydia 2022/06/23
Dim dbTfRate As Double, bolIsHigher As Boolean  'Added by Lydia 2022/06/23 判斷翻譯費折扣率＞30%

   Select Case Index
      Case 1
         Set frm060101_2.fmParent = Me
         frm060101_2.Show
         Me.Hide
         
      Case 2 '確定
         'Screen.MousePointer = vbHourglass 'Remove by Lydia 2018/06/28
         If TxtValidate Then
         
            'Added by Lydia 2024/04/10 分案和工作進度維護點選不可查閱工程師需要彈訊息
            If Left(Trim(cboCP14), 1) > "6" And Left(Trim(cboCP14), 1) < "F" And cboCP14.Text <> cboCP14.Tag Then
               If PUB_ChkCufaByCaseNo(Trim(Left(cboCP14, 6)), Me.Name, pa(1) & pa(2) & pa(3) & pa(4), "2") = False Then
                  Exit Sub
               End If
            End If
            'end 2024/04/10
            
            'Added by Lydia 2023/04/19 外專翻譯分案承辦人不得為翻譯社及外譯人員
            If pa(1) = "P" And text1(1) = "927" And Trim(Left(cboCP14.Text, 6)) <> cboCP14.Tag Then
               'Modified by Lydia 2025/07/01 增加例外案件設定InStr(m_str所內譯例外,  pa(1) & pa(2) & pa(3) & pa(4)) = 0 And
               If InStr(m_str所內譯例外, pa(1) & pa(2) & pa(3) & pa(4)) = 0 And (InStr(m_str所內譯, ChangeCustomerL(pa(26))) > 0 Or InStr(m_str所內譯, ChangeCustomerL(pa(75))) > 0) Then
                  If Trim(Left(cboCP14.Text, 1)) = "F" Then
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
            'Added by Lydia 2018/08/07 外專翻譯管控
            'Modified by Lydia 2022/06/23 判斷有修改承辦人才做檢查 Trim(Left(cboCP14.Text, 6)) <> cboCP14.Tag
            If m_TCT01 <> "" And pa(1) = "P" And text1(1) = "201" And Trim(Left(cboCP14.Text, 6)) <> cboCP14.Tag Then
                'Added Lydia 2023/04/19 外專翻譯分案承辦人不得為翻譯社及外譯人員
                'Modified by Lydia 2025/07/01 增加例外案件設定InStr(m_str所內譯例外,  pa(1) & pa(2) & pa(3) & pa(4)) = 0 And
                If InStr(m_str所內譯例外, pa(1) & pa(2) & pa(3) & pa(4)) = 0 And (InStr(m_str所內譯, ChangeCustomerL(pa(26))) > 0 Or InStr(m_str所內譯, ChangeCustomerL(pa(75))) > 0) Then
                   If Trim(Left(cboCP14.Text, 1)) = "F" Then
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
                   'Added by Lydia 2021/06/23 判斷翻譯費折扣率
                   dbTfRate = PUB_GetTransFeeRate(pa(1), pa(2), pa(3), pa(4), , bolIsHigher, True)
                End If 'Added by Lydia 2023/04/19
                '控制翻譯費折扣率＞30%客戶案件之承辦人只能為所內人員上班譯編號。
                If dbTfRate > 30 Then
                    If Left(cboCP14, 1) = "F" Then
                        '因為案件經過呈報王協理後可以"非所內人員上班譯"，所以改成彈訊息詢問---Sharon 口頭協商
                        If MsgBox("該案件之承辦人只能為所內人員上班譯編號，是否繼續分案？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                           cboCP14.SetFocus
                            Exit Sub
                        End If
                    End If
                ElseIf bolIsHigher = True Then  '折扣率＞30%但是例外控制的客戶
                     '不受限
                End If
                'end 2022/06/23
                
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
                'End If 'Remov by Lydia 2018/09/27
            End If
            'end 2018/08/07
            'Added by Lydia 2018/08/08 若輸入下班翻譯在存檔前先提醒"只能上班翻譯是否繼續存檔"
            'Modified by Lydia 2025/03/13 改用模組取得
            'If Left(mTransKind, 1) = "Y" And Text1(1) = "201" And Trim(Left(cboCP14.Text, 1)) = "F" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, Trim(Left(cboCP14.Text, 6))) = 0 Then
            If Left(mTransKind, 1) = "Y" And text1(1) = "201" And Trim(Left(cboCP14.Text, 1)) = "F" And InStr(Pub_SetF51Order("F", ""), Trim(Left(cboCP14.Text, 6))) = 0 Then
                If MsgBox("只能上班翻譯：" & Mid(mTransKind, 3) & vbCrLf & "，是否繼續存檔？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                      cboCP14.SetFocus
                      Exit Sub
                End If
            End If
            'end 2018/08/07
            
            'Added by Lydia 2021/04/14 外專翻譯承辦及核稿期限控管：
            'Modified by Lydia 2025/03/13 改用模組取得
            'If pa(1) = "P" And Text1(1) = "201" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, Trim(Left(cboCP14.Text, 6))) = 0 Then
            If pa(1) = "P" And text1(1) = "201" And InStr(Pub_SetF51Order("F", ""), Trim(Left(cboCP14.Text, 6))) = 0 Then
                '工程師認領翻譯時，查詢該認領人員，新案翻譯未上完稿日案件,請彈提醒: 尚未完稿案件FCPxxxx , 承辦期限
                strExc(4) = Pub_GetEngEP09List(Trim(Left(cboCP14.Text, 6)))
                If strExc(4) <> "" Then
                    MsgBox "尚未完稿案件：" & strExc(4), vbCritical
                End If
            End If
            'end 2021/04/14
            
            'Add By Sindy 2019/10/17 'Memo by Lydia 2022/06/23 與FCP分案frm060101_1相同判斷
            If pa(1) = "P" And text1(1) = "201" Then
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
                  If DBDATE(text1(23)) > strCP06 Then
                     If MsgBox("注意會稿期限為:" & ChangeWStringToTDateString(strCP06) & "，承辦期限是否確定？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
                        text1(23).SetFocus
                        Exit Sub
                     End If
                  End If
               End If
            End If
            '2019/10/17 END  'Memo by Lydia 2022/06/23 與FCP分案frm060101_1相同判斷
            
            'Added by Lydia 2020/01/21 所有系統別分案，若該收文號存在於下一程序(NP24)，若修改案件性質則彈訊息
            If text1(1).Tag <> text1(1).Text Then
                If Pub_CheckNP24Exists(Label3(8).Caption) = True Then
                End If
            End If
            'end 2020/01/21
            'Added by Lydia 2022/09/29 新增工程師分組控管
            bUpdPA150 = False
            If Trim(Left(cboCP14.Text, 6)) <> cboCP14.Tag And InStr("927,", text1(1)) = 0 Then '除了「其他翻譯927」
                'Added by Lydia 2022/12/21 基本檔無工程師組別需要預設組別; ex.FG-001437
                If (((pa(1) = "P" Or pa(1) = "CFP") And pa(150) = "") Or ((pa(1) = "PS" Or pa(1) = "CPS") And pa(79) = "")) And PUB_GetST03(Trim(Left(cboCP14.Text, 6))) = "F21" Then
                    bUpdPA150 = True
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
         
            Screen.MousePointer = vbHourglass 'Added by Lydia 2018/06/28
            If FormSave = False Then
               ' 設定滑鼠游標為預設
               Screen.MousePointer = vbDefault
               MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
            Else
               'Added by Lydia 2018/06/15 翻譯分案無紙化:設定承辦人後，自動發mail
               Screen.MousePointer = vbDefault
               'Modified by Lydia 2018/06/28 判斷有修改才發mail通知
               'If TypeName(m_PrevForm) = "frm060122" And Trim(cboCP14.Text) <> "" Then
               If TypeName(m_PrevForm) = "frm060122" And Trim(cboCP14.Text) <> "" And Trim(Left(cboCP14.Tag, 6)) <> Trim(Left(cboCP14.Text, 6)) Then
                     m_PrevForm.nKeyNo = Trim(Left(cboCP14.Text, 6))
                     m_PrevForm.bolNextDone = True
               End If
               'end 2018/06/15
               
               If IntNow <> IntTot Then
                  GetData IntNow
                  cboCP14.SetFocus 'Modify By Sindy 2016/7/26
               Else
                  'Added by Lydia 2018/05/21  回到分案前畫面
                  If TypeName(m_PrevForm) = "frm060101" Then
                        Screen.MousePointer = vbHourglass
                  'end 2018/05/21
                        frm060101.Show
                        frm060101.RefreshData
                        frm060101.Show
                        Screen.MousePointer = vbDefault
                  End If
                  Unload Me
               End If
            End If
         End If
         ' 設定滑鼠游標為預設
         'Screen.MousePointer = vbDefault 'Remove by Lydia 2018/05/21
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
   End Select
End Sub

Private Function FormSave() As Boolean
   Dim bolTrans As Boolean
   Dim i As Integer, strTmp(1 To 3) As String
   
On Error GoTo CheckingErr
   
   cnnConnection.BeginTrans
   bolTrans = True
   
   'Modified by Lydia 2022/07/05 更新承辦期限
   'strSql = "update caseprogress set cp14='" & Left(Trim(cboCP14.Text), 5) & "' where cp09='" & strReceiveNo & "'"
   strSql = "update caseprogress set cp14='" & Left(Trim(cboCP14.Text), 5) & "', cp48=" & CNULL(DBDATE(text1(23)), True) & " where cp09='" & strReceiveNo & "'"
   'Added by Lydia 2022/07/05 留記錄
   If Left(Trim(cboCP14.Text), 5) <> Trim(cboCP14.Tag) Then
       '要寫 Log 並改觸發 Trigger 更新修改人員日期時間
       Pub_SeekTbLog strSql
       strSql = "begin user_data.user_enabled:=1; " & strSql & "; end;"
   End If
   'end 2022/07/05
   cnnConnection.Execute strSql, intI
   
   'Added by Morgan 2015/10/13
   '審查意見或核駁修改承辦人時一併修改相關收文號之告代承辦人 Ex.FCP-45516
   If Trim(cboCP14.Text) <> "" And (text1(1) = "1202" Or text1(1) = "1002" Or text1(1) = "1227") Then
      strSql = "update caseprogress set  cp14='" & Left(Trim(cboCP14.Text), 5) & "' where cp43='" & Label3(8) & "' and cp10='901' and cp27 is null"
      cnnConnection.Execute strSql, intI
   End If
   'end 2015/10/13
   
   
   'Added by Lydia 2024/03/12 參考FCP分案; 補上翻譯費用檔
   If (text1(1) = "201" Or text1(1) = "927") Then
      If txtTF05.Enabled = True Then
         If RTrim(txtTF05) = "" Or Val(txtTF05) = 100 Then
            strExc(10) = "Null"
         Else
            strExc(10) = Val(txtTF05)
         End If
         '加成比率
         If RTrim(txtTF18) = "" Or Val(txtTF18) = 100 Then
            strExc(9) = "Null"
         Else
            strExc(9) = Val(txtTF18)
         End If
         
         strSql = "update transfee set TF05=" & strExc(10) & ",TF18=" & strExc(9) & " where TF01='" & strReceiveNo & "' and tf07 is null"
         cnnConnection.Execute strSql, intI
         If intI = 0 Then
            strSql = "insert into transfee(TF01,TF05,TF18) values('" & strReceiveNo & "'," & strExc(10) & "," & strExc(9) & ")"
            cnnConnection.Execute strSql, intI
         End If
      End If
   End If
   'end 2024/03/12
      
   'Added by Lydia 2022/09/29 新增工程師分組控管
   '當承辦人為工程師時，所輸入的工程師組別與原來的工程師組別不一致時，談視窗詢問：是否變更工程師組別，是: 工程師組別改為此次輸入的工程師組別
   If bUpdPA150 = True Then
        If pa(1) = "P" Or pa(1) = "CFP" Then
            strSql = "Update Patent Set PA150=" & CNULL(PUB_GetStaffST16(Trim(Left(cboCP14.Text, 6)))) & " Where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "' "
        Else 'PS
            strSql = "Update ServicePractice Set SP79=" & CNULL(PUB_GetStaffST16(Trim(Left(cboCP14.Text, 6)))) & " Where sp01='" & pa(1) & "' and sp02='" & pa(2) & "' and sp03='" & pa(3) & "' and sp04='" & pa(4) & "' "
        End If
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
    End If
    'end 2022/09/29
    
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
      
      'Add by Morgan 2013/11/22
      '若已開請款單則換承辦人或核稿人時發Mail通知靜芳
      If m_CP60 > "X" Then
         'Modified by Lydia 2019/10/17 本所案號+"-"
         'PUB_PointReAssignInform Label3(9), m_CP60, m_CP14, Left(Trim(cboCP14.Text), 5)
         PUB_PointReAssignInform pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)), m_CP60, m_CP14, Left(Trim(cboCP14.Text), 5)
      End If
      'end 2013/11/22
      
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
   
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         If .TextMatrix(i, 0) = "v" And text1(1) = "404" Then
            strTmp(1) = .TextMatrix(i, 7)
            strTmp(2) = .TextMatrix(i, 8)
            strTmp(3) = .TextMatrix(i, 9)
            
            'Add By Sindy 2021/8/16
            If text1(1) = "404" Then '延期:點選下一程序性質時，若該性質續辦狀態是N，請將恢復期限管制，同時將解除期限日期及原因清除。
               strSql = "UPDATE NEXTPROGRESS SET NP06='',NP11=null,NP12=null WHERE NP01='" & strTmp(1) & "' AND " & _
                  "NP07=" & strTmp(2) & " AND NP22=" & strTmp(3) & " AND NP06='N'"
               cnnConnection.Execute strSql, intI
            End If
            '2021/8/16 END
         End If
      Next
   End With
   
   FormSave = True
   Exit Function
   
CheckingErr:
   If bolTrans Then cnnConnection.RollbackTrans
   FormSave = False
End Function

Private Sub Form_Initialize()
   ReDim pa(1 To TF_PA) As String
End Sub

'Added by Lydia 2018/05/21 改從SetParent傳收文號
'Modified by Lydia 2018/08/08 +只能上班翻譯 pTransKind
'Modified by Lydia 2022/07/05 +交稿期限pLimitDate
Public Sub SetParent(ByRef pForm As Form, ByVal pCnt As Integer, ByVal pCaseNo As String, ByVal pCpNo As String, Optional ByRef pTransKind As String, Optional ByRef pLimitDate As String)
Dim tmpArr As Variant, tmpArr2 As Variant
Dim intA As Variant
    Set m_PrevForm = pForm
    IntTot = pCnt
    tmpArr = Split(pCaseNo, ",")
    tmpArr2 = Split(pCpNo, ",")
    mTransKind = pTransKind 'Added by Lydia 2018/08/08
    mLimitDate = pLimitDate 'Added by Lydia 2022/07/05
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
'      IntTot = 0
'      If .Rows < 2 Then Exit Sub
'      For i = 1 To .Rows - 1
'         If .TextMatrix(i, 0) = "v" Then
'            StrTot1(IntTot) = .TextMatrix(i, 7) '本所案號
'            StrTot2(IntTot) = .TextMatrix(i, 1) '收文號
'            IntTot = IntTot + 1
'         End If
'      Next
'   End With
   'end 2018/05/21
   
   IntNow = 0
   GetData IntNow
   'Modified by Lydia 2025/06/05 更改名稱
   'm_strBASF = Pub_GetSpecMan("外專翻譯分案-BASF") & ","  'Added by Lydia 2023/04/19
   m_str所內譯 = Pub_GetSpecMan("外專翻譯分案-所內譯") & ","
   m_str所內譯例外 = Pub_GetSpecMan("外專翻譯分案-所內譯例外") & "," 'Added by Lydia 2025/07/01
   
   Me.SSTab1.Tab = 0 'Added by Lydia 2021/09/09
   SetCombo2 'Add by Morgan 2008/8/18
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

Private Sub GetData(intSitu As Integer)

   Dim i As Integer, txt As TextBox, Lbl As Object

   For Each txt In text1
      txt.Text = ""
      txt.Locked = True
   Next
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   'Add By Sindy 2021/5/12
   For Each Lbl In Label5
      Lbl.Caption = ""
   Next
   '2021/5/12 END
   
   cboCP14.Tag = "" 'Added by Lydia 2018/06/28
   
   '總收文號
   Label3(8) = StrTot2(intSitu)
   strReceiveNo = StrTot2(intSitu)
   '本所案號
   Label3(9) = StrTot1(intSitu)
   
   i = Len(Label3(9)) - 9
   pa(1) = Left(Label3(9), i)
   pa(2) = Mid(Label3(9), i + 1, 6)
   pa(3) = Mid(Label3(9), i + 7, 1)
   pa(4) = Right(Label3(9), 2)
   m_NA16Na79 = "" 'Added by Lydia 2024/10/04
   Combo1.Clear
   If pa(1) = "P" Or pa(1) = "CFP" Then
      If ClsPDReadPatentDatabase(pa(), intWhere) Then
         If pa(23) = "1" Then
            Combo1.AddItem "中 : " & pa(5)
            Combo1.AddItem "英 : " & pa(6)
            'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
            Combo1.AddItem "外 : " & pa(7)
         Else
            strExc(0) = "SELECT CP37,CP38,CP39 FROM CASEPROGRESS WHERE " & ChgCaseprogress(StrTot1(intSitu))
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
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
         text1(3) = pa(23)
         text1(8) = pa(48)
         
         If pa(57) = "Y" Then
            Label4 = "已閉卷"
            Label1(43).Visible = True
            text1(11).Visible = True
         Else
            Label4 = ""
            Label1(43).Visible = False
            text1(11).Visible = False
         End If
         
         text1(16).Enabled = True
         text1(17).Enabled = True
         For i = 26 To 30
            If pa(i) <> "" Then text1(i - 13) = pa(i): ChgType (i - 13)
         Next
         If pa(75) <> "" Then text1(18) = pa(75): ChgType (18)
         textPA91 = pa(91)
         
         Me.text1(21).Enabled = True
         Me.text1(21).Text = pa(8)
         If pa(1) = "CFP" Or pa(9) = "000" Then
            Me.Label3(11).Caption = "" & PUB_GetPatentKindName(Me.text1(21).Text, 台灣國家代號)
         Else
            Me.Label3(11).Caption = "" & PUB_GetPatentKindName(Me.text1(21).Text, pa(9))
         End If
      End If
   ElseIf pa(1) = "PS" Or pa(1) = "CPS" Then
      If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
         Combo1.AddItem "中 : " & pa(5)
         Combo1.AddItem "英 : " & pa(6)
         'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
         Combo1.AddItem "外 : " & pa(7)
         Combo1.ListIndex = 0
         text1(8) = pa(29)
         If pa(15) = "Y" Then
            Label4 = "已閉卷"
            Label1(43).Visible = True
            text1(11).Visible = True
         Else
            Label4 = ""
            Label1(43).Visible = False
            text1(11).Visible = False
         End If
         
         If pa(8) <> "" Then text1(13) = pa(8): ChgType (13)
         If pa(58) <> "" Then text1(14) = pa(58): ChgType (14)
         If pa(59) <> "" Then text1(15) = pa(59): ChgType (15)
         text1(16).Enabled = False
         text1(17).Enabled = False
         
         If pa(26) <> "" Then text1(18) = pa(26): ChgType (18)
         'pa(26) = pa(8) 'Mark by Lydia 2024/06/19
         textPA91 = pa(18)
      End If
   End If
      
   'Added by Lydia 2024/10/04
   '工程師組別
   If pa(1) = "P" Or pa(1) = "CFP" Then
      Label3(2).Caption = PUB_GetFCPGrpName(pa(150))
   Else
      Label3(2).Caption = PUB_GetFCPGrpName(pa(79))
   End If
   m_NA16Na79 = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) '程序人員
   Label5(11) = GetStaffName(m_NA16Na79, True)
   'end 2024/10/04
   
    'Added by Lydia 2018/08/07 讀取命名作業資料
    strExc(0) = "select TCT01,TCT10,TCT27,TCT28 from TRANSCASETITLE,caseprogress where TCT01=cp09(+) " & _
                      "and cp01= '" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and nvl(tct05,0)> 0 "
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
    
   '銷卷提醒
   CheckCaseDestroy pa(1), pa(2), pa(3), pa(4)
   
   'Modified by Lydia 2024/06/12 +CP31
   strExc(0) = "SELECT CP13,CP14,CP10,CP06,CP07,CP43,CP57,CP26,CP05," & _
      "CP64,CP48,CP60, DC05, DC06, DC07, DC08,CP27,CP20,CP30,CP122,CP31 FROM CASEPROGRESS, DIVISIONCASE " & _
      "WHERE DC01(+)=CP01 AND DC02(+)=CP02 AND DC03(+)=CP03 AND DC04(+)=CP04 AND CP09='" & StrTot2(intSitu) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   With RsTemp
    If intI = 1 Then
       If Not IsNull(.Fields(0)) Then
          text1(24) = .Fields(0)
          ChgType (24)
       End If
       'Add By Sindy 2016/7/26
       SetCboCP10
       'Added by Lydia 2024/03/08 927其他翻譯，承辦人為程序人員
       If mTransKind = "＃" Then
          cboCP14.Text = ""
       Else
       'end 2024/03/08
          If Not IsNull(.Fields("CP14")) Then
             cboCP14.Text = .Fields("CP14")
             CboCP14_Validate False
          End If
          '2016/7/26
       End If 'Added by Lydia 2024/03/08
       cboCP14.Tag = "" & .Fields("CP14") 'Added by Lydia 2018/06/28
       text1(0).Text = "" & .Fields("CP14")  'Added by Lydia 2022/06/23
       
       If Not IsNull(.Fields(2)) Then
          strKind = text1(1)
          text1(1) = .Fields(2)
          ChgType (1)
       End If
       text1(1).Tag = text1(1).Text 'Added by Lydia 2020/01/21
       
       If Not IsNull(.Fields(3)) Then text1(4) = TransDate(.Fields(3), 1)
       If Not IsNull(.Fields(4)) Then text1(5) = TransDate(.Fields(4), 1)
       If Not IsNull(.Fields(5)) Then text1(6) = .Fields(5)
       If Not IsNull(.Fields(6)) Then text1(9) = TransDate(.Fields(6), 1)
       If Not IsNull(.Fields(7)) Then text1(10) = .Fields(7)
       If Not IsNull(.Fields(8)) Then text1(12) = TransDate(.Fields(8), 1)
       If Not IsNull(.Fields(9)) Then textCP64 = .Fields(9)
       
        'Add by Sindy 2022/6/20
        m_CP27 = "" & .Fields("CP27")
        m_CP122 = "" & .Fields("CP122")
        '2022/6/20 END
        
        text1(29) = "" & .Fields("CP20")
        text1(23) = TransDate("" & .Fields("CP48"), 1)
        'Add by Morgan 2008/8/19
        If m_CP27 <> "" Then
            text1(23).Locked = True
            Combo2.Enabled = False
        End If
        'end 2008/8/19
        'Added by Lydia 2022/06/23 新案翻譯開放可編輯
        'Modified by Lydia 2024/03/12 +927其他翻譯
        If Val(m_CP27) = 0 And (text1(1) = "201" Or text1(1) = "927") Then
            text1(23).Locked = False
        End If
        'end 2022/06/23
        
        '分割母案本所案號
        txtDivCaseNo(1) = "" & .Fields("DC05").Value
        txtDivCaseNo(2) = "" & .Fields("DC06").Value
        txtDivCaseNo(3) = "" & .Fields("DC07").Value
        txtDivCaseNo(4) = "" & .Fields("DC08").Value
        If text1(1) = "307" Then
            DivVisibleSwitch True
        Else
            DivVisibleSwitch False
        End If
        m_CP60 = "" & .Fields("CP60"): m_CP14 = "" & .Fields("CP14") 'Added by Morgan 2013/11/22
        m_CP31 = "" & .Fields("CP31") 'Added by Lydia 2024/06/12
    End If
   End With
    
   GetGrid StrTot2(intSitu), 0
   
   IntNow = IntNow + 1
   
   Command2(2).Enabled = False
   'Modified by Lydia 2024/03/12 排除mTransKind <> "＃"; 因應內專支援機械組OA (含P案) 收文: 927其他翻譯
   If Trim(cboCP14.Text) = "" And mTransKind <> "＃" Then
      MsgBox "本程序專利處尚未分案不可作業！"
   ElseIf Left(PUB_GetST03(Left(Trim(cboCP14.Text), 5)), 1) <> "F" And mTransKind <> "＃" Then
      MsgBox "本程序目前承辦人非國外部不可作業！"
   Else
      Command2(2).Enabled = True
      cboCP14.Locked = False
   End If
End Sub

Private Function GetGrid(ByVal strRecive As String, ByVal intSitu As Integer) As Boolean
   GetGrid = True
   If intSitu = 0 Then
      strExc(1) = Label3(9)
   Else
      strExc(1) = Trim(text1(2)) & Trim(text1(25)) & Trim(text1(26)) & Trim(text1(27))
   End If
   '排除程序管制的案件性質
   If pa(1) = "P" Or pa(1) = "CFP" Then
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
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   If intSitu = 1 Then
      If pa(1) = "FCP" Then
         strExc(0) = "SELECT count(*) FROM PATENT WHERE " & ChgPatent(Trim(text1(2)) & Trim(text1(25)) & Trim(text1(26)) & Trim(text1(27)))
      Else
         strExc(0) = "SELECT count(*) FROM SERVICEPRACTICE WHERE " & ChgService(Trim(text1(2)) & Trim(text1(25)) & Trim(text1(26)) & Trim(text1(27)))
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If RsTemp.Fields(0) = 0 Then
         If ClsPDGetMaxNumber(pa(1), strExc(1)) Then
            If text1(2) > pa(1) & String(6 - Len(strExc(1)), "0") & strExc(1) Then
               MsgBox "新本所案號不可大於自動編號，請重新輸入 !", vbCritical
               GetGrid = False
            Else
               If MsgBox("此本所案號不存在 ( " & Trim(text1(2)) & Trim(text1(25)) & Trim(text1(26)) & Trim(text1(27)) & " ) ，請確認 ?", vbQuestion + vbYesNo) = vbNo Then
                  GetGrid = False
               End If
            End If
         End If
      End If
   End If
   GridHead
End Function

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

Private Function ChgType(i As Integer) As Boolean
   Dim strTempName As String
   Dim m_Team As String
   Dim bolIsChina As Boolean
   
   ChgType = False
   Select Case i
      Case 24
         If text1(i) <> "" Then
            If ClsPDGetStaff(text1(i), strTempName) Then
               Label5(10) = strTempName
               ChgType = True
            Else
               Label5(10) = ""
            End If
         End If
      Case 1 '案件性質
         If pa(1) = "CFP" Or pa(9) = "000" Then
            bolIsChina = False
         Else
            bolIsChina = True
         End If
         If ClsPDGetCaseProperty(pa(1), text1(i), strTempName, bolIsChina) Then
            Label3(1) = strTempName
            ChgType = True
         End If
         
      Case 13, 14, 15, 16, 17
         strExc(1) = text1(i).Text
         If ClsPDGetCustomer(strExc(1), strTempName) Then
            text1(i).Text = strExc(1)
            Label5(i - 11) = strTempName
            ChgType = True
         Else
            Label5(i - 11) = ""
         End If
         
      Case 18
         strExc(1) = text1(i).Text
         If ClsPDGetAgent(strExc(1), strTempName) Then
            text1(i) = strExc(1)
            Label5(7) = Replace(strTempName, "&", "&&")
            ChgType = True
         Else
            Label5(7) = ""
         End If
      Case Else
         ChgType = True
   End Select
End Function

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

'Added by Lydia 2018/05/21
Private Sub Form_Unload(Cancel As Integer)

    PUB_SendMailCache , , , True 'Added by Lydia 2024/06/06
    
   '回到前畫面
   If UCase(TypeName(m_PrevForm)) <> "NOTHING" And TypeName(m_PrevForm) <> "frm060101" Then
        m_PrevForm.Show
        If TypeName(m_PrevForm) = "frm060122" Then
            m_PrevForm.cmdState = 0
            Call m_PrevForm.PubShowNextData
        End If
   End If
   
   Set frm060101_3 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
Dim i As Integer
   
   If text1(1) <> "404" Then Exit Sub 'Add By Sindy 2021/8/16
   
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
         text1(4) = ""
         text1(4).Tag = "" 'Add By Sindy 2015/12/16
         .col = 3
         text1(5) = ""
         .col = 7
         text1(6) = ""
         .col = 10
         textCP64 = ""
'         m_CP30 = "" 'Add by Morgan 2011/4/22
      Else
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
         text1(4) = FCDate(.Text)
         text1(4).Tag = text1(4).Text 'Add By Sindy 2015/12/16
         .col = 3
         text1(5) = FCDate(.Text)
         .col = 7
         text1(6) = .Text
         ' 90.07.06 modify by louis (備註帶到進度備註欄位)
         .col = 10
         textCP64 = .Text
'         m_CP30 = .TextMatrix(.row, 9) 'Add by Morgan 2011/4/22
      End If
      'If .Rows > 0 Then
      '   .Row = 1
      '   .Col = 7
      '   Text1(6) = .Text
      'End If
   End With
End Sub

'Added by Lydia 2024/03/12
Private Sub Text1_Change(Index As Integer)
   Select Case Index
      Case 1
         Call SetTF
   End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse text1(Index)
   CloseIme
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim ii As Integer

   If Trim(cboCP14.Text) = "" Then
      MsgBox "承辦人不可空白！"
      cboCP14.SetFocus
      Call CboCP14_GotFocus
      Exit Function
   ElseIf Left(PUB_GetST03(Left(Trim(cboCP14.Text), 5)), 1) <> "F" Then
      MsgBox "承辦人必須為國外部人員！"
      cboCP14.SetFocus
      Call CboCP14_GotFocus
      Exit Function
   End If
   
   'Modify By Sindy 2016/7/26
   'Modified by Lydia 2022/07/05 不改日期
   'Cancel = False
   Cancel = True
   CboCP14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   '2016/7/26 END
   
   'Added by Lydia 2024/06/14 對申請人1~5的重複輸入檢查
   If Pub_ChkAppList(strExc(0), text1(13) & "," & text1(14) & "," & text1(15) & "," & text1(16) & "," & text1(17)) = False Then
      SSTab1.Tab = 0
      text1(Val(strExc(0)) + 12).SetFocus
      Text1_GotFocus Val(strExc(0)) + 12
      Exit Function
   End If
   'end 2024/06/14
   
   'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
   For ii = 13 To 18
      strExc(1) = ChangeCustomerL(text1(ii))
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
               text1(ii).SetFocus
               Text1_GotFocus ii
               Exit Function
            End If
         End If
      Else
         'Added by Lydia 2024/06/19 區分PS案
         If pa(1) = "FCP" Or pa(1) = "CFP" Or pa(1) = "P" Then
            strExc(2) = ChangeCustomerL(pa(75))
         Else
            strExc(2) = ChangeCustomerL(pa(26))
         End If
         If strExc(1) <> "" And strExc(1) <> strExc(2) Then
            If GetAgentAndState(strExc(1), strExc(3), , , , pa(1), strExc(8), False) = False Then
               SSTab1.Tab = 0
               text1(ii).SetFocus
               Text1_GotFocus ii
               Exit Function
            End If
         End If
      End If
   Next ii
   'end 2024/06/13
   
   'Add by Sindy 2021/5/13 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/5/13 END
   
   TxtValidate = True
End Function

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then
      If ChgType(Index) = False Then Cancel = True
   'Add by Sindy 2022/6/20
   ElseIf Index = 1 Then '案件性質
      If text1(Index) <> "" Then
         If ChgType(Index) = False Then
            Cancel = True
         End If
      Else
         MsgBox "案件性質不可空白 !", vbCritical
         Cancel = True
      End If
      If Cancel = False And text1(Index).Text <> text1(Index).Tag Then
'         'Add by Morgan 2008/9/4
'         'Modify by Morgan 2008/9/10 +210 製作中說
'         'Modified by Morgan 2013/11/6 +235核對中說格式
'         'Modified by Lydia 2021/01/06 排除Murgitroyd案的檢視中說
'         'Addd by Lydia 2021/01/29 +系統別判斷 ex.FG-001253在分案時pa(75)非代理人
'         If pa(1) <> "FCP" Then
'             text1(28).Locked = True
'         Else
'         'end 2021/01/29
'             If strMurgitroyd <> "" And pa(75) <> "" And InStr(strMurgitroyd, ChangeCustomerL(pa(75))) > 0 And text1(Index).Text = "209" Then
'                text1(28).Locked = True
'             ElseIf text1(Index).Text = "209" Or text1(Index).Text = "235" Or text1(Index).Text = "210" Then
'             'end 2021/01/06
'                text1(28).Locked = False
'             Else
'                text1(28).Locked = True
'             End If
'         End If 'end 2021/01/29
'         text1(28).Text = m_EP06

         Label3(0) = Trim(Mid(Trim(cboCP14.Text), 6))
         Call PUB_GetFCPsetCP48(Me.Visible, pa, m_CP27, text1(1), text1(0), m_CP122, text1(4), text1(5), text1(23), text1(28), Combo2, text1(12))
      End If
      '2022/6/20 END
   End If
End Sub

Private Sub textCP64_GotFocus()
   TextInverse textCP64
End Sub
Private Sub textPA91_GotFocus()
   TextInverse textPA91
End Sub

'Added by Lydia 2024/03/12 參考FCP分案
'只要是翻譯都要顯示，不必控制是否有輸承辦人，因為有可能會先紀錄相似折扣
Private Sub SetTF()
   If (text1(1) = "201" Or text1(1) = "927" Or text1(1) = "209") Then
      lblTF05.Visible = True
      txtTF05.Visible = True
      lblTF18.Visible = True
      txtTF18.Visible = True
      strExc(0) = "select * from TransFee where TF01='" & Label3(8) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         txtTF05.Text = "" & RsTemp.Fields("TF05")
         txtTF18.Text = "" & RsTemp.Fields("TF18")
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
   End If
End Sub
